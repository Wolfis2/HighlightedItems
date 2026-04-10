using System.Windows.Forms;
using HighlightedItems.Utils;
using ExileCore;
using ExileCore.PoEMemory.Elements.InventoryElements;
using ExileCore.PoEMemory.MemoryObjects;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExileCore.Shared.Enums;
using ImGuiNET;
using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using ExileCore.PoEMemory.Components;
using ExileCore.Shared;
using ExileCore.Shared.Helpers;
using ItemFilterLibrary;
using SharpDX;
using Vector2 = System.Numerics.Vector2;

namespace HighlightedItems;

public class HighlightedItems : BaseSettingsPlugin<Settings>
{
    [DllImport("user32.dll", SetLastError = true)]
    private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, IntPtr dwExtraInfo);

    private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
    private const uint MOUSEEVENTF_RIGHTUP = 0x0010;

    private SyncTask<bool> _currentOperation;
    private string _customStashFilter = "";
    private string _customInventoryFilter = "";

    private record QueryOrException(ItemQuery Query, Exception Exception);

    private readonly ConditionalWeakTable<string, QueryOrException> _queries = [];

    private bool MoveCancellationRequested => Settings.CancelWithRightMouseButton && (Control.MouseButtons & MouseButtons.Right) != 0;
    private IngameState InGameState => GameController.IngameState;
    private SharpDX.Vector2 WindowOffset => GameController.Window.GetWindowRectangleTimeCache.TopLeft;

    public override bool Initialise()
    {
        Graphics.InitImage(Path.Combine(DirectoryFullName, "images\\pick.png").Replace('\\', '/'), false);
        Graphics.InitImage(Path.Combine(DirectoryFullName, "images\\pickL.png").Replace('\\', '/'), false);

        return true;
    }

    public override void AreaChange(AreaInstance area)
    {
        _mouseStateForRect.Clear();
    }

    public override void DrawSettings()
    {
        base.DrawSettings();
        DrawIgnoredCellsSettings();
    }

    private Predicate<Entity> GetPredicate(string windowTitle, ref string filterText, Vector2 defaultPosition)
    {
        if (!Settings.ShowCustomFilterWindow) return null;
        Settings.SavedFilters ??= [];
        ImGui.SetNextWindowPos(defaultPosition, ImGuiCond.FirstUseEver);
        if (ImGui.Begin(windowTitle, ImGuiWindowFlags.NoDecoration | ImGuiWindowFlags.AlwaysAutoResize))
        {
            ImGui.InputTextWithHint("##input", "Filter using IFL syntax", ref filterText, 2000);
            Predicate<Entity> returnValue = null;
            if (!string.IsNullOrWhiteSpace(filterText))
            {
                ImGui.SameLine();
                if (ImGui.Button("Clear"))
                {
                    filterText = "";
                    return null;
                }

                if (!Settings.SavedFilters.Contains(filterText))
                {
                    ImGui.SameLine();
                    if (ImGui.Button("Save"))
                    {
                        Settings.SavedFilters.Add(filterText);
                    }
                }

                var (query, exception) = _queries.GetValue(filterText, s =>
                {
                    try
                    {
                        var itemQuery = ItemQuery.Load(s);
                        if (itemQuery.FailedToCompile)
                        {
                            return new QueryOrException(null, new Exception(itemQuery.Error));
                        }

                        return new QueryOrException(itemQuery, null);
                    }
                    catch (Exception ex)
                    {
                        return new QueryOrException(null, ex);
                    }
                })!;

                if (exception != null)
                {
                    ImGui.TextUnformatted($"{exception.Message}");
                }
                else
                {
                    returnValue = s =>
                    {
                        try
                        {
                            return query.CompiledQuery(new ItemData(s, GameController));
                        }
                        catch (Exception ex)
                        {
                            DebugWindow.LogError($"Failed to match item: {ex}");
                            return false;
                        }
                    };
                }
            }

            // ReSharper disable once AssignmentInConditionalExpression
            if (Settings.SavedFilters.Any() && Settings.UsePopupForFilterSelector
                    ? Settings.OpenSavedFilterList = ImGui.BeginPopupContextItem("saved_filter_popup")
                    : Settings.OpenSavedFilterList = ImGui.TreeNodeEx("Saved filters",
                        Settings.OpenSavedFilterList
                            ? ImGuiTreeNodeFlags.DefaultOpen | ImGuiTreeNodeFlags.NoTreePushOnOpen
                            : ImGuiTreeNodeFlags.NoTreePushOnOpen))
            {
                foreach (var (savedFilter, index) in Settings.SavedFilters.Select((x, i) => (x, i)).ToList())
                {
                    ImGui.PushID($"saved{index}");
                    if (ImGui.Button("Load"))
                    {
                        filterText = savedFilter;
                    }

                    ImGui.SameLine();
                    if (ImGui.Button("Delete"))
                    {
                        if (ImGui.IsKeyDown(ImGuiKey.ModShift))
                        {
                            Settings.SavedFilters.Remove(savedFilter);
                        }
                    }
                    else if (ImGui.IsItemHovered())
                    {
                        ImGui.SetTooltip("Hold Shift");
                    }

                    ImGui.SameLine();
                    ImGui.TextUnformatted(savedFilter);

                    ImGui.PopID();
                }

                if (Settings.UsePopupForFilterSelector)
                {
                    ImGui.EndPopup();
                }
                else
                {
                    ImGui.TreePop();
                }
            }

            if (Settings.UsePopupForFilterSelector)
            {
                if (ImGui.Button("Open Saved Filters"))
                {
                    ImGui.OpenPopup("saved_filter_popup");
                }
            }

            ImGui.End();
            return returnValue;
        }

        return null;
    }

    public override void Render()
    {
        if (_currentOperation != null)
        {
            DebugWindow.LogMsg("Running the inventory dump procedure...");
            TaskUtils.RunOrRestart(ref _currentOperation, () => null);
            if (_itemsToMove is { Count: > 0 } itemsToMove)
            {
                foreach (var (rect, color) in itemsToMove.Skip(1).Select(x => (x, Settings.CustomFilterFrameColor)).Prepend((itemsToMove[0], Color.Green)))
                {
                    Graphics.DrawFrame(rect.TopLeft.ToVector2Num(), rect.BottomRight.ToVector2Num(), color, Settings.CustomFilterFrameThickness);
                }
            }
            return;
        }

        if (!Settings.Enable)
            return;

        var (inventory, rectElement) = (InGameState.IngameUi.StashElement, InGameState.IngameUi.GuildStashElement) switch
        {
            ({ IsVisible: true, VisibleStash: { InventoryUIElement: { } invRect } visibleStash }, _) => (visibleStash, invRect),
            (_, { IsVisible: true, VisibleStash: { InventoryUIElement: { } invRect } visibleStash }) => (visibleStash, invRect),
            _ => (null, null)
        };

        const float buttonSize = 37;
        var highlightedItemsFound = false;
        if (inventory != null)
        {
            var stashRect = rectElement.GetClientRectCache;
            var (itemFilter, isCustomFilter) = GetPredicate("Custom stash filter", ref _customStashFilter, stashRect.BottomLeft.ToVector2Num()) is { } customPredicate
                ? ((Predicate<NormalInventoryItem>)(s => customPredicate(s.Item)), true)
                : (s => s.isHighlighted != Settings.InvertSelection.Value, false);

            //Determine Stash Pickup Button position and draw
            var buttonPos = Settings.UseCustomMoveToInventoryButtonPosition
                ? Settings.CustomMoveToInventoryButtonPosition
                : stashRect.BottomRight.ToVector2Num() + new Vector2(-43, 10);
            var buttonRect = new SharpDX.RectangleF(buttonPos.X, buttonPos.Y, buttonSize, buttonSize);

            Graphics.DrawImage("pick.png", buttonRect);

            var highlightedItems = GetHighlightedItems(inventory, itemFilter);
            highlightedItemsFound = highlightedItems.Any();
            int? stackSizes = 0;
            foreach (var item in highlightedItems)
            {
                stackSizes += item.Item?.GetComponent<Stack>()?.Size;
                if (isCustomFilter)
                {
                    var rect = item.GetClientRectCache;
                    var deflateFactor = Settings.CustomFilterBorderDeflation / 200.0;
                    var deflateWidth = (int)(rect.Width * deflateFactor + Settings.CustomFilterFrameThickness / 2);
                    var deflateHeight = (int)(rect.Height * deflateFactor + Settings.CustomFilterFrameThickness / 2);
                    rect.Inflate(-deflateWidth, -deflateHeight);

                    var topLeft = rect.TopLeft.ToVector2Num();
                    var bottomRight = rect.BottomRight.ToVector2Num();
                    Graphics.DrawFrame(topLeft, bottomRight, Settings.CustomFilterFrameColor, Settings.CustomFilterBorderRounding, Settings.CustomFilterFrameThickness, 0);
                }
            }

            var countText = Settings.ShowStackSizes && highlightedItems.Count != stackSizes && stackSizes != null
                ? Settings.ShowStackCountWithSize
                    ? $"{stackSizes} / {highlightedItems.Count}"
                    : $"{stackSizes}"
                : $"{highlightedItems.Count}";

            var countPos = new Vector2(buttonRect.Left - 2, buttonRect.Center.Y - 11);
            Graphics.DrawText($"{countText}", countPos with { Y = countPos.Y + 2 }, SharpDX.Color.Black, FontAlign.Right);
            Graphics.DrawText($"{countText}", countPos with { X = countPos.X - 2 }, SharpDX.Color.White, FontAlign.Right);

            if (IsButtonPressed(buttonRect) ||
                Input.IsKeyDown(Settings.MoveToInventoryHotkey.Value))
            {
                var orderedItems = highlightedItems
                    .OrderBy(stashItem => stashItem.GetClientRectCache.X)
                    .ThenBy(stashItem => stashItem.GetClientRectCache.Y)
                    .ToList();
                _currentOperation = MoveItemsToInventory(orderedItems);
            }
        }
        else
        {
            if (Settings.ResetCustomFilterOnPanelClose)
            {
                _customStashFilter = "";
            }
        }

        var inventoryPanel = InGameState.IngameUi.InventoryPanel;
        if (inventoryPanel.IsVisible)
        {
            var inventoryRect = inventoryPanel[2].GetClientRectCache;

            var (itemFilter, isCustomFilter) = GetPredicate("Custom inventory filter", ref _customInventoryFilter, inventoryRect.BottomLeft.ToVector2Num()) is { } customPredicate
                ? (customPredicate, true)
                : (_ => true, false);

            if (Settings.DumpButtonEnable && IsStashTargetOpened)
            {
                //Determine Inventory Pickup Button position and draw
                var buttonPos = Settings.UseCustomMoveToStashButtonPosition
                    ? Settings.CustomMoveToStashButtonPosition
                    : inventoryRect.TopLeft.ToVector2Num() + new Vector2(buttonSize / 2, -buttonSize);
                var buttonRect = new SharpDX.RectangleF(buttonPos.X, buttonPos.Y, buttonSize, buttonSize);

                if (isCustomFilter)
                {
                    foreach (var item in GameController.IngameState.ServerData.PlayerInventories[0].Inventory.InventorySlotItems.Where(x => itemFilter(x.Item)))
                    {
                        var rect = item.GetClientRect();
                        var deflateFactor = Settings.CustomFilterBorderDeflation / 200.0;
                        var deflateWidth = (int)(rect.Width * deflateFactor + Settings.CustomFilterFrameThickness / 2);
                        var deflateHeight = (int)(rect.Height * deflateFactor + Settings.CustomFilterFrameThickness / 2);
                        rect.Inflate(-deflateWidth, -deflateHeight);

                        var topLeft = rect.TopLeft.ToVector2Num();
                        var bottomRight = rect.BottomRight.ToVector2Num();
                        Graphics.DrawFrame(topLeft, bottomRight, Settings.CustomFilterFrameColor, Settings.CustomFilterBorderRounding, Settings.CustomFilterFrameThickness, 0);
                    }
                }

                Graphics.DrawImage("pickL.png", buttonRect);
                if (IsButtonPressed(buttonRect) ||
                    Input.IsKeyDown(Settings.MoveToStashHotkey.Value) ||
                    Settings.UseMoveToInventoryAsMoveToStashWhenNoHighlights &&
                    !highlightedItemsFound &&
                    Input.IsKeyDown(Settings.MoveToInventoryHotkey.Value))
                {
                    var inventoryItems = GameController.IngameState.ServerData.PlayerInventories[0].Inventory.InventorySlotItems
                        .Where(x => !IsInIgnoreCell(x))
                        .Where(x => itemFilter(x.Item))
                        .OrderBy(x => x.PosX)
                        .ThenBy(x => x.PosY)
                        .ToList();

                    _currentOperation = MoveItemsToStash(inventoryItems);
                }
            }
        }
        else
        {
            if (Settings.ResetCustomFilterOnPanelClose)
            {
                _customInventoryFilter = "";
            }
        }
    }

    private async SyncTask<bool> MoveItemsCommonPreamble()
    {
        while (Control.MouseButtons == MouseButtons.Left || MoveCancellationRequested)
        {
            if (MoveCancellationRequested)
            {
                return false;
            }

            await TaskUtils.NextFrame();
        }

        if (Settings.IdleMouseDelay.Value == 0)
        {
            return true;
        }

        var mousePos = Mouse.GetCursorPosition();
        var sw = Stopwatch.StartNew();
        await TaskUtils.NextFrame();
        while (true)
        {
            if (MoveCancellationRequested)
            {
                return false;
            }

            var newPos = Mouse.GetCursorPosition();
            if (mousePos != newPos)
            {
                mousePos = newPos;
                sw.Restart();
            }
            else if (sw.ElapsedMilliseconds >= Settings.IdleMouseDelay.Value)
            {
                return true;
            }
            else
            {
                await TaskUtils.NextFrame();
            }
        }
    }

    private async SyncTask<bool> MoveItemsToStash(List<ServerInventory.InventSlotItem> items)
    {
        if (!await MoveItemsCommonPreamble())
        {
            return false;
        }

        _prevMousePos = Mouse.GetCursorPosition();
        var processedIndices = new HashSet<int>();
        for (var i = 0; i < items.Count; i++)
        {
            if (processedIndices.Contains(i))
                continue;

            var item = items[i];
            if (MoveCancellationRequested) 
            {
                await StopMovingItems();
                return false;
            }

            if (!InGameState.IngameUi.InventoryPanel.IsVisible)
            {
                DebugWindow.LogMsg("HighlightedItems: Inventory Panel closed, aborting loop");
                break;
            }

            if (!IsStashTargetOpened)
            {
                DebugWindow.LogMsg("HighlightedItems: Target inventory closed, aborting loop");
                break;
            }

            var isStackable = item.Item?.GetComponent<Stack>() != null;
            if (isStackable)
            {
                var itemPath = item.Item?.Path ?? "";
                for (var j = i + 1; j < items.Count; j++)
                {
                    if (items[j].Item?.Path == itemPath)
                        processedIndices.Add(j);
                }
            }
            processedIndices.Add(i);
            _itemsToMove = items.Where((_, idx) => !processedIndices.Contains(idx)).Select(x => x.GetClientRect()).ToList();

            Keyboard.KeyDown(Keys.LControlKey);
            await Wait(KeyDelay, true);
            await MoveItem(item.GetClientRect().Center, isStackable);
            Keyboard.KeyUp(Keys.LControlKey);
            await Wait(KeyDelay, true);
        }

        await StopMovingItems();
        return true;
    }

    private bool IsStashTargetOpened =>
        !Settings.VerifyTargetInventoryIsOpened
        || InGameState.IngameUi.StashElement.IsVisible
        || InGameState.IngameUi.SellWindow.IsVisible
        || InGameState.IngameUi.TradeWindow.IsVisible
        || InGameState.IngameUi.GuildStashElement.IsVisible;

    private bool IsStashSourceOpened =>
        !Settings.VerifyTargetInventoryIsOpened
        || InGameState.IngameUi.StashElement.IsVisible
        || InGameState.IngameUi.GuildStashElement.IsVisible;

    private List<RectangleF> _itemsToMove = null;
    private Point _prevMousePos = Point.Zero;

    private async SyncTask<bool> MoveItemsToInventory(List<NormalInventoryItem> items)
    {
        if (!await MoveItemsCommonPreamble())
        {
            return false;
        }

        _prevMousePos = Mouse.GetCursorPosition();
        var processedIndices = new HashSet<int>();
        for (var i = 0; i < items.Count; i++)
        {
            if (processedIndices.Contains(i))
                continue;

            var item = items[i];
            if (MoveCancellationRequested)
            {
                await StopMovingItems();
                return false;
            }

            if (!IsStashSourceOpened)
            {
                DebugWindow.LogMsg("HighlightedItems: Stash Panel closed, aborting loop");
                break;
            }

            if (!InGameState.IngameUi.InventoryPanel.IsVisible)
            {
                DebugWindow.LogMsg("HighlightedItems: Inventory Panel closed, aborting loop");
                break;
            }

            if (IsInventoryFull())
            {
                DebugWindow.LogMsg("HighlightedItems: Inventory full, aborting loop");
                break;
            }

            var isStackable = item.Item?.GetComponent<Stack>() != null;
            if (isStackable)
            {
                var itemPath = item.Item?.Path ?? "";
                for (var j = i + 1; j < items.Count; j++)
                {
                    if (items[j].Item?.Path == itemPath)
                        processedIndices.Add(j);
                }
            }
            processedIndices.Add(i);
            _itemsToMove = items.Where((_, idx) => !processedIndices.Contains(idx)).Select(x => x.GetClientRectCache).ToList();

            Keyboard.KeyDown(Keys.LControlKey);
            await Wait(KeyDelay, true);
            await MoveItem(item.GetClientRect().Center, isStackable);
            Keyboard.KeyUp(Keys.LControlKey);
            await Wait(KeyDelay, true);
        }

        await StopMovingItems();
        return true;
    }

    private async SyncTask<bool> StopMovingItems() {
        Keyboard.KeyUp(Keys.LControlKey);
        await Wait(KeyDelay, false);
        Mouse.moveMouse(_prevMousePos);
        _prevMousePos = Point.Zero;
        _itemsToMove = null;
        DebugWindow.LogMsg("HighlightedItems: Stopped moving items");
        return true;
    }

    private List<NormalInventoryItem> GetHighlightedItems(Inventory stash, Predicate<NormalInventoryItem> filter)
    {
        try
        {
            var stashItems = stash.VisibleInventoryItems;

            var highlightedItems = stashItems
                .Where(stashItem => filter(stashItem))
                .ToList();

            return highlightedItems;
        }
        catch
        {
            return [];
        }
    }

    private bool IsInventoryFull()
    {
        var inventoryItems = GameController.IngameState.ServerData.PlayerInventories[0].Inventory.InventorySlotItems;

        // quick sanity check
        if (inventoryItems.Count < 12)
        {
            return false;
        }

        // track each inventory slot
        bool[,] inventorySlot = new bool[12, 5];

        // iterate through each item in the inventory and mark used slots
        foreach (var inventoryItem in inventoryItems)
        {
            int x = inventoryItem.PosX;
            int y = inventoryItem.PosY;
            int height = inventoryItem.SizeY;
            int width = inventoryItem.SizeX;
            for (int row = x; row < x + width; row++)
            {
                for (int col = y; col < y + height; col++)
                {
                    inventorySlot[row, col] = true;
                }
            }
        }

        // check for any empty slots
        for (int x = 0; x < 12; x++)
        {
            for (int y = 0; y < 5; y++)
            {
                if (inventorySlot[x, y] == false)
                {
                    return false;
                }
            }
        }

        // no empty slots, so inventory is full
        return true;
    }

    private static readonly TimeSpan KeyDelay = TimeSpan.FromMilliseconds(10);
    private static readonly TimeSpan MouseMoveDelay = TimeSpan.FromMilliseconds(20);
    private TimeSpan MouseDownDelay => TimeSpan.FromMilliseconds(25 + Settings.ExtraDelay.Value);
    private static readonly TimeSpan MouseUpDelay = TimeSpan.FromMilliseconds(5);

    private async SyncTask<bool> MoveItem(SharpDX.Vector2 itemPosition, bool isStackable = false)
    {
        itemPosition += WindowOffset;
        Mouse.moveMouse(itemPosition);
        await Wait(MouseMoveDelay, true);
        if (isStackable)
        {
            uint x = (uint)itemPosition.X;
            uint y = (uint)itemPosition.Y;
            mouse_event(MOUSEEVENTF_RIGHTDOWN, x, y, 0, IntPtr.Zero);
            await Wait(MouseDownDelay, true);
            mouse_event(MOUSEEVENTF_RIGHTUP, x, y, 0, IntPtr.Zero);
        }
        else
        {
            Mouse.LeftDown();
            await Wait(MouseDownDelay, true);
            Mouse.LeftUp();
        }
        await Wait(MouseUpDelay, true);
        return true;
    }

    private async SyncTask<bool> Wait(TimeSpan period, bool canUseThreadSleep)
    {
        if (canUseThreadSleep && Settings.UseThreadSleep)
        {
            Thread.Sleep(period);
            return true;
        }

        var sw = Stopwatch.StartNew();
        while (sw.Elapsed < period)
        {
            await TaskUtils.NextFrame();
        }

        return true;
    }

    private readonly ConcurrentDictionary<RectangleF, bool?> _mouseStateForRect = [];

    private bool IsButtonPressed(RectangleF buttonRect)
    {
        var prevState = _mouseStateForRect.GetValueOrDefault(buttonRect);
        var isHovered = buttonRect.Contains(Mouse.GetCursorPosition() - WindowOffset);
        if (!isHovered)
        {
            _mouseStateForRect[buttonRect] = null;
            return false;
        }

        var isPressed = Control.MouseButtons == MouseButtons.Left && CanClickButtons;
        _mouseStateForRect[buttonRect] = isPressed;
        return isPressed &&
               prevState == false;
    }

    private bool CanClickButtons => !Settings.VerifyButtonIsNotObstructed || !ImGui.GetIO().WantCaptureMouse;

    private bool IsInIgnoreCell(ServerInventory.InventSlotItem inventItem)
    {
        var inventPosX = inventItem.PosX;
        var inventPosY = inventItem.PosY;

        if (inventPosX < 0 || inventPosX >= 12)
            return true;
        if (inventPosY < 0 || inventPosY >= 5)
            return true;

        return Settings.IgnoredCells[inventPosY, inventPosX]; //No need to check all item size
    }

    private void DrawIgnoredCellsSettings()
    {
        ImGui.BeginChild("##IgnoredCellsMain", new Vector2(ImGui.GetContentRegionAvail().X, 204f), ImGuiChildFlags.Border,
            ImGuiWindowFlags.NoScrollWithMouse);
        ImGui.Text("Ignored Inventory Slots (checked = ignored)");

        var contentRegionAvail = ImGui.GetContentRegionAvail();
        ImGui.BeginChild("##IgnoredCellsCels", new Vector2(contentRegionAvail.X, contentRegionAvail.Y), ImGuiChildFlags.Border,
            ImGuiWindowFlags.NoScrollWithMouse);

        for (int y = 0; y < 5; ++y)
        {
            for (int x = 0; x < 12; ++x)
            {
                bool isCellIgnored = Settings.IgnoredCells[y, x];
                if (ImGui.Checkbox($"##{y}_{x}IgnoredCells", ref isCellIgnored))
                    Settings.IgnoredCells[y, x] = isCellIgnored;
                if (x < 11)
                    ImGui.SameLine();
            }
        }

        ImGui.EndChild();
        ImGui.EndChild();
    }
}
