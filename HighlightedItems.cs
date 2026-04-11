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

    private readonly object _logLock = new();
    private string _logFilePath;

    private void Log(string message)
    {
        if (!Settings.DebugLog) return;
        // Lazily resolve the path in case Log is called before Initialise
        var path = _logFilePath ?? Path.Combine(DirectoryFullName, "HighlightedItems_debug.log");
        var line = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
        DebugWindow.LogMsg($"[HI] {message}");
        try
        {
            lock (_logLock)
            {
                File.AppendAllText(path, line + Environment.NewLine);
            }
        }
        catch
        {
            // Swallow file write errors so they never crash the plugin
        }
    }

    private record QueryOrException(ItemQuery Query, Exception Exception);

    private readonly ConditionalWeakTable<string, QueryOrException> _queries = [];

    private bool MoveCancellationRequested => Settings.CancelWithRightMouseButton && (Control.MouseButtons & MouseButtons.Right) != 0;
    private IngameState InGameState => GameController.IngameState;
    private SharpDX.Vector2 WindowOffset => GameController.Window.GetWindowRectangleTimeCache.TopLeft;

    public override bool Initialise()
    {
        _logFilePath = Path.Combine(DirectoryFullName, "HighlightedItems_debug.log");
        Log($"Initialise — DirectoryFullName={DirectoryFullName}");

        var pickPath = Path.Combine(DirectoryFullName, "images\\pick.png").Replace('\\', '/');
        var pickLPath = Path.Combine(DirectoryFullName, "images\\pickL.png").Replace('\\', '/');
        Log($"Loading images: {pickPath}  |  {pickLPath}");
        Graphics.InitImage(pickPath, false);
        Graphics.InitImage(pickLPath, false);
        Log("Initialise complete");

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
                // When no items are highlighted and no custom filter is active, move ALL visible stash items
                var itemsToMove = (isCustomFilter || highlightedItems.Any())
                    ? highlightedItems
                    : inventory.VisibleInventoryItems.ToList();
                var orderedItems = itemsToMove
                    .OrderBy(stashItem => stashItem.GetClientRectCache.X)
                    .ThenBy(stashItem => stashItem.GetClientRectCache.Y)
                    .ToList();
                Log($"Trigger MoveToInventory — customFilter={isCustomFilter} highlighted={highlightedItems.Count} total={orderedItems.Count} stashRect={stashRect}");
                _currentOperation = MoveItemsToInventory(orderedItems);
            }

            if (Settings.SortStashButtonEnable)
            {
                // Sort button: to the right of the stash dump icon
                var sortStashPos = Settings.UseCustomMoveToInventoryButtonPosition
                    ? Settings.CustomMoveToInventoryButtonPosition + new Vector2(buttonSize + 4, 0)
                    : stashRect.BottomRight.ToVector2Num() + new Vector2(-43 + buttonSize + 4, 10);
                var sortStashRect = new SharpDX.RectangleF(sortStashPos.X, sortStashPos.Y, buttonSize, buttonSize);

                Graphics.DrawFrame(sortStashRect.TopLeft.ToVector2Num(), sortStashRect.BottomRight.ToVector2Num(), SharpDX.Color.White, 2);
                var slp = new Vector2(sortStashRect.Center.X, sortStashRect.Center.Y - 10);
                Graphics.DrawText("Sort", slp with { X = slp.X + 1 }, SharpDX.Color.Black, FontAlign.Center);
                Graphics.DrawText("Sort", slp, SharpDX.Color.White, FontAlign.Center);

                if (IsButtonPressed(sortStashRect) || Input.IsKeyDown(Settings.SortStashHotkey.Value))
                {
                    Log($"Trigger SortStash — stashRect={stashRect}");
                    _currentOperation = SortStash(stashRect);
                }
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

                    Log($"Trigger MoveToStash — customFilter={isCustomFilter} items={inventoryItems.Count} inventoryRect={inventoryRect}");
                    _currentOperation = MoveItemsToStash(inventoryItems);
                }
            }

            if (Settings.SortButtonEnable)
            {
                // Sort button: to the right of the move-to-stash button
                var sortButtonPos = inventoryRect.TopLeft.ToVector2Num() + new Vector2(buttonSize / 2 + buttonSize + 4, -buttonSize);
                var sortButtonRect = new SharpDX.RectangleF(sortButtonPos.X, sortButtonPos.Y, buttonSize, buttonSize);

                Graphics.DrawFrame(sortButtonRect.TopLeft.ToVector2Num(), sortButtonRect.BottomRight.ToVector2Num(), SharpDX.Color.White, 2);
                var sortLabelPos = new Vector2(sortButtonRect.Center.X, sortButtonRect.Center.Y - 10);
                Graphics.DrawText("Sort", sortLabelPos with { X = sortLabelPos.X + 1 }, SharpDX.Color.Black, FontAlign.Center);
                Graphics.DrawText("Sort", sortLabelPos, SharpDX.Color.White, FontAlign.Center);

                if (IsButtonPressed(sortButtonRect) || Input.IsKeyDown(Settings.SortInventoryHotkey.Value))
                {
                    Log($"Trigger SortInventory — inventoryRect={inventoryRect}");
                    _currentOperation = SortInventory(inventoryRect);
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
            Log("MoveItemsToStash: preamble returned false (cancelled before start)");
            return false;
        }

        Log($"MoveItemsToStash: starting — {items.Count} items");
        _prevMousePos = Mouse.GetCursorPosition();
        var processedIndices = new HashSet<int>();
        for (var i = 0; i < items.Count; i++)
        {
            if (processedIndices.Contains(i))
                continue;

            var item = items[i];
            if (MoveCancellationRequested)
            {
                Log($"MoveItemsToStash: cancelled by right-click at item {i}");
                await StopMovingItems();
                return false;
            }

            if (!InGameState.IngameUi.InventoryPanel.IsVisible)
            {
                Log("MoveItemsToStash: Inventory Panel closed, aborting loop");
                DebugWindow.LogMsg("HighlightedItems: Inventory Panel closed, aborting loop");
                break;
            }

            if (!IsStashTargetOpened)
            {
                Log("MoveItemsToStash: Target inventory closed, aborting loop");
                DebugWindow.LogMsg("HighlightedItems: Target inventory closed, aborting loop");
                break;
            }

            var isStackable = item.Item?.GetComponent<Stack>() != null;
            var itemRect = item.GetClientRect();
            Log($"MoveItemsToStash: [{i}/{items.Count}] path={item.Item?.Path} pos=({item.PosX},{item.PosY}) size=({item.SizeX}x{item.SizeY}) " +
                $"stackable={isStackable} rect=({itemRect.X:F0},{itemRect.Y:F0},{itemRect.Width:F0}x{itemRect.Height:F0})");

            if (isStackable)
            {
                var itemPath = item.Item?.Path ?? "";
                for (var j = i + 1; j < items.Count; j++)
                {
                    if (items[j].Item?.Path == itemPath)
                        processedIndices.Add(j);
                }
            }
            _itemsToMove = items.Where((_, idx) => !processedIndices.Contains(idx)).Select(x => x.GetClientRect()).ToList();

            Keyboard.KeyDown(Keys.LControlKey);
            await Wait(KeyDelay, true);
            await MoveItem(itemRect.Center, isStackable);
            Keyboard.KeyUp(Keys.LControlKey);
            await Wait(KeyDelay, true);
            processedIndices.Add(i);
        }

        Log("MoveItemsToStash: done");
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
            Log("MoveItemsToInventory: preamble returned false (cancelled before start)");
            return false;
        }

        Log($"MoveItemsToInventory: starting — {items.Count} items");
        _prevMousePos = Mouse.GetCursorPosition();
        for (var i = 0; i < items.Count; i++)
        {
            var item = items[i];
            _itemsToMove = items[i..].Select(x => x.GetClientRectCache).ToList();
            if (MoveCancellationRequested)
            {
                Log($"MoveItemsToInventory: cancelled by right-click at item {i}");
                await StopMovingItems();
                return false;
            }

            if (!IsStashSourceOpened)
            {
                Log("MoveItemsToInventory: Stash Panel closed, aborting loop");
                DebugWindow.LogMsg("HighlightedItems: Stash Panel closed, aborting loop");
                break;
            }

            if (!InGameState.IngameUi.InventoryPanel.IsVisible)
            {
                Log("MoveItemsToInventory: Inventory Panel closed, aborting loop");
                DebugWindow.LogMsg("HighlightedItems: Inventory Panel closed, aborting loop");
                break;
            }

            if (IsInventoryFull())
            {
                Log("MoveItemsToInventory: Inventory full, aborting loop");
                DebugWindow.LogMsg("HighlightedItems: Inventory full, aborting loop");
                break;
            }

            var itemRect = item.GetClientRect();
            Log($"MoveItemsToInventory: [{i}/{items.Count}] path={item.Item?.Path} " +
                $"rect=({itemRect.X:F0},{itemRect.Y:F0},{itemRect.Width:F0}x{itemRect.Height:F0})");

            Keyboard.KeyDown(Keys.LControlKey);
            await Wait(KeyDelay, true);
            await MoveItem(itemRect.Center);
            Keyboard.KeyUp(Keys.LControlKey);
            await Wait(KeyDelay, true);
        }

        Log("MoveItemsToInventory: done");
        await StopMovingItems();
        return true;
    }

    private async SyncTask<bool> StopMovingItems()
    {
        Keyboard.KeyUp(Keys.LControlKey);
        await Wait(KeyDelay, false);
        Mouse.moveMouse(_prevMousePos);
        _prevMousePos = Point.Zero;
        _itemsToMove = null;
        Log("StopMovingItems");
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
        try
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
            // clamp loop bounds to avoid IndexOutOfRangeException for items with invalid positions
            // (ExileCore can return items with PosX/PosY outside the 12x5 grid, e.g. cursor-held items)
            foreach (var inventoryItem in inventoryItems)
            {
                int x = inventoryItem.PosX;
                int y = inventoryItem.PosY;
                int height = inventoryItem.SizeY;
                int width = inventoryItem.SizeX;
                for (int row = Math.Max(0, x); row < Math.Min(12, x + width); row++)
                {
                    for (int col = Math.Max(0, y); col < Math.Min(5, y + height); col++)
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

            // no empty slots; log the full grid for diagnostics
            var gridLines = new System.Text.StringBuilder();
            for (int row = 0; row < 5; row++)
            {
                var line = new System.Text.StringBuilder("  ");
                for (int col = 0; col < 12; col++)
                    line.Append(inventorySlot[col, row] ? "X" : ".");
                gridLines.AppendLine(line.ToString());
            }
            Log($"IsInventoryFull: FULL — item count={inventoryItems.Count}{Environment.NewLine}{gridLines}");
            return true;
        }
        catch (Exception ex)
        {
            Log($"IsInventoryFull: exception — {ex.Message}");
            DebugWindow.LogError($"HighlightedItems: IsInventoryFull check failed: {ex.Message}");
            return false;
        }
    }

    private static readonly TimeSpan KeyDelay = TimeSpan.FromMilliseconds(10);
    private static readonly TimeSpan MouseMoveDelay = TimeSpan.FromMilliseconds(20);
    private TimeSpan MouseDownDelay => TimeSpan.FromMilliseconds(25 + Settings.ExtraDelay.Value);
    private static readonly TimeSpan MouseUpDelay = TimeSpan.FromMilliseconds(5);

    private async SyncTask<bool> MoveItem(SharpDX.Vector2 itemPosition, bool isStackable = false)
    {
        var offset = WindowOffset;
        var absolutePos = itemPosition + offset;
        Log($"MoveItem: pos=({itemPosition.X:F0},{itemPosition.Y:F0}) offset=({offset.X:F0},{offset.Y:F0}) absolute=({absolutePos.X:F0},{absolutePos.Y:F0}) stackable={isStackable}");
        Mouse.moveMouse(absolutePos);
        await Wait(MouseMoveDelay, true);
        if (isStackable)
        {
            mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, IntPtr.Zero);
            await Wait(MouseDownDelay, true);
            mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, IntPtr.Zero);
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

    private async SyncTask<bool> SortInventory(SharpDX.RectangleF inventoryRect)
    {
        if (!await MoveItemsCommonPreamble())
        {
            Log("SortInventory: preamble returned false (cancelled before start)");
            return false;
        }
        _prevMousePos = Mouse.GetCursorPosition();

        // ── PHASE 1: Scan inventory and compute the complete target layout ────────────
        // Read the inventory once, assign every item a fixed target slot using column-first
        // First-Fit Decreasing bin-packing, and store the plan keyed by item address.
        // Computing the plan upfront (rather than re-computing on every iteration) ensures
        // target assignments never change mid-sort, preventing the algorithm from chasing
        // a moving goal and placing items on top of each other.

        var initialItems = GameController.IngameState.ServerData.PlayerInventories[0].Inventory.InventorySlotItems
            .Where(x => !IsInIgnoreCell(x))
            .ToList();

        Log($"SortInventory: starting — inventoryRect={inventoryRect} items={initialItems.Count}");

        if (!initialItems.Any())
        {
            Log("SortInventory: no items to sort. Aborting.");
            await StopMovingItems();
            return true;
        }

        // Derive pixel grid geometry once from on-screen item rects.
        var (cellW, cellH, gridOriginX, gridOriginY) = CalibrateGridGeometry(initialItems, inventoryRect);
        Log($"SortInventory: grid geometry — cellW={cellW:F2} cellH={cellH:F2} originX={gridOriginX:F2} originY={gridOriginY:F2} windowOffset=({WindowOffset.X:F0},{WindowOffset.Y:F0})");

        // Sort order: area asc → width asc → category → path → address.
        // Using area as the PRIMARY key packs the smallest items (1×1 rings/jewels) into the
        // leftmost columns first, so the left side of the inventory always favours compact items
        // and large items (2×2 armours, 2×3 quivers) land on the right side.
        // Within the same area, narrower items pack more tightly (e.g. a 1×2 before a 2×1).
        // Category as tertiary key keeps items of the same type clustered within each size band.
        // Address as final tiebreaker gives a completely stable ordering.
        var packingOrder = initialItems
            .OrderBy(x => x.SizeX * x.SizeY)   // PRIMARY: area ascending (1×1 items pack to left)
            .ThenBy(x => x.SizeX)               // secondary: narrower items first within same area
            .ThenBy(GetItemCategory)            // tertiary: group same-size items by category
            .ThenBy(x => x.Item?.Path ?? "")
            .ThenBy(x => x.Item?.Address ?? 0)
            .ToList();

        // Assign target slots and record them in plan[itemAddress] = (col, row, sizeX, sizeY).
        var targetGrid = BuildIgnoredCellsGrid();
        var plan = new Dictionary<long, (int col, int row, int sizeX, int sizeY)>();
        foreach (var it in packingOrder)
        {
            if (it.Item == null) continue;
            var fit = FindFirstFit(targetGrid, it.SizeX, it.SizeY);
            if (fit == null)
            {
                Log($"SortInventory: NOFIT — {it.Item.Path} ({it.SizeX}x{it.SizeY}) at ({it.PosX},{it.PosY}) — no slot found, leaving in place");
                continue;
            }
            plan[it.Item.Address] = (fit.Value.col, fit.Value.row, it.SizeX, it.SizeY);
            Log($"SortInventory: plan — {it.Item.Path} ({it.SizeX}x{it.SizeY}) ({it.PosX},{it.PosY}) → ({fit.Value.col},{fit.Value.row})");
            for (var dy = 0; dy < it.SizeY; dy++)
                for (var dx = 0; dx < it.SizeX; dx++)
                    targetGrid[fit.Value.col + dx, fit.Value.row + dy] = true;
        }

        Log($"SortInventory: plan complete — {plan.Count}/{initialItems.Count} items assigned targets");

        // ── PHASE 2: Execute the plan move by move ────────────────────────────────────
        // Hard upper bound: 120 moves covers any realistic 12×5 inventory (60 items × 2 moves
        // in the absolute worst case where every item must be parked before its target is free).
        const int MaxMoves = 120;
        var moveCount = 0;

        // PoE creates a new InventSlotItem (new Item.Address) each time an item is placed in a
        // slot, so a moved item's address no longer matches the key we stored in plan.  Track
        // the last move's old address and destination so we can re-key the plan entry once the
        // server reflects the placement at the start of the next loop iteration.
        long pendingOldAddress = 0;
        int pendingDestCol = -1, pendingDestRow = -1;
        int pendingRetryCount = 0; // guard against stale-server infinite wait

        while (moveCount < MaxMoves)
        {
            if (MoveCancellationRequested)
            {
                Log("SortInventory: cancelled by right-click");
                await StopMovingItems();
                return false;
            }
            if (!InGameState.IngameUi.InventoryPanel.IsVisible)
            {
                Log("SortInventory: inventory panel no longer visible — stopping");
                break;
            }

            // Re-read current positions every iteration so we always have an accurate view.
            var currentItems = GameController.IngameState.ServerData.PlayerInventories[0].Inventory.InventorySlotItems
                .Where(x => !IsInIgnoreCell(x))
                .ToList();

            if (!currentItems.Any())
            {
                Log("SortInventory: currentItems is empty — stopping");
                break;
            }

            // Use the raw server inventory rather than the filtered currentItems so that a plan
            // item currently held on the cursor (which appears with out-of-bounds PosX/PosY and
            // is therefore excluded from currentItems) is also considered.
            var rawInventoryItems = GameController.IngameState.ServerData.PlayerInventories[0].Inventory.InventorySlotItems;

            // ── Plan-key reconciliation ─────────────────────────────────────────────────────
            // PoE assigns a new Item.Address when an item lands in a new slot.  Re-key the plan
            // entry for the item we placed last iteration: find whatever is sitting at the
            // destination now and, if its address differs from the one we moved, transfer the
            // plan entry to the new address so the item remains trackable.
            if (pendingOldAddress != 0)
            {
                var newItem = rawInventoryItems.FirstOrDefault(x =>
                    x.Item != null && x.PosX == pendingDestCol && x.PosY == pendingDestRow);
                if (newItem != null)
                {
                    if (newItem.Item.Address != pendingOldAddress && plan.ContainsKey(pendingOldAddress))
                    {
                        var pendingEntry = plan[pendingOldAddress];
                        // Verify dimensions match to guard against re-keying to the wrong item
                        // if PoE performed an implicit swap and the destination transiently holds
                        // a different item.
                        if (newItem.SizeX == pendingEntry.sizeX && newItem.SizeY == pendingEntry.sizeY)
                        {
                            Log($"SortInventory: re-keyed plan entry 0x{pendingOldAddress:X} → 0x{newItem.Item.Address:X} at ({pendingDestCol},{pendingDestRow})");
                            plan[newItem.Item.Address] = pendingEntry;
                            plan.Remove(pendingOldAddress);
                        }
                    }
                    pendingOldAddress = 0; // confirmed — item found at destination
                    pendingRetryCount = 0;
                }
                // If the item is not yet visible at the destination (server still catching up),
                // keep pendingOldAddress set so we retry the re-key next iteration.
            }

            // ── Stale-data guard ────────────────────────────────────────────────────────────────
            // If the server has not yet confirmed the last placement, the occupancy grid built
            // below would be stale.  Acting on stale data can cause a move to land on a slot that
            // is still logically occupied, triggering a PoE implicit swap that displaces the
            // existing item to the cursor with a fresh address unknown to the plan — making it
            // permanently invisible to all subsequent move and convergence checks.
            // Waiting here (instead of proceeding) ensures occupancy is always up-to-date.
            if (pendingOldAddress != 0)
            {
                if (++pendingRetryCount > 10)
                {
                    Log($"SortInventory: server did not confirm last placement after {pendingRetryCount} retries — stopping");
                    break;
                }
                await Wait(KeyDelay, true);
                continue;
            }
            pendingRetryCount = 0;

            // ── Secondary re-key: displaced cursor items ───────────────────────────────────────
            // If a PoE implicit swap did somehow occur (e.g. network hiccup), the displaced item
            // lands on the cursor with a new address not present in the plan, while the original
            // plan entry becomes stale (its address is no longer in the inventory).  Detect this
            // by finding plan keys absent from the live inventory, then match them by item size to
            // any cursor-held item not already tracked.  This rescues the item so the sort can
            // continue rather than silently stalling.
            {
                var liveAddresses = new HashSet<long>(
                    rawInventoryItems.Where(x => x.Item != null).Select(x => x.Item.Address));
                var staleKeys = plan.Keys.Where(k => !liveAddresses.Contains(k)).ToList();
                if (staleKeys.Count > 0)
                {
                    var cursorItems = rawInventoryItems.Where(x =>
                        x.Item != null &&
                        !plan.ContainsKey(x.Item.Address) &&
                        (x.PosX < 0 || x.PosX >= 12 || x.PosY < 0 || x.PosY >= 5)).ToList();
                    foreach (var cursorItem in cursorItems)
                    {
                        long matchKey = 0;
                        foreach (var k in staleKeys)
                        {
                            if (plan[k].sizeX == cursorItem.SizeX && plan[k].sizeY == cursorItem.SizeY)
                            {
                                matchKey = k;
                                break;
                            }
                        }
                        if (matchKey != 0)
                        {
                            Log($"SortInventory: re-keyed displaced cursor item 0x{matchKey:X} → 0x{cursorItem.Item.Address:X}");
                            plan[cursorItem.Item.Address] = plan[matchKey];
                            plan.Remove(matchKey);
                            staleKeys.Remove(matchKey);
                        }
                    }
                }
            }

            // Convergence check: every plan entry must have its item (by address) sitting at the
            // plan target position.  Using plan.All(...) rather than a filtered Where(...).All(...)
            // prevents vacuous-true exits when parked items have stale plan keys.
            // pendingOldAddress is always 0 here because the stale-data guard above would have
            // continued the loop without reaching this point if it were non-zero.
            var done = plan.All(kv => rawInventoryItems.Any(x =>
                x.Item != null
                && x.Item.Address == kv.Key
                && x.PosX == kv.Value.col
                && x.PosY == kv.Value.row));
            if (done)
            {
                Log($"SortInventory: all items at target — done in {moveCount} moves");
                break;
            }

            // Find a "safe" move: item not at its target AND the target area is currently free.
            var occupancy = BuildOccupancyGrid(currentItems);

            // ── Cursor-held item recovery ──────────────────────────────────────────────────────
            // PoE can perform an implicit swap during a placement (e.g. when the destination click
            // lands on a different item), leaving a plan item on the cursor.  Cursor-held items
            // appear in the raw server inventory with out-of-bounds PosX/PosY.  Detect this and
            // issue a destination-only click (the item is already on the cursor; no pickup click
            // is needed) to place it and keep the sort running.
            {
                var rawForCursor = GameController.IngameState.ServerData.PlayerInventories[0].Inventory.InventorySlotItems;
                var heldItem = rawForCursor.FirstOrDefault(x =>
                    x.Item != null && plan.ContainsKey(x.Item.Address) &&
                    (x.PosX < 0 || x.PosX >= 12 || x.PosY < 0 || x.PosY >= 5));
                if (heldItem != null)
                {
                    var hTarget = plan[heldItem.Item.Address];
                    int hCol, hRow;
                    // Prefer placing directly at the planned target; fall back to any free slot.
                    if (IsTargetAreaFree(occupancy, hTarget.col, hTarget.row,
                            heldItem.SizeX, heldItem.SizeY, -1, -1, 0, 0))
                    {
                        hCol = hTarget.col;
                        hRow = hTarget.row;
                    }
                    else
                    {
                        var parkSlot = FindFirstFitNotAt(occupancy, heldItem.SizeX, heldItem.SizeY, -1, -1);
                        if (parkSlot == null)
                        {
                            Log($"SortInventory: cursor-held {heldItem.Item.Path} ({heldItem.SizeX}x{heldItem.SizeY}) — no free slot, stopping");
                            break;
                        }
                        hCol = parkSlot.Value.col;
                        hRow = parkSlot.Value.row;
                    }
                    var hDst = new SharpDX.Vector2(
                        gridOriginX + (hCol + heldItem.SizeX / 2 + 0.5f) * cellW,
                        gridOriginY + (hRow + heldItem.SizeY / 2 + 0.5f) * cellH);
                    Log($"SortInventory: move #{moveCount + 1} — cursor-held {heldItem.Item.Path} " +
                        $"({heldItem.SizeX}x{heldItem.SizeY}) → ({hCol},{hRow}) | " +
                        $"dst=({hDst.X:F0},{hDst.Y:F0}) " +
                        $"dst+offset=({hDst.X + WindowOffset.X:F0},{hDst.Y + WindowOffset.Y:F0})");
                    Mouse.moveMouse(hDst + WindowOffset);
                    await Wait(MouseMoveDelay, true);
                    Mouse.LeftDown();
                    await Wait(MouseDownDelay, true);
                    Mouse.LeftUp();
                    await Wait(KeyDelay, true);
                    pendingOldAddress = heldItem.Item.Address;
                    pendingDestCol = hCol;
                    pendingDestRow = hRow;
                    pendingRetryCount = 0; // fresh retry budget for this placement
                    moveCount++;
                    continue;
                }
            }

            ServerInventory.InventSlotItem itemToMove = null;
            int destCol = 0, destRow = 0;
            foreach (var it in currentItems)
            {
                if (it.Item == null || !plan.TryGetValue(it.Item.Address, out var target)) continue;
                if (it.PosX == target.col && it.PosY == target.row) continue;
                if (!IsTargetAreaFree(occupancy, target.col, target.row, it.SizeX, it.SizeY,
                        it.PosX, it.PosY, it.SizeX, it.SizeY)) continue;
                itemToMove = it;
                destCol = target.col;
                destRow = target.row;
                break;
            }

            // Deadlock: every pending item's target is blocked by another item.
            // Break the cycle by parking one blocked item in any free area that is NOT its own cell.
            if (itemToMove == null)
            {
                Log("SortInventory: deadlock detected — searching for park slot");
                foreach (var it in currentItems)
                {
                    if (it.Item == null || !plan.TryGetValue(it.Item.Address, out var target)) continue;
                    if (it.PosX == target.col && it.PosY == target.row) continue;

                    var tempGrid = BuildOccupancyGrid(currentItems);

                    // Mark ignored cells as off-limits so items are never parked there.
                    // Items in ignored cells are excluded from currentItems and therefore
                    // invisible to BuildOccupancyGrid; without this step the park algorithm
                    // would treat those cells as free and park items into them, where the
                    // convergence check cannot see them.
                    for (var row = 0; row < 5; row++)
                        for (var col = 0; col < 12; col++)
                            if (Settings.IgnoredCells[row, col]) tempGrid[col, row] = true;

                    for (var dy = 0; dy < it.SizeY; dy++)
                        for (var dx = 0; dx < it.SizeX; dx++)
                            tempGrid[it.PosX + dx, it.PosY + dy] = false;

                    // Prefer a park slot that doesn't overlap any PENDING plan target (a target
                    // whose owner hasn't reached it yet).  Parking on a pending target causes a
                    // cascading chain: the target's owner later bumps into the parked item and
                    // PoE performs an implicit swap that discards the displaced item's plan key,
                    // making it invisible to all subsequent moves.
                    // Build a mask of pending target cells, overlay it on tempGrid, and try to
                    // find a "clean" slot first.  If no clean slot exists (inventory nearly full),
                    // fall back to any free slot so the sort never hard-deadlocks.
                    var pendingTargetGrid = new bool[12, 5];
                    foreach (var kv in plan)
                    {
                        // Skip items already at their target — those cells are occupied in
                        // tempGrid anyway, so they're already blocked by BuildOccupancyGrid.
                        if (currentItems.Any(ci =>
                                ci.Item?.Address == kv.Key &&
                                ci.PosX == kv.Value.col &&
                                ci.PosY == kv.Value.row))
                            continue;
                        for (var dy2 = 0; dy2 < kv.Value.sizeY; dy2++)
                            for (var dx2 = 0; dx2 < kv.Value.sizeX; dx2++)
                            {
                                var tc = kv.Value.col + dx2;
                                var tr = kv.Value.row + dy2;
                                if (tc >= 0 && tc < 12 && tr >= 0 && tr < 5)
                                    pendingTargetGrid[tc, tr] = true;
                            }
                    }
                    var tempGridNonTarget = (bool[,])tempGrid.Clone();
                    for (var tr2 = 0; tr2 < 5; tr2++)
                        for (var tc2 = 0; tc2 < 12; tc2++)
                            if (pendingTargetGrid[tc2, tr2]) tempGridNonTarget[tc2, tr2] = true;

                    // Try a slot that avoids pending plan targets; fall back to any free slot.
                    var parkSlot = FindFirstFitNotAt(tempGridNonTarget, it.SizeX, it.SizeY, it.PosX, it.PosY)
                                   ?? FindFirstFitNotAt(tempGrid, it.SizeX, it.SizeY, it.PosX, it.PosY);
                    if (parkSlot == null)
                    {
                        Log($"SortInventory: no park slot for {it.Item.Path} ({it.SizeX}x{it.SizeY}) at ({it.PosX},{it.PosY})");
                        continue;
                    }

                    Log($"SortInventory: parking {it.Item.Path} ({it.SizeX}x{it.SizeY}) from ({it.PosX},{it.PosY}) → park ({parkSlot.Value.col},{parkSlot.Value.row})");
                    itemToMove = it;
                    destCol = parkSlot.Value.col;
                    destRow = parkSlot.Value.row;
                    break;
                }
            }

            if (itemToMove == null)
            {
                Log("SortInventory: no moveable item found (deadlock unresolvable) — stopping");
                break;
            }

            // Source: click the centre of the item's anchor cell.
            // PoE attaches the held item to the cursor at the clicked cell (the "anchor").
            // We use floor(SizeX/2), floor(SizeY/2) as the anchor — the same cell used for
            // the destination click — so pickup and placement are fully consistent.
            // Example: 2×3 item → anchor at (1,1); click centre of that cell both to grab
            // and to place, so the item's top-left always lands exactly at the target slot.
            var itemRect = itemToMove.GetClientRect();
            var srcCellW = itemToMove.SizeX > 0 ? itemRect.Width / itemToMove.SizeX : cellW;
            var srcCellH = itemToMove.SizeY > 0 ? itemRect.Height / itemToMove.SizeY : cellH;
            var srcCenter = new SharpDX.Vector2(
                itemRect.X + (itemToMove.SizeX / 2 + 0.5f) * srcCellW,
                itemRect.Y + (itemToMove.SizeY / 2 + 0.5f) * srcCellH);

            // Destination: click the centre of the anchor cell at the target position.
            // Integer division gives floor(SizeX/2); +0.5f centres within that cell.
            var dstCenter = new SharpDX.Vector2(
                gridOriginX + (destCol + itemToMove.SizeX / 2 + 0.5f) * cellW,
                gridOriginY + (destRow + itemToMove.SizeY / 2 + 0.5f) * cellH);

            Log($"SortInventory: move #{moveCount + 1} — {itemToMove.Item?.Path} ({itemToMove.SizeX}x{itemToMove.SizeY}) " +
                $"grid ({itemToMove.PosX},{itemToMove.PosY}) → ({destCol},{destRow}) | " +
                $"itemRect=({itemRect.X:F0},{itemRect.Y:F0},{itemRect.Width:F0}x{itemRect.Height:F0}) " +
                $"srcCellWH=({srcCellW:F2},{srcCellH:F2}) " +
                $"src=({srcCenter.X:F0},{srcCenter.Y:F0}) dst=({dstCenter.X:F0},{dstCenter.Y:F0}) " +
                $"src+offset=({srcCenter.X + WindowOffset.X:F0},{srcCenter.Y + WindowOffset.Y:F0}) " +
                $"dst+offset=({dstCenter.X + WindowOffset.X:F0},{dstCenter.Y + WindowOffset.Y:F0})");

            // Pick up (left-click on source).
            Mouse.moveMouse(srcCenter + WindowOffset);
            await Wait(MouseMoveDelay, true);
            Mouse.LeftDown();
            await Wait(MouseDownDelay, true);
            Mouse.LeftUp();
            await Wait(KeyDelay, true);

            // Place (left-click on destination).
            Mouse.moveMouse(dstCenter + WindowOffset);
            await Wait(MouseMoveDelay, true);
            Mouse.LeftDown();
            await Wait(MouseDownDelay, true);
            Mouse.LeftUp();
            await Wait(KeyDelay, true);

            // Record the old address and destination so the plan key can be re-keyed
            // at the start of the next iteration once the server reflects the placement.
            pendingOldAddress = itemToMove.Item.Address;
            pendingDestCol = destCol;
            pendingDestRow = destRow;
            pendingRetryCount = 0; // fresh retry budget for this placement

            moveCount++;
        }

        if (moveCount >= MaxMoves)
            Log($"SortInventory: reached move limit ({MaxMoves}) — stopping");

        Log($"SortInventory: finished — total moves={moveCount}");
        await StopMovingItems();
        return true;
    }

    /// <summary>
    /// Derives the inventory grid's per-cell dimensions and top-left origin by averaging
    /// the values implied by every item's actual on-screen rect and known grid position.
    /// This is more accurate than dividing the outer panel rect by 12×5, because the
    /// panel typically includes padding/borders outside the playable cell area.
    /// Falls back to the panel rect when no valid item rects are available.
    /// </summary>
    private static (float cellW, float cellH, float originX, float originY) CalibrateGridGeometry(
        List<ServerInventory.InventSlotItem> items, SharpDX.RectangleF fallbackPanelRect)
    {
        double sumW = 0, sumH = 0, sumOX = 0, sumOY = 0;
        var n = 0;
        foreach (var item in items)
        {
            var r = item.GetClientRect();
            if (r.Width <= 0 || r.Height <= 0 || item.SizeX <= 0 || item.SizeY <= 0) continue;
            var w = r.Width / item.SizeX;
            var h = r.Height / item.SizeY;
            sumW += w;
            sumH += h;
            sumOX += r.X - item.PosX * w;
            sumOY += r.Y - item.PosY * h;
            n++;
        }
        if (n > 0)
            return ((float)(sumW / n), (float)(sumH / n), (float)(sumOX / n), (float)(sumOY / n));

        // Fallback: derive from the panel rect (less accurate but better than nothing).
        return (fallbackPanelRect.Width / 12f, fallbackPanelRect.Height / 5f,
                fallbackPanelRect.X, fallbackPanelRect.Y);
    }

    /// <summary>
    /// Extracts a logical grouping category from an item's path so that the sort algorithm
    /// keeps items of the same type (Maps, Currency, Rings, etc.) clustered together.
    /// PoE item paths follow the pattern "Metadata/Items/[Category]/...".
    /// </summary>
    private static string GetItemCategory(ServerInventory.InventSlotItem item)
    {
        var path = item.Item?.Path ?? "";
        var segments = path.Split('/');
        // segments[0]="Metadata", segments[1]="Items", segments[2]="Maps"|"Currency"|"Rings"|...
        return segments.Length > 2 ? segments[2] : "Misc";
    }

    /// <summary>
    /// Builds a bool[12,5] occupancy grid where true = cell is occupied by an inventory item.
    /// grid[col, row] -- col is x (0-11), row is y (0-4).
    /// </summary>
    private static bool[,] BuildOccupancyGrid(IEnumerable<ServerInventory.InventSlotItem> items)
    {
        var grid = new bool[12, 5];
        foreach (var item in items)
        {
            for (var dy = 0; dy < item.SizeY; dy++)
                for (var dx = 0; dx < item.SizeX; dx++)
                {
                    var c = item.PosX + dx;
                    var r = item.PosY + dy;
                    if (c >= 0 && c < 12 && r >= 0 && r < 5)
                        grid[c, r] = true;
                }
        }
        return grid;
    }

    /// <summary>
    /// Builds a bool[12,5] grid pre-populated only with ignored cells (as per settings).
    /// Used as the starting grid for bin-packing layout computation.
    /// </summary>
    private bool[,] BuildIgnoredCellsGrid()
    {
        var grid = new bool[12, 5];
        for (var r = 0; r < 5; r++)
            for (var c = 0; c < 12; c++)
                if (Settings.IgnoredCells[r, c]) grid[c, r] = true;
        return grid;
    }

    /// <summary>
    /// Returns true if the rectangular region [targetCol..+sizeX, targetRow..+sizeY] in the
    /// occupancy grid is entirely free, except for cells that belong to the source item's own
    /// current footprint (which will vacate when the item moves).
    /// </summary>
    private static bool IsTargetAreaFree(bool[,] occupancy,
        int targetCol, int targetRow, int sizeX, int sizeY,
        int srcCol, int srcRow, int srcSizeX, int srcSizeY)
    {
        if (targetCol < 0 || targetRow < 0 || targetCol + sizeX > 12 || targetRow + sizeY > 5)
            return false;

        for (var dy = 0; dy < sizeY; dy++)
        {
            for (var dx = 0; dx < sizeX; dx++)
            {
                var c = targetCol + dx;
                var r = targetRow + dy;
                if (!occupancy[c, r]) continue;

                // Cell is occupied - is it within the source item's own footprint?
                if (c >= srcCol && c < srcCol + srcSizeX && r >= srcRow && r < srcRow + srcSizeY)
                    continue; // will be vacated when the item moves, so it's fine

                return false; // occupied by a different item
            }
        }
        return true;
    }

    private static (int col, int row)? FindFirstFit(bool[,] grid, int width, int height)
    {
        // Scan column-first so items are packed towards the left side of the inventory.
        // col is the outer loop (0..11) and row is the inner loop (0..4), which fills
        // the leftmost column top-to-bottom before moving to the next column.
        for (var col = 0; col <= 12 - width; col++)
        {
            for (var row = 0; row <= 5 - height; row++)
            {
                var fits = true;
                for (var dr = 0; dr < height && fits; dr++)
                    for (var dc = 0; dc < width && fits; dc++)
                        if (grid[col + dc, row + dr]) fits = false;
                if (fits) return (col, row);
            }
        }
        return null;
    }

    /// <summary>
    /// Like <see cref="FindFirstFit"/> but rejects the slot at (excludeCol, excludeRow) so
    /// that an item is never "parked" back at its own current position (a no-op move).
    /// </summary>
    private static (int col, int row)? FindFirstFitNotAt(
        bool[,] grid, int width, int height, int excludeCol, int excludeRow)
    {
        // Column-first scan mirrors FindFirstFit for consistent left-side-first packing.
        for (var col = 0; col <= 12 - width; col++)
        {
            for (var row = 0; row <= 5 - height; row++)
            {
                if (col == excludeCol && row == excludeRow) continue;
                var fits = true;
                for (var dr = 0; dr < height && fits; dr++)
                    for (var dc = 0; dc < width && fits; dc++)
                        if (grid[col + dc, row + dr]) fits = false;
                if (fits) return (col, row);
            }
        }
        return null;
    }

    // ── Stash sort helpers ────────────────────────────────────────────────────────────

    /// <summary>
    /// Calibrates the stash grid's per-cell pixel dimensions and top-left origin by
    /// averaging values implied by each item's screen rect and logical size.
    /// Because <see cref="NormalInventoryItem.InventPosX"/> / <see cref="NormalInventoryItem.InventPosY"/>
    /// are obsolete and return 0 in current ExileCore builds, grid positions are bootstrapped
    /// by rounding each item's screen-rect offset against the panel rect, then refining the
    /// origin once cellW/cellH are known.
    /// </summary>
    private static (float cellW, float cellH, float originX, float originY, int cols, int rows)
        CalibrateStashGrid(IEnumerable<NormalInventoryItem> items, SharpDX.RectangleF panelRect)
    {
        var itemList = items.Where(x => x.ItemWidth > 0 && x.ItemHeight > 0).ToList();
        if (!itemList.Any())
            return (panelRect.Width / 12f, panelRect.Height / 12f, panelRect.X, panelRect.Y, 12, 12);

        // Step 1: estimate cell size from item screen rects and item sizes.
        double sumW = 0, sumH = 0;
        var n = 0;
        foreach (var item in itemList)
        {
            var r = item.GetClientRectCache;
            if (r.Width <= 0 || r.Height <= 0) continue;
            sumW += r.Width / (double)item.ItemWidth;
            sumH += r.Height / (double)item.ItemHeight;
            n++;
        }
        if (n == 0)
            return (panelRect.Width / 12f, panelRect.Height / 12f, panelRect.X, panelRect.Y, 12, 12);

        var cellW = (float)(sumW / n);
        var cellH = (float)(sumH / n);

        // Step 2: for each item estimate its grid position (using the panel's left/top edge as
        // a first approximation of the origin), then compute the precise origin.
        double sumOX = 0, sumOY = 0;
        var maxCol = 0;
        var maxRow = 0;
        foreach (var item in itemList)
        {
            var r = item.GetClientRectCache;
            var estPosX = (int)Math.Round((r.X - panelRect.X) / cellW);
            var estPosY = (int)Math.Round((r.Y - panelRect.Y) / cellH);
            sumOX += r.X - estPosX * cellW;
            sumOY += r.Y - estPosY * cellH;
            maxCol = Math.Max(maxCol, estPosX + item.ItemWidth);
            maxRow = Math.Max(maxRow, estPosY + item.ItemHeight);
        }

        return (cellW, cellH,
                (float)(sumOX / itemList.Count),
                (float)(sumOY / itemList.Count),
                Math.Max(maxCol, 12),
                Math.Max(maxRow, 12));
    }

    /// <summary>
    /// Like <see cref="FindFirstFit"/> but works on an arbitrarily-sized grid.
    /// Scans column-first (left-to-right outer, top-to-bottom inner) for consistent
    /// left-side-first bin-packing, matching the inventory sort order.
    /// </summary>
    private static (int col, int row)? FindFirstFitInGrid(bool[,] grid, int maxCols, int maxRows, int w, int h)
    {
        for (var col = 0; col <= maxCols - w; col++)
            for (var row = 0; row <= maxRows - h; row++)
            {
                var fits = true;
                for (var dr = 0; dr < h && fits; dr++)
                    for (var dc = 0; dc < w && fits; dc++)
                        if (grid[col + dc, row + dr]) fits = false;
                if (fits) return (col, row);
            }
        return null;
    }

    /// <summary>
    /// Like <see cref="IsTargetAreaFree"/> but works on an arbitrarily-sized grid.
    /// Pass srcSizeX = srcSizeY = 0 when there is no source footprint to exclude.
    /// </summary>
    private static bool IsTargetAreaFreeInGrid(
        bool[,] occupancy, int maxCols, int maxRows,
        int targetCol, int targetRow, int sizeX, int sizeY,
        int srcCol, int srcRow, int srcSizeX, int srcSizeY)
    {
        if (targetCol < 0 || targetRow < 0 || targetCol + sizeX > maxCols || targetRow + sizeY > maxRows)
            return false;
        for (var dy = 0; dy < sizeY; dy++)
            for (var dx = 0; dx < sizeX; dx++)
            {
                var c = targetCol + dx;
                var r = targetRow + dy;
                if (!occupancy[c, r]) continue;
                if (srcSizeX > 0 && srcSizeY > 0 &&
                    c >= srcCol && c < srcCol + srcSizeX &&
                    r >= srcRow && r < srcRow + srcSizeY)
                    continue; // within source footprint — will be vacated
                return false;
            }
        return true;
    }

    // ── SortStash ─────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Sorts the currently-visible stash tab using the same column-first, area-ascending
    /// bin-packing order as <see cref="SortInventory"/>.  Because stash tabs are typically
    /// full, the player inventory is used as a staging area: when all remaining stash items
    /// are deadlocked (every target is occupied by another item that also needs to move),
    /// one item is Ctrl+Clicked to the inventory to free its stash cell, enabling other
    /// direct stash→stash moves, after which the staged item is returned to its target.
    /// This cycle repeats until all items are at their sorted positions.
    /// </summary>
    private async SyncTask<bool> SortStash(SharpDX.RectangleF stashPanelRect)
    {
        if (!await MoveItemsCommonPreamble())
        {
            Log("SortStash: preamble returned false");
            return false;
        }
        _prevMousePos = Mouse.GetCursorPosition();

        if (!InGameState.IngameUi.InventoryPanel.IsVisible)
        {
            Log("SortStash: inventory panel must be open (needed as staging area) — aborting");
            DebugWindow.LogMsg("HighlightedItems: Open your inventory before sorting the stash.");
            await StopMovingItems();
            return false;
        }

        // ── PHASE 1: Read stash and build sort plan ───────────────────────────────────
        var stashEl = (InGameState.IngameUi.StashElement, InGameState.IngameUi.GuildStashElement) switch
        {
            ({ IsVisible: true, VisibleStash: { } vs }, _) => vs,
            (_, { IsVisible: true, VisibleStash: { } vs }) => vs,
            _ => null
        };
        if (stashEl == null)
        {
            Log("SortStash: no visible stash — aborting");
            await StopMovingItems();
            return false;
        }

        var initialItems = stashEl.VisibleInventoryItems
            .Where(x => x?.Item != null && x.ItemWidth > 0 && x.ItemHeight > 0)
            .ToList();

        Log($"SortStash: starting — {initialItems.Count} items, panelRect={stashPanelRect}");
        if (!initialItems.Any())
        {
            Log("SortStash: stash is empty — nothing to sort");
            await StopMovingItems();
            return true;
        }

        // Calibrate stash grid geometry from visible items' screen rects.
        var (cellW, cellH, originX, originY, stashCols, stashRows) =
            CalibrateStashGrid(initialItems, stashPanelRect);
        Log($"SortStash: grid {stashCols}×{stashRows}, cellW={cellW:F2} cellH={cellH:F2} origin=({originX:F2},{originY:F2})");

        // Local helper: derive grid position for a NormalInventoryItem from its screen rect.
        (int posX, int posY) GridPos(NormalInventoryItem item)
        {
            var r = item.GetClientRectCache;
            return ((int)Math.Round((r.X - originX) / cellW),
                    (int)Math.Round((r.Y - originY) / cellH));
        }

        // Build sort plan: same ordering as SortInventory (area asc → narrower first → category → path → address).
        var packingOrder = initialItems
            .OrderBy(x => x.ItemWidth * x.ItemHeight)
            .ThenBy(x => x.ItemWidth)
            .ThenBy(x => x.Item?.Path?.Split('/') is { Length: > 2 } seg ? seg[2] : "Misc")
            .ThenBy(x => x.Item?.Path ?? "")
            .ThenBy(x => x.Item?.Address ?? 0L)
            .ToList();

        var targetGrid = new bool[stashCols, stashRows];
        // plan[currentAddress] = (targetCol, targetRow, sizeX, sizeY)
        var plan = new Dictionary<long, (int col, int row, int sizeX, int sizeY)>();
        foreach (var it in packingOrder)
        {
            if (it.Item == null) continue;
            var fit = FindFirstFitInGrid(targetGrid, stashCols, stashRows, it.ItemWidth, it.ItemHeight);
            if (fit == null)
            {
                Log($"SortStash: NOFIT — {it.Item.Path} ({it.ItemWidth}×{it.ItemHeight})");
                continue;
            }
            var (px, py) = GridPos(it);
            plan[it.Item.Address] = (fit.Value.col, fit.Value.row, it.ItemWidth, it.ItemHeight);
            Log($"SortStash: plan — {it.Item.Path} ({it.ItemWidth}×{it.ItemHeight}) ({px},{py}) → ({fit.Value.col},{fit.Value.row})");
            for (var dy = 0; dy < it.ItemHeight; dy++)
                for (var dx = 0; dx < it.ItemWidth; dx++)
                    targetGrid[fit.Value.col + dx, fit.Value.row + dy] = true;
        }
        Log($"SortStash: plan complete — {plan.Count}/{initialItems.Count} items assigned targets");

        // ── PHASE 2: Execute plan ─────────────────────────────────────────────────────
        // stagingSet: addresses of plan items currently in the player inventory (not in stash).
        var stagingSet = new HashSet<long>();

        // Pending-move tracking: mirrors the scheme used in SortInventory.
        long pendingOldAddress = 0;
        int pendingDestCol = -1, pendingDestRow = -1;
        bool pendingIsStashToInventory = false; // true when a Ctrl+Click staging move is pending
        HashSet<long> preStageInvAddresses = null; // inventory snapshot taken before staging ctrl+click
        int pendingRetryCount = 0;

        // Upper bound: each of the N stash cells can require at most 2 moves in a sort (one
        // staging move stash→inventory and one return move inventory→stash), plus a direct
        // stash→stash placement, giving 3 moves per cell.  The +60 covers the extra round-trip
        // moves for inventory staging items (up to one full inventory worth of items in flight).
        var maxMoves = stashCols * stashRows * 3 + 60;
        var moveCount = 0;

        while (moveCount < maxMoves)
        {
            if (MoveCancellationRequested)
            {
                Log("SortStash: cancelled by right-click");
                await StopMovingItems();
                return false;
            }

            var loopStashEl = (InGameState.IngameUi.StashElement, InGameState.IngameUi.GuildStashElement) switch
            {
                ({ IsVisible: true, VisibleStash: { } vs }, _) => vs,
                (_, { IsVisible: true, VisibleStash: { } vs }) => vs,
                _ => null
            };
            if (loopStashEl == null)
            {
                Log("SortStash: stash panel closed — stopping");
                break;
            }
            if (!InGameState.IngameUi.InventoryPanel.IsVisible)
            {
                Log("SortStash: inventory panel closed — stopping");
                break;
            }

            var currentStashItems = loopStashEl.VisibleInventoryItems
                .Where(x => x?.Item != null && x.ItemWidth > 0 && x.ItemHeight > 0)
                .ToList();
            var rawInvItems = GameController.IngameState.ServerData.PlayerInventories[0].Inventory.InventorySlotItems;
            var currentInvItems = rawInvItems.Where(x => !IsInIgnoreCell(x)).ToList();

            // ── Plan re-keying (PoE reassigns Item.Address on each slot placement) ──────
            if (pendingOldAddress != 0)
            {
                var rekeyed = false;

                if (pendingIsStashToInventory)
                {
                    // Stash→inventory (Ctrl+Click): find the new inventory item that appeared.
                    if (plan.TryGetValue(pendingOldAddress, out var pendingEntry))
                    {
                        var newInvItem = rawInvItems.FirstOrDefault(x =>
                            x.Item != null &&
                            !plan.ContainsKey(x.Item.Address) &&
                            preStageInvAddresses != null &&
                            !preStageInvAddresses.Contains(x.Item.Address) &&
                            x.SizeX == pendingEntry.sizeX &&
                            x.SizeY == pendingEntry.sizeY);
                        if (newInvItem != null)
                        {
                            Log($"SortStash: staged 0x{pendingOldAddress:X} → inv 0x{newInvItem.Item.Address:X}");
                            plan[newInvItem.Item.Address] = pendingEntry;
                            plan.Remove(pendingOldAddress);
                            stagingSet.Add(newInvItem.Item.Address);
                            pendingOldAddress = 0;
                            pendingRetryCount = 0;
                            rekeyed = true;
                        }
                    }
                    else
                    {
                        // Entry already re-keyed by a previous iteration.
                        pendingOldAddress = 0;
                        rekeyed = true;
                    }
                }
                else
                {
                    // Stash→stash or inventory→stash: find item at target stash position.
                    var newStashItem = currentStashItems.FirstOrDefault(x =>
                    {
                        var (px, py) = GridPos(x);
                        return px == pendingDestCol && py == pendingDestRow;
                    });
                    if (newStashItem != null)
                    {
                        if (newStashItem.Item.Address != pendingOldAddress && plan.ContainsKey(pendingOldAddress))
                        {
                            var e = plan[pendingOldAddress];
                            if (newStashItem.ItemWidth == e.sizeX && newStashItem.ItemHeight == e.sizeY)
                            {
                                Log($"SortStash: re-keyed 0x{pendingOldAddress:X} → 0x{newStashItem.Item.Address:X} at ({pendingDestCol},{pendingDestRow})");
                                plan[newStashItem.Item.Address] = e;
                                plan.Remove(pendingOldAddress);
                                stagingSet.Remove(pendingOldAddress); // no-op for stash→stash; removes for inv→stash
                            }
                        }
                        pendingOldAddress = 0;
                        pendingRetryCount = 0;
                        rekeyed = true;
                    }
                }

                if (!rekeyed)
                {
                    if (++pendingRetryCount > 10)
                    {
                        Log($"SortStash: server confirmation timeout after {pendingRetryCount} retries — stopping");
                        break;
                    }
                    await Wait(KeyDelay, true);
                    continue;
                }
            }
            pendingRetryCount = 0;

            // ── Convergence check ─────────────────────────────────────────────────────
            if (stagingSet.Count == 0 && plan.All(kv =>
            {
                var item = currentStashItems.FirstOrDefault(x => x.Item?.Address == kv.Key);
                if (item == null) return false;
                var (px, py) = GridPos(item);
                return px == kv.Value.col && py == kv.Value.row;
            }))
            {
                Log($"SortStash: all items at target — done in {moveCount} moves");
                break;
            }

            // Build stash and inventory occupancy grids.
            var stashOcc = new bool[stashCols, stashRows];
            foreach (var it in currentStashItems)
            {
                var (px, py) = GridPos(it);
                for (var dy = 0; dy < it.ItemHeight; dy++)
                    for (var dx = 0; dx < it.ItemWidth; dx++)
                    {
                        var c = px + dx; var r = py + dy;
                        if (c >= 0 && c < stashCols && r >= 0 && r < stashRows)
                            stashOcc[c, r] = true;
                    }
            }
            var invOcc = BuildOccupancyGrid(currentInvItems);

            var moved = false;

            // ── Priority 1: Return staged inventory items to their stash targets ──────
            foreach (var addr in stagingSet.ToList())
            {
                if (!plan.TryGetValue(addr, out var target)) continue;
                var invItem = rawInvItems.FirstOrDefault(x => x.Item?.Address == addr);
                if (invItem == null) continue;

                if (!IsTargetAreaFreeInGrid(stashOcc, stashCols, stashRows,
                        target.col, target.row, target.sizeX, target.sizeY, -1, -1, 0, 0))
                    continue;

                var invR = invItem.GetClientRect();
                var iCW = invItem.SizeX > 0 ? invR.Width / invItem.SizeX : invR.Width;
                var iCH = invItem.SizeY > 0 ? invR.Height / invItem.SizeY : invR.Height;
                var srcCenter = new SharpDX.Vector2(
                    invR.X + (invItem.SizeX / 2 + 0.5f) * iCW,
                    invR.Y + (invItem.SizeY / 2 + 0.5f) * iCH);
                var dstCenter = new SharpDX.Vector2(
                    originX + (target.col + target.sizeX / 2 + 0.5f) * cellW,
                    originY + (target.row + target.sizeY / 2 + 0.5f) * cellH);

                Log($"SortStash: move #{moveCount + 1} [inv→stash] {invItem.Item?.Path} ({target.sizeX}×{target.sizeY}) → stash({target.col},{target.row})");
                Mouse.moveMouse(srcCenter + WindowOffset);
                await Wait(MouseMoveDelay, true);
                Mouse.LeftDown(); await Wait(MouseDownDelay, true); Mouse.LeftUp();
                await Wait(KeyDelay, true);
                Mouse.moveMouse(dstCenter + WindowOffset);
                await Wait(MouseMoveDelay, true);
                Mouse.LeftDown(); await Wait(MouseDownDelay, true); Mouse.LeftUp();
                await Wait(KeyDelay, true);

                pendingOldAddress = addr;
                pendingDestCol = target.col;
                pendingDestRow = target.row;
                pendingIsStashToInventory = false;
                pendingRetryCount = 0;
                moveCount++;
                moved = true;
                break;
            }
            if (moved) continue;

            // ── Priority 2: Direct stash→stash moves (target area is free) ───────────
            foreach (var it in currentStashItems)
            {
                if (it.Item == null || !plan.TryGetValue(it.Item.Address, out var target)) continue;
                var (px, py) = GridPos(it);
                if (px == target.col && py == target.row) continue; // already at target
                if (!IsTargetAreaFreeInGrid(stashOcc, stashCols, stashRows,
                        target.col, target.row, target.sizeX, target.sizeY,
                        px, py, it.ItemWidth, it.ItemHeight))
                    continue;

                var ir = it.GetClientRectCache;
                var sCW = it.ItemWidth > 0 ? ir.Width / it.ItemWidth : cellW;
                var sCH = it.ItemHeight > 0 ? ir.Height / it.ItemHeight : cellH;
                var srcCenter = new SharpDX.Vector2(
                    ir.X + (it.ItemWidth / 2 + 0.5f) * sCW,
                    ir.Y + (it.ItemHeight / 2 + 0.5f) * sCH);
                var dstCenter = new SharpDX.Vector2(
                    originX + (target.col + target.sizeX / 2 + 0.5f) * cellW,
                    originY + (target.row + target.sizeY / 2 + 0.5f) * cellH);

                Log($"SortStash: move #{moveCount + 1} [stash→stash] {it.Item?.Path} ({it.ItemWidth}×{it.ItemHeight}) ({px},{py})→({target.col},{target.row})");
                Mouse.moveMouse(srcCenter + WindowOffset);
                await Wait(MouseMoveDelay, true);
                Mouse.LeftDown(); await Wait(MouseDownDelay, true); Mouse.LeftUp();
                await Wait(KeyDelay, true);
                Mouse.moveMouse(dstCenter + WindowOffset);
                await Wait(MouseMoveDelay, true);
                Mouse.LeftDown(); await Wait(MouseDownDelay, true); Mouse.LeftUp();
                await Wait(KeyDelay, true);

                pendingOldAddress = it.Item.Address;
                pendingDestCol = target.col;
                pendingDestRow = target.row;
                pendingIsStashToInventory = false;
                pendingRetryCount = 0;
                moveCount++;
                moved = true;
                break;
            }
            if (moved) continue;

            // ── Priority 3: Stage one stash item to inventory to break the deadlock ───
            NormalInventoryItem toStage = null;
            foreach (var it in currentStashItems)
            {
                if (it.Item == null || !plan.TryGetValue(it.Item.Address, out var target)) continue;
                var (px, py) = GridPos(it);
                if (px == target.col && py == target.row) continue; // already at target
                // Only stage if inventory has space for this item's footprint.
                if (FindFirstFit(invOcc, it.ItemWidth, it.ItemHeight) == null) continue;
                toStage = it;
                break;
            }
            if (toStage == null)
            {
                Log("SortStash: deadlock — no item can be staged (inventory full or no unsettled items)");
                break;
            }

            // Snapshot inventory addresses before the Ctrl+Click so we can identify the new item.
            preStageInvAddresses = new HashSet<long>(
                rawInvItems.Where(x => x.Item != null).Select(x => x.Item.Address));

            var sr = toStage.GetClientRectCache;
            var ssCW = toStage.ItemWidth > 0 ? sr.Width / toStage.ItemWidth : cellW;
            var ssCH = toStage.ItemHeight > 0 ? sr.Height / toStage.ItemHeight : cellH;
            var stageCenter = new SharpDX.Vector2(
                sr.X + (toStage.ItemWidth / 2 + 0.5f) * ssCW,
                sr.Y + (toStage.ItemHeight / 2 + 0.5f) * ssCH);
            var (spx, spy) = GridPos(toStage);
            Log($"SortStash: staging {toStage.Item?.Path} ({toStage.ItemWidth}×{toStage.ItemHeight}) from stash({spx},{spy}) to inventory");

            Keyboard.KeyDown(Keys.LControlKey);
            await Wait(KeyDelay, true);
            Mouse.moveMouse(stageCenter + WindowOffset);
            await Wait(MouseMoveDelay, true);
            Mouse.LeftDown(); await Wait(MouseDownDelay, true); Mouse.LeftUp();
            Keyboard.KeyUp(Keys.LControlKey);
            await Wait(KeyDelay, true);

            pendingOldAddress = toStage.Item.Address;
            pendingIsStashToInventory = true;
            pendingRetryCount = 0;
            moveCount++;
        }

        if (moveCount >= maxMoves)
            Log($"SortStash: reached move limit ({maxMoves}) — stopping");

        Log($"SortStash: finished — total moves={moveCount}");
        await StopMovingItems();
        return true;
    }
}
