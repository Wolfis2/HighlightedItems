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
                // When no items are highlighted and no custom filter is active, move ALL visible stash items
                var itemsToMove = (isCustomFilter || highlightedItems.Any())
                    ? highlightedItems
                    : inventory.VisibleInventoryItems.ToList();
                var orderedItems = itemsToMove
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
            _itemsToMove = items.Where((_, idx) => !processedIndices.Contains(idx)).Select(x => x.GetClientRect()).ToList();

            Keyboard.KeyDown(Keys.LControlKey);
            await Wait(KeyDelay, true);
            await MoveItem(item.GetClientRect().Center, isStackable);
            Keyboard.KeyUp(Keys.LControlKey);
            await Wait(KeyDelay, true);
            processedIndices.Add(i);
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
        for (var i = 0; i < items.Count; i++)
        {
            var item = items[i];
            _itemsToMove = items[i..].Select(x => x.GetClientRectCache).ToList();
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

            Keyboard.KeyDown(Keys.LControlKey);
            await Wait(KeyDelay, true);
            await MoveItem(item.GetClientRect().Center);
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

            // no empty slots, so inventory is full
            return true;
        }
        catch (Exception ex)
        {
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
        itemPosition += WindowOffset;
        Mouse.moveMouse(itemPosition);
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
        if (!await MoveItemsCommonPreamble()) return false;
        _prevMousePos = Mouse.GetCursorPosition();

        // A 12×5 inventory has 60 cells. In the worst case each iteration moves exactly one
        // item one cell closer to its target, so 60 iterations is a safe upper bound that
        // guarantees termination while covering any realistic starting arrangement.
        const int MaxIterations = 60;
        for (var iteration = 0; iteration < MaxIterations; iteration++)
        {
            if (MoveCancellationRequested) { await StopMovingItems(); return false; }
            if (!InGameState.IngameUi.InventoryPanel.IsVisible) break;

            // Re-read inventory state fresh every iteration so we always have the true layout.
            var currentItems = GameController.IngameState.ServerData.PlayerInventories[0].Inventory.InventorySlotItems
                .Where(x => !IsInIgnoreCell(x))
                .ToList();

            if (!currentItems.Any()) break;

            // --- Planning phase: compute the ideal target layout ---
            // Sort: primary by item category (groups maps/currency/rings/etc. together),
            //       then largest items first within each group (FFD packs most efficiently).
            var packingOrder = currentItems
                .OrderBy(x => GetItemCategory(x))
                .ThenByDescending(x => x.SizeX * x.SizeY)
                .ThenByDescending(x => x.SizeX)
                .ThenBy(x => x.Item?.Path ?? "")
                .ToList();

            var targetGrid = BuildIgnoredCellsGrid();
            // targetMap key = item's CURRENT (PosX, PosY); value = computed target (col, row)
            var targetMap = new Dictionary<(int, int), (int col, int row)>();
            foreach (var item in packingOrder)
            {
                var fit = FindFirstFit(targetGrid, item.SizeX, item.SizeY);
                if (fit == null) continue;
                var (tc, tr) = fit.Value;
                for (var dy = 0; dy < item.SizeY; dy++)
                    for (var dx = 0; dx < item.SizeX; dx++)
                        targetGrid[tc + dx, tr + dy] = true;
                targetMap[(item.PosX, item.PosY)] = (tc, tr);
            }

            // --- Execution phase: move one item per iteration to an empty target ---
            // Build current occupancy for collision checks.
            var occupancy = BuildOccupancyGrid(currentItems);

            // Try to find a "safe" move: item whose entire target area is currently free.
            ServerInventory.InventSlotItem itemToMove = null;
            int destCol = 0, destRow = 0;

            foreach (var item in packingOrder)
            {
                if (!targetMap.TryGetValue((item.PosX, item.PosY), out var target)) continue;
                var (tc, tr) = target;
                if (item.PosX == tc && item.PosY == tr) continue; // already in place

                if (IsTargetAreaFree(occupancy, tc, tr, item.SizeX, item.SizeY, item.PosX, item.PosY, item.SizeX, item.SizeY))
                {
                    itemToMove = item;
                    destCol = tc;
                    destRow = tr;
                    break;
                }
            }

            if (itemToMove == null)
            {
                // All moves are blocked (cycle). Break the deadlock by parking one item in
                // any free region, which will free up space for other items next iteration.
                var parkTarget = FindFirstFit(occupancy, 1, 1);
                if (parkTarget == null) break; // Inventory truly full - nothing to do

                // Pick the first item that still needs to move and can fit in a free region.
                foreach (var candidate in packingOrder)
                {
                    if (!targetMap.TryGetValue((candidate.PosX, candidate.PosY), out var t)) continue;
                    if (candidate.PosX == t.col && candidate.PosY == t.row) continue;

                    // Temporarily unmark the candidate's own cells so FindFirstFit can find
                    // a truly empty region elsewhere (not the item's own current position).
                    var tempGrid = BuildOccupancyGrid(currentItems);
                    for (var dy = 0; dy < candidate.SizeY; dy++)
                        for (var dx = 0; dx < candidate.SizeX; dx++)
                            tempGrid[candidate.PosX + dx, candidate.PosY + dy] = false;

                    var parkFit = FindFirstFit(tempGrid, candidate.SizeX, candidate.SizeY);
                    if (parkFit == null) continue;

                    itemToMove = candidate;
                    destCol = parkFit.Value.col;
                    destRow = parkFit.Value.row;
                    break;
                }

                if (itemToMove == null) break; // Can't break the deadlock
            }

            // Execute the single move: pick up then place.
            var srcCenter = GetInventorySlotCenter(inventoryRect, itemToMove.PosX, itemToMove.PosY);
            var dstCenter = GetInventorySlotCenter(inventoryRect, destCol, destRow);

            Mouse.moveMouse(srcCenter + WindowOffset);
            await Wait(MouseMoveDelay, true);
            Mouse.LeftDown();
            await Wait(MouseDownDelay, true);
            Mouse.LeftUp();
            await Wait(KeyDelay, true);

            Mouse.moveMouse(dstCenter + WindowOffset);
            await Wait(MouseMoveDelay, true);
            Mouse.LeftDown();
            await Wait(MouseDownDelay, true);
            Mouse.LeftUp();
            await Wait(KeyDelay, true);
        }

        await StopMovingItems();
        return true;
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
    /// grid[col, row] — col is x (0–11), row is y (0–4).
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

                // Cell is occupied – is it within the source item's own footprint?
                if (c >= srcCol && c < srcCol + srcSizeX && r >= srcRow && r < srcRow + srcSizeY)
                    continue; // will be vacated when the item moves, so it's fine

                return false; // occupied by a different item
            }
        }
        return true;
    }

    private static (int col, int row)? FindFirstFit(bool[,] grid, int width, int height)
    {
        for (var row = 0; row <= 5 - height; row++)
        {
            for (var col = 0; col <= 12 - width; col++)
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

    private static SharpDX.Vector2 GetInventorySlotCenter(SharpDX.RectangleF inventoryRect, int col, int row)
    {
        var cellW = inventoryRect.Width / 12f;
        var cellH = inventoryRect.Height / 5f;
        return new SharpDX.Vector2(
            inventoryRect.X + (col + 0.5f) * cellW,
            inventoryRect.Y + (row + 0.5f) * cellH
        );
    }
}
