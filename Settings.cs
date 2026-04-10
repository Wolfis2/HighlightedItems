using System.Collections.Generic;
using System.Windows.Forms;
using ExileCore.Shared.Attributes;
using ExileCore.Shared.Interfaces;
using ExileCore.Shared.Nodes;
using SharpDX;
using Vector2 = System.Numerics.Vector2;

namespace HighlightedItems;

public class Settings : ISettings
{
    [IgnoreMenu]
    public bool[,] IgnoredCells { get; set; } = new bool[5, 12];

    public ToggleNode Enable { get; set; } = new(true);

    [Menu("Enable Inventory Dump Button")]
    public ToggleNode DumpButtonEnable { get; set; } = new(true);

    [Menu("Enable Inventory Sort Button")]
    public ToggleNode SortButtonEnable { get; set; } = new(true);

    public ToggleNode ShowStackSizes { get; set; } = new(true);

    [Menu("Show Stack Count Next to Stack Size")]
    public ToggleNode ShowStackCountWithSize { get; set; } = new(true);

    public HotkeyNode MoveToInventoryHotkey { get; set; } = new(Keys.F1);
    public ToggleNode UseMoveToInventoryAsMoveToStashWhenNoHighlights { get; set; } = new(false);
    public HotkeyNode MoveToStashHotkey { get; set; } = new(Keys.None);
    public HotkeyNode SortInventoryHotkey { get; set; } = new(Keys.None);
    public ToggleNode InvertSelection { get; set; } = new(false);
    public ToggleNode ShowCustomFilterWindow { get; set; } = new(true);
    public ToggleNode ResetCustomFilterOnPanelClose { get; set; } = new(true);
    public ToggleNode UsePopupForFilterSelector { get; set; } = new(false);
    public RangeNode<int> CustomFilterFrameThickness { get; set; } = new(2, 1, 20);
    public ColorNode CustomFilterFrameColor { get; set; } = new(Color.Violet);
    public RangeNode<float> CustomFilterBorderRounding { get; set; } = new(0, 0, 25);
    public RangeNode<int> CustomFilterBorderDeflation { get; set; } = new(8, 0, 100);

    public RangeNode<int> ExtraDelay { get; set; } = new(20, 0, 100);

    [Menu("Use Thread.Sleep", "Is a little faster, but HUD will hang while clicking")]
    public ToggleNode UseThreadSleep { get; set; } = new(false);

    [Menu("Idle mouse delay", "Wait this long after the user lets go of the button and stops moving the mouse")]
    public RangeNode<int> IdleMouseDelay { get; set; } = new(200, 0, 1000);

    public ToggleNode CancelWithRightMouseButton { get; set; } = new(true);

    public ToggleNode VerifyTargetInventoryIsOpened { get; set; } = new(true);

    public ToggleNode VerifyButtonIsNotObstructed { get; set; } = new(true);

    public ToggleNode UseCustomMoveToStashButtonPosition { get; set; } = new(false);

    [ConditionalDisplay(nameof(UseCustomMoveToStashButtonPosition))]
    public RangeNode<Vector2> CustomMoveToStashButtonPosition { get; set; } = new(Vector2.Zero, Vector2.Zero, Vector2.One * 6000);

    public ToggleNode UseCustomMoveToInventoryButtonPosition { get; set; } = new(false);

    [ConditionalDisplay(nameof(UseCustomMoveToInventoryButtonPosition))]
    public RangeNode<Vector2> CustomMoveToInventoryButtonPosition { get; set; } = new(Vector2.Zero, Vector2.Zero, Vector2.One * 6000);

    [IgnoreMenu]
    public List<string> SavedFilters { get; set; } = [];

    [IgnoreMenu]
    public bool OpenSavedFilterList { get; set; } = true;
}