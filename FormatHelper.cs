using System;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace decksterity
{
    // Applies fill, line and text colors and line weight to the current selection (used by the color pane)
    // Colors are passed as 0xRRGGBB; null means "no fill" / "no line"
    // Groups are applied member by member; tables get fill and text color on the selected cells (or all cells)
    public static class FormatHelper
    {
        private static Application App => Globals.ThisAddIn.Application;

        public static void ApplyFill(int? rgb)
        {
            App.StartNewUndoEntry();
            foreach (var shape in SelectionHelper.ExpandGroups(SelectionHelper.GetSelectedShapes()))
            {
                if (SelectionHelper.IsTable(shape))
                {
                    foreach (var cell in SelectionHelper.GetTargetCells(shape.Table))
                    {
                        SetFill(cell.Shape.Fill, rgb);
                    }
                }
                else
                {
                    Try(() => SetFill(shape.Fill, rgb)); // lines and some placeholders have no fill
                }
            }
        }

        public static void ApplyLine(int? rgb)
        {
            App.StartNewUndoEntry();
            foreach (var shape in SelectionHelper.ExpandGroups(SelectionHelper.GetSelectedShapes()))
            {
                if (SelectionHelper.IsTable(shape)) continue;
                Try(() =>
                {
                    if (rgb.HasValue)
                    {
                        shape.Line.Visible = Office.MsoTriState.msoTrue;
                        shape.Line.ForeColor.RGB = ElementHelper.ConvertRgbToBgr(rgb.Value);
                    }
                    else
                    {
                        shape.Line.Visible = Office.MsoTriState.msoFalse;
                    }
                });
            }
        }

        // While editing text, colors only the highlighted characters; otherwise all text in the selected shapes / cells
        public static void ApplyTextColor(int rgb)
        {
            int bgr = ElementHelper.ConvertRgbToBgr(rgb);
            var selection = App.ActiveWindow.Selection;

            App.StartNewUndoEntry();
            if (selection.Type == PpSelectionType.ppSelectionText)
            {
                selection.TextRange.Font.Color.RGB = bgr;
                return;
            }

            foreach (var shape in SelectionHelper.ExpandGroups(SelectionHelper.GetSelectedShapes()))
            {
                if (SelectionHelper.IsTable(shape))
                {
                    foreach (var cell in SelectionHelper.GetTargetCells(shape.Table))
                    {
                        cell.Shape.TextFrame.TextRange.Font.Color.RGB = bgr;
                    }
                }
                else if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    shape.TextFrame.TextRange.Font.Color.RGB = bgr;
                }
            }
        }

        // Sets the outline weight in points; 0 hides the line
        // Doesn't start an undo entry so a slider drag can be undone in one step (the pane starts one when the drag begins)
        public static void ApplyLineWeight(float points)
        {
            foreach (var shape in SelectionHelper.ExpandGroups(SelectionHelper.GetSelectedShapes()))
            {
                if (SelectionHelper.IsTable(shape)) continue;
                Try(() =>
                {
                    if (points <= 0)
                    {
                        shape.Line.Visible = Office.MsoTriState.msoFalse;
                    }
                    else
                    {
                        shape.Line.Visible = Office.MsoTriState.msoTrue;
                        shape.Line.Weight = points;
                    }
                });
            }
        }

        // Line weight of the first selected shape (0 if its line is hidden), or null if nothing with a line is selected
        public static float? GetSelectedLineWeight()
        {
            try
            {
                var shape = SelectionHelper.ExpandGroups(SelectionHelper.GetSelectedShapes())
                    .FirstOrDefault(s => !SelectionHelper.IsTable(s));
                if (shape == null) return null;
                return shape.Line.Visible == Office.MsoTriState.msoTrue ? shape.Line.Weight : 0f;
            }
            catch
            {
                return null; // no window, or a selection type without shapes
            }
        }

        public static void StartUndoEntry() => App.StartNewUndoEntry();

        private static void SetFill(FillFormat fill, int? rgb)
        {
            if (rgb.HasValue)
            {
                fill.Visible = Office.MsoTriState.msoTrue;
                fill.Solid();
                fill.ForeColor.RGB = ElementHelper.ConvertRgbToBgr(rgb.Value);
            }
            else
            {
                fill.Visible = Office.MsoTriState.msoFalse;
            }
        }

        private static void Try(Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Could not apply format: {ex.Message}");
            }
        }
    }
}
