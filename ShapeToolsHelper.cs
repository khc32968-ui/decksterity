using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Office = Microsoft.Office.Core;

namespace decksterity
{
    // Layout and text tools: arrange in grid, copy/paste position and size, text margins, autofit, swap and merge text
    // Each operation starts a new undo entry so a single Ctrl+Z reverts it
    public static class ShapeToolsHelper
    {
        // PowerPoint's default text box insets: 0.1" left/right, 0.05" top/bottom
        private const float DefaultHorizontalMargin = 7.2f;
        private const float DefaultVerticalMargin = 3.6f;

        private static Application App => Globals.ThisAddIn.Application;

        #region Arrange in grid

        // Tiles the selected shapes into rows and columns inside the selection's current width
        // Shapes are taken in reading order (top to bottom, then left to right) and keep their heights
        public static void ArrangeInGrid()
        {
            var shapes = GetSelectedShapes();
            if (shapes.Length < 2)
            {
                ShowInfo("Select at least two shapes to arrange in a grid.");
                return;
            }

            int defaultColumns = (int)Math.Ceiling(Math.Sqrt(shapes.Length));
            string columnsInput = Interaction.InputBox(
                $"Arrange {shapes.Length} shapes in how many columns?", "Arrange in grid", defaultColumns.ToString());
            if (string.IsNullOrEmpty(columnsInput)) return; // cancelled
            if (!int.TryParse(columnsInput, out int columns) || columns < 1)
            {
                ShowError("Please enter a whole number of columns (1 or more).");
                return;
            }
            columns = Math.Min(columns, shapes.Length);

            string gapInput = Interaction.InputBox("Gap between shapes (in points):", "Arrange in grid", "10");
            if (string.IsNullOrEmpty(gapInput)) return; // cancelled
            if (!float.TryParse(gapInput, NumberStyles.Float, CultureInfo.CurrentCulture, out float gap) || gap < 0)
            {
                ShowError("Please enter a gap of 0 or more points.");
                return;
            }

            float left = shapes.Min(s => s.Left);
            float top = shapes.Min(s => s.Top);
            float right = shapes.Max(s => s.Left + s.Width);
            float cellWidth = (right - left - (columns - 1) * gap) / columns;
            if (cellWidth <= 0)
            {
                ShowError("The selection is too narrow for that many columns and that gap.");
                return;
            }

            App.StartNewUndoEntry();

            var ordered = shapes.OrderBy(s => s.Top).ThenBy(s => s.Left).ToArray();
            float rowTop = top;
            for (int start = 0; start < ordered.Length; start += columns)
            {
                var row = ordered.Skip(start).Take(columns).OrderBy(s => s.Left).ToArray();
                float rowHeight = 0;
                for (int column = 0; column < row.Length; column++)
                {
                    var shape = row[column];
                    SetSize(shape, cellWidth, null);
                    shape.Left = left + column * (cellWidth + gap);
                    shape.Top = rowTop;
                    rowHeight = Math.Max(rowHeight, shape.Height); // read after resizing: autofit may change height
                }
                rowTop += rowHeight + gap;
            }
        }

        #endregion

        #region Copy / paste position and size

        private static bool hasCopiedBounds;
        private static float copiedLeft, copiedTop, copiedWidth, copiedHeight;

        // Remembers the position and size of the (first) selected shape
        public static void CopyPositionAndSize()
        {
            var shapes = GetSelectedShapes();
            if (shapes.Length == 0)
            {
                ShowInfo("Select a shape to copy its position and size.");
                return;
            }

            var source = shapes[0];
            copiedLeft = source.Left;
            copiedTop = source.Top;
            copiedWidth = source.Width;
            copiedHeight = source.Height;
            hasCopiedBounds = true;
        }

        // Applies the copied position and/or size to every selected shape (works across slides and presentations)
        public static void PastePositionAndSize(bool position, bool size)
        {
            if (!hasCopiedBounds)
            {
                ShowInfo("Copy a shape's position and size first.");
                return;
            }

            var shapes = GetSelectedShapes();
            if (shapes.Length == 0) return;

            App.StartNewUndoEntry();
            foreach (var shape in shapes)
            {
                if (size)
                {
                    SetSize(shape, copiedWidth, copiedHeight);
                }
                if (position)
                {
                    shape.Left = copiedLeft;
                    shape.Top = copiedTop;
                }
            }
        }

        #endregion

        #region Text tools

        // Sets all text insets to zero; if they're all zero already, restores PowerPoint's defaults
        // Applies to text shapes, shapes inside groups, and selected table cells (or the whole table)
        public static void ToggleZeroMargins()
        {
            var frames = GetSelectedTextFrames();
            if (frames.Count == 0)
            {
                ShowInfo("Select text boxes, shapes with text, or table cells.");
                return;
            }

            bool allZero = frames.All(f => f.MarginLeft == 0 && f.MarginRight == 0 && f.MarginTop == 0 && f.MarginBottom == 0);
            float horizontal = allZero ? DefaultHorizontalMargin : 0;
            float vertical = allZero ? DefaultVerticalMargin : 0;

            App.StartNewUndoEntry();
            foreach (var frame in frames)
            {
                frame.MarginLeft = horizontal;
                frame.MarginRight = horizontal;
                frame.MarginTop = vertical;
                frame.MarginBottom = vertical;
            }
        }

        // Turns off "shrink text on overflow" and "resize shape to fit text" so boxes stay the size you set
        public static void AutofitOff()
        {
            var shapes = GetSelectedTextShapes();
            if (shapes.Count == 0)
            {
                ShowInfo("Select text boxes or shapes with text.");
                return;
            }

            App.StartNewUndoEntry();
            foreach (var shape in shapes)
            {
                shape.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeNone;
            }
        }

        // Swaps the text of exactly two shapes; each shape keeps its own formatting
        public static void SwapText()
        {
            var shapes = GetSelectedTextShapes();
            if (shapes.Count != 2)
            {
                ShowInfo("Select exactly two shapes with text to swap their text.");
                return;
            }

            var first = shapes[0].TextFrame.TextRange;
            var second = shapes[1].TextFrame.TextRange;
            string firstText = first.Text;
            string secondText = second.Text;

            App.StartNewUndoEntry();
            first.Text = secondText;
            second.Text = firstText;
        }

        // Merges the selected text shapes into the top-most one, as separate paragraphs in reading order
        // Character formatting, alignment, bullets and indent levels are carried over; the other shapes are deleted
        public static void MergeTextBoxes()
        {
            var shapes = GetSelectedTextShapes();
            if (shapes.Count < 2)
            {
                ShowInfo("Select two or more shapes with text to merge.");
                return;
            }

            var ordered = shapes.OrderBy(s => s.Top).ThenBy(s => s.Left).ToList();
            var target = ordered[0];
            float bottom = ordered.Max(s => s.Top + s.Height);

            App.StartNewUndoEntry();
            foreach (var source in ordered.Skip(1))
            {
                AppendWithFormatting(target, source.TextFrame.TextRange);
                source.Delete();
            }

            // Grow the merged box to cover the area the merged shapes used, unless it sizes itself to its text
            if (target.TextFrame2.AutoSize == Office.MsoAutoSize.msoAutoSizeNone && target.Top + target.Height < bottom)
            {
                target.Height = bottom - target.Top;
            }
            target.Select();
        }

        // Appends the source text as new paragraphs at the end of the target shape's text
        private static void AppendWithFormatting(Shape target, TextRange source)
        {
            string sourceText = source.Text.TrimEnd('\r');
            if (sourceText.Length == 0) return;

            // The whole-text range is re-read after each insert, since an earlier range doesn't grow to include new text
            bool targetEmpty = target.TextFrame.TextRange.Text.Length == 0;
            int paragraphOffset = 0;
            if (!targetEmpty)
            {
                paragraphOffset = target.TextFrame.TextRange.Paragraphs().Count;
                target.TextFrame.TextRange.InsertAfter("\r");
            }

            int runCount = source.Runs().Count;
            for (int i = 1; i <= runCount; i++)
            {
                var run = source.Runs(i, 1);
                string runText = run.Text;
                if (i == runCount) runText = runText.TrimEnd('\r');
                if (runText.Length == 0) continue;

                var inserted = target.TextFrame.TextRange.InsertAfter(runText);
                CopyFont(run.Font, inserted.Font);
            }

            int sourceParagraphs = source.Paragraphs().Count;
            var targetText = target.TextFrame.TextRange;
            for (int p = 1; p <= sourceParagraphs; p++)
            {
                int targetIndex = paragraphOffset + p;
                if (targetIndex > targetText.Paragraphs().Count) break;

                var from = source.Paragraphs(p, 1);
                var to = targetText.Paragraphs(targetIndex, 1);
                TryCopy(() => to.IndentLevel = from.IndentLevel);
                TryCopy(() => to.ParagraphFormat.Alignment = from.ParagraphFormat.Alignment);
                TryCopy(() => to.ParagraphFormat.Bullet.Visible = from.ParagraphFormat.Bullet.Visible);
            }
        }

        private static void CopyFont(Font from, Font to)
        {
            TryCopy(() => to.Name = from.Name);
            TryCopy(() => to.Size = from.Size);
            TryCopy(() => to.Bold = from.Bold);
            TryCopy(() => to.Italic = from.Italic);
            TryCopy(() => to.Underline = from.Underline);
            TryCopy(() => to.Color.RGB = from.Color.RGB);
        }

        // Mixed or unsupported values throw via COM; skip that property rather than abort the merge
        private static void TryCopy(Action copy)
        {
            try
            {
                copy();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Could not copy formatting: {ex.Message}");
            }
        }

        #endregion

        #region Selection helpers

        // Selected shapes in selection order; shapes picked inside a group are used individually
        private static Shape[] GetSelectedShapes()
        {
            var selection = App.ActiveWindow.Selection;
            if (selection.Type != PpSelectionType.ppSelectionShapes && selection.Type != PpSelectionType.ppSelectionText)
            {
                return new Shape[0];
            }

            var range = selection.HasChildShapeRange ? selection.ChildShapeRange : selection.ShapeRange;
            var shapes = new Shape[range.Count];
            for (int i = 0; i < range.Count; i++)
            {
                shapes[i] = range[i + 1];
            }
            return shapes;
        }

        // Selected shapes that can hold text, with whole groups expanded into their members
        private static List<Shape> GetSelectedTextShapes()
        {
            var result = new List<Shape>();
            foreach (var shape in ExpandGroups(GetSelectedShapes()))
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.HasTable != Office.MsoTriState.msoTrue)
                {
                    result.Add(shape);
                }
            }
            return result;
        }

        // Text frames of the selected text shapes plus table cells (selected cells, or all cells if none are)
        private static List<TextFrame> GetSelectedTextFrames()
        {
            var frames = new List<TextFrame>();
            foreach (var shape in ExpandGroups(GetSelectedShapes()))
            {
                if (shape.HasTable == Office.MsoTriState.msoTrue)
                {
                    var cells = GetTableCells(shape.Table);
                    var selectedCells = cells.Where(c => c.Selected).ToList();
                    foreach (var cell in selectedCells.Count > 0 ? selectedCells : cells)
                    {
                        frames.Add(cell.Shape.TextFrame);
                    }
                }
                else if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    frames.Add(shape.TextFrame);
                }
            }
            return frames;
        }

        private static List<Cell> GetTableCells(Table table)
        {
            var cells = new List<Cell>();
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                for (int column = 1; column <= table.Columns.Count; column++)
                {
                    cells.Add(table.Cell(row, column));
                }
            }
            return cells;
        }

        private static IEnumerable<Shape> ExpandGroups(IEnumerable<Shape> shapes)
        {
            foreach (var shape in shapes)
            {
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    foreach (Shape member in shape.GroupItems)
                    {
                        yield return member;
                    }
                }
                else
                {
                    yield return shape;
                }
            }
        }

        // Resizes without the aspect-ratio lock interfering; zero-size dimensions (straight lines) are left alone
        private static void SetSize(Shape shape, float? width, float? height)
        {
            var lockAspectRatio = shape.LockAspectRatio;
            shape.LockAspectRatio = Office.MsoTriState.msoFalse;
            if (width.HasValue && shape.Width > 0) shape.Width = width.Value;
            if (height.HasValue && shape.Height > 0) shape.Height = height.Value;
            shape.LockAspectRatio = lockAspectRatio;
        }

        private static void ShowInfo(string message)
        {
            MessageBox.Show(message, "Decksterity", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static void ShowError(string message)
        {
            MessageBox.Show(message, "Decksterity", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        #endregion
    }
}
