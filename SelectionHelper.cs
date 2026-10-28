using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace decksterity
{
    // Shared helpers for reading the current selection
    internal static class SelectionHelper
    {
        // Selected shapes in selection order; shapes picked inside a group are used individually
        // While editing text, this is the shape that contains the text
        internal static Shape[] GetSelectedShapes()
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
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

        // Replaces each group with its member shapes
        internal static IEnumerable<Shape> ExpandGroups(IEnumerable<Shape> shapes)
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

        internal static bool IsTable(Shape shape) => shape.HasTable == Office.MsoTriState.msoTrue;

        // The table's selected cells, or every cell if none are selected (the whole table is selected)
        internal static List<Cell> GetTargetCells(Table table)
        {
            var cells = new List<Cell>();
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                for (int column = 1; column <= table.Columns.Count; column++)
                {
                    cells.Add(table.Cell(row, column));
                }
            }

            var selected = cells.Where(c => c.Selected).ToList();
            return selected.Count > 0 ? selected : cells;
        }
    }
}
