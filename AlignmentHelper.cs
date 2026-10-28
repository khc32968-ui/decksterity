using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Office = Microsoft.Office.Core;

namespace decksterity
{
    // Comprehensive alignment, distribution, and spacing utilities for PowerPoint shapes
    // Handles both built-in Office alignment and advanced custom spacing algorithms
    public static class AlignmentHelper
    {
        // Gets active PowerPoint application instance via COM interop
        private static Application GetPowerPointApplication()
        {
            try
            {
                return (Application)Marshal.GetActiveObject("PowerPoint.Application");
            }
            catch
            {
                throw new InvalidOperationException("PowerPoint is not running or accessible.");
            }
        }

        // Gets current active selection from PowerPoint window
        private static Selection GetActiveSelection()
        {
            var app = GetPowerPointApplication();
            if (app.ActiveWindow?.Selection != null)
            {
                return app.ActiveWindow.Selection;
            }
            throw new InvalidOperationException("No active selection in PowerPoint.");
        }

        // Aligns selected shapes to left edge (multiple shapes) or slide edge (single shape)
        public static void AlignLeft()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                selection.ShapeRange.Align(Office.MsoAlignCmd.msoAlignLefts, Office.MsoTriState.msoFalse);
            }
            else if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 1)
            {
                selection.ShapeRange.Align(Office.MsoAlignCmd.msoAlignLefts, Office.MsoTriState.msoTrue);
            }
        }

        // Centers selected shapes horizontally (multiple shapes) or to slide center (single shape)
        public static void AlignCenter()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                selection.ShapeRange.Align(Office.MsoAlignCmd.msoAlignCenters, Office.MsoTriState.msoFalse);
            }
            else if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 1)
            {
                selection.ShapeRange.Align(Office.MsoAlignCmd.msoAlignCenters, Office.MsoTriState.msoTrue);
            }
        }

        // Aligns selected shapes to right edge (multiple shapes) or slide edge (single shape)
        public static void AlignRight()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                selection.ShapeRange.Align(Office.MsoAlignCmd.msoAlignRights, Office.MsoTriState.msoFalse);
            }
            else if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 1)
            {
                selection.ShapeRange.Align(Office.MsoAlignCmd.msoAlignRights, Office.MsoTriState.msoTrue);
            }
        }

        // Aligns selected shapes to top edge (multiple shapes) or slide edge (single shape)
        public static void AlignTop()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                selection.ShapeRange.Align(Office.MsoAlignCmd.msoAlignTops, Office.MsoTriState.msoFalse);
            }
            else if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 1)
            {
                selection.ShapeRange.Align(Office.MsoAlignCmd.msoAlignTops, Office.MsoTriState.msoTrue);
            }
        }

        // Centers selected shapes vertically (multiple shapes) or to slide center (single shape)
        public static void AlignMiddle()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                selection.ShapeRange.Align(Office.MsoAlignCmd.msoAlignMiddles, Office.MsoTriState.msoFalse);
            }
            else if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 1)
            {
                selection.ShapeRange.Align(Office.MsoAlignCmd.msoAlignMiddles, Office.MsoTriState.msoTrue);
            }
        }

        // Aligns selected shapes to bottom edge (multiple shapes) or slide edge (single shape)
        public static void AlignBottom()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                selection.ShapeRange.Align(Office.MsoAlignCmd.msoAlignBottoms, Office.MsoTriState.msoFalse);
            }
            else if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 1)
            {
                selection.ShapeRange.Align(Office.MsoAlignCmd.msoAlignBottoms, Office.MsoTriState.msoTrue);
            }
        }

        // Distributes 3+ shapes evenly across horizontal space, centers single shape
        public static void DistributeHorizontally()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 2)
            {
                selection.ShapeRange.Distribute(Office.MsoDistributeCmd.msoDistributeHorizontally, Office.MsoTriState.msoFalse);
            }
            else if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 1)
            {
                selection.ShapeRange.Align(Office.MsoAlignCmd.msoAlignCenters, Office.MsoTriState.msoTrue);
            }
        }

        // Distributes 3+ shapes evenly across vertical space, centers single shape
        public static void DistributeVertically()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 2)
            {
                selection.ShapeRange.Distribute(Office.MsoDistributeCmd.msoDistributeVertically, Office.MsoTriState.msoFalse);
            }
            else if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 1)
            {
                selection.ShapeRange.Align(Office.MsoAlignCmd.msoAlignMiddles, Office.MsoTriState.msoTrue);
            }
        }

        // Makes all selected shapes the same width as the first selected shape
        public static void SameWidth()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                var referenceWidth = selection.ShapeRange[1].Width;
                for (int i = 2; i <= selection.ShapeRange.Count; i++)
                {
                    selection.ShapeRange[i].Width = referenceWidth;
                }
            }
        }

        // Makes all selected shapes the same height as the first selected shape
        public static void SameHeight()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                var referenceHeight = selection.ShapeRange[1].Height;
                for (int i = 2; i <= selection.ShapeRange.Count; i++)
                {
                    selection.ShapeRange[i].Height = referenceHeight;
                }
            }
        }

        // Aligns all shapes to the left edge of the first selected shape
        public static void PrimaryAlignLeft()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                var referenceLeft = selection.ShapeRange[1].Left;
                for (int i = 2; i <= selection.ShapeRange.Count; i++)
                {
                    selection.ShapeRange[i].Left = referenceLeft;
                }
            }
        }

        // Centers all shapes horizontally on the first selected shape
        public static void PrimaryAlignCenter()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                var referenceCenter = selection.ShapeRange[1].Left + (selection.ShapeRange[1].Width / 2);
                for (int i = 2; i <= selection.ShapeRange.Count; i++)
                {
                    selection.ShapeRange[i].Left = referenceCenter - (selection.ShapeRange[i].Width / 2);
                }
            }
        }

        // Aligns all shapes to the right edge of the first selected shape
        public static void PrimaryAlignRight()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                var referenceRight = selection.ShapeRange[1].Left + selection.ShapeRange[1].Width;
                for (int i = 2; i <= selection.ShapeRange.Count; i++)
                {
                    selection.ShapeRange[i].Left = referenceRight - selection.ShapeRange[i].Width;
                }
            }
        }

        // Aligns all shapes to the top edge of the first selected shape
        public static void PrimaryAlignTop()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                var referenceTop = selection.ShapeRange[1].Top;
                for (int i = 2; i <= selection.ShapeRange.Count; i++)
                {
                    selection.ShapeRange[i].Top = referenceTop;
                }
            }
        }

        // Centers all shapes vertically on the first selected shape
        public static void PrimaryAlignMiddle()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                var referenceMiddle = selection.ShapeRange[1].Top + (selection.ShapeRange[1].Height / 2);
                for (int i = 2; i <= selection.ShapeRange.Count; i++)
                {
                    selection.ShapeRange[i].Top = referenceMiddle - (selection.ShapeRange[i].Height / 2);
                }
            }
        }

        // Aligns all shapes to the bottom edge of the first selected shape
        public static void PrimaryAlignBottom()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                var referenceBottom = selection.ShapeRange[1].Top + selection.ShapeRange[1].Height;
                for (int i = 2; i <= selection.ShapeRange.Count; i++)
                {
                    selection.ShapeRange[i].Top = referenceBottom - selection.ShapeRange[i].Height;
                }
            }
        }

        // Swaps the positions of exactly two selected shapes
        public static void ObjectsSwapPositionCentered()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 2)
            {
                var shape1 = selection.ShapeRange[1];
                var shape2 = selection.ShapeRange[2];

                var tempLeft = shape1.Left;
                var tempTop = shape1.Top;

                shape1.Left = shape2.Left;
                shape1.Top = shape2.Top;
                shape2.Left = tempLeft;
                shape2.Top = tempTop;
            }
        }

        // Brings selected shapes to the front of the z-order
        public static void BringToFront()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                selection.ShapeRange.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            }
        }

        // Sends selected shapes to the back of the z-order
        public static void SendToBack()
        {
            var selection = GetActiveSelection();
            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                selection.ShapeRange.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
            }
        }

        // Resize increments in points: 0.05" for Ctrl+Arrow, 0.01" (PowerPoint's displayed precision) with Shift
        public const float ResizeStep = 3.6f;
        public const float FineResizeStep = 0.72f;

        // Grows or shrinks each selected shape by the given deltas, keeping its top-left corner fixed
        // Returns false when no shapes are selected so the keystroke can fall through to PowerPoint
        public static bool ResizeSelection(float deltaWidth, float deltaHeight)
        {
            var selection = GetActiveSelection();
            if (selection.Type != PpSelectionType.ppSelectionShapes) return false;

            var shapeRange = selection.HasChildShapeRange ? selection.ChildShapeRange : selection.ShapeRange;
            foreach (Shape shape in shapeRange)
            {
                var lockAspectRatio = shape.LockAspectRatio;
                shape.LockAspectRatio = Office.MsoTriState.msoFalse;

                // Zero-size dimensions (e.g. straight lines) are left alone so lines don't turn diagonal
                if (deltaWidth != 0 && shape.Width > 0)
                {
                    shape.Width = Math.Max(FineResizeStep, shape.Width + deltaWidth);
                }
                if (deltaHeight != 0 && shape.Height > 0)
                {
                    shape.Height = Math.Max(FineResizeStep, shape.Height + deltaHeight);
                }

                shape.LockAspectRatio = lockAspectRatio;
            }
            return true;
        }

        // Advanced resize and spacing with multiple algorithms and user input
        // Supports even spacing, preserve first/last, horizontal/vertical modes
        public static void ResizeAndSpaceEvenly(string spacingType)
        {
            var selection = GetActiveSelection();
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            ShapeRange shapeRange;
            if (selection.HasChildShapeRange)
            {
                shapeRange = selection.ChildShapeRange;
            }
            else
            {
                shapeRange = selection.ShapeRange;
            }

            if (shapeRange.Count < 2) return;

            // Step 1: Convert ShapeRange to array for sorting
            var shapes = new Shape[shapeRange.Count];
            for (int i = 0; i < shapeRange.Count; i++)
            {
                shapes[i] = shapeRange[i + 1];
            }

            // Step 2: Get user input for spacing value
            string userInput = Interaction.InputBox(
                "Enter desired spacing between objects (in points):",
                "Spacing Input",
                "30");

            if (string.IsNullOrEmpty(userInput) || !float.TryParse(userInput, out float spaceValue))
            {
                MessageBox.Show("Invalid input. Please enter a numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            switch (spacingType)
            {
                case "evenhorizontal":
                    ProcessEvenHorizontal(shapes, spaceValue);
                    break;
                case "evenhorizontalpreservefirst":
                    ProcessEvenHorizontalPreserveFirst(shapes, spaceValue);
                    break;
                case "evenhorizontalpreservelast":
                    ProcessEvenHorizontalPreserveLast(shapes, spaceValue);
                    break;
                case "evenvertical":
                    ProcessEvenVertical(shapes, spaceValue);
                    break;
                case "evenverticalpreservefirst":
                    ProcessEvenVerticalPreserveFirst(shapes, spaceValue);
                    break;
                case "evenverticalpreservelast":
                    ProcessEvenVerticalPreserveLast(shapes, spaceValue);
                    break;
            }
        }

        // Even horizontal spacing - all shapes get equal size and spacing
        private static void ProcessEvenHorizontal(Shape[] shapes, float spaceWidth)
        {
            SortShapesByLeftPosition(shapes);
            
            float totalWidth = shapes.Last().Left + shapes.Last().Width - shapes.First().Left;
            float shapeSize = (totalWidth - (shapes.Length - 1) * spaceWidth) / shapes.Length;
            
            shapes[0].Width = shapeSize;
            
            for (int i = 1; i < shapes.Length; i++)
            {
                shapes[i].Width = shapeSize;
                shapes[i].Left = shapes[0].Left + (shapeSize + spaceWidth) * i;
            }
        }

        // Preserve first shape size, adjust others proportionally
        private static void ProcessEvenHorizontalPreserveFirst(Shape[] shapes, float spaceWidth)
        {
            SortShapesByLeftPosition(shapes);
            
            float totalWidth = shapes.Last().Left + shapes.Last().Width - shapes.First().Left;
            float totalShapeWidth = 0;
            
            for (int i = 1; i < shapes.Length; i++)
            {
                totalShapeWidth += shapes[i].Width;
            }
            
            float shapeSizeIncrease = ((totalWidth - (shapes.Length - 1) * spaceWidth) - shapes[0].Width - totalShapeWidth) / (shapes.Length - 1);
            
            for (int i = 1; i < shapes.Length; i++)
            {
                shapes[i].Width += shapeSizeIncrease;
                shapes[i].Left = shapes[i - 1].Left + shapes[i - 1].Width + spaceWidth;
            }
        }

        // Preserve last shape size, adjust others proportionally
        private static void ProcessEvenHorizontalPreserveLast(Shape[] shapes, float spaceWidth)
        {
            SortShapesByLeftPosition(shapes);
            
            float totalWidth = shapes.Last().Left + shapes.Last().Width - shapes.First().Left;
            float totalShapeWidth = 0;
            
            for (int i = 0; i < shapes.Length - 1; i++)
            {
                totalShapeWidth += shapes[i].Width;
            }
            
            float shapeSizeIncrease = ((totalWidth - (shapes.Length - 1) * spaceWidth) - shapes.Last().Width - totalShapeWidth) / (shapes.Length - 1);
            
            shapes[0].Width += shapeSizeIncrease;
            
            for (int i = 1; i < shapes.Length - 1; i++)
            {
                shapes[i].Width += shapeSizeIncrease;
                shapes[i].Left = shapes[i - 1].Left + shapes[i - 1].Width + spaceWidth;
            }
        }

        // Even vertical spacing - all shapes get equal size and spacing
        private static void ProcessEvenVertical(Shape[] shapes, float spaceHeight)
        {
            SortShapesByTopPosition(shapes);
            
            float totalHeight = shapes.Last().Top + shapes.Last().Height - shapes.First().Top;
            float shapeSize = (totalHeight - (shapes.Length - 1) * spaceHeight) / shapes.Length;
            
            shapes[0].Height = shapeSize;
            
            for (int i = 1; i < shapes.Length; i++)
            {
                shapes[i].Height = shapeSize;
                shapes[i].Top = shapes[0].Top + (shapeSize + spaceHeight) * i;
            }
        }

        // Preserve first shape size vertically, adjust others proportionally
        private static void ProcessEvenVerticalPreserveFirst(Shape[] shapes, float spaceHeight)
        {
            SortShapesByTopPosition(shapes);
            
            float totalHeight = shapes.Last().Top + shapes.Last().Height - shapes.First().Top;
            float totalShapeHeight = 0;
            
            for (int i = 1; i < shapes.Length; i++)
            {
                totalShapeHeight += shapes[i].Height;
            }
            
            float shapeSizeIncrease = ((totalHeight - (shapes.Length - 1) * spaceHeight) - shapes[0].Height - totalShapeHeight) / (shapes.Length - 1);
            
            for (int i = 1; i < shapes.Length; i++)
            {
                shapes[i].Height += shapeSizeIncrease;
                shapes[i].Top = shapes[i - 1].Top + shapes[i - 1].Height + spaceHeight;
            }
        }

        // Preserve last shape size vertically, adjust others proportionally
        private static void ProcessEvenVerticalPreserveLast(Shape[] shapes, float spaceHeight)
        {
            SortShapesByTopPosition(shapes);
            
            float totalHeight = shapes.Last().Top + shapes.Last().Height - shapes.First().Top;
            float totalShapeHeight = 0;
            
            for (int i = 0; i < shapes.Length - 1; i++)
            {
                totalShapeHeight += shapes[i].Height;
            }
            
            float shapeSizeIncrease = ((totalHeight - (shapes.Length - 1) * spaceHeight) - shapes.Last().Height - totalShapeHeight) / (shapes.Length - 1);
            
            shapes[0].Height += shapeSizeIncrease;
            
            for (int i = 1; i < shapes.Length - 1; i++)
            {
                shapes[i].Height += shapeSizeIncrease;
                shapes[i].Top = shapes[i - 1].Top + shapes[i - 1].Height + spaceHeight;
            }
        }

        // Sorts shape array by left position for horizontal operations
        private static void SortShapesByLeftPosition(Shape[] shapes)
        {
            Array.Sort(shapes, (x, y) => x.Left.CompareTo(y.Left));
        }

        // Sorts shape array by top position for vertical operations
        private static void SortShapesByTopPosition(Shape[] shapes)
        {
            Array.Sort(shapes, (x, y) => x.Top.CompareTo(y.Top));
        }
    }
}