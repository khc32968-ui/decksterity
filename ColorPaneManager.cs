using System;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using Office = Microsoft.Office.Core;

namespace decksterity
{
    // Shows / hides the color pane. PowerPoint attaches task panes to a single window,
    // so each presentation window gets its own pane, created the first time it's toggled there
    internal static class ColorPaneManager
    {
        private const string Title = "Decksterity colors";
        private const int DefaultWidth = 270;

        public static void Toggle()
        {
            var app = Globals.ThisAddIn.Application;
            if (app.Windows.Count == 0) return; // no presentation open

            var window = app.ActiveWindow;
            var pane = FindPane(window);
            if (pane == null)
            {
                pane = Globals.ThisAddIn.CustomTaskPanes.Add(new ColorPaneControl(), Title, window);
                pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                pane.Width = DefaultWidth;
                pane.Visible = true;
            }
            else
            {
                pane.Visible = !pane.Visible;
            }

            if (pane.Visible)
            {
                ((ColorPaneControl)pane.Control).SyncToSelection();
            }
        }

        // Keeps the line-weight slider of visible panes in step with what's selected
        public static void OnSelectionChanged()
        {
            foreach (var pane in Globals.ThisAddIn.CustomTaskPanes.ToList())
            {
                if (pane.Control is ColorPaneControl control && pane.Visible)
                {
                    control.SyncToSelection();
                }
            }
        }

        private static CustomTaskPane FindPane(DocumentWindow window)
        {
            RemoveClosedPanes();
            return Globals.ThisAddIn.CustomTaskPanes
                .FirstOrDefault(pane => pane.Control is ColorPaneControl && SameWindow(pane.Window, window));
        }

        private static bool SameWindow(object paneWindow, DocumentWindow window)
        {
            if (ReferenceEquals(paneWindow, window)) return true;
            try
            {
                return paneWindow is DocumentWindow other
                    && other.Caption == window.Caption
                    && other.Presentation.FullName == window.Presentation.FullName;
            }
            catch
            {
                return false;
            }
        }

        // A pane whose window was closed throws when its window is touched; drop it
        private static void RemoveClosedPanes()
        {
            foreach (var pane in Globals.ThisAddIn.CustomTaskPanes.ToList())
            {
                if (!(pane.Control is ColorPaneControl)) continue;
                try
                {
                    var caption = ((DocumentWindow)pane.Window).Caption;
                }
                catch (Exception)
                {
                    Globals.ThisAddIn.CustomTaskPanes.Remove(pane);
                }
            }
        }
    }
}
