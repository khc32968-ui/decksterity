using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace decksterity
{
    // Separate window for managing the color pane's palette: import a theme, or add / change / reorder / remove colors
    // Changes apply when you click OK
    public class PaletteEditorForm : Form
    {
        private const int SwatchSize = 30;

        private readonly List<int> colors;
        private string source;
        private int selectedIndex = -1;

        private readonly ToolTip toolTip = new ToolTip();
        private readonly FlowLayoutPanel swatchPanel;
        private readonly Label sourceLabel;
        private readonly Button changeButton;
        private readonly Button removeButton;
        private readonly Button moveLeftButton;
        private readonly Button moveRightButton;

        public static void ShowEditor()
        {
            using (var form = new PaletteEditorForm())
            {
                form.ShowDialog(new PowerPointWindow());
            }
        }

        // Replaces the palette with the active presentation's theme without opening the editor
        public static void ImportFromActivePresentation()
        {
            var app = Globals.ThisAddIn.Application;
            if (app.Presentations.Count == 0) return;
            var presentation = app.ActivePresentation;
            PaletteStore.Save(PaletteStore.ReadThemeColors(presentation), "Theme of " + presentation.Name);
        }

        public PaletteEditorForm()
        {
            colors = PaletteStore.Colors.ToList();
            source = PaletteStore.Source;

            Text = "Decksterity palette";
            Font = new Font("Segoe UI", 9f);
            AutoScaleMode = AutoScaleMode.Font;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(10);

            var layout = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true };
            Controls.Add(layout);

            layout.Controls.Add(new Label
            {
                Text = "These colors appear in the color pane. Click a color to select it.",
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 6)
            });

            swatchPanel = new FlowLayoutPanel
            {
                Size = new Size(12 * (SwatchSize + 4) + 8, 2 * (SwatchSize + 4) + 10),
                BorderStyle = BorderStyle.FixedSingle,
                AutoScroll = true,
                WrapContents = true,
                Padding = new Padding(2)
            };
            layout.Controls.Add(swatchPanel);

            sourceLabel = new Label { AutoSize = true, ForeColor = SystemColors.GrayText, Margin = new Padding(0, 4, 0, 8) };
            layout.Controls.Add(sourceLabel);

            var editRow = Row(layout);
            editRow.Controls.Add(NewButton("Add...", (s, e) => AddColor()));
            changeButton = NewButton("Change...", (s, e) => ChangeColor());
            removeButton = NewButton("Remove", (s, e) => RemoveColor());
            moveLeftButton = NewButton("Move left", (s, e) => MoveColor(-1));
            moveRightButton = NewButton("Move right", (s, e) => MoveColor(1));
            editRow.Controls.AddRange(new Control[] { changeButton, removeButton, moveLeftButton, moveRightButton });
            var fromShape = NewButton("Add from selected shape", (s, e) => AddFromSelection());
            toolTip.SetToolTip(fromShape, "Adds the fill color (or, without a fill, the line color) of the first selected shape");
            editRow.Controls.Add(fromShape);

            var importRow = Row(layout);
            importRow.Controls.Add(NewButton("Use this presentation's theme", (s, e) => ImportActive()));
            importRow.Controls.Add(NewButton("Use theme from file...", (s, e) => ImportFile()));
            importRow.Controls.Add(NewButton("Office default", (s, e) => SetColors(PaletteStore.DefaultColors, "Office default theme")));

            var okCancel = Row(layout);
            okCancel.Anchor = AnchorStyles.Right;
            okCancel.Margin = new Padding(0, 12, 0, 0);
            var ok = NewButton("OK", (s, e) => Save());
            var cancel = NewButton("Cancel", null);
            cancel.DialogResult = DialogResult.Cancel;
            okCancel.Controls.AddRange(new Control[] { ok, cancel });
            AcceptButton = ok;
            CancelButton = cancel;

            Redraw();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) toolTip.Dispose();
            base.Dispose(disposing);
        }

        #region Actions

        private void AddColor()
        {
            int? picked = ColorPaneControl.PickColor(colors);
            if (!picked.HasValue) return;
            colors.Add(picked.Value);
            selectedIndex = colors.Count - 1;
            MarkEdited();
        }

        private void ChangeColor()
        {
            if (selectedIndex < 0) return;
            int? picked = ColorPaneControl.PickColor(colors, colors[selectedIndex]);
            if (!picked.HasValue) return;
            colors[selectedIndex] = picked.Value;
            MarkEdited();
        }

        private void RemoveColor()
        {
            if (selectedIndex < 0) return;
            colors.RemoveAt(selectedIndex);
            selectedIndex = Math.Min(selectedIndex, colors.Count - 1);
            MarkEdited();
        }

        private void MoveColor(int offset)
        {
            int target = selectedIndex + offset;
            if (selectedIndex < 0 || target < 0 || target >= colors.Count) return;
            int moved = colors[selectedIndex];
            colors[selectedIndex] = colors[target];
            colors[target] = moved;
            selectedIndex = target;
            MarkEdited();
        }

        private void AddFromSelection()
        {
            int? color = null;
            try
            {
                var shape = SelectionHelper.GetSelectedShapes().FirstOrDefault();
                if (shape != null)
                {
                    if (shape.Fill.Visible == Office.MsoTriState.msoTrue) color = ElementHelper.ConvertRgbToBgr(shape.Fill.ForeColor.RGB);
                    else if (shape.Line.Visible == Office.MsoTriState.msoTrue) color = ElementHelper.ConvertRgbToBgr(shape.Line.ForeColor.RGB);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Could not read the selected shape's color: {ex.Message}");
            }

            if (!color.HasValue)
            {
                MessageBox.Show(this, "Select a shape with a fill or line color first.", "Decksterity",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            colors.Add(color.Value);
            selectedIndex = colors.Count - 1;
            MarkEdited();
        }

        private void ImportActive()
        {
            var app = Globals.ThisAddIn.Application;
            if (app.Presentations.Count == 0)
            {
                MessageBox.Show(this, "Open a presentation first.", "Decksterity", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var presentation = app.ActivePresentation;
            SetColors(PaletteStore.ReadThemeColors(presentation), "Theme of " + presentation.Name);
        }

        private void ImportFile()
        {
            using (var dialog = new OpenFileDialog
            {
                Title = "Choose a presentation or template",
                Filter = "PowerPoint files (*.pptx;*.potx;*.pptm;*.potm;*.ppt;*.pot)|*.pptx;*.potx;*.pptm;*.potm;*.ppt;*.pot|All files (*.*)|*.*"
            })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;

                Cursor = Cursors.WaitCursor;
                try
                {
                    SetColors(PaletteStore.ReadThemeColors(dialog.FileName), "Theme of " + Path.GetFileName(dialog.FileName));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Couldn't read the theme from that file.\n\n" + ex.Message, "Decksterity",
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
        }

        private void SetColors(IEnumerable<int> newColors, string newSource)
        {
            colors.Clear();
            colors.AddRange(newColors);
            source = newSource;
            selectedIndex = -1;
            Redraw();
        }

        private void Save()
        {
            if (colors.Count == 0)
            {
                MessageBox.Show(this, "The palette needs at least one color.", "Decksterity", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            PaletteStore.Save(colors, source);
            DialogResult = DialogResult.OK;
            Close();
        }

        private void MarkEdited()
        {
            if (!source.EndsWith(" (edited)")) source += " (edited)";
            Redraw();
        }

        #endregion

        #region Drawing

        private void Redraw()
        {
            swatchPanel.SuspendLayout();
            foreach (Control old in swatchPanel.Controls.Cast<Control>().ToList())
            {
                old.Dispose();
            }
            swatchPanel.Controls.Clear();

            for (int i = 0; i < colors.Count; i++)
            {
                int index = i;
                int color = colors[i];
                var swatch = new Button
                {
                    Size = new Size(SwatchSize, SwatchSize),
                    Margin = new Padding(2),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = PaletteStore.ToDrawingColor(color),
                    Cursor = Cursors.Hand
                };
                bool selected = index == selectedIndex;
                swatch.FlatAppearance.BorderSize = selected ? 3 : 1;
                swatch.FlatAppearance.BorderColor = selected ? Color.Black : Color.FromArgb(160, 160, 160);
                swatch.Click += (s, e) =>
                {
                    selectedIndex = index;
                    Redraw();
                };
                toolTip.SetToolTip(swatch, PaletteStore.ToHex(color));
                swatchPanel.Controls.Add(swatch);
            }
            swatchPanel.ResumeLayout();

            sourceLabel.Text = $"{colors.Count} colors · {source}";
            bool hasSelection = selectedIndex >= 0;
            changeButton.Enabled = removeButton.Enabled = hasSelection;
            moveLeftButton.Enabled = hasSelection && selectedIndex > 0;
            moveRightButton.Enabled = hasSelection && selectedIndex < colors.Count - 1;
        }

        private static FlowLayoutPanel Row(FlowLayoutPanel parent)
        {
            var row = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 2, 0, 2) };
            parent.Controls.Add(row);
            return row;
        }

        private static Button NewButton(string text, EventHandler onClick)
        {
            var button = new Button { Text = text, AutoSize = true, Margin = new Padding(0, 0, 6, 0) };
            if (onClick != null) button.Click += onClick;
            return button;
        }

        #endregion

        // Lets dialogs use the PowerPoint window as their owner so they stay in front of it
        private class PowerPointWindow : IWin32Window
        {
            public IntPtr Handle => new IntPtr(Globals.ThisAddIn.Application.HWND);
        }
    }
}
