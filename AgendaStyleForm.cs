using System;
using System.Drawing;
using System.Windows.Forms;

namespace decksterity
{
    // Dialog for the agenda colors used by Ctrl+Shift+F: saved as defaults for new agendas,
    // and optionally applied to the agenda slides already in the active presentation
    public class AgendaStyleForm : Form
    {
        private readonly int[] colors;
        private readonly Button[] swatches = new Button[AgendaStyle.Labels.Length];
        private readonly CheckBox applyToPresentation;
        private readonly ToolTip toolTip = new ToolTip();

        public static void ShowEditor()
        {
            using (var form = new AgendaStyleForm())
            {
                form.ShowDialog(new PowerPointWindow());
            }
        }

        public AgendaStyleForm()
        {
            colors = (int[])AgendaStyle.Load().Colors.Clone();

            Text = "Decksterity agenda colors";
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
                Text = "Colors for new agenda slides (Ctrl+Shift+F). Updating an existing agenda keeps " +
                       "the colors of the agenda slide you run it from. Click a color to change it.",
                AutoSize = true,
                MaximumSize = new Size(380, 0),
                Margin = new Padding(0, 0, 0, 8)
            });

            var grid = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Margin = new Padding(0, 0, 0, 6) };
            layout.Controls.Add(grid);
            for (int i = 0; i < swatches.Length; i++)
            {
                int index = i;
                grid.Controls.Add(new Label
                {
                    Text = AgendaStyle.Labels[i],
                    AutoSize = true,
                    Anchor = AnchorStyles.Left,
                    Margin = new Padding(0, 4, 12, 4)
                });

                var swatch = new Button
                {
                    Size = new Size(64, 24),
                    FlatStyle = FlatStyle.Flat,
                    Cursor = Cursors.Hand,
                    Margin = new Padding(0, 2, 0, 2)
                };
                swatch.FlatAppearance.BorderColor = Color.FromArgb(160, 160, 160);
                swatch.Click += (s, e) => PickColor(index);
                swatches[i] = swatch;
                grid.Controls.Add(swatch);
            }

            bool hasAgenda = AgendaHelper.PresentationHasAgenda();
            applyToPresentation = new CheckBox
            {
                Text = "Also recolor the agenda slides in this presentation",
                AutoSize = true,
                Checked = hasAgenda,
                Enabled = hasAgenda,
                Margin = new Padding(0, 4, 0, 4)
            };
            layout.Controls.Add(applyToPresentation);

            var buttons = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 10, 0, 0) };
            layout.Controls.Add(buttons);
            buttons.Controls.Add(NewButton("Reset to defaults", (s, e) => ResetDefaults()));
            var ok = NewButton("OK", (s, e) => Save());
            var cancel = NewButton("Cancel", null);
            cancel.DialogResult = DialogResult.Cancel;
            buttons.Controls.AddRange(new Control[] { ok, cancel });
            AcceptButton = ok;
            CancelButton = cancel;

            Redraw();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) toolTip.Dispose();
            base.Dispose(disposing);
        }

        private void PickColor(int index)
        {
            int? picked = ColorPaneControl.PickColor(PaletteStore.Colors, colors[index]);
            if (!picked.HasValue) return;
            colors[index] = picked.Value;
            Redraw();
        }

        private void ResetDefaults()
        {
            Array.Copy(AgendaStyle.Defaults, colors, colors.Length);
            Redraw();
        }

        private void Save()
        {
            var style = new AgendaStyle(colors);
            style.Save();

            if (applyToPresentation.Checked)
            {
                try
                {
                    AgendaHelper.ApplyStyleToPresentation(style);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Saved the colors, but couldn't recolor this presentation's agenda slides.\n\n" + ex.Message,
                        "Decksterity", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        private void Redraw()
        {
            for (int i = 0; i < swatches.Length; i++)
            {
                swatches[i].BackColor = PaletteStore.ToDrawingColor(colors[i]);
                toolTip.SetToolTip(swatches[i], PaletteStore.ToHex(colors[i]));
            }
        }

        private static Button NewButton(string text, EventHandler onClick)
        {
            var button = new Button { Text = text, AutoSize = true, Margin = new Padding(0, 0, 6, 0) };
            if (onClick != null) button.Click += onClick;
            return button;
        }

        // Lets the dialog use the PowerPoint window as its owner so it stays in front of it
        private class PowerPointWindow : IWin32Window
        {
            public IntPtr Handle => new IntPtr(Globals.ThisAddIn.Application.HWND);
        }
    }
}
