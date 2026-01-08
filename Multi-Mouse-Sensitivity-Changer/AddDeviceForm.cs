using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MultiMouseSensitivityChanger
{
    class AddDeviceForm : Form
    {
        readonly TextBox _nameTextBox;
        readonly TextBox _pathTextBox;
        readonly NumericUpDown _speedSelector;
        readonly NumericUpDown _scrollLinesSelector;
        readonly NumericUpDown _scrollCharsSelector;
        readonly CheckBox _enhancePrecisionCheckBox;
        readonly CheckBox _swapButtonsCheckBox;
        readonly CheckBox _applyOnStartupCheckBox;
        readonly CheckBox _autoApplyCheckBox;
        readonly Label _statusLabel;
        readonly Panel _colorPreview;
        Color _selectedColor = Color.Gray;
        RawInputCaptureWindow _captureWindow;

        public Program.DeviceProfile NewProfile { get; private set; }

        public AddDeviceForm(Program.DeviceProfile existing = null, IEnumerable<Color> existingColors = null)
        {
            Text = "Add new device";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(720, 560);
            MinimumSize = new Size(720, 560);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(12)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var identityGroup = new GroupBox { Text = "Device identity", Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
            var identityLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 3,
                Padding = new Padding(10),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            identityLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            identityLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));

            identityLayout.Controls.Add(new Label { Text = "Device path", AutoSize = true }, 0, 0);
            _pathTextBox = new TextBox { Dock = DockStyle.Fill };
            identityLayout.Controls.Add(_pathTextBox, 1, 0);
            var capturePanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, AutoSize = true, WrapContents = false };
            var captureButton = new Button { Text = "Capture from movement", AutoSize = true };
            captureButton.Click += (_, __) => BeginCapture();
            capturePanel.Controls.Add(captureButton);
            capturePanel.Controls.Add(new Label { Text = "Move the device after clicking.", AutoSize = true, ForeColor = Color.DimGray, Padding = new Padding(6, 6, 0, 0) });
            identityLayout.Controls.Add(capturePanel, 1, 1);

            identityLayout.Controls.Add(new Label { Text = "Name", AutoSize = true }, 0, 2);
            _nameTextBox = new TextBox { Dock = DockStyle.Fill };
            identityLayout.Controls.Add(_nameTextBox, 1, 2);
            identityGroup.Controls.Add(identityLayout);

            var pointerGroup = new GroupBox { Text = "Pointer settings", Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
            var pointerLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 3,
                Padding = new Padding(10),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            pointerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            pointerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));

            pointerLayout.Controls.Add(new Label { Text = "Speed", AutoSize = true }, 0, 0);
            _speedSelector = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 20,
                Value = 10,
                Dock = DockStyle.Left,
                Width = 80
            };
            var speedPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, AutoSize = true, WrapContents = false };
            speedPanel.Controls.Add(_speedSelector);
            var testButton = new Button { Text = "Apply now", AutoSize = true, Margin = new Padding(12, 0, 0, 0) };
            testButton.Click += (_, __) => TestSettings();
            speedPanel.Controls.Add(testButton);
            pointerLayout.Controls.Add(speedPanel, 1, 0);

            _enhancePrecisionCheckBox = new CheckBox { Text = "Enhance pointer precision", AutoSize = true };
            pointerLayout.Controls.Add(_enhancePrecisionCheckBox, 1, 1);

            _swapButtonsCheckBox = new CheckBox { Text = "Swap primary and secondary buttons", AutoSize = true };
            pointerLayout.Controls.Add(_swapButtonsCheckBox, 1, 2);
            pointerGroup.Controls.Add(pointerLayout);

            var scrollGroup = new GroupBox { Text = "Scrolling", Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
            var scrollLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                Padding = new Padding(10),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            scrollLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            scrollLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));

            scrollLayout.Controls.Add(new Label { Text = "Vertical wheel lines", AutoSize = true }, 0, 0);
            _scrollLinesSelector = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 100,
                Value = 3,
                Dock = DockStyle.Left,
                Width = 80
            };
            scrollLayout.Controls.Add(_scrollLinesSelector, 1, 0);

            scrollLayout.Controls.Add(new Label { Text = "Horizontal wheel chars", AutoSize = true }, 0, 1);
            _scrollCharsSelector = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 100,
                Value = 3,
                Dock = DockStyle.Left,
                Width = 80
            };
            scrollLayout.Controls.Add(_scrollCharsSelector, 1, 1);
            scrollGroup.Controls.Add(scrollLayout);

            var behaviorGroup = new GroupBox { Text = "Behavior & icon", Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
            var behaviorLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 3,
                Padding = new Padding(10),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            behaviorLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            behaviorLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));

            behaviorLayout.Controls.Add(new Label { Text = "Icon color", AutoSize = true }, 0, 0);
            var colorButton = new Button { Text = "Choose color", AutoSize = true };
            colorButton.Click += (_, __) => ChooseColor();
            var colorPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Fill, WrapContents = false, AutoSize = true };
            _colorPreview = new Panel { Width = 32, Height = 32, BackColor = _selectedColor, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(8, 0, 0, 0) };
            colorPanel.Controls.Add(colorButton);
            colorPanel.Controls.Add(_colorPreview);
            behaviorLayout.Controls.Add(colorPanel, 1, 0);

            _applyOnStartupCheckBox = new CheckBox { Text = "Apply these settings on app startup", AutoSize = true };
            behaviorLayout.Controls.Add(_applyOnStartupCheckBox, 1, 1);

            _autoApplyCheckBox = new CheckBox { Text = "Apply automatically when this device is used", AutoSize = true };
            behaviorLayout.Controls.Add(_autoApplyCheckBox, 1, 2);
            behaviorGroup.Controls.Add(behaviorLayout);

            _statusLabel = new Label { Text = "Move your device to capture its path.", AutoSize = true, ForeColor = Color.DimGray };

            layout.Controls.Add(identityGroup);
            layout.Controls.Add(pointerGroup);
            layout.Controls.Add(scrollGroup);
            layout.Controls.Add(behaviorGroup);
            layout.Controls.Add(_statusLabel);

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Bottom,
                Padding = new Padding(12, 8, 12, 12),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            var okButton = new Button { Text = "Save", DialogResult = DialogResult.OK, AutoSize = true };
            var cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
            buttons.Controls.Add(okButton);
            buttons.Controls.Add(cancelButton);

            Controls.Add(layout);
            Controls.Add(buttons);

            AcceptButton = okButton;
            CancelButton = cancelButton;

            okButton.Click += (_, __) => SaveDevice();

            if (existing != null)
            {
                _nameTextBox.Text = existing.Name;
                _pathTextBox.Text = existing.DevicePath;
                _speedSelector.Value = existing.Speed;
                _selectedColor = existing.IconColor;
                _enhancePrecisionCheckBox.Checked = existing.EnhancePointerPrecision;
                _scrollLinesSelector.Value = existing.VerticalScrollLines;
                _scrollCharsSelector.Value = existing.HorizontalScrollChars;
                _swapButtonsCheckBox.Checked = existing.SwapButtons;
                _applyOnStartupCheckBox.Checked = existing.ApplyOnStartup;
                _autoApplyCheckBox.Checked = existing.AutoApply;
                _colorPreview.BackColor = _selectedColor;
                Text = "Edit device";
            }
            else
            {
                var defaults = Program.GetSystemMouseSettings();
                _speedSelector.Value = defaults.Speed;
                _enhancePrecisionCheckBox.Checked = defaults.EnhancePointerPrecision;
                _scrollLinesSelector.Value = defaults.VerticalScrollLines;
                _scrollCharsSelector.Value = defaults.HorizontalScrollChars;
                _swapButtonsCheckBox.Checked = defaults.SwapButtons;
                _applyOnStartupCheckBox.Checked = false;
                _autoApplyCheckBox.Checked = true;
                _selectedColor = GetDefaultColor(existingColors);
                _colorPreview.BackColor = _selectedColor;
            }
        }

        void BeginCapture()
        {
            _captureWindow?.Dispose();
            _captureWindow = new RawInputCaptureWindow(OnDeviceCaptured);
            _statusLabel.Text = "Listening for mouse movement... move the new device now.";
        }

        void OnDeviceCaptured(string devicePath)
        {
            if (string.IsNullOrWhiteSpace(devicePath))
                return;

            _pathTextBox.Text = devicePath;
            _statusLabel.Text = "Captured device path. You can rename it and choose a default speed.";
            _captureWindow?.Dispose();
            _captureWindow = null;
            Program.EnsureRawInputRegistration();
        }

        void TestSettings()
        {
            var preview = BuildProfileFromForm();
            Program.ApplyProfileSettings(preview, applySpeed: true, forceSpeed: true);
            _statusLabel.Text = "Applied these settings for preview.";
        }

        void ChooseColor()
        {
            using (var dialog = new ColorDialog { Color = _selectedColor })
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    _selectedColor = dialog.Color;
                    _colorPreview.BackColor = _selectedColor;
                }
            }
        }

        void SaveDevice()
        {
            if (string.IsNullOrWhiteSpace(_pathTextBox.Text))
            {
                MessageBox.Show(this, "Please capture or enter a device path before saving.", "Missing device path", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            if (string.IsNullOrWhiteSpace(_nameTextBox.Text))
            {
                MessageBox.Show(this, "Please provide a name for the device.", "Missing name", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            var profile = BuildProfileFromForm();
            profile.Name = _nameTextBox.Text.Trim();
            profile.DevicePath = _pathTextBox.Text.Trim();
            NewProfile = profile;
        }

        Program.DeviceProfile BuildProfileFromForm()
        {
            return new Program.DeviceProfile(
                _nameTextBox.Text.Trim(),
                _pathTextBox.Text.Trim(),
                (int)_speedSelector.Value,
                _selectedColor,
                _enhancePrecisionCheckBox.Checked,
                (int)_scrollLinesSelector.Value,
                (int)_scrollCharsSelector.Value,
                _swapButtonsCheckBox.Checked,
                _applyOnStartupCheckBox.Checked,
                _autoApplyCheckBox.Checked);
        }

        Color GetDefaultColor(IEnumerable<Color> existingColors)
        {
            var palette = new List<Color>
            {
                Color.SteelBlue,
                Color.ForestGreen,
                Color.OrangeRed,
                Color.SlateBlue,
                Color.DarkCyan,
                Color.Goldenrod,
                Color.Crimson,
                Color.Teal
            };

            var existingSet = new HashSet<Color>(existingColors ?? Enumerable.Empty<Color>());
            var unused = palette.FirstOrDefault(c => !existingSet.Contains(c));
            return unused == default ? palette[0] : unused;
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _captureWindow?.Dispose();
            Program.EnsureRawInputRegistration();
            base.OnFormClosed(e);
        }

        class RawInputCaptureWindow : NativeWindow, IDisposable
        {
            readonly Action<string> _callback;

            public RawInputCaptureWindow(Action<string> callback)
            {
                _callback = callback;
                CreateHandle(new CreateParams());

                Program.RAWINPUTDEVICE[] rid = new[]
                {
                    new Program.RAWINPUTDEVICE
                    {
                        usUsagePage = 0x01,
                        usUsage = 0x02,
                        dwFlags = Program.NativeMethods.RIDEV_INPUTSINK,
                        hwndTarget = Handle
                    }
                };

                Program.NativeMethods.RegisterRawInputDevices(rid, (uint)rid.Length, (uint)Marshal.SizeOf(typeof(Program.RAWINPUTDEVICE)));
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == Program.NativeMethods.WM_INPUT)
                {
                    string devicePath = GetDeviceNameFromMessage(m.LParam);
                    if (!string.IsNullOrEmpty(devicePath))
                        _callback(devicePath);
                }

                base.WndProc(ref m);
            }

            string GetDeviceNameFromMessage(IntPtr lParam)
            {
                uint size = 0;
                Program.NativeMethods.GetRawInputData(lParam, Program.NativeMethods.RID_INPUT, IntPtr.Zero, ref size, (uint)Marshal.SizeOf(typeof(Program.RAWINPUTHEADER)));
                if (size == 0)
                    return string.Empty;

                IntPtr buffer = Marshal.AllocHGlobal((int)size);
                try
                {
                    if (Program.NativeMethods.GetRawInputData(lParam, Program.NativeMethods.RID_INPUT, buffer, ref size, (uint)Marshal.SizeOf(typeof(Program.RAWINPUTHEADER))) != size)
                        return string.Empty;

                    Program.RAWINPUT raw = (Program.RAWINPUT)Marshal.PtrToStructure(buffer, typeof(Program.RAWINPUT));
                    if (raw.header.dwType != Program.NativeMethods.RIM_TYPEMOUSE)
                        return string.Empty;

                    if (raw.data.lLastX == 0 && raw.data.lLastY == 0)
                        return string.Empty;

                    return GetDeviceName(raw.header.hDevice);
                }
                finally
                {
                    Marshal.FreeHGlobal(buffer);
                }
            }

            string GetDeviceName(IntPtr hDevice)
            {
                uint size = 0;
                Program.NativeMethods.GetRawInputDeviceInfo(hDevice, Program.NativeMethods.RIDI_DEVICENAME, IntPtr.Zero, ref size);
                if (size == 0)
                    return string.Empty;

                IntPtr data = Marshal.AllocHGlobal((int)size);
                try
                {
                    if (Program.NativeMethods.GetRawInputDeviceInfo(hDevice, Program.NativeMethods.RIDI_DEVICENAME, data, ref size) == uint.MaxValue)
                        return string.Empty;

                    return Marshal.PtrToStringAnsi(data);
                }
                finally
                {
                    Marshal.FreeHGlobal(data);
                }
            }

            public void Dispose()
            {
                DestroyHandle();
            }
        }
    }
}
