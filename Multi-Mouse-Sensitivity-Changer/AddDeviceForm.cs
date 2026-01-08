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
        readonly NumericUpDown _verticalScrollSelector;
        readonly NumericUpDown _horizontalScrollSelector;
        readonly CheckBox _enhancePrecisionCheckBox;
        readonly CheckBox _swapButtonsCheckBox;
        readonly CheckBox _enabledCheckBox;
        readonly CheckBox _autoApplyCheckBox;
        readonly CheckBox _startupApplyCheckBox;
        readonly Label _statusLabel;
        readonly Panel _colorPreview;
        Color _selectedColor = Color.Gray;
        RawInputCaptureWindow _captureWindow;

        public Program.DeviceProfile NewProfile { get; private set; }

        public AddDeviceForm(Program.DeviceProfile existing = null, IEnumerable<Color> existingColors = null)
        {
            Text = existing == null ? "Add new device" : "Edit device";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            MinimumSize = new Size(560, 0);

            var mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(12),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            for (int i = 0; i < mainLayout.RowCount; i++)
            {
                mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            var deviceGroup = BuildDeviceGroup();
            var pointerGroup = BuildPointerGroup();
            var scrollGroup = BuildScrollGroup();
            var activationGroup = BuildActivationGroup();

            mainLayout.Controls.Add(deviceGroup, 0, 0);
            mainLayout.Controls.Add(pointerGroup, 0, 1);
            mainLayout.Controls.Add(scrollGroup, 0, 2);
            mainLayout.Controls.Add(activationGroup, 0, 3);

            _statusLabel = new Label { Text = "Move your device to capture its path.", AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(0, 8, 0, 0) };
            mainLayout.Controls.Add(_statusLabel, 0, 4);

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 8, 0, 0),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            var okButton = new Button { Text = "Save", DialogResult = DialogResult.OK, AutoSize = true };
            var cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
            buttons.Controls.Add(okButton);
            buttons.Controls.Add(cancelButton);
            mainLayout.Controls.Add(buttons, 0, 5);

            Controls.Add(mainLayout);

            AcceptButton = okButton;
            CancelButton = cancelButton;

            okButton.Click += (_, __) => SaveDevice();

            if (existing != null)
            {
                _nameTextBox.Text = existing.Name;
                _pathTextBox.Text = existing.DevicePath;
                _speedSelector.Value = existing.Speed;
                _selectedColor = existing.IconColor;
                _colorPreview.BackColor = _selectedColor;
                _enhancePrecisionCheckBox.Checked = existing.EnhancePointerPrecision;
                _swapButtonsCheckBox.Checked = existing.SwapButtons;
                _verticalScrollSelector.Value = existing.VerticalScrollLines;
                _horizontalScrollSelector.Value = existing.HorizontalScrollChars;
                _enabledCheckBox.Checked = existing.Enabled;
                _autoApplyCheckBox.Checked = existing.AutoApply;
                _startupApplyCheckBox.Checked = existing.ApplyOnStartup;
            }
            else
            {
                _selectedColor = GetDefaultColor(existingColors);
                _colorPreview.BackColor = _selectedColor;
            }
        }

        GroupBox BuildDeviceGroup()
        {
            var group = new GroupBox
            {
                Text = "Device identification",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(12)
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 4,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));
            for (int i = 0; i < layout.RowCount; i++)
            {
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            layout.Controls.Add(new Label { Text = "Device path", AutoSize = true }, 0, 0);
            _pathTextBox = new TextBox { Dock = DockStyle.Fill };
            layout.Controls.Add(_pathTextBox, 1, 0);

            layout.Controls.Add(new Label { Text = "Capture path", AutoSize = true }, 0, 1);
            var captureButton = new Button { Text = "Capture from movement", AutoSize = true };
            captureButton.Click += (_, __) => BeginCapture();
            layout.Controls.Add(captureButton, 1, 1);

            layout.Controls.Add(new Label { Text = "Device name", AutoSize = true }, 0, 2);
            _nameTextBox = new TextBox { Dock = DockStyle.Fill };
            layout.Controls.Add(_nameTextBox, 1, 2);

            layout.Controls.Add(new Label { Text = "Icon color", AutoSize = true }, 0, 3);
            var colorButton = new Button { Text = "Choose color", AutoSize = true };
            colorButton.Click += (_, __) => ChooseColor();
            var colorPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Fill, WrapContents = false, AutoSize = true };
            _colorPreview = new Panel { Width = 32, Height = 32, BackColor = _selectedColor, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(8, 0, 0, 0) };
            colorPanel.Controls.Add(colorButton);
            colorPanel.Controls.Add(_colorPreview);
            layout.Controls.Add(colorPanel, 1, 3);

            group.Controls.Add(layout);
            return group;
        }

        GroupBox BuildPointerGroup()
        {
            var group = new GroupBox
            {
                Text = "Pointer and buttons",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(12)
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 4,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));
            for (int i = 0; i < layout.RowCount; i++)
            {
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            layout.Controls.Add(new Label { Text = "Pointer speed", AutoSize = true }, 0, 0);
            var speedPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Fill, AutoSize = true };
            _speedSelector = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 20,
                Value = 10,
                Width = 80
            };
            var testButton = new Button { Text = "Test speed", AutoSize = true };
            testButton.Click += (_, __) => TestSpeed();
            speedPanel.Controls.Add(_speedSelector);
            speedPanel.Controls.Add(testButton);
            layout.Controls.Add(speedPanel, 1, 0);

            _enhancePrecisionCheckBox = new CheckBox { Text = "Enhance pointer precision", AutoSize = true, Checked = true };
            layout.Controls.Add(_enhancePrecisionCheckBox, 1, 1);

            _swapButtonsCheckBox = new CheckBox { Text = "Swap left/right buttons", AutoSize = true };
            layout.Controls.Add(_swapButtonsCheckBox, 1, 2);

            var pointerHint = new Label
            {
                Text = "Settings apply when this device becomes active.",
                AutoSize = true,
                ForeColor = Color.DimGray
            };
            layout.Controls.Add(pointerHint, 1, 3);

            group.Controls.Add(layout);
            return group;
        }

        GroupBox BuildScrollGroup()
        {
            var group = new GroupBox
            {
                Text = "Scrolling",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(12)
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 3,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));
            for (int i = 0; i < layout.RowCount; i++)
            {
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            layout.Controls.Add(new Label { Text = "Vertical scroll lines", AutoSize = true }, 0, 0);
            _verticalScrollSelector = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 100,
                Value = 3,
                Width = 80
            };
            layout.Controls.Add(_verticalScrollSelector, 1, 0);

            layout.Controls.Add(new Label { Text = "Horizontal scroll characters", AutoSize = true }, 0, 1);
            _horizontalScrollSelector = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 100,
                Value = 3,
                Width = 80
            };
            layout.Controls.Add(_horizontalScrollSelector, 1, 1);

            var scrollHint = new Label
            {
                Text = "Use 0 to scroll one screen at a time, similar to Windows settings.",
                AutoSize = true,
                ForeColor = Color.DimGray
            };
            layout.Controls.Add(scrollHint, 1, 2);

            group.Controls.Add(layout);
            return group;
        }

        GroupBox BuildActivationGroup()
        {
            var group = new GroupBox
            {
                Text = "Activation",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(12)
            };

            var layout = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false
            };

            _enabledCheckBox = new CheckBox { Text = "Enable this device profile", AutoSize = true, Checked = true };
            _autoApplyCheckBox = new CheckBox { Text = "Apply automatically when this device is used", AutoSize = true, Checked = true };
            _startupApplyCheckBox = new CheckBox { Text = "Apply at app startup", AutoSize = true };

            layout.Controls.Add(_enabledCheckBox);
            layout.Controls.Add(_autoApplyCheckBox);
            layout.Controls.Add(_startupApplyCheckBox);

            group.Controls.Add(layout);
            return group;
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
            _statusLabel.Text = "Captured device path. You can rename it and configure settings.";
            _captureWindow?.Dispose();
            _captureWindow = null;
            Program.EnsureRawInputRegistration();
        }

        void TestSpeed()
        {
            Program.NativeMethods.SetMouseSpeed((int)_speedSelector.Value);
            _statusLabel.Text = $"Applied speed {(int)_speedSelector.Value} for testing.";
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

            NewProfile = new Program.DeviceProfile(
                _nameTextBox.Text.Trim(),
                _pathTextBox.Text.Trim(),
                (int)_speedSelector.Value,
                _selectedColor,
                _enhancePrecisionCheckBox.Checked,
                (int)_verticalScrollSelector.Value,
                (int)_horizontalScrollSelector.Value,
                _swapButtonsCheckBox.Checked,
                _enabledCheckBox.Checked,
                _autoApplyCheckBox.Checked,
                _startupApplyCheckBox.Checked);
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
