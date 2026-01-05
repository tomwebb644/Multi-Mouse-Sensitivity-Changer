using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace MultiMouseSensitivityChanger
{
    static class Program
    {
        // ====== SET THESE AFTER FIRST RUN ======
        static string X8_DEVICE_PATH =
        @"\\?\HID#VID_1997&PID_2433&MI_01&Col01#9&354159c2&0&0000#{378de44c-56ef-11d1-bc8c-00a0c91405dd}";

        static string MOUSE_DEVICE_PATH =
        @"\\?\HID#VID_046D&PID_C539&MI_01&Col01#8&10c9e4b2&0&0000#{378de44c-56ef-11d1-bc8c-00a0c91405dd}";

        static int X8_SPEED = 20;    // 1..20
        static int MOUSE_SPEED = 10; // 1..20

        const string SettingsKeyPath = "Software\\MultiMouseSensitivityChanger";
        const string DevicesValueName = "Devices";

        static int MIN_SWITCH_MS = 200;
        // ======================================

        static int _lastSpeed = -1;
        static long _lastSwitchMs = 0;
        static readonly Stopwatch _sw = Stopwatch.StartNew();

        static NotifyIcon _notifyIcon;
        static ContextMenuStrip _menu;
        static ToolStripMenuItem _activeDeviceItem;
        static ToolStripMenuItem _startupItem;
        static ToolStripSeparator _devicesSeparator;
        static ToolStripSeparator _afterDevicesSeparator;
        static ToolStripMenuItem _addDeviceItem;

        static Icon _mouseIcon;
        static Icon _x8Icon;
        static Icon _defaultIcon;
        static readonly Dictionary<string, Icon> _deviceIcons = new Dictionary<string, Icon>(StringComparer.OrdinalIgnoreCase);

        static DeviceConfig _activeDevice;

        const string X8_KEY = "X8";
        const string MOUSE_KEY = "Mouse";

        static readonly List<DeviceConfig> _devices = new List<DeviceConfig>();
        static readonly Color[] DeviceColors = new[]
        {
            Color.SeaGreen,
            Color.Teal,
            Color.IndianRed,
            Color.DarkOrange,
            Color.MediumSlateBlue,
            Color.SteelBlue
        };

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TrayApplicationContext());
        }

        static void OnDeviceChanged(string devicePath)
        {
            if (string.IsNullOrEmpty(devicePath))
                return;

            DeviceConfig device = _devices.FirstOrDefault(d => devicePath.Equals(d.DevicePath, StringComparison.OrdinalIgnoreCase));

            if (device == null)
                return;

            int targetSpeed = device.Speed;
            long now = _sw.ElapsedMilliseconds;
            if (targetSpeed != _lastSpeed && (now - _lastSwitchMs) >= MIN_SWITCH_MS)
            {
                NativeMethods.SetMouseSpeed(targetSpeed);
                _lastSpeed = targetSpeed;
                _lastSwitchMs = now;
                UpdateActiveDevice(device, targetSpeed);
            }
        }

        static void LoadDevices()
        {
            _devices.Clear();

            string rawList = ReadDeviceList();
            if (!string.IsNullOrWhiteSpace(rawList))
            {
                _devices.AddRange(DeviceSerializer.Deserialize(rawList));
            }

            if (_devices.Count == 0)
            {
                MOUSE_SPEED = ReadSpeed(MOUSE_KEY, MOUSE_SPEED);
                X8_SPEED = ReadSpeed(X8_KEY, X8_SPEED);

                _devices.Add(new DeviceConfig(MOUSE_KEY, MOUSE_DEVICE_PATH, MOUSE_SPEED));
                _devices.Add(new DeviceConfig(X8_KEY, X8_DEVICE_PATH, X8_SPEED));
                SaveDevices();
            }
        }

        static int ReadSpeed(string deviceKey, int defaultValue)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(SettingsKeyPath, false))
            {
                object value = key?.GetValue(deviceKey);

                if (value is int intValue)
                    return ClampSpeed(intValue, defaultValue);

                if (value is string str && int.TryParse(str, out int parsed))
                    return ClampSpeed(parsed, defaultValue);
            }

            return defaultValue;
        }

        static void SaveSpeed(string deviceKey, int speed)
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(SettingsKeyPath))
            {
                key.SetValue(deviceKey, ClampSpeed(speed, speed), RegistryValueKind.DWord);
            }
        }

        static void SaveDevices()
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(SettingsKeyPath))
            {
                key.SetValue(DevicesValueName, DeviceSerializer.Serialize(_devices), RegistryValueKind.String);
            }
        }

        static string ReadDeviceList()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(SettingsKeyPath, false))
            {
                return key?.GetValue(DevicesValueName) as string;
            }
        }

        static int ClampSpeed(int speed, int fallback)
        {
            if (speed < 1 || speed > 20)
                return fallback;

            return speed;
        }

        static void UpdateActiveDevice(DeviceConfig device, int speed)
        {
            _activeDevice = device;

            if (device == null)
            {
                _activeDeviceItem.Text = "Active: none";
                _notifyIcon.Icon = _defaultIcon;
                _notifyIcon.Text = "Multi-Mouse Sensitivity Changer";
                return;
            }

            _activeDeviceItem.Text = $"Active: {device.Name} (speed {speed})";
            _notifyIcon.Icon = GetIconForDevice(device);
            _notifyIcon.Text = $"{device.Name} speed {speed}";
        }

        static void InitializeTrayIcon()
        {
            _mouseIcon = CreateIcon(Color.DodgerBlue, "M");
            _x8Icon = CreateIcon(Color.MediumPurple, "X8");
            _defaultIcon = CreateIcon(Color.Gray, "MM");

            _menu = new ContextMenuStrip();
            _activeDeviceItem = new ToolStripMenuItem("Active: none") { Enabled = false };
            _menu.Items.Add(_activeDeviceItem);
            _menu.Items.Add(new ToolStripSeparator());

            _addDeviceItem = new ToolStripMenuItem("Add device...");
            _addDeviceItem.Click += (_, __) => ShowAddDeviceWizard();
            _menu.Items.Add(_addDeviceItem);

            _devicesSeparator = new ToolStripSeparator();
            _afterDevicesSeparator = new ToolStripSeparator();
            _menu.Items.Add(_devicesSeparator);
            _menu.Items.Add(_afterDevicesSeparator);

            _startupItem = new ToolStripMenuItem("Start with Windows")
            {
                CheckOnClick = true,
                Checked = StartupManager.IsEnabled()
            };
            _startupItem.CheckedChanged += (_, __) => StartupManager.SetEnabled(_startupItem.Checked);
            _menu.Items.Add(_startupItem);

            _menu.Items.Add(new ToolStripSeparator());
            var exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += (_, __) => Application.Exit();
            _menu.Items.Add(exitItem);

            RefreshDeviceMenus();

            _notifyIcon = new NotifyIcon
            {
                Visible = true,
                Icon = _defaultIcon,
                Text = "Multi-Mouse Sensitivity Changer",
                ContextMenuStrip = _menu
            };
        }

        static void RefreshDeviceMenus()
        {
            if (_menu == null)
                return;

            int startIndex = _menu.Items.IndexOf(_devicesSeparator) + 1;
            int endIndex = _menu.Items.IndexOf(_afterDevicesSeparator);

            while (endIndex > startIndex)
            {
                _menu.Items.RemoveAt(startIndex);
                endIndex--;
            }

            foreach (DeviceConfig device in _devices)
            {
                var menu = BuildSpeedMenu(device);
                _menu.Items.Insert(startIndex++, menu);
            }
        }

        static void AddOrUpdateDevice(DeviceConfig device)
        {
            if (device == null || string.IsNullOrWhiteSpace(device.DevicePath))
                return;

            DeviceConfig existing = _devices.FirstOrDefault(d => device.DevicePath.Equals(d.DevicePath, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                existing.Name = device.Name;
                existing.Speed = device.Speed;
            }
            else
            {
                _devices.Add(device);
            }

            SaveDevices();
            RefreshDeviceMenus();

            if (_activeDevice != null && device.DevicePath.Equals(_activeDevice.DevicePath, StringComparison.OrdinalIgnoreCase))
                UpdateActiveDevice(existing ?? device, (existing ?? device).Speed);
        }

        static Icon GetIconForDevice(DeviceConfig device)
        {
            if (device == null)
                return _defaultIcon;

            if (device.DevicePath.Equals(X8_DEVICE_PATH, StringComparison.OrdinalIgnoreCase))
                return _x8Icon;

            if (device.DevicePath.Equals(MOUSE_DEVICE_PATH, StringComparison.OrdinalIgnoreCase))
                return _mouseIcon;

            if (_deviceIcons.TryGetValue(device.DevicePath, out Icon icon))
                return icon;

            string name = device.Name ?? string.Empty;
            Color color = DeviceColors[Math.Abs(name.Aggregate(0, (acc, c) => acc + c)) % DeviceColors.Length];
            string label = CreateIconLabel(device.Name);
            icon = CreateIcon(color, label);
            _deviceIcons[device.DevicePath] = icon;
            return icon;
        }

        static string CreateIconLabel(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "MM";

            string letters = new string(name.Where(char.IsLetterOrDigit).Take(3).ToArray());
            return string.IsNullOrEmpty(letters) ? "MM" : letters.ToUpperInvariant();
        }

        static void ShowAddDeviceWizard()
        {
            using (var form = new AddDeviceForm(AddOrUpdateDevice))
            {
                form.StartPosition = FormStartPosition.CenterScreen;
                form.ShowDialog();
            }
        }

        static ToolStripMenuItem BuildSpeedMenu(DeviceConfig device)
        {
            var menu = new ToolStripMenuItem(device.Name);
            for (int i = 1; i <= 20; i++)
            {
                var item = new ToolStripMenuItem(i.ToString())
                {
                    Tag = new SpeedTag(device, i),
                    Checked = i == device.Speed
                };
                item.Click += OnSpeedMenuClick;
                menu.DropDownItems.Add(item);
            }

            menu.DropDownItems.Add(new ToolStripSeparator());
            menu.DropDownItems.Add(new ToolStripMenuItem("Device path") { Enabled = false, ToolTipText = device.DevicePath });
            menu.DropDownItems.Add(new ToolStripMenuItem(device.DevicePath) { Enabled = false });
            return menu;
        }

        static void OnSpeedMenuClick(object sender, EventArgs e)
        {
            if (sender is ToolStripMenuItem item && item.Tag is SpeedTag tag)
            {
                tag.Device.Speed = tag.Speed;
                SaveDevices();
                UpdateMenuChecks(item.OwnerItem as ToolStripMenuItem, tag.Speed);

                if (_activeDevice == tag.Device)
                {
                    NativeMethods.SetMouseSpeed(tag.Speed);
                    _lastSpeed = tag.Speed;
                    UpdateActiveDevice(tag.Device, tag.Speed);
                }
            }
        }

        static void UpdateMenuChecks(ToolStripMenuItem menu, int activeSpeed)
        {
            if (menu == null)
                return;

            foreach (ToolStripItem child in menu.DropDownItems)
            {
                if (child is ToolStripMenuItem childItem && childItem.Tag is SpeedTag)
                    childItem.Checked = childItem.Text == activeSpeed.ToString();
            }
        }

        static Icon CreateIcon(Color color, string label)
        {
            using (var bmp = new Bitmap(32, 32))
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                using (var brush = new SolidBrush(color))
                {
                    g.FillEllipse(brush, 0, 0, 31, 31);
                }

                using (var pen = new Pen(Color.White, 2))
                {
                    g.DrawEllipse(pen, 1, 1, 29, 29);
                }

                using (var font = new Font(FontFamily.GenericSansSerif, label.Length > 2 ? 7 : 10, FontStyle.Bold))
                using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                using (var textBrush = new SolidBrush(Color.White))
                {
                    g.DrawString(label, font, textBrush, new RectangleF(0, 0, 32, 32), format);
                }

                IntPtr hIcon = bmp.GetHicon();
                try
                {
                    using (var icon = Icon.FromHandle(hIcon))
                    {
                        return (Icon)icon.Clone();
                    }
                }
                finally
                {
                    NativeMethods.DestroyIcon(hIcon);
                }
            }
        }

        static string GetDevicePathFromMessage(IntPtr lParam)
        {
            uint size = 0;
            NativeMethods.GetRawInputData(lParam, NativeMethods.RID_INPUT, IntPtr.Zero, ref size, (uint)Marshal.SizeOf(typeof(RAWINPUTHEADER)));
            if (size == 0)
                return string.Empty;

            IntPtr buffer = Marshal.AllocHGlobal((int)size);
            try
            {
                if (NativeMethods.GetRawInputData(lParam, NativeMethods.RID_INPUT, buffer, ref size, (uint)Marshal.SizeOf(typeof(RAWINPUTHEADER))) != size)
                    return string.Empty;

                RAWINPUT raw = (RAWINPUT)Marshal.PtrToStructure(buffer, typeof(RAWINPUT));
                if (raw.header.dwType != NativeMethods.RIM_TYPEMOUSE)
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

        static string GetDeviceName(IntPtr hDevice)
        {
            uint size = 0;
            NativeMethods.GetRawInputDeviceInfo(hDevice, NativeMethods.RIDI_DEVICENAME, IntPtr.Zero, ref size);
            if (size == 0)
                return string.Empty;

            IntPtr data = Marshal.AllocHGlobal((int)size);
            try
            {
                if (NativeMethods.GetRawInputDeviceInfo(hDevice, NativeMethods.RIDI_DEVICENAME, data, ref size) == uint.MaxValue)
                    return string.Empty;

                return Marshal.PtrToStringAnsi(data);
            }
            finally
            {
                Marshal.FreeHGlobal(data);
            }
        }

        class TrayApplicationContext : ApplicationContext
        {
            readonly RawInputWindow _window;

            public TrayApplicationContext()
            {
                LoadDevices();
                InitializeTrayIcon();
                _window = new RawInputWindow(OnDeviceChanged);
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    _notifyIcon?.Dispose();
                    _menu?.Dispose();
                    _mouseIcon?.Dispose();
                    _x8Icon?.Dispose();
                    _defaultIcon?.Dispose();
                    foreach (var icon in _deviceIcons.Values)
                        icon?.Dispose();
                    _deviceIcons.Clear();
                    _window?.Dispose();
                }
                base.Dispose(disposing);
            }
        }

        class RawInputWindow : NativeWindow, IDisposable
        {
            readonly Action<string> _deviceCallback;

            public RawInputWindow(Action<string> deviceCallback)
            {
                _deviceCallback = deviceCallback;
                CreateHandle(new CreateParams());

                RAWINPUTDEVICE[] rid = new[]
                {
                    new RAWINPUTDEVICE
                    {
                        usUsagePage = 0x01,
                        usUsage = 0x02,
                        dwFlags = NativeMethods.RIDEV_INPUTSINK,
                        hwndTarget = Handle
                    }
                };

                if (!NativeMethods.RegisterRawInputDevices(rid, (uint)rid.Length, (uint)Marshal.SizeOf(typeof(RAWINPUTDEVICE))))
                    throw new InvalidOperationException("RegisterRawInputDevices failed, Win32=" + Marshal.GetLastWin32Error());
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == NativeMethods.WM_INPUT)
                {
                    string devicePath = GetDevicePathFromMessage(m.LParam);
                    if (!string.IsNullOrEmpty(devicePath))
                        _deviceCallback(devicePath);
                }

                base.WndProc(ref m);
            }

            public void Dispose()
            {
                DestroyHandle();
            }
        }

        class SpeedTag
        {
            public DeviceConfig Device { get; }
            public int Speed { get; }

            public SpeedTag(DeviceConfig device, int speed)
            {
                Device = device;
                Speed = speed;
            }
        }

        class DeviceConfig
        {
            public string Name { get; set; }
            public string DevicePath { get; set; }
            public int Speed { get; set; }

            public DeviceConfig()
            {
            }

            public DeviceConfig(string name, string devicePath, int speed)
            {
                Name = string.IsNullOrWhiteSpace(name) ? "Device" : name.Trim();
                DevicePath = devicePath ?? string.Empty;
                Speed = ClampSpeed(speed, 10);
            }
        }

        static class DeviceSerializer
        {
            public static string Serialize(IEnumerable<DeviceConfig> devices)
            {
                var sb = new StringBuilder();
                foreach (DeviceConfig device in devices)
                {
                    string name = Convert.ToBase64String(Encoding.UTF8.GetBytes(device.Name ?? string.Empty));
                    string path = Convert.ToBase64String(Encoding.UTF8.GetBytes(device.DevicePath ?? string.Empty));
                    sb.Append(name)
                        .Append('|')
                        .Append(path)
                        .Append('|')
                        .Append(ClampSpeed(device.Speed, 10))
                        .Append('\n');
                }

                return sb.ToString().TrimEnd('\n');
            }

            public static IEnumerable<DeviceConfig> Deserialize(string raw)
            {
                if (string.IsNullOrWhiteSpace(raw))
                    yield break;

                string[] lines = raw.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string line in lines)
                {
                    string[] parts = line.Split('|');
                    if (parts.Length != 3)
                        continue;

                    try
                    {
                        string name = Encoding.UTF8.GetString(Convert.FromBase64String(parts[0]));
                        string path = Encoding.UTF8.GetString(Convert.FromBase64String(parts[1]));
                        int speed = int.TryParse(parts[2], out int parsed) ? parsed : 10;
                        yield return new DeviceConfig(name, path, speed);
                    }
                    catch
                    {
                        // ignore malformed entries
                    }
                }
            }
        }

        class AddDeviceForm : Form
        {
            readonly Action<DeviceConfig> _saveCallback;

            readonly TextBox _nameBox;
            readonly TextBox _pathBox;
            readonly NumericUpDown _speedUpDown;
            readonly Button _captureButton;
            readonly Button _testButton;
            readonly Label _statusLabel;

            bool _listening;

            public AddDeviceForm(Action<DeviceConfig> saveCallback)
            {
                _saveCallback = saveCallback;
                Text = "Add device";
                Width = 440;
                Height = 260;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;

                var nameLabel = new Label { Text = "Device name", AutoSize = true, Left = 15, Top = 20 };
                _nameBox = new TextBox { Left = 130, Top = 16, Width = 270, Text = "New device" };

                var pathLabel = new Label { Text = "Device path", AutoSize = true, Left = 15, Top = 55 };
                _pathBox = new TextBox { Left = 130, Top = 51, Width = 270, ReadOnly = true };

                _captureButton = new Button { Text = "Start capture", Left = 130, Top = 82, Width = 120 };
                _captureButton.Click += (_, __) => StartListening();

                var instructions = new Label
                {
                    Left = 260,
                    Top = 86,
                    Width = 150,
                    Text = "Move the mouse you want to add while capture is active.",
                    AutoSize = true
                };

                var speedLabel = new Label { Text = "Default speed", AutoSize = true, Left = 15, Top = 130 };
                _speedUpDown = new NumericUpDown
                {
                    Left = 130,
                    Top = 126,
                    Minimum = 1,
                    Maximum = 20,
                    Value = 10,
                    Width = 60
                };

                _testButton = new Button { Text = "Test speed", Left = 200, Top = 124, Width = 90 };
                _testButton.Click += (_, __) => TestSpeed();

                var saveButton = new Button { Text = "Save device", Left = 130, Top = 170, Width = 120 };
                saveButton.Click += (_, __) => SaveDevice();

                var cancelButton = new Button { Text = "Cancel", Left = 260, Top = 170, Width = 80, DialogResult = DialogResult.Cancel };

                _statusLabel = new Label { Left = 15, Top = 205, Width = 390, Text = "Click Start capture to detect a device." };

                Controls.AddRange(new Control[]
                {
                    nameLabel, _nameBox,
                    pathLabel, _pathBox,
                    _captureButton, instructions,
                    speedLabel, _speedUpDown, _testButton,
                    saveButton, cancelButton,
                    _statusLabel
                });
            }

            void StartListening()
            {
                RAWINPUTDEVICE[] rid = new[]
                {
                    new RAWINPUTDEVICE
                    {
                        usUsagePage = 0x01,
                        usUsage = 0x02,
                        dwFlags = NativeMethods.RIDEV_INPUTSINK,
                        hwndTarget = Handle
                    }
                };

                if (!NativeMethods.RegisterRawInputDevices(rid, (uint)rid.Length, (uint)Marshal.SizeOf(typeof(RAWINPUTDEVICE))))
                {
                    MessageBox.Show("Unable to start capture. Please try again.", "Capture failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                _listening = true;
                _statusLabel.Text = "Capturing input. Move the device you want to add.";
            }

            void SaveDevice()
            {
                if (string.IsNullOrWhiteSpace(_pathBox.Text))
                {
                    MessageBox.Show("No device movement detected yet. Move the device while capture is active.", "Missing device", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string name = string.IsNullOrWhiteSpace(_nameBox.Text) ? "New device" : _nameBox.Text.Trim();
                int speed = (int)_speedUpDown.Value;

                _saveCallback?.Invoke(new DeviceConfig(name, _pathBox.Text, speed));
                DialogResult = DialogResult.OK;
                Close();
            }

            void TestSpeed()
            {
                NativeMethods.SetMouseSpeed((int)_speedUpDown.Value);
                _statusLabel.Text = $"Applied speed {(int)_speedUpDown.Value} for testing.";
            }

            protected override void WndProc(ref Message m)
            {
                if (_listening && m.Msg == NativeMethods.WM_INPUT)
                {
                    string devicePath = GetDevicePathFromMessage(m.LParam);
                    if (!string.IsNullOrEmpty(devicePath))
                    {
                        _pathBox.Text = devicePath;
                        _statusLabel.Text = "Device detected. You can now save it.";
                        _listening = false;
                    }
                }

                base.WndProc(ref m);
            }
        }

        static class StartupManager
        {
            const string RunKeyPath = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";
            static readonly string AppName = "Multi-Mouse-Sensitivity-Changer";

            public static bool IsEnabled()
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RunKeyPath, false))
                {
                    string value = key?.GetValue(AppName) as string;
                    return !string.IsNullOrEmpty(value);
                }
            }

            public static void SetEnabled(bool enabled)
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RunKeyPath, true) ?? Registry.CurrentUser.CreateSubKey(RunKeyPath))
                {
                    if (enabled)
                    {
                        string exePath = Process.GetCurrentProcess().MainModule.FileName;
                        key.SetValue(AppName, '"' + exePath + '"');
                    }
                    else
                    {
                        key.DeleteValue(AppName, false);
                    }
                }
            }
        }

        internal static class NativeMethods
        {
            public const uint WM_INPUT = 0x00FF;
            public const uint RID_INPUT = 0x10000003;
            public const uint RIM_TYPEMOUSE = 0;
            public const uint RIDI_DEVICENAME = 0x20000007;
            public const uint RIDEV_INPUTSINK = 0x00000100;
            public const uint SPI_SETMOUSESPEED = 0x0071;
            public const uint SPIF_SENDCHANGE = 0x02;

            public static void SetMouseSpeed(int speed)
            {
                if (speed < 1) speed = 1;
                if (speed > 20) speed = 20;
                SystemParametersInfo(SPI_SETMOUSESPEED, 0, (IntPtr)speed, SPIF_SENDCHANGE);
            }

            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool SystemParametersInfo(uint uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool RegisterRawInputDevices(RAWINPUTDEVICE[] pRawInputDevices, uint uiNumDevices, uint cbSize);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern uint GetRawInputData(IntPtr hRawInput, uint uiCommand, IntPtr pData, ref uint pcbSize, uint cbSizeHeader);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern uint GetRawInputDeviceInfo(IntPtr hDevice, uint uiCommand, IntPtr pData, ref uint pcbSize);

            [DllImport("user32.dll")]
            public static extern bool DestroyIcon(IntPtr handle);
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RAWINPUTDEVICE
        {
            public ushort usUsagePage;
            public ushort usUsage;
            public uint dwFlags;
            public IntPtr hwndTarget;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RAWINPUTHEADER
        {
            public uint dwType;
            public uint dwSize;
            public IntPtr hDevice;
            public IntPtr wParam;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct RAWINPUT
        {
            [FieldOffset(0)] public RAWINPUTHEADER header;
            [FieldOffset(16)] public RAWMOUSE data;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RAWMOUSE
        {
            public ushort usFlags;
            public uint ulButtons;
            public ushort usButtonFlags;
            public ushort usButtonData;
            public uint ulRawButtons;
            public int lLastX;
            public int lLastY;
            public uint ulExtraInformation;
        }
    }
}
