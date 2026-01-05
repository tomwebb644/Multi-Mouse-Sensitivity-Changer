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

        static int MIN_SWITCH_MS = 200;
        // ======================================

        static int _lastSpeed = -1;
        static long _lastSwitchMs = 0;
        static readonly Stopwatch _sw = Stopwatch.StartNew();

        static NotifyIcon _notifyIcon;
        static ContextMenuStrip _menu;
        static ToolStripMenuItem _activeDeviceItem;
        static ToolStripMenuItem _startupItem;

        static Icon _defaultIcon;
        static readonly Dictionary<string, Icon> _deviceIcons = new Dictionary<string, Icon>(StringComparer.OrdinalIgnoreCase);

        static string _activeDevicePath = string.Empty;

        static readonly List<DeviceProfile> _devices = new List<DeviceProfile>();

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

            var device = _devices.FirstOrDefault(d => d.DevicePath.Equals(devicePath, StringComparison.OrdinalIgnoreCase));
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

        static void LoadSpeedSettings()
        {
            MOUSE_SPEED = ReadSpeed("Mouse", MOUSE_SPEED);
            X8_SPEED = ReadSpeed("X8", X8_SPEED);
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

        static int ClampSpeed(int speed, int fallback)
        {
            if (speed < 1 || speed > 20)
                return fallback;

            return speed;
        }

        static void UpdateActiveDevice(DeviceProfile device, int speed)
        {
            _activeDevicePath = device.DevicePath;
            _activeDeviceItem.Text = $"Active: {device.Name} (speed {speed})";

            _notifyIcon.Icon = GetDeviceIcon(device);
            _notifyIcon.Text = device.Name + " speed " + speed;
        }

        static void InitializeTrayIcon()
        {
            _defaultIcon = CreateIcon(Color.Gray, "MM");

            RebuildContextMenu();

            _notifyIcon = new NotifyIcon
            {
                Visible = true,
                Icon = _defaultIcon,
                Text = "Multi-Mouse Sensitivity Changer",
                ContextMenuStrip = _menu
            };
        }

        static void RebuildContextMenu()
        {
            _menu?.Dispose();
            _menu = new ContextMenuStrip();
            foreach (var icon in _deviceIcons.Values)
            {
                icon.Dispose();
            }
            _deviceIcons.Clear();

            _activeDeviceItem = new ToolStripMenuItem("Active: none") { Enabled = false };
            _menu.Items.Add(_activeDeviceItem);
            _menu.Items.Add(new ToolStripSeparator());

            foreach (DeviceProfile device in _devices)
            {
                _menu.Items.Add(BuildSpeedMenu(device));
                _menu.Items.Add(new ToolStripMenuItem(device.DevicePath) { Enabled = false });
                _menu.Items.Add(new ToolStripSeparator());
            }

            var addDeviceItem = new ToolStripMenuItem("Add device...");
            addDeviceItem.Click += (_, __) => ShowAddDeviceWizard();
            _menu.Items.Add(addDeviceItem);
            _menu.Items.Add(new ToolStripSeparator());

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

            if (_notifyIcon != null)
                _notifyIcon.ContextMenuStrip = _menu;
        }

        static ToolStripMenuItem BuildSpeedMenu(DeviceProfile device)
        {
            var menu = new ToolStripMenuItem(device.Name + " Speed");
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
            return menu;
        }

        static void OnSpeedMenuClick(object sender, EventArgs e)
        {
            if (sender is ToolStripMenuItem item && item.Tag is SpeedTag tag)
            {
                tag.Device.Speed = tag.Speed;
                DeviceRepository.SaveDevices(_devices);
                UpdateMenuChecks(item.OwnerItem as ToolStripMenuItem, tag.Speed);

                if (_activeDevicePath == tag.Device.DevicePath)
                {
                    NativeMethods.SetMouseSpeed(tag.Speed);
                    _lastSpeed = tag.Speed;
                    UpdateActiveDevice(tag.Device, tag.Speed);
                }
            }
        }

        static void UpdateMenuChecks(ToolStripMenuItem menu, int activeSpeed)
        {
            foreach (ToolStripMenuItem child in menu.DropDownItems)
            {
                child.Checked = child.Text == activeSpeed.ToString();
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

        static Icon GetDeviceIcon(DeviceProfile device)
        {
            if (_deviceIcons.TryGetValue(device.DevicePath, out Icon existing))
                return existing;

            Color[] palette = new[]
            {
                Color.DodgerBlue, Color.MediumPurple, Color.SeaGreen, Color.OrangeRed, Color.Goldenrod,
                Color.CadetBlue, Color.SlateBlue, Color.Teal, Color.SandyBrown
            };

            int colorIndex = Math.Abs(device.DevicePath.GetHashCode()) % palette.Length;
            string label = new string(device.Name.Take(3).ToArray());
            var icon = CreateIcon(palette[colorIndex], label);
            _deviceIcons[device.DevicePath] = icon;
            return icon;
        }

        class TrayApplicationContext : ApplicationContext
        {
            readonly RawInputWindow _window;

            public TrayApplicationContext()
            {
                LoadSpeedSettings();
                DeviceRepository.LoadDevices(_devices, MOUSE_SPEED, X8_SPEED, MOUSE_DEVICE_PATH, X8_DEVICE_PATH);
                InitializeTrayIcon();
                _window = new RawInputWindow(OnDeviceChanged);
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    _notifyIcon?.Dispose();
                    _menu?.Dispose();
                    _defaultIcon?.Dispose();
                    _window?.Dispose();
                    foreach (var icon in _deviceIcons.Values)
                    {
                        icon.Dispose();
                    }
                }
                base.Dispose(disposing);
            }
        }

        static void ShowAddDeviceWizard()
        {
            using (var form = new AddDeviceForm(_devices.Select(d => d.DevicePath)))
            {
                if (form.ShowDialog() == DialogResult.OK && form.CreatedDevice != null)
                {
                    AddOrUpdateDevice(form.CreatedDevice);
                }
            }
        }

        static void AddOrUpdateDevice(DeviceProfile newDevice)
        {
            var existing = _devices.FirstOrDefault(d => d.DevicePath.Equals(newDevice.DevicePath, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                existing.Name = newDevice.Name;
                existing.Speed = newDevice.Speed;
            }
            else
            {
                _devices.Add(newDevice);
            }

            DeviceRepository.SaveDevices(_devices);
            RebuildContextMenu();
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
                    string devicePath = GetDeviceNameFromMessage(m.LParam);
                    if (!string.IsNullOrEmpty(devicePath))
                        _deviceCallback(devicePath);
                }

                base.WndProc(ref m);
            }

            string GetDeviceNameFromMessage(IntPtr lParam)
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

            string GetDeviceName(IntPtr hDevice)
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

            public void Dispose()
            {
                DestroyHandle();
            }
        }

        class SpeedTag
        {
            public DeviceProfile Device { get; }
            public int Speed { get; }

            public SpeedTag(DeviceProfile device, int speed)
            {
                Device = device;
                Speed = speed;
            }
        }

        class DeviceProfile
        {
            public string Name { get; set; }
            public string DevicePath { get; set; }
            public int Speed { get; set; }
        }

        static class DeviceRepository
        {
            const string DevicesValueName = "Devices";

            public static void LoadDevices(List<DeviceProfile> devices, int defaultMouseSpeed, int defaultX8Speed, string defaultMousePath, string defaultX8Path)
            {
                devices.Clear();

                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(SettingsKeyPath, false))
                {
                    if (key != null)
                    {
                        object value = key.GetValue(DevicesValueName);
                        if (value is string[] entries && entries.Length > 0)
                        {
                            foreach (string entry in entries)
                            {
                                if (TryParse(entry, out DeviceProfile device))
                                    devices.Add(device);
                            }
                        }
                    }
                }

                if (devices.Count == 0)
                {
                    devices.Add(new DeviceProfile { Name = "Mouse", DevicePath = defaultMousePath, Speed = defaultMouseSpeed });
                    devices.Add(new DeviceProfile { Name = "X8", DevicePath = defaultX8Path, Speed = defaultX8Speed });
                }
            }

            public static void SaveDevices(IEnumerable<DeviceProfile> devices)
            {
                var entries = devices
                    .Select(d => Serialize(d))
                    .Where(s => !string.IsNullOrEmpty(s))
                    .ToArray();

                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(SettingsKeyPath))
                {
                    key.SetValue(DevicesValueName, entries, RegistryValueKind.MultiString);
                }
            }

            static string Serialize(DeviceProfile device)
            {
                if (string.IsNullOrEmpty(device.DevicePath) || string.IsNullOrEmpty(device.Name))
                    return string.Empty;

                string Encode(string value) => Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
                return string.Join("|", new[] { Encode(device.Name), Encode(device.DevicePath), device.Speed.ToString() });
            }

            static bool TryParse(string entry, out DeviceProfile device)
            {
                device = null;
                if (string.IsNullOrEmpty(entry))
                    return false;

                string[] parts = entry.Split('|');
                if (parts.Length != 3)
                    return false;

                string Decode(string value)
                {
                    try
                    {
                        return Encoding.UTF8.GetString(Convert.FromBase64String(value));
                    }
                    catch
                    {
                        return string.Empty;
                    }
                }

                if (!int.TryParse(parts[2], out int speed))
                    return false;

                device = new DeviceProfile
                {
                    Name = Decode(parts[0]),
                    DevicePath = Decode(parts[1]),
                    Speed = ClampSpeed(speed, speed)
                };

                return !string.IsNullOrEmpty(device.Name) && !string.IsNullOrEmpty(device.DevicePath);
            }
        }

        class AddDeviceForm : Form
        {
            readonly Label _captureLabel;
            readonly TextBox _nameBox;
            readonly TextBox _pathBox;
            readonly NumericUpDown _speedUpDown;
            readonly Button _listenButton;
            readonly Button _testButton;
            readonly HashSet<string> _existingPaths;
            RawInputWindow _listener;

            public DeviceProfile CreatedDevice { get; private set; }

            public AddDeviceForm(IEnumerable<string> existingPaths)
            {
                _existingPaths = new HashSet<string>(existingPaths ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);

                Text = "Add Device";
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                StartPosition = FormStartPosition.CenterScreen;
                ClientSize = new Size(420, 260);

                var step1Label = new Label
                {
                    Text = "Step 1: Listen for device movement to capture its path.",
                    AutoSize = true,
                    Location = new Point(12, 12)
                };

                _captureLabel = new Label
                {
                    Text = "Listening is off.",
                    AutoSize = true,
                    Location = new Point(12, 40)
                };

                _listenButton = new Button
                {
                    Text = "Start listening",
                    Location = new Point(12, 65),
                    Size = new Size(120, 28)
                };
                _listenButton.Click += (_, __) => ToggleListening();

                var pathLabel = new Label
                {
                    Text = "Captured device path:",
                    AutoSize = true,
                    Location = new Point(12, 105)
                };

                _pathBox = new TextBox
                {
                    ReadOnly = true,
                    Location = new Point(15, 125),
                    Width = 380
                };

                var nameLabel = new Label
                {
                    Text = "Step 2: Name the device:",
                    AutoSize = true,
                    Location = new Point(12, 155)
                };

                _nameBox = new TextBox
                {
                    Location = new Point(15, 175),
                    Width = 200
                };

                var speedLabel = new Label
                {
                    Text = "Step 3: Choose a default speed (1-20):",
                    AutoSize = true,
                    Location = new Point(230, 155)
                };

                _speedUpDown = new NumericUpDown
                {
                    Minimum = 1,
                    Maximum = 20,
                    Value = 10,
                    Location = new Point(233, 175),
                    Width = 60
                };

                _testButton = new Button
                {
                    Text = "Test speed",
                    Location = new Point(305, 172),
                    Size = new Size(90, 28)
                };
                _testButton.Click += (_, __) => NativeMethods.SetMouseSpeed((int)_speedUpDown.Value);

                var instructions = new Label
                {
                    Text = "Step 4: Click OK to save. You can move the device to test switching.",
                    AutoSize = true,
                    Location = new Point(12, 205)
                };

                var okButton = new Button
                {
                    Text = "OK",
                    DialogResult = DialogResult.OK,
                    Location = new Point(240, 225),
                    Size = new Size(75, 25)
                };
                okButton.Click += OnOkClicked;

                var cancelButton = new Button
                {
                    Text = "Cancel",
                    DialogResult = DialogResult.Cancel,
                    Location = new Point(320, 225),
                    Size = new Size(75, 25)
                };

                Controls.Add(step1Label);
                Controls.Add(_captureLabel);
                Controls.Add(_listenButton);
                Controls.Add(pathLabel);
                Controls.Add(_pathBox);
                Controls.Add(nameLabel);
                Controls.Add(_nameBox);
                Controls.Add(speedLabel);
                Controls.Add(_speedUpDown);
                Controls.Add(_testButton);
                Controls.Add(instructions);
                Controls.Add(okButton);
                Controls.Add(cancelButton);

                AcceptButton = okButton;
                CancelButton = cancelButton;
            }

            void ToggleListening()
            {
                if (_listener == null)
                {
                    _listener = new RawInputWindow(OnDeviceCaptured);
                    _captureLabel.Text = "Listening... move the target mouse.";
                    _listenButton.Text = "Stop listening";
                }
                else
                {
                    _listener.Dispose();
                    _listener = null;
                    _captureLabel.Text = "Listening is off.";
                    _listenButton.Text = "Start listening";
                }
            }

            void OnDeviceCaptured(string path)
            {
                if (string.IsNullOrEmpty(path))
                    return;

                BeginInvoke(new Action(() =>
                {
                    if (_existingPaths.Contains(path))
                    {
                        _captureLabel.Text = "Device already added; captured existing path.";
                    }
                    else
                    {
                        _captureLabel.Text = "Captured new device path.";
                    }

                    _pathBox.Text = path;
                    if (string.IsNullOrWhiteSpace(_nameBox.Text))
                    {
                        _nameBox.Text = "Mouse " + (_existingPaths.Count + 1);
                    }
                }));
            }

            void OnOkClicked(object sender, EventArgs e)
            {
                if (string.IsNullOrWhiteSpace(_pathBox.Text))
                {
                    MessageBox.Show("Please capture a device by listening and moving it first.", "Device required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                    return;
                }

                if (string.IsNullOrWhiteSpace(_nameBox.Text))
                {
                    MessageBox.Show("Please provide a device name.", "Name required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                    return;
                }

                CreatedDevice = new DeviceProfile
                {
                    Name = _nameBox.Text.Trim(),
                    DevicePath = _pathBox.Text.Trim(),
                    Speed = (int)_speedUpDown.Value
                };
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    _listener?.Dispose();
                }
                base.Dispose(disposing);
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
