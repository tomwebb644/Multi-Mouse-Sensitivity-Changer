using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

namespace MultiMouseSensitivityChanger
{
    static class Program
    {
        const string SettingsKeyPath = "Software\\MultiMouseSensitivityChanger";

        static int MIN_SWITCH_MS = 200;
        // ======================================

        static int _lastSpeed = -1;
        static long _lastSwitchMs = 0;
        static readonly Stopwatch _sw = Stopwatch.StartNew();

        static NotifyIcon _notifyIcon;
        static ContextMenuStrip _menu;
        static readonly Dictionary<string, ToolStripMenuItem> _speedMenus = new Dictionary<string, ToolStripMenuItem>(StringComparer.OrdinalIgnoreCase);
        static ToolStripMenuItem _activeDeviceItem;
        static ToolStripMenuItem _startupItem;
        static ToolStripMenuItem _pathsMenu;

        static Icon _defaultIcon;
        static readonly Dictionary<string, Icon> _deviceIcons = new Dictionary<string, Icon>(StringComparer.OrdinalIgnoreCase);

        static string _activeDeviceKey = string.Empty;

        static readonly List<DeviceProfile> _deviceProfiles = new List<DeviceProfile>();

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

            var profile = _deviceProfiles.FirstOrDefault(p => devicePath.Equals(p.DevicePath, StringComparison.OrdinalIgnoreCase));
            if (profile == null)
                return;

            int targetSpeed = profile.Speed;
            long now = _sw.ElapsedMilliseconds;
            if (targetSpeed != _lastSpeed && (now - _lastSwitchMs) >= MIN_SWITCH_MS)
            {
                NativeMethods.SetMouseSpeed(targetSpeed);
                _lastSpeed = targetSpeed;
                _lastSwitchMs = now;
                UpdateActiveDevice(profile, targetSpeed);
            }
        }

        static void InitializeDevices()
        {
            _deviceProfiles.Clear();
            _deviceProfiles.AddRange(LoadDevices());
        }

        static IEnumerable<DeviceProfile> LoadDevices()
        {
            var devices = new List<DeviceProfile>();

            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(SettingsKeyPath, false))
            {
                string[] stored = key?.GetValue("Devices") as string[];
                if (stored != null)
                {
                    foreach (string entry in stored)
                    {
                        if (string.IsNullOrWhiteSpace(entry))
                            continue;

                        string[] parts = entry.Split(new[] { '|' }, 3);
                        if (parts.Length < 3)
                            continue;

                        string name = parts[0];
                        string path = parts[1];
                        if (!int.TryParse(parts[2], out int speed))
                            continue;

                        Color iconColor = Color.Gray;
                        if (parts.Length >= 4)
                        {
                            try
                            {
                                iconColor = Color.FromArgb(Convert.ToInt32(parts[3], 16));
                            }
                            catch
                            {
                                iconColor = Color.Gray;
                            }
                        }

                        if (string.IsNullOrWhiteSpace(path))
                            continue;

                        var existing = devices.FirstOrDefault(p => p.DevicePath.Equals(path, StringComparison.OrdinalIgnoreCase));
                        if (existing != null)
                        {
                            existing.Name = existing.Name ?? name;
                            existing.Speed = ClampSpeed(speed, existing.Speed);
                            existing.IconColor = iconColor;
                        }
                        else
                        {
                            devices.Add(new DeviceProfile(name, path, ClampSpeed(speed, speed), iconColor));
                        }
                    }
                }
            }

            return devices;
        }

        static void SaveDevices()
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(SettingsKeyPath))
            {
                string[] serialized = _deviceProfiles
                    .Select(p => $"{p.Name}|{p.DevicePath}|{ClampSpeed(p.Speed, p.Speed)}|{p.IconColor.ToArgb():X8}")
                    .ToArray();
                key.SetValue("Devices", serialized, RegistryValueKind.MultiString);
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

        static int ClampSpeed(int speed, int fallback)
        {
            if (speed < 1 || speed > 20)
                return fallback;

            return speed;
        }

        static void UpdateActiveDevice(DeviceProfile profile, int speed)
        {
            _activeDeviceKey = profile.Name;
            _activeDeviceItem.Text = $"Active: {profile.Name} (speed {speed})";

            _notifyIcon.Icon = GetIconForProfile(profile);
            _notifyIcon.Text = $"{profile.Name} speed {speed}";
        }

        static void InitializeTrayIcon()
        {
            _defaultIcon = CreateIcon(Color.Gray, "MM");

            _menu = new ContextMenuStrip();
            _activeDeviceItem = new ToolStripMenuItem("Active: none") { Enabled = false };

            _startupItem = new ToolStripMenuItem("Start with Windows")
            {
                CheckOnClick = true,
                Checked = StartupManager.IsEnabled()
            };
            _startupItem.CheckedChanged += (_, __) => StartupManager.SetEnabled(_startupItem.Checked);

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
            _menu.Items.Clear();
            _speedMenus.Clear();
            foreach (var icon in _deviceIcons.Values)
                icon?.Dispose();
            _deviceIcons.Clear();

            _menu.Items.Add(_activeDeviceItem);
            _menu.Items.Add(new ToolStripSeparator());

            if (_deviceProfiles.Count == 0)
            {
                _menu.Items.Add(new ToolStripMenuItem("No devices configured") { Enabled = false });
            }
            else
            {
                foreach (var profile in _deviceProfiles)
                {
                    var speedMenu = BuildSpeedMenu(profile);
                    _speedMenus[profile.Name] = speedMenu;
                    _menu.Items.Add(speedMenu);
                }
            }

            _menu.Items.Add(new ToolStripSeparator());
            _pathsMenu = BuildPathsMenu();
            _menu.Items.Add(_pathsMenu);

            _menu.Items.Add(new ToolStripSeparator());
            var addDeviceItem = new ToolStripMenuItem("Add new device...");
            addDeviceItem.Click += (_, __) => ShowAddDeviceDialog();
            _menu.Items.Add(addDeviceItem);

            var manageDevicesItem = new ToolStripMenuItem("Manage devices...");
            manageDevicesItem.Click += (_, __) => ShowManageDevicesDialog();
            _menu.Items.Add(manageDevicesItem);

            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(_startupItem);

            _menu.Items.Add(new ToolStripSeparator());
            var exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += (_, __) => Application.Exit();
            _menu.Items.Add(exitItem);
        }

        static ToolStripMenuItem BuildSpeedMenu(DeviceProfile profile)
        {
            var menu = new ToolStripMenuItem($"{profile.Name} Speed")
            {
                Tag = profile.DevicePath,
                ToolTipText = profile.DevicePath
            };
            for (int i = 1; i <= 20; i++)
            {
                var item = new ToolStripMenuItem(i.ToString())
                {
                    Tag = new SpeedTag(profile.Name, i),
                    Checked = i == profile.Speed
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
                var profile = _deviceProfiles.FirstOrDefault(p => p.Name.Equals(tag.DeviceName, StringComparison.OrdinalIgnoreCase));
                if (profile == null)
                    return;

                profile.Speed = tag.Speed;
                SaveDevices();

                if (_speedMenus.TryGetValue(profile.Name, out var menu))
                    UpdateMenuChecks(menu, tag.Speed);

                if (_activeDeviceKey.Equals(profile.Name, StringComparison.OrdinalIgnoreCase))
                {
                    NativeMethods.SetMouseSpeed(tag.Speed);
                    _lastSpeed = tag.Speed;
                    UpdateActiveDevice(profile, tag.Speed);
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

        static ToolStripMenuItem BuildPathsMenu()
        {
            var menu = new ToolStripMenuItem("Device paths") { Enabled = true };
            foreach (var profile in _deviceProfiles)
            {
                menu.DropDownItems.Add(new ToolStripMenuItem(profile.Name) { Enabled = false, ToolTipText = profile.DevicePath });
                menu.DropDownItems.Add(new ToolStripMenuItem(profile.DevicePath) { Enabled = false });
                menu.DropDownItems.Add(new ToolStripSeparator());
            }

            if (menu.DropDownItems.Count > 0 && menu.DropDownItems[menu.DropDownItems.Count - 1] is ToolStripSeparator)
                menu.DropDownItems.RemoveAt(menu.DropDownItems.Count - 1);

            if (menu.DropDownItems.Count == 0)
                menu.DropDownItems.Add(new ToolStripMenuItem("No devices configured") { Enabled = false });

            return menu;
        }

        static void ShowAddDeviceDialog()
        {
            using (var form = new AddDeviceForm(null, name => _deviceProfiles.Any(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase))))
            {
                if (form.ShowDialog() == DialogResult.OK && form.NewProfile != null)
                {
                    var existing = _deviceProfiles.FirstOrDefault(p => p.DevicePath.Equals(form.NewProfile.DevicePath, StringComparison.OrdinalIgnoreCase));
                    if (existing != null)
                    {
                        existing.Name = form.NewProfile.Name;
                        existing.Speed = form.NewProfile.Speed;
                        existing.IconColor = form.NewProfile.IconColor;
                    }
                    else
                    {
                        _deviceProfiles.Add(form.NewProfile);
                    }

                    SaveDevices();
                    RebuildContextMenu();
                }
            }
        }

        static void ShowManageDevicesDialog()
        {
            using (var form = new ManageDevicesForm(_deviceProfiles))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    _deviceProfiles.Clear();
                    _deviceProfiles.AddRange(form.UpdatedProfiles);
                    SaveDevices();
                    RebuildContextMenu();
                }
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

        static Icon GetIconForProfile(DeviceProfile profile)
        {
            if (string.IsNullOrWhiteSpace(profile?.Name))
                return _defaultIcon;

            if (_deviceIcons.TryGetValue(profile.Name, out var existing))
                return existing;

            var label = profile.Name.Length <= 2 ? profile.Name : profile.Name.Substring(0, 2);
            var icon = CreateIcon(profile.IconColor, label);
            _deviceIcons[profile.Name] = icon;
            return icon;
        }

        class TrayApplicationContext : ApplicationContext
        {
            readonly RawInputWindow _window;

            public TrayApplicationContext()
            {
                InitializeDevices();
                InitializeTrayIcon();
                _window = new RawInputWindow(OnDeviceChanged);
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    _notifyIcon?.Dispose();
                    _menu?.Dispose();
                    foreach (var icon in _deviceIcons.Values)
                        icon?.Dispose();
                    _defaultIcon?.Dispose();
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
            public string DeviceName { get; }
            public int Speed { get; }

            public SpeedTag(string deviceName, int speed)
            {
                DeviceName = deviceName;
                Speed = speed;
            }
        }

        public class DeviceProfile
        {
            public DeviceProfile(string name, string devicePath, int speed, Color? iconColor = null)
            {
                Name = name;
                DevicePath = devicePath;
                Speed = speed;
                IconColor = iconColor ?? Color.Gray;
            }

            public string Name { get; set; }
            public string DevicePath { get; set; }
            public int Speed { get; set; }
            public Color IconColor { get; set; }
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
