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

        static AppliedSettings _lastAppliedSettings;
        static long _lastSwitchMs = 0;
        static readonly Stopwatch _sw = Stopwatch.StartNew();

        static NotifyIcon _notifyIcon;
        static ContextMenuStrip _menu;
        static readonly Dictionary<string, ToolStripMenuItem> _speedMenus = new Dictionary<string, ToolStripMenuItem>(StringComparer.OrdinalIgnoreCase);
        static ToolStripMenuItem _activeDeviceItem;
        static ToolStripMenuItem _startupItem;

        static Icon _defaultIcon;
        static readonly Dictionary<string, Icon> _profileIcons = new Dictionary<string, Icon>(StringComparer.OrdinalIgnoreCase);

        static string _activeDeviceKey = string.Empty;
        static string _lastAppliedDeviceKey = string.Empty;

        static readonly List<DeviceProfile> _deviceProfiles = new List<DeviceProfile>();
        static RawInputWindow _rawInputWindow;
        static int DefaultScrollLines => SystemInformation.MouseWheelScrollLines > 0 ? SystemInformation.MouseWheelScrollLines : 3;

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

            ApplyProfileSettings(profile, false);
        }

        static void InitializeDevices()
        {
            _deviceProfiles.Clear();
            _deviceProfiles.AddRange(LoadDevices());
        }

        static IEnumerable<DeviceProfile> LoadDevices()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(SettingsKeyPath, false))
            {
                string[] stored = key?.GetValue("Devices") as string[];
                if (stored != null)
                {
                    var devices = new List<DeviceProfile>();
                    foreach (string entry in stored)
                    {
                        if (string.IsNullOrWhiteSpace(entry))
                            continue;

                        string[] parts = entry.Split(new[] { '|' }, 4);
                        if (parts.Length < 3)
                            continue;

                        string name = parts[0];
                        string path = parts[1];
                        if (!int.TryParse(parts[2], out int speed))
                            continue;

                        if (string.IsNullOrWhiteSpace(path))
                            continue;

                        Color color = Color.Gray;
                        if (parts.Length >= 4 && int.TryParse(parts[3], out int argb))
                            color = Color.FromArgb(argb);

                        bool isEnabled = ReadBool(parts, 4, true);
                        bool applyAutomatically = ReadBool(parts, 5, true);
                        bool applyOnStartup = ReadBool(parts, 6, false);
                        bool enhancePrecision = ReadBool(parts, 7, true);
                        int scrollLines = ReadInt(parts, 8, DefaultScrollLines);
                        int scrollChars = ReadInt(parts, 9, 3);
                        bool swapButtons = ReadBool(parts, 10, false);
                        int doubleClickTime = ReadInt(parts, 11, SystemInformation.DoubleClickTime);

                        devices.Add(new DeviceProfile(
                            name,
                            path,
                            ClampSpeed(speed, speed),
                            color,
                            isEnabled,
                            applyAutomatically,
                            applyOnStartup,
                            enhancePrecision,
                            scrollLines,
                            scrollChars,
                            swapButtons,
                            ClampDoubleClickTime(doubleClickTime)));
                    }

                    return devices;
                }
            }

            return Enumerable.Empty<DeviceProfile>();
        }

        static void SaveDevices()
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(SettingsKeyPath))
            {
                string[] serialized = _deviceProfiles
                    .Select(p => string.Join("|", new[]
                    {
                        p.Name,
                        p.DevicePath,
                        ClampSpeed(p.Speed, p.Speed).ToString(),
                        p.IconColor.ToArgb().ToString(),
                        p.IsEnabled ? "1" : "0",
                        p.ApplyAutomatically ? "1" : "0",
                        p.ApplyOnStartup ? "1" : "0",
                        p.EnhancePointerPrecision ? "1" : "0",
                        p.ScrollLines.ToString(),
                        p.ScrollChars.ToString(),
                        p.SwapButtons ? "1" : "0",
                        ClampDoubleClickTime(p.DoubleClickTime).ToString()
                    }))
                    .ToArray();
                key.SetValue("Devices", serialized, RegistryValueKind.MultiString);
            }
        }

        static int ClampSpeed(int speed, int fallback)
        {
            if (speed < 1 || speed > 20)
                return fallback;

            return speed;
        }

        static int ClampDoubleClickTime(int time)
        {
            if (time < 200)
                return 200;
            if (time > 900)
                return 900;

            return time;
        }

        static int ReadInt(string[] parts, int index, int fallback)
        {
            if (parts.Length <= index)
                return fallback;

            return int.TryParse(parts[index], out int value) ? value : fallback;
        }

        static bool ReadBool(string[] parts, int index, bool fallback)
        {
            if (parts.Length <= index)
                return fallback;

            if (bool.TryParse(parts[index], out bool value))
                return value;

            if (int.TryParse(parts[index], out int intValue))
                return intValue != 0;

            return fallback;
        }

        static void UpdateActiveDevice(DeviceProfile profile, int speed)
        {
            _activeDeviceKey = profile.DevicePath ?? profile.Name;
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
            ClearProfileIcons();

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
            string label = profile.IsEnabled ? profile.Name : $"{profile.Name} (disabled)";
            var menu = new ToolStripMenuItem($"{label} Speed")
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

                var activeKey = profile.DevicePath ?? profile.Name;
                if (_activeDeviceKey.Equals(activeKey, StringComparison.OrdinalIgnoreCase))
                    ApplyProfileSettings(profile, true);
            }
        }

        static void UpdateMenuChecks(ToolStripMenuItem menu, int activeSpeed)
        {
            foreach (ToolStripMenuItem child in menu.DropDownItems)
            {
                child.Checked = child.Text == activeSpeed.ToString();
            }
        }

        static void ShowAddDeviceDialog()
        {
            using (var form = new AddDeviceForm(existingColors: _deviceProfiles.Select(p => p.IconColor)))
            {
                if (form.ShowDialog() == DialogResult.OK && form.NewProfile != null)
                {
                    var existing = _deviceProfiles.FirstOrDefault(p => p.DevicePath.Equals(form.NewProfile.DevicePath, StringComparison.OrdinalIgnoreCase));
                    if (existing != null)
                    {
                        existing.Name = form.NewProfile.Name;
                        existing.Speed = form.NewProfile.Speed;
                    }
                    else
                    {
                        _deviceProfiles.Add(form.NewProfile);
                    }

                    SaveDevices();
                    RebuildContextMenu();
                }
            }

            EnsureRawInputRegistration();
        }

        static void ShowManageDevicesDialog()
        {
            using (var form = new ManageDevicesForm(_deviceProfiles))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    _deviceProfiles.Clear();
                    _deviceProfiles.AddRange(form.Devices);
                    SaveDevices();
                    RebuildContextMenu();
                }
            }
        }

        static void ApplyStartupProfile()
        {
            var profile = _deviceProfiles.FirstOrDefault(p => p.IsEnabled && p.ApplyOnStartup);
            if (profile == null)
                return;

            ApplyProfileSettings(profile, true);
        }

        static void ApplyProfileSettings(DeviceProfile profile, bool force)
        {
            if (profile == null || !profile.IsEnabled)
                return;

            if (!profile.ApplyAutomatically && !force)
                return;

            long now = _sw.ElapsedMilliseconds;
            var settings = new AppliedSettings(profile);
            bool deviceChanged = !string.Equals(_lastAppliedDeviceKey, settings.DeviceKey, StringComparison.OrdinalIgnoreCase);
            bool settingsChanged = _lastAppliedSettings == null || !_lastAppliedSettings.Equals(settings);
            bool canSwitch = !settingsChanged || force || (now - _lastSwitchMs) >= MIN_SWITCH_MS;

            if (settingsChanged && canSwitch)
            {
                NativeMethods.SetMouseSpeed(settings.Speed);
                SetEnhancePointerPrecision(settings.EnhancePointerPrecision);
                SetScrollSettings(settings.ScrollLines, settings.ScrollChars);
                SetSwapButtons(settings.SwapButtons);
                SetDoubleClickTime(settings.DoubleClickTime);

                _lastAppliedSettings = settings;
                _lastAppliedDeviceKey = settings.DeviceKey;
                _lastSwitchMs = now;
            }

            if ((settingsChanged && canSwitch) || deviceChanged)
            {
                UpdateActiveDevice(profile, settings.Speed);
                if (deviceChanged)
                    _lastAppliedDeviceKey = settings.DeviceKey;
                if (_lastAppliedSettings == null)
                    _lastAppliedSettings = settings;
            }
        }

        static void SetEnhancePointerPrecision(bool enabled)
        {
            int[] values = enabled ? new[] { 6, 10, 1 } : new[] { 0, 0, 0 };
            NativeMethods.SystemParametersInfo(NativeMethods.SPI_SETMOUSE, 0, values, NativeMethods.SPIF_SENDCHANGE);
        }

        static void SetScrollSettings(int scrollLines, int scrollChars)
        {
            int lines = Math.Max(0, scrollLines);
            int chars = Math.Max(0, scrollChars);
            NativeMethods.SystemParametersInfo(NativeMethods.SPI_SETWHEELSCROLLLINES, 0, (IntPtr)lines, NativeMethods.SPIF_SENDCHANGE);
            NativeMethods.SystemParametersInfo(NativeMethods.SPI_SETWHEELSCROLLCHARS, 0, (IntPtr)chars, NativeMethods.SPIF_SENDCHANGE);
        }

        static void SetSwapButtons(bool swapButtons)
        {
            NativeMethods.SwapMouseButton(swapButtons ? 1 : 0);
        }

        static void SetDoubleClickTime(int time)
        {
            NativeMethods.SetDoubleClickTime((uint)ClampDoubleClickTime(time));
        }

        static Icon GetIconForProfile(DeviceProfile profile)
        {
            if (profile == null)
                return _defaultIcon;

            string key = profile.DevicePath ?? profile.Name;
            if (string.IsNullOrWhiteSpace(key))
                return _defaultIcon;

            if (_profileIcons.TryGetValue(key, out var icon))
                return icon;

            string label = GetIconLabel(profile.Name);
            icon = CreateIcon(profile.IconColor, label);
            _profileIcons[key] = icon;
            return icon;
        }

        static string GetIconLabel(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "MM";

            string trimmed = new string(name.Where(char.IsLetterOrDigit).ToArray());
            if (trimmed.Length == 0)
                trimmed = name.Trim();

            if (trimmed.Length > 2)
                return trimmed.Substring(0, 2).ToUpperInvariant();

            return trimmed.ToUpperInvariant();
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

        static void ClearProfileIcons()
        {
            foreach (var icon in _profileIcons.Values)
            {
                icon?.Dispose();
            }

            _profileIcons.Clear();
        }

        class TrayApplicationContext : ApplicationContext
        {
            public TrayApplicationContext()
            {
                InitializeDevices();
                InitializeTrayIcon();
                ApplyStartupProfile();
                _rawInputWindow = new RawInputWindow(OnDeviceChanged);
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    _notifyIcon?.Dispose();
                    _menu?.Dispose();
                    _defaultIcon?.Dispose();
                    ClearProfileIcons();
                    _rawInputWindow?.Dispose();
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

                RegisterForRawInput();
            }

            public void RegisterForRawInput()
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

        public static void EnsureRawInputRegistration()
        {
            _rawInputWindow?.RegisterForRawInput();
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
            public DeviceProfile(string name, string devicePath, int speed) : this(
                name,
                devicePath,
                speed,
                Color.Gray,
                true,
                true,
                false,
                true,
                DefaultScrollLines,
                3,
                false,
                SystemInformation.DoubleClickTime)
            {
            }

            public DeviceProfile(
                string name,
                string devicePath,
                int speed,
                Color iconColor,
                bool isEnabled,
                bool applyAutomatically,
                bool applyOnStartup,
                bool enhancePointerPrecision,
                int scrollLines,
                int scrollChars,
                bool swapButtons,
                int doubleClickTime)
            {
                Name = name;
                DevicePath = devicePath;
                Speed = speed;
                IconColor = iconColor;
                IsEnabled = isEnabled;
                ApplyAutomatically = applyAutomatically;
                ApplyOnStartup = applyOnStartup;
                EnhancePointerPrecision = enhancePointerPrecision;
                ScrollLines = scrollLines;
                ScrollChars = scrollChars;
                SwapButtons = swapButtons;
                DoubleClickTime = doubleClickTime;
            }

            public string Name { get; set; }
            public string DevicePath { get; set; }
            public int Speed { get; set; }
            public Color IconColor { get; set; }
            public bool IsEnabled { get; set; }
            public bool ApplyAutomatically { get; set; }
            public bool ApplyOnStartup { get; set; }
            public bool EnhancePointerPrecision { get; set; }
            public int ScrollLines { get; set; }
            public int ScrollChars { get; set; }
            public bool SwapButtons { get; set; }
            public int DoubleClickTime { get; set; }
        }

        class AppliedSettings
        {
            public AppliedSettings(DeviceProfile profile)
            {
                Speed = profile.Speed;
                EnhancePointerPrecision = profile.EnhancePointerPrecision;
                ScrollLines = profile.ScrollLines;
                ScrollChars = profile.ScrollChars;
                SwapButtons = profile.SwapButtons;
                DoubleClickTime = profile.DoubleClickTime;
                DeviceKey = profile.DevicePath ?? profile.Name ?? string.Empty;
            }

            public int Speed { get; }
            public bool EnhancePointerPrecision { get; }
            public int ScrollLines { get; }
            public int ScrollChars { get; }
            public bool SwapButtons { get; }
            public int DoubleClickTime { get; }
            public string DeviceKey { get; }

            public bool Equals(AppliedSettings other)
            {
                if (other == null)
                    return false;

                return Speed == other.Speed
                    && EnhancePointerPrecision == other.EnhancePointerPrecision
                    && ScrollLines == other.ScrollLines
                    && ScrollChars == other.ScrollChars
                    && SwapButtons == other.SwapButtons
                    && DoubleClickTime == other.DoubleClickTime
                    && string.Equals(DeviceKey, other.DeviceKey, StringComparison.OrdinalIgnoreCase);
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
            public const uint SPI_SETMOUSE = 0x0004;
            public const uint SPI_SETWHEELSCROLLLINES = 0x0069;
            public const uint SPI_SETWHEELSCROLLCHARS = 0x006D;
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
            public static extern bool SystemParametersInfo(uint uiAction, uint uiParam, int[] pvParam, uint fWinIni);

            [DllImport("user32.dll")]
            public static extern int SwapMouseButton(int fSwap);

            [DllImport("user32.dll")]
            public static extern bool SetDoubleClickTime(uint ms);

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
