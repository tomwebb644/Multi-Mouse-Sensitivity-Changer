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
        const int DefaultScrollLines = 3;
        const int DefaultScrollChars = 3;
        // ======================================

        static int _lastSpeed = -1;
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

        static readonly List<DeviceProfile> _deviceProfiles = new List<DeviceProfile>();
        static RawInputWindow _rawInputWindow;

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

            if (!profile.Enabled || !profile.AutoApply)
                return;

            int targetSpeed = profile.Speed;
            long now = _sw.ElapsedMilliseconds;

            string deviceKey = profile.DevicePath ?? profile.Name;
            bool deviceChanged = !string.Equals(_activeDeviceKey, deviceKey, StringComparison.OrdinalIgnoreCase);
            bool speedChanged = targetSpeed != _lastSpeed;
            bool canSwitchSpeed = !speedChanged || (now - _lastSwitchMs) >= MIN_SWITCH_MS;

            if (deviceChanged)
            {
                ApplyProfileSettings(profile, canSwitchSpeed);
                if (canSwitchSpeed)
                {
                    _lastSpeed = targetSpeed;
                    _lastSwitchMs = now;
                }
                UpdateActiveDevice(profile, targetSpeed);
                return;
            }

            if (speedChanged && canSwitchSpeed)
            {
                ApplyProfileSettings(profile, true);
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

                        string[] parts = entry.Split(new[] { '|' }, 11);
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

                        bool enabled = parts.Length >= 5 ? ParseBool(parts[4], true) : true;
                        bool enhancePrecision = parts.Length >= 6 ? ParseBool(parts[5], true) : true;
                        int verticalScroll = parts.Length >= 7 && int.TryParse(parts[6], out int vScroll) ? ClampScroll(vScroll, DefaultScrollLines) : DefaultScrollLines;
                        int horizontalScroll = parts.Length >= 8 && int.TryParse(parts[7], out int hScroll) ? ClampScroll(hScroll, DefaultScrollChars) : DefaultScrollChars;
                        bool swapButtons = parts.Length >= 9 ? ParseBool(parts[8], false) : false;
                        bool applyOnStartup = parts.Length >= 10 ? ParseBool(parts[9], false) : false;
                        bool autoApply = parts.Length >= 11 ? ParseBool(parts[10], true) : true;

                        devices.Add(new DeviceProfile(
                            name,
                            path,
                            ClampSpeed(speed, speed),
                            color,
                            enhancePrecision,
                            verticalScroll,
                            horizontalScroll,
                            swapButtons,
                            enabled,
                            autoApply,
                            applyOnStartup));
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
                        p.Enabled.ToString(),
                        p.EnhancePointerPrecision.ToString(),
                        ClampScroll(p.VerticalScrollLines, DefaultScrollLines).ToString(),
                        ClampScroll(p.HorizontalScrollChars, DefaultScrollChars).ToString(),
                        p.SwapButtons.ToString(),
                        p.ApplyOnStartup.ToString(),
                        p.AutoApply.ToString()
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

        static int ClampScroll(int value, int fallback)
        {
            if (value < 0 || value > 100)
                return fallback;

            return value;
        }

        static bool ParseBool(string value, bool fallback)
        {
            if (bool.TryParse(value, out bool parsed))
                return parsed;

            return fallback;
        }

        static void UpdateActiveDevice(DeviceProfile profile, int speed)
        {
            _activeDeviceKey = profile.DevicePath ?? profile.Name;
            _activeDeviceItem.Text = $"Active: {profile.Name} (speed {speed})";

            _notifyIcon.Icon = GetIconForProfile(profile);
            _notifyIcon.Text = $"{profile.Name} speed {speed}";
        }

        static void ApplyProfileSettings(DeviceProfile profile, bool applySpeed)
        {
            if (profile == null)
                return;

            if (applySpeed)
                NativeMethods.SetMouseSpeed(profile.Speed);

            SetPointerPrecision(profile.EnhancePointerPrecision);
            SetScrollLines(profile.VerticalScrollLines);
            SetScrollChars(profile.HorizontalScrollChars);
            SetSwapButtons(profile.SwapButtons);
        }

        static void SetPointerPrecision(bool enabled)
        {
            int[] mouseParams = new int[3];
            if (!NativeMethods.SystemParametersInfo(NativeMethods.SPI_GETMOUSE, 0, mouseParams, 0))
            {
                mouseParams[0] = 6;
                mouseParams[1] = 10;
            }

            mouseParams[2] = enabled ? 1 : 0;
            NativeMethods.SystemParametersInfo(NativeMethods.SPI_SETMOUSE, 0, mouseParams, NativeMethods.SPIF_SENDCHANGE);
        }

        static void SetScrollLines(int lines)
        {
            NativeMethods.SystemParametersInfo(NativeMethods.SPI_SETWHEELSCROLLLINES, (uint)ClampScroll(lines, DefaultScrollLines), IntPtr.Zero, NativeMethods.SPIF_SENDCHANGE);
        }

        static void SetScrollChars(int chars)
        {
            NativeMethods.SystemParametersInfo(NativeMethods.SPI_SETWHEELSCROLLCHARS, (uint)ClampScroll(chars, DefaultScrollChars), IntPtr.Zero, NativeMethods.SPIF_SENDCHANGE);
        }

        static void SetSwapButtons(bool swap)
        {
            NativeMethods.SystemParametersInfo(NativeMethods.SPI_SETMOUSEBUTTONSWAP, swap ? 1u : 0u, IntPtr.Zero, NativeMethods.SPIF_SENDCHANGE);
        }

        static void ClearStartupProfiles()
        {
            foreach (var profile in _deviceProfiles)
            {
                profile.ApplyOnStartup = false;
            }
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

        static void ApplyStartupProfile()
        {
            var profile = _deviceProfiles.FirstOrDefault(p => p.ApplyOnStartup && p.Enabled);
            if (profile == null)
                return;

            ApplyProfileSettings(profile, true);
            _lastSpeed = profile.Speed;
            _lastSwitchMs = _sw.ElapsedMilliseconds;
            UpdateActiveDevice(profile, profile.Speed);
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

                var activeKey = profile.DevicePath ?? profile.Name;
                if (_activeDeviceKey.Equals(activeKey, StringComparison.OrdinalIgnoreCase))
                {
                    ApplyProfileSettings(profile, true);
                    _lastSpeed = tag.Speed;
                    _lastSwitchMs = _sw.ElapsedMilliseconds;
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

        static void ShowAddDeviceDialog()
        {
            using (var form = new AddDeviceForm(existingColors: _deviceProfiles.Select(p => p.IconColor)))
            {
                if (form.ShowDialog() == DialogResult.OK && form.NewProfile != null)
                {
                    if (form.NewProfile.ApplyOnStartup)
                        ClearStartupProfiles();

                    var existing = _deviceProfiles.FirstOrDefault(p => p.DevicePath.Equals(form.NewProfile.DevicePath, StringComparison.OrdinalIgnoreCase));
                    if (existing != null)
                    {
                        existing.Name = form.NewProfile.Name;
                        existing.Speed = form.NewProfile.Speed;
                        existing.IconColor = form.NewProfile.IconColor;
                        existing.EnhancePointerPrecision = form.NewProfile.EnhancePointerPrecision;
                        existing.VerticalScrollLines = form.NewProfile.VerticalScrollLines;
                        existing.HorizontalScrollChars = form.NewProfile.HorizontalScrollChars;
                        existing.SwapButtons = form.NewProfile.SwapButtons;
                        existing.Enabled = form.NewProfile.Enabled;
                        existing.AutoApply = form.NewProfile.AutoApply;
                        existing.ApplyOnStartup = form.NewProfile.ApplyOnStartup;
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
            public DeviceProfile(string name, string devicePath, int speed) : this(name, devicePath, speed, Color.Gray)
            {
            }

            public DeviceProfile(string name, string devicePath, int speed, Color iconColor)
                : this(name, devicePath, speed, iconColor, true, DefaultScrollLines, DefaultScrollChars, false, true, true, false)
            {
            }

            public DeviceProfile(
                string name,
                string devicePath,
                int speed,
                Color iconColor,
                bool enhancePointerPrecision,
                int verticalScrollLines,
                int horizontalScrollChars,
                bool swapButtons,
                bool enabled,
                bool autoApply,
                bool applyOnStartup)
            {
                Name = name;
                DevicePath = devicePath;
                Speed = speed;
                IconColor = iconColor;
                EnhancePointerPrecision = enhancePointerPrecision;
                VerticalScrollLines = verticalScrollLines;
                HorizontalScrollChars = horizontalScrollChars;
                SwapButtons = swapButtons;
                Enabled = enabled;
                AutoApply = autoApply;
                ApplyOnStartup = applyOnStartup;
            }

            public string Name { get; set; }
            public string DevicePath { get; set; }
            public int Speed { get; set; }
            public Color IconColor { get; set; }
            public bool EnhancePointerPrecision { get; set; }
            public int VerticalScrollLines { get; set; }
            public int HorizontalScrollChars { get; set; }
            public bool SwapButtons { get; set; }
            public bool Enabled { get; set; }
            public bool AutoApply { get; set; }
            public bool ApplyOnStartup { get; set; }
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
            public const uint SPI_GETMOUSE = 0x0003;
            public const uint SPI_SETMOUSE = 0x0004;
            public const uint SPI_SETMOUSESPEED = 0x0071;
            public const uint SPI_SETMOUSEBUTTONSWAP = 0x0021;
            public const uint SPI_SETWHEELSCROLLLINES = 0x0069;
            public const uint SPI_SETWHEELSCROLLCHARS = 0x006C;
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
