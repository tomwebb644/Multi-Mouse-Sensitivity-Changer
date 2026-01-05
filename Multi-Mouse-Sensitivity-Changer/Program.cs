using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
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

        static int MIN_SWITCH_MS = 200;
        // ======================================

        static int _lastSpeed = -1;
        static long _lastSwitchMs = 0;
        static readonly Stopwatch _sw = Stopwatch.StartNew();

        static NotifyIcon _notifyIcon;
        static ContextMenuStrip _menu;
        static ToolStripMenuItem _mouseSpeedMenu;
        static ToolStripMenuItem _x8SpeedMenu;
        static ToolStripMenuItem _activeDeviceItem;
        static ToolStripMenuItem _startupItem;

        static Icon _mouseIcon;
        static Icon _x8Icon;
        static Icon _defaultIcon;

        static string _activeDeviceKey = string.Empty;

        const string X8_KEY = "X8";
        const string MOUSE_KEY = "Mouse";

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

            string deviceKey =
                devicePath.Equals(X8_DEVICE_PATH, StringComparison.OrdinalIgnoreCase) ? X8_KEY :
                devicePath.Equals(MOUSE_DEVICE_PATH, StringComparison.OrdinalIgnoreCase) ? MOUSE_KEY :
                string.Empty;

            if (string.IsNullOrEmpty(deviceKey))
                return;

            int targetSpeed = deviceKey == X8_KEY ? X8_SPEED : MOUSE_SPEED;
            long now = _sw.ElapsedMilliseconds;
            if (targetSpeed != _lastSpeed && (now - _lastSwitchMs) >= MIN_SWITCH_MS)
            {
                NativeMethods.SetMouseSpeed(targetSpeed);
                _lastSpeed = targetSpeed;
                _lastSwitchMs = now;
                UpdateActiveDevice(deviceKey, targetSpeed);
            }
        }

        static void UpdateActiveDevice(string deviceKey, int speed)
        {
            _activeDeviceKey = deviceKey;
            _activeDeviceItem.Text = deviceKey == X8_KEY
                ? $"Active: X8 (speed {speed})"
                : $"Active: Mouse (speed {speed})";

            if (deviceKey == X8_KEY)
            {
                _notifyIcon.Icon = _x8Icon;
                _notifyIcon.Text = "X8 speed " + speed;
            }
            else
            {
                _notifyIcon.Icon = _mouseIcon;
                _notifyIcon.Text = "Mouse speed " + speed;
            }
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

            _mouseSpeedMenu = BuildSpeedMenu("Mouse Speed", MOUSE_KEY);
            _x8SpeedMenu = BuildSpeedMenu("X8 Speed", X8_KEY);

            _menu.Items.Add(_mouseSpeedMenu);
            _menu.Items.Add(_x8SpeedMenu);
            _menu.Items.Add(new ToolStripSeparator());

            _menu.Items.Add(new ToolStripMenuItem("Mouse device path") { Enabled = false, ToolTipText = MOUSE_DEVICE_PATH });
            _menu.Items.Add(new ToolStripMenuItem(MOUSE_DEVICE_PATH) { Enabled = false });
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(new ToolStripMenuItem("X8 device path") { Enabled = false, ToolTipText = X8_DEVICE_PATH });
            _menu.Items.Add(new ToolStripMenuItem(X8_DEVICE_PATH) { Enabled = false });
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

            _notifyIcon = new NotifyIcon
            {
                Visible = true,
                Icon = _defaultIcon,
                Text = "Multi-Mouse Sensitivity Changer",
                ContextMenuStrip = _menu
            };
        }

        static ToolStripMenuItem BuildSpeedMenu(string label, string deviceKey)
        {
            var menu = new ToolStripMenuItem(label);
            for (int i = 1; i <= 20; i++)
            {
                var item = new ToolStripMenuItem(i.ToString())
                {
                    Tag = new SpeedTag(deviceKey, i),
                    Checked = (deviceKey == MOUSE_KEY && i == MOUSE_SPEED) || (deviceKey == X8_KEY && i == X8_SPEED)
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
                if (tag.DeviceKey == MOUSE_KEY)
                {
                    MOUSE_SPEED = tag.Speed;
                    UpdateMenuChecks(_mouseSpeedMenu, tag.Speed);
                }
                else if (tag.DeviceKey == X8_KEY)
                {
                    X8_SPEED = tag.Speed;
                    UpdateMenuChecks(_x8SpeedMenu, tag.Speed);
                }

                if (_activeDeviceKey == tag.DeviceKey)
                {
                    NativeMethods.SetMouseSpeed(tag.Speed);
                    _lastSpeed = tag.Speed;
                    UpdateActiveDevice(tag.DeviceKey, tag.Speed);
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

        class TrayApplicationContext : ApplicationContext
        {
            readonly RawInputWindow _window;

            public TrayApplicationContext()
            {
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
            public string DeviceKey { get; }
            public int Speed { get; }

            public SpeedTag(string deviceKey, int speed)
            {
                DeviceKey = deviceKey;
                Speed = speed;
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
