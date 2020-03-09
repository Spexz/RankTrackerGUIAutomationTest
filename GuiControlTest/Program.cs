using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GuiControlTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var t = new Thread((ThreadStart)(() => {
                RankTrackerProcess();
            }));

            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();

            Console.ReadLine();
        }


        public static void RankTrackerProcess()
        {
            string keywords = @"
            amritsar hotel room\n
            amritsar hotel rooms\n
            hotel room amritsar\n
            hotel rooms amritsar\n
            hotel room in amritsar\n
            ";



            var handles = FindWindowsWithText("Rank Tracker v");
            Console.WriteLine(handles.Count());

            IntPtr target_hwnd = handles.First();



            // Get the target window's handle.
            //IntPtr target_hwnd = FindWindowByCaption(IntPtr.Zero, "test - Rank Tracker v8.32.9");


            if (target_hwnd == IntPtr.Zero)
            {
                MessageBox.Show(
                    "Could not find a window with the title \"" +
                    "Notepad" + "\"");
                return;
            }

            ShowWindowAsync(new HandleRef(null, target_hwnd), SW_RESTORE);

            SetForegroundWindow(target_hwnd);

            // Set the window's position.
            int width = 1100;
            int height = 630;
            int x = 10;
            int y = 10;
            SetWindowPos(target_hwnd, IntPtr.Zero,
                x, y, width, height, 0);

            Clipboard.SetText(keywords);


            //ClickOnPointTool.ClickOnPoint(target_hwnd, new Point(382, 146));
            LeftMouseClick(382, 146);
            Thread.Sleep(1000);

            LeftMouseClick(335, 303);
            Thread.Sleep(1000);

            PressPaste();

            // Next
            LeftMouseClick(686, 543);
            Thread.Sleep(2000);


            // Finish
            LeftMouseClick(770, 542);
            //Thread.Sleep(2000);

            string pColor = "#61A032";

            var result = LoopUntil(() => {
                var c = HexConverter(GetColorAt(180, 596)); //184, 596
                if (c.Equals(pColor))
                    return true;
                return false;
            }, TimeSpan.FromSeconds(30)); // should be way longer in release

            Thread.Sleep(1000); // should be longer in release

            //if (result)
            //{
                // Download
                LeftMouseClick(1057, 145);
                Thread.Sleep(1000);
            //}

            // Save csv file
            //Keywords & rankings - test.csv
            var csvFileName = $"keywords_rankings_{DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond}.csv";
            SaveCsv(csvFileName);
            Thread.Sleep(1000);

            // Click in the keywords table and Delete keywords
            DeleteKeywords();

            //Color color = GetColorAt(686, 543);
            //Console.WriteLine(HexConverter(color));

            //Thread.Sleep(10000);
        }


        public static bool LoopUntil(Func<bool> task, TimeSpan timeSpan)
        {
            bool success = false;
            int elapsed = 0;
            while ((!success) && (elapsed < timeSpan.TotalMilliseconds))
            {
                Thread.Sleep(100);
                elapsed += 100;
                success = task();
            }
            return success;
        }


        private static String HexConverter(System.Drawing.Color c)
        {
            return "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        }

        private static String RGBConverter(System.Drawing.Color c)
        {
            return "RGB(" + c.R.ToString() + "," + c.G.ToString() + "," + c.B.ToString() + ")";
        }


        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetDesktopWindow();
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetWindowDC(IntPtr window);
        [DllImport("gdi32.dll", SetLastError = true)]
        public static extern uint GetPixel(IntPtr dc, int x, int y);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern int ReleaseDC(IntPtr window, IntPtr dc);

        public static Color GetColorAt(int x, int y)
        {
            IntPtr desk = GetDesktopWindow();
            IntPtr dc = GetWindowDC(desk);
            int a = (int)GetPixel(dc, x, y);
            ReleaseDC(desk, dc);
            return Color.FromArgb(255, (a >> 0) & 0xff, (a >> 8) & 0xff, (a >> 16) & 0xff);
        }






        [DllImport("user32.dll", SetLastError = true)]
        static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

        public const int KEYEVENTF_KEYDOWN = 0x0000; // New definition
        public const int KEYEVENTF_EXTENDEDKEY = 0x0001; //Key down flag
        public const int KEYEVENTF_KEYUP = 0x0002; //Key up flag
        public const int VK_LCONTROL = 0xA2; //Left Control key code
        public const int VK_MENU = 0x12; //ALT key code
        public const int VK_DELETE = 0x2E; //DEL key code
        public const int VK_RETURN = 0x0D; //ENTER key code

        public const int A = 0x41; //A key code
        public const int C = 0x43; //C key code
        public const int N = 0x4E; //N key code
        public const int V = 0x56; //V key code
        public const int Y = 0x59; //Y key code

        public static void PressPaste()
        {
            // Hold Control down and press V
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(V, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(V, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYUP, 0);
        }

        public static void SaveCsv(string filename)
        {
            AltN(); Thread.Sleep(500);
            CtrlA(); Thread.Sleep(500);

            Clipboard.SetText(filename);
            PressPaste(); Thread.Sleep(500);
            PressEnter(); Thread.Sleep(500);
            // Do not open folder
            AltN(); Thread.Sleep(500);
        }

        public static void DeleteKeywords()
        {
            LeftMouseClick(600, 222); Thread.Sleep(500);
            CtrlA(); Thread.Sleep(500);
            PressDelete(); Thread.Sleep(500);
            AltY(); Thread.Sleep(500);
        }

        public static void AltN()
        {
            // Hold ALT down and press N
            keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(N, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(N, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
        }

        public static void AltY()
        {
            // Hold ALT down and press N
            keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(Y, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(Y, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
        }

        public static void CtrlA()
        {
            // Hold Ctrl down and press A
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(A, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(A, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYUP, 0);
        }

        public static void PressKeys()
        {
            // Hold Control down and press A
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(A, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(A, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYUP, 0);

            // Hold Control down and press C
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(C, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(C, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYUP, 0);
        }


        public static void PressEnter()
        {
            keybd_event(VK_RETURN, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0);
        }

        public static void PressDelete()
        {
            keybd_event(VK_DELETE, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(VK_DELETE, 0, KEYEVENTF_KEYUP, 0);
        }



        //This is a replacement for Cursor.Position in WinForms
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool SetCursorPos(int x, int y);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        public const int MOUSEEVENTF_LEFTDOWN = 0x02;
        public const int MOUSEEVENTF_LEFTUP = 0x04;

        //This simulates a left mouse click
        public static void LeftMouseClick(int xpos, int ypos)
        {
            SetCursorPos(xpos, ypos);
            mouse_event(MOUSEEVENTF_LEFTDOWN, xpos, ypos, 0, 0);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTUP, xpos, ypos, 0, 0);
        }






        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder strText, int maxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        // Delegate to filter which windows to include 
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        /// <summary> Get the text for the window pointed to by hWnd </summary>
        public static string GetWindowText(IntPtr hWnd)
        {
            int size = GetWindowTextLength(hWnd);
            if (size > 0)
            {
                var builder = new StringBuilder(size + 1);
                GetWindowText(hWnd, builder, builder.Capacity);
                return builder.ToString();
            }

            return String.Empty;
        }

        /// <summary> Find all windows that match the given filter </summary>
        /// <param name="filter"> A delegate that returns true for windows
        ///    that should be returned and false for windows that should
        ///    not be returned </param>
        public static IEnumerable<IntPtr> FindWindows(EnumWindowsProc filter)
        {
            IntPtr found = IntPtr.Zero;
            List<IntPtr> windows = new List<IntPtr>();

            EnumWindows(delegate (IntPtr wnd, IntPtr param)
            {
                if (filter(wnd, param))
                {
                    // only add the windows that pass the filter
                    windows.Add(wnd);
                }

                // but return true here so that we iterate all windows
                return true;
            }, IntPtr.Zero);

            return windows;
        }

        /// <summary> Find all windows that contain the given title text </summary>
        /// <param name="titleText"> The text that the window title must contain. </param>
        public static IEnumerable<IntPtr> FindWindowsWithText(string titleText)
        {
            return FindWindows(delegate (IntPtr wnd, IntPtr param)
            {
                return GetWindowText(wnd).Contains(titleText);
            });
        }










        // Define the FindWindow API function.
        [DllImport("user32.dll", EntryPoint = "FindWindow",
            SetLastError = true)]
        static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly,
            string lpWindowName);

        [DllImport("user32.dll")]
        public static extern bool ShowWindowAsync(HandleRef hWnd, int nCmdShow);

        public const int SW_RESTORE = 9;

        // Define the SetWindowPos API function.
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetWindowPos(IntPtr hWnd,
            IntPtr hWndInsertAfter, int X, int Y, int cx, int cy,
            SetWindowPosFlags uFlags);

        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        // Define the SetWindowPosFlags enumeration.
        [Flags()]
        private enum SetWindowPosFlags : uint
        {
            SynchronousWindowPosition = 0x4000,
            DeferErase = 0x2000,
            DrawFrame = 0x0020,
            FrameChanged = 0x0020,
            HideWindow = 0x0080,
            DoNotActivate = 0x0010,
            DoNotCopyBits = 0x0100,
            IgnoreMove = 0x0002,
            DoNotChangeOwnerZOrder = 0x0200,
            DoNotRedraw = 0x0008,
            DoNotReposition = 0x0200,
            DoNotSendChangingEvent = 0x0400,
            IgnoreResize = 0x0001,
            IgnoreZOrder = 0x0004,
            ShowWindow = 0x0040,
        }
    }




    public class ClickOnPointTool
    {

        [DllImport("user32.dll")]
        static extern bool ClientToScreen(IntPtr hWnd, ref Point lpPoint);

        [DllImport("user32.dll")]
        internal static extern uint SendInput(uint nInputs, [MarshalAs(UnmanagedType.LPArray), In] INPUT[] pInputs, int cbSize);

#pragma warning disable 649
        internal struct INPUT
        {
            public UInt32 Type;
            public MOUSEKEYBDHARDWAREINPUT Data;
        }

        [StructLayout(LayoutKind.Explicit)]
        internal struct MOUSEKEYBDHARDWAREINPUT
        {
            [FieldOffset(0)]
            public MOUSEINPUT Mouse;
        }

        internal struct MOUSEINPUT
        {
            public Int32 X;
            public Int32 Y;
            public UInt32 MouseData;
            public UInt32 Flags;
            public UInt32 Time;
            public IntPtr ExtraInfo;
        }

#pragma warning restore 649


        public static void ClickOnPoint(IntPtr wndHandle, Point clientPoint)
        {
            var oldPos = Cursor.Position;

            /// get screen coordinates
            ClientToScreen(wndHandle, ref clientPoint);

            /// set cursor on coords, and press mouse
            Cursor.Position = new Point(clientPoint.X, clientPoint.Y);

            var inputMouseDown = new INPUT();
            inputMouseDown.Type = 0; /// input type mouse
            inputMouseDown.Data.Mouse.Flags = 0x0002; /// left button down

            var inputMouseUp = new INPUT();
            inputMouseUp.Type = 0; /// input type mouse
            inputMouseUp.Data.Mouse.Flags = 0x0004; /// left button up

            var inputs = new INPUT[] { inputMouseDown, inputMouseUp };
            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));

            /// return mouse 
            Cursor.Position = oldPos;
        }

    }
}
