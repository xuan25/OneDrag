using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace OneDrag
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        //定义API函数
        [DllImport("kernel32.dll")]
        static extern uint SetThreadExecutionState(uint esFlags);
        const uint ES_SYSTEM_REQUIRED = 0x00000001;
        const uint ES_DISPLAY_REQUIRED = 0x00000002;
        const uint ES_CONTINUOUS = 0x80000000;

        //Sets window attributes
        [DllImport("USER32.DLL")]
        public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        //Gets window attributes
        [DllImport("USER32.DLL")]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        //assorted constants needed
        public static int GWL_STYLE = -16;
        public static int WS_CHILD = 0x40000000; //child window
        public static int WS_BORDER = 0x00800000; //window with border
        public static int WS_DLGFRAME = 0x00400000; //window with double border but no title
        public static int WS_CAPTION = WS_BORDER | WS_DLGFRAME; //window with a title bar

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            IntPtr windowHandle = new WindowInteropHelper(this).Handle;
            int style = GetWindowLong(windowHandle, GWL_STYLE);
            SetWindowLong(windowHandle, GWL_STYLE, (style | WS_CAPTION));
        }

        public byte[] StreamToBytes(Stream stream)
        {
            byte[] bytes = new byte[stream.Length];
            stream.Read(bytes, 0, bytes.Length);
            // 设置当前流的位置为流的开始
            stream.Seek(0, SeekOrigin.Begin);
            return bytes;
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        /// <param name="path">释放路径</param>
        /// <param name="source">释放源</param>
        private void ExportResource(string path, string source)
        {
            if (!File.Exists(path))
            {
                //释放资源到磁盘
                String projectName = Assembly.GetExecutingAssembly().GetName().Name.ToString();
                Stream gzStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(projectName + "." + source);
                GZipStream stream = new GZipStream(gzStream, CompressionMode.Decompress);
                FileStream decompressedFile = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write);
                stream.CopyTo(decompressedFile);
                decompressedFile.Close();
                stream.Close();
                gzStream.Close();
            }
        }

        private string VerifyResource(string source)
        {
            String projectName = Assembly.GetExecutingAssembly().GetName().Name.ToString();
            Stream gzStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(projectName + "." + source);
            string md5 = GetMD5(StreamToBytes(gzStream));
            gzStream.Close();
            return md5;
        }

        /// <summary>
        /// MD5验证
        /// </summary>
        /// <param name="fromData">输入</param>
        /// <returns></returns>
        public static string GetMD5(byte[] fromData)
        {
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] targetData = md5.ComputeHash(fromData);
            string byte2String = null;

            for (int i = 0; i < targetData.Length; i++)
            {
                byte2String += targetData[i].ToString("x");
            }

            return byte2String;
        }

        public MainWindow()
        {
            InitializeComponent();

            Blur.Interop.WindowBlur.SetIsEnabled(this, true);

            Process currentProcess = Process.GetCurrentProcess();
            currentProcess.PriorityClass = ProcessPriorityClass.RealTime;
            Thread primaryThread = Thread.CurrentThread;
            primaryThread.Priority = ThreadPriority.Highest;
            //UI初始化
            AboutBox.Text = "by  瑄  |  Version 1.1";
            Description.Visibility = Visibility.Hidden;
            Initialization.Visibility = Visibility.Hidden;
            ES = ES_REQUIRE.OFF;

            //校验UI
            //Console.WriteLine(GetMD5(Encoding.Unicode.GetBytes(TitleBox.Text)));
            if (GetMD5(Encoding.Unicode.GetBytes(TitleBox.Text)) != "151451921c456c1cc152e86e52da6651")
            {
                ProcessLog("VERIFICATION FAILED at  " + DateTime.Now.ToString());
                Environment.Exit(0);
            }
            //Console.WriteLine(GetMD5(Encoding.Unicode.GetBytes(AboutBox.Text)));
            if (GetMD5(Encoding.Unicode.GetBytes(AboutBox.Text)) != "be8afb9eb9e067e97e289e9e3f3172")
            {
                ProcessLog("VERIFICATION FAILED at  " + DateTime.Now.ToString());
                Environment.Exit(0);
            }

            //功能实例化
            P_Vocal = new ProcessClass(null, ProcessClass.ProcessMode.Vocal, 0, 0, "");
            P_Music = new ProcessClass(null, ProcessClass.ProcessMode.Music, 0, 0, "");
            P_Mixed = new ProcessClass(null, ProcessClass.ProcessMode.Mixed, 0, 0, "");

            //释放资源
            if(!File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\OneDrag\\Bin\\ffmpeg.exe") || !File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\OneDrag\\Bin\\lame.exe") || !File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\OneDrag\\Bin\\neroAacEnc.exe"))
            {
                Initialization.Visibility = Visibility.Visible;
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\OneDrag\\Bin");
                Thread ExportResource_T = new Thread(delegate ()
                {
                    if(!File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\OneDrag\\Bin\\ffmpeg.exe"))
                    {
                        //Console.WriteLine(VerifyResource("gz.ffmpeg.exe"));
                        if(VerifyResource("gz.ffmpeg.exe") == "93c97daed6a266c388c4335746941d40")
                        {
                            ExportResource(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\OneDrag\\Bin\\ffmpeg.exe", "gz.ffmpeg.exe");
                        }
                        else
                        {
                            ProcessLog("VERIFICATION FAILED at  " + DateTime.Now.ToString());
                            Environment.Exit(0);
                        }
                    }
                    if (!File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\OneDrag\\Bin\\lame.exe"))
                    {
                        //Console.WriteLine(VerifyResource("gz.lame.exe"));
                        if (VerifyResource("gz.lame.exe") == "eb3cd03670397c7d39973532ac3ef5d8")
                        {
                            ExportResource(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\OneDrag\\Bin\\lame.exe", "gz.lame.exe");
                        }
                        else
                        {
                            ProcessLog("VERIFICATION FAILED at  " + DateTime.Now.ToString());
                            Environment.Exit(0);
                        }
                    }
                    if (!File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\OneDrag\\Bin\\neroAacEnc.exe"))
                    {
                        //Console.WriteLine(VerifyResource("gz.neroAacEnc.exe"));
                        if (VerifyResource("gz.neroAacEnc.exe") == "91551576525a89afa55979e632254a")
                        {
                            ExportResource(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\OneDrag\\Bin\\neroAacEnc.exe", "gz.neroAacEnc.exe");
                        }
                        else
                        {
                            ProcessLog("VERIFICATION FAILED at  " + DateTime.Now.ToString());
                            Environment.Exit(0);
                        }
                    }

                    double O = 1;
                    while (O > 0)
                    {
                        O = O - 0.1;
                        Dispatcher.Invoke(new Action(() =>
                        {
                            Initialization.Opacity = O;
                        }));
                        Thread.Sleep(20);
                    }
                    Dispatcher.Invoke(new Action(() =>
                    {
                        Initialization.Visibility = Visibility.Hidden;
                    }));
                });
                ExportResource_T.Start();
            }

            //配置检测
            SampleRate = 44100;
            CodeRate = 320;
            FileFormat = "mp3";
            if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\OneDrag\\config.ini"))
            {
                StreamReader SR = new StreamReader(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\OneDrag\\config.ini");
                string configFile = SR.ReadToEnd();
                SR.Close();
                string[] config = configFile.Split('\n');
                foreach(string c in config)
                {
                    if(c.IndexOf("SampleRate") != -1)
                    {
                        int.TryParse(c.Substring(c.IndexOf('=') + 1).Trim(), out SampleRate);
                        PropertyBox_Vocal1.Text = SampleRate + "Hz 16Bit";
                        PropertyBox_Music1.Text = SampleRate + "Hz 16Bit";
                        PropertyBox_Mixed1.Text = SampleRate + "Hz 16Bit";
                    }
                    if (c.IndexOf("CodeRate") != -1)
                    {
                        int.TryParse(c.Substring(c.IndexOf('=') + 1).Trim(), out CodeRate);
                        PropertyBox_Mixed2.Text = CodeRate + "kbps " + FileFormat.ToUpper();
                    }
                    if (c.IndexOf("FileFormat") != -1)
                    {
                        FileFormat = c.Substring(c.IndexOf('=') + 1).Trim();
                        PropertyBox_Mixed2.Text = CodeRate + "kbps " + FileFormat.ToUpper();
                    }
                }
            }
            else
            {

            }

            //UI初始化
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            //this.Topmost = true;

            VocalProgressBar.BarOpacity = 0;
            MusicProgressBar.BarOpacity = 0;
            MixedProgressBar.BarOpacity = 0;

            if (!File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\OneDrag\\DescriptionHided"))
            {
                Description.Visibility = Visibility.Visible;
            }

            log += "[Software] Initial Complete at  " + DateTime.Now.ToString();
            //LogBox.AppendText("[Software] Initial Complete at  " + DateTime.Now.ToString());

            //显示UI
            MainGrid.Opacity = 0;
            Thread T = new Thread(delegate ()
            {
                double O = 0;
                while (O < 1)
                {
                    Dispatcher.Invoke(new Action(() =>
                    {
                        MainGrid.Opacity = O;
                    }));
                    O = O + 0.1;
                    Thread.Sleep(10);
                }
            });
            T.Start();

            VocalProgressBar.Reseted += delegate { Dispatcher.Invoke(new Action(() => { VocalGrid.AllowDrop = true; })); };
            MusicProgressBar.Reseted += delegate { Dispatcher.Invoke(new Action(() => { MusicGrid.AllowDrop = true; })); };
            MixedProgressBar.Reseted += delegate { Dispatcher.Invoke(new Action(() => { MixedGrid.AllowDrop = true; })); };
        }

        string log = "";

        ProcessClass P_Vocal;
        private void VocalGrid_PreviewDrop(object sender, DragEventArgs e)
        {
            Sys_ProcessOn();
            VocalGrid.AllowDrop = false;
            Array pathArr = (Array)e.Data.GetData(DataFormats.FileDrop);
            P_Vocal = new ProcessClass(pathArr, ProcessClass.ProcessMode.Vocal, SampleRate, CodeRate, "");
            P_Vocal.OutputReceived += P_Vocal_OutputReceivedEvent;
            P_Vocal.ProcessExited += Sys_ProcessOff;
            P_Vocal.ProgressUpdate += P_Vocal_ProgressUpdate;
            P_Vocal.Start();
        }

        private void P_Vocal_ProgressUpdate(double value)
        {
            VocalProgressBar.ChangeValue(value);
        }

        private void P_Vocal_OutputReceivedEvent(string message)
        {
            log += "\r\n[Vocal] " + message;
            //Dispatcher.Invoke(new Action(() =>
            //{
                
            //    LogBox.AppendText("\r\n[Vocal] " + message);
            //    LogBox.ScrollToEnd();
            //}));
        }

        ProcessClass P_Music;
        private void MusicGrid_PreviewDrop(object sender, DragEventArgs e)
        {
            Sys_ProcessOn();
            MusicGrid.AllowDrop = false;
            Array pathArr = (Array)e.Data.GetData(DataFormats.FileDrop);
            P_Music = new ProcessClass(pathArr, ProcessClass.ProcessMode.Music, SampleRate, CodeRate, "");
            P_Music.OutputReceived += P_Music_OutputReceivedEvent;
            P_Music.ProcessExited += Sys_ProcessOff;
            P_Music.ProgressUpdate += P_Music_ProgressUpdate;
            P_Music.Start();
        }

        private void P_Music_ProgressUpdate(double value)
        {
            MusicProgressBar.ChangeValue(value);
        }

        private void P_Music_OutputReceivedEvent(string message)
        {
            log += "\r\n[Music] " + message;
            //Dispatcher.Invoke(new Action(() =>
            //{
            //    LogBox.AppendText("\r\n[Music] " + message);
            //    LogBox.ScrollToEnd();
            //}));
        }

        ProcessClass P_Mixed;
        private void MixedGrid_PreviewDrop(object sender, DragEventArgs e)
        {
            Sys_ProcessOn();
            MixedGrid.AllowDrop = false;
            Array pathArr = (Array)e.Data.GetData(DataFormats.FileDrop);
            P_Mixed = new ProcessClass(pathArr, ProcessClass.ProcessMode.Mixed, SampleRate, CodeRate, FileFormat);
            P_Mixed.OutputReceived += P_Mixed_OutputReceived;
            P_Mixed.ProcessExited += Sys_ProcessOff;
            P_Mixed.ProgressUpdate += P_Mixed_ProgressUpdate;
            P_Mixed.Start();
        }

        private void P_Mixed_ProgressUpdate(double value)
        {
            MixedProgressBar.ChangeValue(value);
        }

        private void P_Mixed_OutputReceived(string message)
        {
            log += "\r\n[Mixed] " + message;
            //Dispatcher.Invoke(new Action(() =>
            //{
            //    LogBox.AppendText("\r\n[Mixed] " + message);
            //    LogBox.ScrollToEnd();
            //}));
        }

        private enum ES_REQUIRE
        {
            OFF = 0,
            ON = 1
        }

        ES_REQUIRE ES;
        private void Sys_ProcessOn()
        {
            if (ES == ES_REQUIRE.OFF)
            {
                SetThreadExecutionState(ES_CONTINUOUS | ES_DISPLAY_REQUIRED | ES_SYSTEM_REQUIRED);  //阻止休眠
                ES = ES_REQUIRE.ON;
                //Console.WriteLine("休眠已阻止");
            }
        }

        private void Sys_ProcessOff()
        {
            if(P_Vocal.ProgressValue == -1 && P_Music.ProgressValue == -1 && P_Mixed.ProgressValue == -1)
            {
                SetThreadExecutionState(ES_CONTINUOUS);  //恢复休眠
                ES = ES_REQUIRE.OFF;
                //Console.WriteLine("休眠已恢复");
            }
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyboardDevice.Modifiers == ModifierKeys.Control && e.Key == Key.R)
            {
                LogBox.Text = log;
                LogBox.ScrollToEnd();
                if (LogBox.Visibility == Visibility.Hidden)
                {
                    LogBox.Visibility = Visibility.Visible;
                    Storyboard S = Resources["ShowLogBox"] as Storyboard;
                    S.Begin();
                }
                else
                {
                    Storyboard S = Resources["HideLogBox"] as Storyboard;
                    S.Completed += delegate { LogBox.Visibility = Visibility.Hidden; LogBox.Text = ""; };
                    S.Begin();
                }
            }
            if(e.Key == Key.Escape)
            {
                this.Close();
            }
        }

        private enum TitleFlag
        {
            DragMove = 0,
            Minimize = 1,
            Close = 2
        }
        private TitleFlag titleflag;

        private void Header_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (CloseBtn.IsMouseOver == false && MinimizeBtn.IsMouseOver == false)
            {
                this.DragMove();
                titleflag = TitleFlag.DragMove;
            }
            else if (MinimizeBtn.IsMouseOver == true)
            {
                titleflag = TitleFlag.Minimize;
            }
            else if (CloseBtn.IsMouseOver == true)
            {
                titleflag = TitleFlag.Close;
            }
        }

        private void Header_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (MinimizeBtn.IsMouseOver == true && titleflag == TitleFlag.Minimize)
            {
                this.WindowState = WindowState.Minimized;
            }
            else if (CloseBtn.IsMouseOver == true && titleflag == TitleFlag.Close)
            {
                this.Close();
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;

            StreamWriter SW = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\OneDrag\\config.ini", false);
            SW.WriteLine("SampleRate = " + SampleRate);
            SW.WriteLine("CodeRate = " + CodeRate);
            SW.WriteLine("FileFormate = " + FileFormat);
            SW.Close();


            if (P_Vocal.ProgressValue != -1 || P_Music.ProgressValue != -1 || P_Mixed.ProgressValue != -1)
            {
                if (MessageBox.Show("尚有任务正在进行，是否终止", "软件即将关闭", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
                {
                    return;
                }
            }

            P_Vocal.Abort();
            P_Music.Abort();
            P_Mixed.Abort();

            ProcessLog(log);
            //ProcessLog(LogBox.Text);

            Thread T = new Thread(delegate ()
            {
                double O = 1;
                while (O > 0)
                {
                    O = O - 0.1;
                    Dispatcher.Invoke(new Action(() =>
                    {
                        MainGrid.Opacity = O;
                    }));
                    Thread.Sleep(10);
                }
                Dispatcher.Invoke(new Action(() =>
                {
                    //this.ShowInTaskbar = false;
                    this.Hide();
                    Environment.Exit(0);
                }));
            });
            T.Start();
        }

        private void ProcessLog(string message)
        {
            string Folder = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\OneDrag\\";
            Directory.CreateDirectory(Folder);
            int i = 0;
            while (i < 100)
            {
                string filename = "ProcessLog";
                if (i != 0)
                {
                    filename = filename + " (" + i + ")";
                }
                File.Delete(Folder + filename + ".log");
                if (!File.Exists(Folder + filename + ".log"))
                {
                    try
                    {
                        StreamWriter SW = new StreamWriter(Folder + filename + ".log", false);
                        SW.WriteLine(message);
                        SW.Close();
                        break;
                    }
                    catch { }
                }
                i++;
            }
        }

        private void CloseDescriptionBtn_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if(DescriptionCkb.IsChecked == true)
            {
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\OneDrag");
                StreamWriter SW = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\OneDrag\\DescriptionHided", false);
                SW.Close();
            }
            Thread T = new Thread(delegate ()
            {
                double O = 1;
                while (O > 0)
                {
                    O = O - 0.1;
                    Dispatcher.Invoke(new Action(() =>
                    {
                        Description.Opacity = O;
                    }));
                    Thread.Sleep(20);
                }
                Dispatcher.Invoke(new Action(() =>
                {
                    Description.Visibility = Visibility.Hidden;
                }));
            });
            T.Start();
        }

        int SampleRate;
        private void SampleRateChange_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if(SampleRate == 44100)
            {
                SampleRate = 48000;
            }
            else
            {
                SampleRate = 44100;
            }
            
            PropertyBox_Vocal1.Text = SampleRate + "Hz 16Bit";
            PropertyBox_Music1.Text = SampleRate + "Hz 16Bit";
            PropertyBox_Mixed1.Text = SampleRate + "Hz 16Bit";
        }

        int CodeRate;
        private void CodeRateChange_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (CodeRate == 320)
            {
                CodeRate = 192;
            }
            else if(CodeRate == 192)
            {
                CodeRate = 128;
            }
            else if (CodeRate == 128)
            {
                CodeRate = 960;
            }
            else
            {
                CodeRate = 320;
            }

            PropertyBox_Mixed2.Text = CodeRate + "kbps " + FileFormat.ToUpper();
        }

        string FileFormat;
        private void FileFormatChange_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (FileFormat == "mp3")
            {
                FileFormat = "m4a";
            }
            else
            {
                FileFormat = "mp3";
            }

            PropertyBox_Mixed2.Text = CodeRate + "kbps " + FileFormat.ToUpper();
        }

    }
}
