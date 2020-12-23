using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Threading;
using System.Windows;
using FileAutomation_1._0.Views;
using FileAutomation_1._0.Library;
using FileAutomation_1._0.Model;
using System.Runtime.InteropServices;
using FileAutomation_1._0.Helper;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Windows.Input;
using System.Windows.Xps.Packaging;
using Microsoft.Office.Interop.Word;
using System.IO;
using GdPicture14;
using System.Text.RegularExpressions;
using System.Linq;
using IronOcr;
using System.Windows.Documents;
using System.Drawing;
using System.Windows.Media.Imaging;
using System.Drawing.Imaging;

namespace FileAutomation_1._0
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        #region Private Properties
        private KeyboardHook keyboardHook = new KeyboardHook();
        private MouseHook mouseHook = new MouseHook();
        private int count = 0;
        private bool isHooked = false;
        private List<Record> recordList;
        private List<Filess> myFile;
        #endregion

        private volatile bool m_StopThread = false;
        public MainWindow()
        {
            InitializeComponent();
            LicenseManager lm = new LicenseManager();
            lm.RegisterKEY("0481484039779637361392356");
            myFile = new List<Filess>();
            recordList = new List<Record>();
            ((INotifyCollectionChanged)listView.Items).CollectionChanged += ListView_CollectionChanged;

            lastInPutNfo = new LASTINPUTINFO();
            lastInPutNfo.cbSize = (uint)Marshal.SizeOf(lastInPutNfo);

            uint idleTime = 0;
            Thread thread = new Thread(new ThreadStart(delegate ()
            {
                while (!m_StopThread)
                {
                    if (idleTime == 0)
                    {
                        if (g != null)
                        {
                            try
                            {
                                RedrawWindow(IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
                                g.Dispose();
                                ReleaseDC(desktop);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e);
                            }
                        }
                    }
                    idleTime = GetLastInputTime();
                    if (idleTime < 1)
                        drawOutlineElement();
                    Console.WriteLine(idleTime);
                    Thread.Sleep(1000);
                }
                //if(GetLastInputTime() > 1000)

            }));
            //thread.Start();
        }



        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            m_StopThread = true;
        }

        private void ListView_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            // Scroll to the last item
            if (e.Action == NotifyCollectionChangedAction.Add)
            {
                listView.ScrollIntoView(e.NewItems[0]);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //// Setup keyboard hook
            //keyboardHook.OnKeyboardEvent += KeyboardHook_OnKeyboardEvent;
            //keyboardHook.Install();

            //// Setup mouse hook
            //mouseHook.OnMouseEvent += MouseHook_OnMouseEvent;
            //mouseHook.OnMouseMove += MouseHook_OnMouseMove;
            //mouseHook.OnMouseWheelEvent += MouseHook_OnMouseWheelEvent;
            //mouseHook.Install();
            keyboardHook.OnKeyboardEvent += KeyboardHook_OnKeyboardEvent;

            mouseHook.OnMouseEvent += MouseHook_OnMouseEvent;
            mouseHook.OnMouseMove += MouseHook_OnMouseMove;
            mouseHook.OnMouseWheelEvent += MouseHook_OnMouseWheelEvent;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            keyboardHook.Uninstall();
            mouseHook.Uninstall();
        }

        #region Mouse events
        private void ProcessMouseEvent(MouseHook.MouseEvents mAction, int mValue)
        {
            CursorPoint mPoint = GetCurrentMousePosition();
            MouseEvent mEvent = new MouseEvent
            {
                Location = mPoint,
                Action = mAction,
                Value = mValue
            };

            LogMouseEvents(mEvent);
        }

        private bool MouseHook_OnMouseWheelEvent(int wheelValue)
        {
            ProcessMouseEvent((MouseHook.MouseEvents)wheelValue, 120);
            return false;
        }

        private bool MouseHook_OnMouseEvent(int mouseEvent)
        {
            ProcessMouseEvent((MouseHook.MouseEvents)mouseEvent, 0);
            return false;
        }


        private bool MouseHook_OnMouseMove(int x, int y)
        {
            ProcessMouseEvent(MouseHook.MouseEvents.MouseMove, 0);
            return false;
        }
        #endregion

        #region Keyboard events
        private bool KeyboardHook_OnKeyboardEvent(uint key, BaseHook.KeyState keyState)
        {
            KeyboardEvent kEvent = new KeyboardEvent
            {
                Key = (Keys)key,
                Action = (keyState == BaseHook.KeyState.Keydown) ? Constants.KEY_DOWN : Constants.KEY_UP
            };
            LogKeyboardEvents(kEvent);
            return false;
        }
        #endregion

        #region Record/Stop
        private void BtnRecord_Click(object sender, RoutedEventArgs e)
        {
            if (isHooked)
                return;
            if (listView.Items.Count > 0)
            {
                //MessageBoxResult result = System.Windows.MessageBox.Show("Do you want to record again?",
                //                          "Confirmation",
                //                          MessageBoxButton.YesNo,
                //                          MessageBoxImage.Question);
                //switch (result)
                //{
                //    case MessageBoxResult.Yes:
                //        listView.Items.Clear();
                //        recordList = new List<Record>();
                //        count = 0;
                //        break;
                //    default:
                //        return;
                //}

                listView.Items.Clear();
                recordList = new List<Record>();
                count = 0;
            }

            keyboardHook.Install();
            mouseHook.Install();
            isHooked = true;

            LaunchApp();
        }

        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            keyboardHook.Uninstall();
            mouseHook.Uninstall();
            isHooked = false;
        }
        #endregion

        #region Helper + Logging methods

        private void LaunchApp()
        {
            // An app is supposed to launch
            if (appPath.IsEnabled == false)
            {
                System.Diagnostics.Process.Start(appPath.Text);
            }
        }

        [DllImport("User32.dll")]
        static extern IntPtr GetDC(IntPtr hwnd);

        [DllImport("User32.dll", CallingConvention = CallingConvention.FastCall)]
        static extern void ReleaseDC(IntPtr dc);

        IntPtr desktop;

        private void TrackAutomationElement(Record item)
        {
            if (item.Type == Constants.MOUSE
                && item.EventMouse.Action == MouseHook.MouseEvents.LeftUp)
            {
                var windowTitle = Win32Utils.GetActiveWindowTitle();
                var position = Control.MousePosition;

                System.Windows.Point coordinates = new System.Windows.Point(position.X, position.Y);
                inspectBox.Text = string.Format("Title: {0}", windowTitle);

                try
                {
                    AutomationElement targetApp = AutomationElement.FromPoint(coordinates);

                    inspectBox.Text += "\n";
                    inspectBox.Text += string.Format("" +
                        "Name: {0}\n" +
                        "Automation Id: {1}\n" +
                        "Text: {2}\n" +
                        "Control Type: {3}",
                        targetApp.Current.Name,
                        targetApp.Current.AutomationId,
                        GetText(targetApp),
                        targetApp.Current.ControlType);

                    //desktop = GetDC(IntPtr.Zero);
                    //using (System.Drawing.Graphics g = System.Drawing.Graphics.FromHdc(desktop))
                    //{
                    //	System.Drawing.Pen pen = new System.Drawing.Pen(System.Drawing.Color.Red, 5);
                    //	Rect rect = targetApp.Current.BoundingRectangle;
                    //	Point point = rect.TopLeft;
                    //	g.DrawRectangle(pen, (float)point.X, (float)point.Y, (float)rect.Width, (float)rect.Height);
                    //}
                }
                catch (Exception)
                {
                    Console.WriteLine("Invalid UI element");
                }
            }
        }

        private void drawOutlineElement()
        {
            try
            {
                var position = Control.MousePosition;

                System.Windows.Point coordinates = new System.Windows.Point(position.X, position.Y);
                AutomationElement targetApp = AutomationElement.FromPoint(coordinates);

                Console.WriteLine(targetApp.Current.Name);
                string appName = targetApp.Current.Name;
                List<string> forbiden = new List<string>
                {
                    "",
                    "0",
                    "App Recorder"
                };
                foreach (string s in forbiden)
                {
                    if (appName.Equals(s))
                        return;
                }

                desktop = GetDC(IntPtr.Zero);
                using (System.Drawing.Graphics g = System.Drawing.Graphics.FromHdc(desktop))
                {
                    System.Drawing.Pen pen = new System.Drawing.Pen(System.Drawing.Color.Red, 5);
                    Rect rect = targetApp.Current.BoundingRectangle;
                    System.Windows.Point point = rect.TopLeft;
                    g.DrawRectangle(pen, (float)point.X, (float)point.Y, (float)rect.Width, (float)rect.Height);
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Invalid UI element");
            }
        }

        private string GetText(AutomationElement element)
        {
            object patternObj;
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out patternObj))
            {
                var valuePattern = (ValuePattern)patternObj;
                return valuePattern.Current.Value;
            }
            else if (element.TryGetCurrentPattern(TextPattern.Pattern, out patternObj))
            {
                var textPattern = (TextPattern)patternObj;
                return textPattern.DocumentRange.GetText(-1).TrimEnd('\r'); // often there is an extra '\r' hanging off the end.
            }
            else
            {
                return element.Current.Name;
            }
        }

        private CursorPoint GetCurrentMousePosition()
        {
            var position = Control.MousePosition;
            return new CursorPoint(position.X, position.Y);
        }

        private void LogMouseEvents(MouseEvent mEvent)
        {
            count++;
            Record item = new Record
            {
                Id = count,
                EventMouse = mEvent,
                Type = Constants.MOUSE,
                Content = String.Format("{0} was triggered at ({1}, {2})", mEvent.Action, mEvent.Location.X, mEvent.Location.Y)
            };

            AddRecordItem(item);
        }

        private void LogKeyboardEvents(KeyboardEvent kEvent)
        {
            count++;
            Record item = new Record
            {
                Id = count,
                Type = Constants.KEYBOARD,
                EventKey = kEvent,
                Content = String.Format("{0} was {1}", kEvent.Key.ToString(),
                    (kEvent.Action == Constants.KEY_DOWN) ? "pressed" : "released")
            };

            AddRecordItem(item);
        }

        private void LogWaitEvent(Record record)
        {
            count++;
            record.Id = count;
            record.Content = $"Wait {record.WaitMs} ms.";
            AddRecordItem(record);
        }

        private void AddRecordItem(Record item)
        {
            TrackAutomationElement(item);

            AddToListView(item);
            //this.listView.Items.Add(item);
            this.recordList.Add(item);
            countRecord.Content = String.Format("{0} records", count.ToString());
        }

        private void AddToListView(Record item)
        {
            // Check if two last records are similar
            if (listView.Items.Count > 0)
            {
                var lastItem = (Record)listView.Items[listView.Items.Count - 1];
                if (lastItem.Type == item.Type)
                {
                    switch (item.Type)
                    {
                        case Constants.MOUSE:
                            var lastAction = lastItem.EventMouse.Action;
                            if (lastAction == MouseHook.MouseEvents.MouseMove
                                && item.EventMouse.Action == lastAction)
                                this.listView.Items.RemoveAt(this.listView.Items.Count - 1);
                            break;
                        case Constants.KEYBOARD:
                            break;
                    }
                }
            }

            // satisfy every condition
            this.listView.Items.Add(item);
        }
        #endregion

        #region Playback
        private void BtnPlayback_Click(object sender, RoutedEventArgs e)
        {
            if (isHooked)
                return;

            int num;
            if (int.TryParse(repeatTime.Text, out num))
            {
                for (int i = 0; i < num; ++i)
                {
                    LaunchApp();

                    foreach (var record in recordList)
                    {
                        switch (record.Type)
                        {
                            case Constants.MOUSE:
                                PlaybackMouse(record);
                                break;
                            case Constants.KEYBOARD:
                                PlaybackKeyboard(record);
                                break;
                            case Constants.WAIT:
                                Thread.Sleep(record.WaitMs);
                                break;
                            default:
                                break;
                        }
                        Thread.Sleep(4);
                    }

                    Thread.Sleep(10);
                }
            }
            else
            {
                System.Windows.MessageBox.Show("Repeat time is not valid!");
            }

        }

        private void PlaybackMouse(Record record)
        {
            CursorPoint newPos = record.EventMouse.Location;
            MouseHook.MouseEvents mEvent = record.EventMouse.Action;
            MouseUtils.PerformMouseEvent(mEvent, newPos);
        }

        private void PlaybackKeyboard(Record record)
        {
            Keys key = record.EventKey.Key;
            string action = record.EventKey.Action;

            KeyboardUtils.PerformKeyEvent(key, action);
        }
        #endregion

        private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.FileName = "";
            openDialog.Title = "Application";
            openDialog.Filter = "All applications|*.exe";
            openDialog.ShowDialog();

            var appName = openDialog.FileName.ToString();
            if (!String.IsNullOrEmpty(appName))
            {
                appPath.Text = appName;
                appPath.IsEnabled = false;
            }
        }

        private void Label_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            appPath.Clear();
            appPath.IsEnabled = true;
        }

        private void BtnCreateClick_Click(object sender, RoutedEventArgs e)
        {
            var window = new CreateManualClickWindow();
            window.ShowDialog();

            if (window.mouseEvents != null)
            {
                window.mouseEvents.ForEach(me => LogMouseEvents(me));
            }
        }

        private void BtnCreateText_Click(object sender, RoutedEventArgs e)
        {
            var window = new CreateManualTypeKeyWindow();
            window.ShowDialog();

            string text = window.text;
            if (text != null)
            {
                text = text.ToUpper();
                foreach (char c in text)
                {
                    int code = c;
                    var key = (Keys)Enum.Parse(typeof(Keys), code.ToString());
                    LogKeyboardEvents(new KeyboardEvent { Key = key, Action = Constants.KEY_DOWN });
                    LogKeyboardEvents(new KeyboardEvent { Key = key, Action = Constants.KEY_UP });
                }
            }
        }

        const int RDW_INVALIDATE = 0x0001;
        const int RDW_ALLCHILDREN = 0x0080;
        const int RDW_UPDATENOW = 0x0100;
        [DllImport("User32.dll", CallingConvention = CallingConvention.StdCall)]
        static extern bool RedrawWindow(IntPtr hwnd, IntPtr rcUpdate, IntPtr regionUpdate, int flags);

        System.Drawing.Graphics g;
        private void ListView_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var item = listView.SelectedItem as Record;
            if (!listView.HasItems || item.EventMouse == null)
                return;
            try
            {
                if (g != null)
                {
                    RedrawWindow(IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
                    g.Dispose();
                    //ReleaseDC(desktop);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            int id = item.Id;
            System.Drawing.Pen pen = new System.Drawing.Pen(System.Drawing.Color.Red, 3);

            if (item.EventMouse.Action == MouseHook.MouseEvents.MouseMove)
            {
                Record last = recordList.FindLast(r =>
                {
                    if (r.EventMouse == null)
                        return false;
                    return r.Id < id && r.EventMouse.Action != MouseHook.MouseEvents.MouseMove;
                });
                if (last == null)
                {
                    last = recordList[0];
                }
                List<Record> list = recordList.FindAll(r => r.Id <= id && r.Id > last.Id);

                desktop = GetDC(IntPtr.Zero);
                g = System.Drawing.Graphics.FromHdc(desktop);
                System.Drawing.Point[] points = list.ConvertAll(new Converter<Record, System.Drawing.Point>(RecordToPoint)).ToArray();
                System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
                path.AddLines(points);
                g.DrawPath(pen, path);
                //g.Clear(System.Drawing.Color.Transparent);
            }
            else if (item.Type == Constants.MOUSE)
            {
                int lengthLine = 40;
                desktop = GetDC(IntPtr.Zero);
                g = System.Drawing.Graphics.FromHdc(desktop);
                System.Drawing.Point point1 = new System.Drawing.Point(
                    (int)item.EventMouse.Location.X, (int)item.EventMouse.Location.Y - lengthLine);
                System.Drawing.Point point2 = new System.Drawing.Point(
                    (int)item.EventMouse.Location.X, (int)item.EventMouse.Location.Y + lengthLine);
                g.DrawLine(pen, point1, point2);

                System.Drawing.Point point3 = new System.Drawing.Point(
                    (int)item.EventMouse.Location.X - lengthLine, (int)item.EventMouse.Location.Y);
                System.Drawing.Point point4 = new System.Drawing.Point(
                    (int)item.EventMouse.Location.X + lengthLine, (int)item.EventMouse.Location.Y);

                g.DrawLine(pen, point3, point4);
            }
        }

        private struct LASTINPUTINFO
        {
            public uint cbSize;
            public uint dwTime;
        }

        private static LASTINPUTINFO lastInPutNfo;
        [DllImport("user32.dll")]
        static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        public static uint GetLastInputTime()
        {
            uint idleTime = 0;
            LASTINPUTINFO lastInputInfo = new LASTINPUTINFO();
            lastInputInfo.cbSize = (uint)Marshal.SizeOf(lastInputInfo);
            lastInputInfo.dwTime = 0;

            uint envTicks = (uint)Environment.TickCount;

            if (GetLastInputInfo(ref lastInputInfo))
            {
                uint lastInputTick = lastInputInfo.dwTime;

                idleTime = envTicks - lastInputTick;
            }

            return ((idleTime > 0) ? (idleTime / 1000) : 0);
        }

        private System.Drawing.Point RecordToPoint(Record r)
        {
            return new System.Drawing.Point((int)r.EventMouse.Location.X, (int)r.EventMouse.Location.Y);
        }

        private void ListView_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            try
            {
                if (g != null)
                {
                    RedrawWindow(IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
                    g.Dispose();
                    ReleaseDC(desktop);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private void BtnInsertKey_Click(object sender, RoutedEventArgs e)
        {
            CreateInsertKeyWindow window = new CreateInsertKeyWindow();
            window.ShowDialog();

            if (window.keyboardEvents != null)
            {
                window.keyboardEvents.ForEach(me => LogKeyboardEvents(me));
            }
        }

        private void BtnWait_Click(object sender, RoutedEventArgs e)
        {
            CreateWaitWindow window = new CreateWaitWindow();
            window.ShowDialog();

            Record record = window.waitEvent;
            if (record != null)
            {
                LogWaitEvent(record);
            }
        }


        /*private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                if (ofd.FileName.Length > 0)

                {

                    selectedFileTextBox.Text = ofd.FileName;

                    XpsDocument document1 = new XpsDocument(ofd.FileName, System.IO.FileAccess.Read);

                    documentViewer1.Document = document1.GetFixedDocumentSequence();

                    //string newXPSDocumentName = String.Concat(System.IO.Path.GetDirectoryName(ofd.FileName), "\\",

                    //System.IO.Path.GetFileNameWithoutExtension(ofd.FileName), ".xps");


                    // Set DocumentViewer.Document to XPS document
                    //documentViewer1.Document = ConvertWordDocToXPSDoc(ofd.FileName, newXPSDocumentName).GetFixedDocumentSequence();

                }
            }
        }

        private XpsDocument ConvertWordDocToXPSDoc(string wordDocName, string xpsDocName)

        {

            // Create a WordApplication and add Document to it

            Microsoft.Office.Interop.Word.Application wordApplication = new Microsoft.Office.Interop.Word.Application();
            wordApplication.Documents.Add(wordDocName);

            Document doc = wordApplication.ActiveDocument;

            // You must ensure you have Microsoft.Office.Interop.Word.Dll version 12.

            // Version 11 or previous versions do not have WdSaveFormat.wdFormatXPS option

            try

            {

                doc.SaveAs(xpsDocName, WdSaveFormat.wdFormatXPS);
                wordApplication.Quit();
                XpsDocument xpsDoc = new XpsDocument(xpsDocName, System.IO.FileAccess.Read);

                return xpsDoc;

            }

            catch (Exception exp)

            {
                string str = exp.Message;
            }

            return null;

        }
        */

        private void AddFile_Click(object sender, RoutedEventArgs e)
        {

            OpenFileDialog ofd = new OpenFileDialog();

            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                FileInfo fi = new FileInfo(ofd.FileName);

                //var lastItem = (Filess)listView.Items[fileLIstView.Items.Count - 1];

                fileLIstView.Items.Add(new Filess() { Name = fi.Name, FilePath = fi.FullName, Type = fi.Extension });
                FileViewer.DisplayFromFile(ofd.FileName);

            }

        }

        private void fileLIstView_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            Filess f = (Filess)fileLIstView.SelectedItem;
            FileViewer.DisplayFromFile(f.FilePath);

        }

        /*private void searcWord(string fileName)
        {
            string filePath = fileName;
            DocumentCore dc = DocumentCore.Load(filePath);

            string searchText = "Tom";

            Regex regex = new Regex("(?i)" + searchText);
            int count = dc.Content.Find(regex).Count();

        }*/



        private void extractText_Click(object sender, RoutedEventArgs e)
        {
            var ocr = new IronTesseract();
            Filess f = (Filess)fileLIstView.SelectedItem;

            if (!string.IsNullOrEmpty(f.FilePath))
            {
                var res = ocr.Read(f.FilePath);

                ExtracctText Window = new ExtracctText(res.Text);
                Window.ShowDialog();
            }

        }

        private void saveFile_Click(object sender, RoutedEventArgs e)
        {
            List<string> content = new List<string>();

            foreach (var x in recordList)
            {

                var result = x.Id + " " + x.Content + " " + x.Type;
                content.Add(result);

            }

            var result2 = string.Join("\n", content);
            System.Windows.MessageBox.Show(result2, "My App", MessageBoxButton.OK, MessageBoxImage.Information);

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Text file(*.txt)|*.txt";

            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                File.WriteAllText(sfd.FileName, result2);
            }
        }

        private void openApp(object sender, MouseButtonEventArgs e)
        {
            Filess f = (Filess)fileLIstView.SelectedItem;
            System.Diagnostics.Process.Start(f.FilePath);
        }

        private void Take_Screenshot_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();

            System.Threading.Thread.Sleep(500);

            System.Drawing.Rectangle bounds = Screen.GetBounds(System.Drawing.Point.Empty);

            Bitmap bitmap_Screen = new Bitmap(bounds.Width, bounds.Height);
            Graphics g = Graphics.FromImage(bitmap_Screen);
            g.CopyFromScreen(0, 0, 0, 0, new System.Drawing.Size(bounds.Width, bounds.Height));

            Image image = (Image)bitmap_Screen;

            BitmapImage bi = new BitmapImage();

            bi.BeginInit();

            MemoryStream ms = new MemoryStream();

            // Save to a memory stream...

            image.Save(ms, ImageFormat.Bmp);

            // Rewind the stream...

            ms.Seek(0, SeekOrigin.Begin);

            // Tell the WPF image to use this stream...

            bi.StreamSource = ms;

            bi.EndInit();
            screenShot_picture.Source = bi;

            /**
            SendKeys.SendWait("{PRTSC}");
            Image myImage = System.Windows.Forms.Clipboard.GetImage();

            BitmapImage bi = new BitmapImage();

            bi.BeginInit();

            MemoryStream ms = new MemoryStream();

            // Save to a memory stream...

            myImage.Save(ms, ImageFormat.Bmp);

            // Rewind the stream...

            ms.Seek(0, SeekOrigin.Begin);

            // Tell the WPF image to use this stream...

            bi.StreamSource = ms;

            bi.EndInit();
            screenShot_picture.Source = bi;
            **/


            this.Show();
        }

        private void Extract_Screenshot_Button_Click(object sender, RoutedEventArgs e)
        {
            var ocr = new IronTesseract();



            if (screenShot_picture.Source != null)
            {
                Image myImage = SourceImageToImage(screenShot_picture.Source);
                var res = ocr.Read(myImage);
                ExtracctText Window = new ExtracctText(res.Text);
                Window.ShowDialog();
            }
        }

        private System.Drawing.Image SourceImageToImage(System.Windows.Media.ImageSource image)
        {
            MemoryStream ms = new MemoryStream();
            var encoder = new System.Windows.Media.Imaging.BmpBitmapEncoder();
            encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(image as System.Windows.Media.Imaging.BitmapSource));
            encoder.Save(ms);
            ms.Flush();
            return System.Drawing.Image.FromStream(ms);
        }

        private void Capture_Click(object sender, RoutedEventArgs e)
        {
            SelectArea area = new SelectArea();
            //this.WindowState = WindowState.Minimized;
            area.ShowDialog();

            ScreenShot shotData = area.shotData;

            if (shotData != null)
            {
                this.WindowState = WindowState.Normal;
                System.Drawing.Rectangle rect = new System.Drawing.Rectangle(shotData.x, shotData.y, shotData.width, shotData.height);

                Bitmap bmp = new Bitmap(rect.Width, rect.Height, PixelFormat.Format32bppArgb);

                Graphics g = Graphics.FromImage(bmp);
                g.CopyFromScreen(rect.Left, rect.Top, 0, 0, shotData.size, CopyPixelOperation.SourceCopy);

                Image image = (Image)bmp;

                BitmapImage bi = new BitmapImage();

                bi.BeginInit();

                MemoryStream ms = new MemoryStream();

                // Save to a memory stream...

                image.Save(ms, ImageFormat.Bmp);

                // Rewind the stream...

                ms.Seek(0, SeekOrigin.Begin);

                // Tell the WPF image to use this stream...

                bi.StreamSource = ms;

                bi.EndInit();

                screenShot_picture.Source = bi;
            }

        }

        private void saveScreenshot_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.CheckPathExists = true;
            sfd.FileName = "Capture";
            sfd.Filter = "PNG Image(*.png)|*.png|JPG Image(*.jpg)|*.jpg|BMP Image(*.bmp)|*.bmp";

            if (screenShot_picture.Source != null)
            {
                Image image = SourceImageToImage(screenShot_picture.Source);
                if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    image.Save(sfd.FileName);
                }
            }

        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            screenShot_picture.Source = null;
        }

        /*private void richbutton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            if(ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                object readOnly = false;
                object visible = true;
                object save = false;
                object fileName = ofd.FileName;
                object newTemplate = false;
                object docType = 0;
                object missing = Type.Missing;
                Microsoft.Office.Interop.Word._Document document;
                Microsoft.Office.Interop.Word._Application application = new Microsoft.Office.Interop.Word.Application();
                application.Visible = false;

                document = application.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref visible, ref missing, ref missing, ref missing, ref missing);
                document.ActiveWindow.Selection.WholeStory();
                document.ActiveWindow.Selection.Copy();
                System.Windows.Forms.IDataObject dataObject = System.Windows.Forms.Clipboard.GetDataObject();
                richtextbox.Rtf = dataObject.GetData(System.Windows.Forms.DataFormats.Rtf).ToString();


            }
        }*/
    }
}
