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
using System.IO;
using GdPicture14;
using System.Linq;
using IronOcr;
using System.Windows.Documents;
using System.Drawing;
using System.Windows.Media.Imaging;
using System.Drawing.Imaging;
using System.Text;
using System.Windows.Media;

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

            LicenseManager lm = new LicenseManager();
            lm.RegisterKEY("0402451051843103834661532");

            keyboardHook.OnKeyboardEvent += KeyboardHook_OnKeyboardEvent;

            mouseHook.OnMouseEvent += MouseHook_OnMouseEvent;
            mouseHook.OnMouseMove += MouseHook_OnMouseMove;
            mouseHook.OnMouseWheelEvent += MouseHook_OnMouseWheelEvent;

            string folder = @"C:\Users\Acer\Documents";
            string pathString = System.IO.Path.Combine(folder, "automation");

            string[] file;

            if (Directory.Exists(pathString))
            {
                file = Directory.GetFiles(pathString);

                foreach (string x in file)
                {
                    FileInfo fi = new FileInfo(x);
                    myFile.Add(new Filess() { Name = fi.Name, FilePath = fi.FullName, Type = fi.Extension });
                    fileLIstView.Items.Add(new Filess() { Name = fi.Name, FilePath = fi.FullName, Type = fi.Extension });
                }
            }


        }

        private void Window_Closed(object sender, EventArgs e)
        {
            keyboardHook.Uninstall();
            mouseHook.Uninstall();

            string folder = @"C:\Users\Acer\Documents";
            string pathString = System.IO.Path.Combine(folder, "automation");

            string[] fileName;

            if (!Directory.Exists(pathString))
            {
                Directory.CreateDirectory(pathString);
                foreach (var x in myFile)
                {

                    string pathFile = System.IO.Path.Combine(pathString, x.Name);
                    File.Copy(x.FilePath, pathFile);
                }
            }
            else
            {
                fileName = Directory.GetFiles(pathString);
                bool found = false;
                foreach (string x in fileName)
                {
                    foreach (var j in myFile)
                    {
                        if (!x.Equals(j.FilePath))
                        {
                            string pathFile = System.IO.Path.Combine(pathString, j.Name);
                            if (!j.FilePath.Equals(pathFile))
                            {
                                File.Copy(j.FilePath, pathFile);
                                found = true;
                                break;
                            }

                        }
                    }
                    if (found == true)
                    {
                        break;
                    }
                }

            }
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
                //inspectBox.Text = string.Format("Title: {0}", windowTitle);

                try
                {
                    AutomationElement targetApp = AutomationElement.FromPoint(coordinates);
                    /*
                    inspectBox.Text += "\n";
                    inspectBox.Text += string.Format("" +
                        "Name: {0}\n" +
                        "Automation Id: {1}\n" +
                        "Text: {2}\n" +
                        "Control Type: {3}",
                        targetApp.Current.Name,
                        targetApp.Current.AutomationId,
                        GetText(targetApp),
                        targetApp.Current.ControlType);*/

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
                myFile.Add(new Filess() { Name = fi.Name, FilePath = fi.FullName, Type = fi.Extension });
                fileLIstView.Items.Add(new Filess() { Name = fi.Name, FilePath = fi.FullName, Type = fi.Extension });

                FileViewer.DisplayFromFile(ofd.FileName);
                FileViewer.Focus();
                //FileViewer.DisplayFirstPage();
                UpdatePageCount();

            }

        }

        private void fileLIstView_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            Filess f = (Filess)fileLIstView.SelectedItem;
            if (f != null)
            {
                FileViewer.DisplayFromFile(f.FilePath);
                FileViewer.Focus();
                DisplayFirstPage();
                UpdatePageCount();
            }

        }

        private void DisplayFirstPage()
        {
            //FileViewer.DisplayFirstPage();
            if (FileViewer.GetStat() != GdPictureStatus.OK)
            {
                System.Windows.MessageBox.Show("Error: " + FileViewer.GetStat());
            }
        }
        /*private void searcWord(string fileName)
        {
            string filePath = fileName;
            DocumentCore dc = DocumentCore.Load(filePath);

            string searchText = "Tom";

            Regex regex = new Regex("(?i)" + searchText);
            int count = dc.Content.Find(regex).Count();

        }*/


        private void PreviousPageCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            if (FileViewer != null)
                if (FileViewer.PageCount > 0 && FileViewer.CurrentPage > 1)
                    e.CanExecute = true;
                else
                    e.CanExecute = false;
        }

        private void PreviousPageCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            FileViewer.DisplayPreviousPage();
            UpdatePageCount();
            if (FileViewer.GetStat() != GdPictureStatus.OK)
            {
                System.Windows.MessageBox.Show("Error: " + FileViewer.GetStat());
            }
        }

        private void NextPageCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            if (FileViewer != null)
                if (FileViewer.PageCount > 0 && FileViewer.CurrentPage != FileViewer.PageCount)
                    e.CanExecute = true;

                else
                    e.CanExecute = false;
        }

        private void NextPageCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            FileViewer.DisplayNextPage();
            UpdatePageCount();
            if (FileViewer.GetStat() != GdPictureStatus.OK)
            {
                System.Windows.MessageBox.Show("Error: " + FileViewer.GetStat());
            }
        }



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

                Bitmap bmp = new Bitmap(rect.Width, rect.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);

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


        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
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
        }

        private void btn_saveAutomation_Click(object sender, RoutedEventArgs e)
        {
            if (listView.Items.Count > 0)
            {

                List<string> content = new List<string>();
                var result = "";
                foreach (var x in recordList)
                {

                    switch (x.Type)
                    {
                        case Constants.MOUSE:
                            result = x.EventMouse.Action + " " + x.EventMouse.Location.X + " " + x.EventMouse.Location.Y;
                            content.Add(result);
                            break;
                        case Constants.KEYBOARD:
                            result = x.EventKey.Key.ToString() + " " + x.EventKey.Action;
                            content.Add(result);
                            break;
                        case Constants.WAIT:
                            Thread.Sleep(x.WaitMs);
                            break;
                        default:
                            break;
                    }

                }

                var result2 = string.Join("\n", content);
                //System.Windows.MessageBox.Show(result2, "My App", MessageBoxButton.OK, MessageBoxImage.Information);

                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Text file(*.txt)|*.txt";

                if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    File.WriteAllText(sfd.FileName, result2);
                }

            }
        }

        private void btn_loadAutomation_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            string[] lines;
            var list = new List<string>();
            string word;
            int offset = 0;
            List<MouseEvent> mouseEvent = new List<MouseEvent>();
            List<KeyboardEvent> keyEvent = new List<KeyboardEvent>();

            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                var fileStream = new FileStream(ofd.FileName, FileMode.Open, FileAccess.Read);
                using (var streamReader = new StreamReader(fileStream, Encoding.UTF8))
                {
                    string line;
                    while ((line = streamReader.ReadLine()) != null)
                    {
                        list.Add(line);
                    }
                }
                lines = list.ToArray();

                word = string.Join("\n", lines);
                //rtb_loadAutomationText.Document.Blocks.Clear();
                //rtb_loadAutomationText.Document.Blocks.Add(new System.Windows.Documents.Paragraph(new Run(word)));

                foreach (var w in lines)
                {
                    string[] text = w.Split(null);


                    if (!w.Contains("KEY"))
                    {
                        string a = text[0];
                        double x = Convert.ToDouble(text[1]);
                        double y = Convert.ToDouble(text[2]);

                        CursorPoint point = new CursorPoint(x, y);
                        if (a.Contains("Left"))
                        {
                            offset = 0;
                            if (a.Equals("LeftUp"))
                            {
                                mouseEvent.Add(new MouseEvent
                                {
                                    Location = point,
                                    Action = MouseHook.MouseEvents.LeftUp + offset
                                });
                                LogMouseEvents(new MouseEvent { Location = point, Action = MouseHook.MouseEvents.LeftUp + offset });
                            }


                            if (a.Equals("LeftDown"))
                            {
                                mouseEvent.Add(new MouseEvent
                                {
                                    Location = point,
                                    Action = MouseHook.MouseEvents.LeftDown + offset
                                });
                                LogMouseEvents(new MouseEvent { Location = point, Action = MouseHook.MouseEvents.LeftDown + offset });
                            }


                        }

                        if (a.Contains("Right"))
                        {
                            offset = 0;
                            if (a.Equals("RightUp"))
                            {
                                mouseEvent.Add(new MouseEvent
                                {
                                    Location = point,
                                    Action = MouseHook.MouseEvents.RightUp + offset
                                });

                                LogMouseEvents(new MouseEvent { Location = point, Action = MouseHook.MouseEvents.RightUp + offset });
                            }


                            if (a.Equals("RightDown"))
                            {
                                mouseEvent.Add(new MouseEvent
                                {
                                    Location = point,
                                    Action = MouseHook.MouseEvents.RightDown + offset
                                });

                                LogMouseEvents(new MouseEvent { Location = point, Action = MouseHook.MouseEvents.RightDown + offset });
                            }


                        }

                        if (a.Contains("Middle"))
                        {
                            offset = 6;
                        }

                        if (a.Contains("Move"))
                        {
                            mouseEvent.Add(new MouseEvent
                            {
                                Location = point,
                                Action = MouseHook.MouseEvents.MouseMove
                            });
                            LogMouseEvents(new MouseEvent { Location = point, Action = MouseHook.MouseEvents.MouseMove });
                        }


                        if (a.Equals("ScrollDown"))
                        {
                            mouseEvent.Add(new MouseEvent
                            {
                                Location = point,
                                Action = MouseHook.MouseEvents.ScrollDown + offset
                            });

                            LogMouseEvents(new MouseEvent { Location = point, Action = MouseHook.MouseEvents.ScrollDown + offset });
                        }

                        if (a.Equals("ScrollUp"))
                        {
                            mouseEvent.Add(new MouseEvent
                            {
                                Location = point,
                                Action = MouseHook.MouseEvents.ScrollUp + offset
                            });

                            LogMouseEvents(new MouseEvent { Location = point, Action = MouseHook.MouseEvents.ScrollUp + offset });
                        }

                    }


                    if (w.Contains("KEY"))
                    {
                        string a = text[0];
                        //a = a.ToUpper();
                        //char c = a[0];
                        string x = text[1];

                        if (x.Equals("KEY_DOWN"))
                        {
                            //int code = c;
                            var key = (Keys)Enum.Parse(typeof(Keys), a);
                            LogKeyboardEvents(new KeyboardEvent { Key = key, Action = Constants.KEY_DOWN });

                            keyEvent.Add(new KeyboardEvent { Key = key, Action = Constants.KEY_DOWN });
                        }

                        if (x.Equals("KEY_UP"))
                        {
                            //int code = c;
                            var key = (Keys)Enum.Parse(typeof(Keys), a);
                            LogKeyboardEvents(new KeyboardEvent { Key = key, Action = Constants.KEY_UP });
                            keyEvent.Add(new KeyboardEvent { Key = key, Action = Constants.KEY_UP });
                        }

                    }


                }

                /* if (mouseEvent != null)
                 {
                     mouseEvent.ForEach(me => LogMouseEvents(me));
                 }
                */
            }

        }

        private void tbCurrentPage_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            int page = 0;
            if (int.TryParse(tbCurrentPage.Text, out page))
            {
                if (page > 0 & page <= FileViewer.PageCount)
                {
                    FileViewer.DisplayPage(page);
                    UpdateaNavigationToolbar();
                }
            }
        }

        private void UpdateaNavigationToolbar()
        {
            int currentPage = FileViewer.CurrentPage;
            tbCurrentPage.Text = currentPage.ToString();

            lblPageCount.Content = "/ " + FileViewer.PageCount;


        }

        private void UpdatePageCount()
        {
            if (FileViewer.PageCount == 0)
            {
                tbCurrentPage.Text = "0";
                lblPageCount.Content = "/ 0";
                lblPageCount.Content = "/ " + FileViewer.PageCount;

            }
            else
            {
                UpdateaNavigationToolbar();

            }
        }


        private string GetDocumentTypeLabel()
        {
            return FileViewer.GetDocumentType().ToString().Replace("DocumentType", "");
        }

        private void btn_deleteFile_Click(object sender, RoutedEventArgs e)
        {
            Filess file = (Filess)fileLIstView.SelectedItem;
            if (file != null)
            {
                int x = fileLIstView.Items.IndexOf(fileLIstView.SelectedItem);
                myFile.RemoveAt(x);
                fileLIstView.Items.RemoveAt(fileLIstView.Items.IndexOf(fileLIstView.SelectedItem));

                //File.Delete(file.FilePath);
                FileViewer.Clear();
            }

        }

        private void btn_refresh_Click(object sender, RoutedEventArgs e)
        {
            string folder = @"C:\Users\Acer\Documents";
            string pathString = System.IO.Path.Combine(folder, "automation");
            fileLIstView.Items.Clear();
            myFile.Clear();

            string[] file;

            if (Directory.Exists(pathString))
            {
                file = Directory.GetFiles(pathString);

                foreach (string x in file)
                {
                    FileInfo fi = new FileInfo(x);
                    myFile.Add(new Filess() { Name = fi.Name, FilePath = fi.FullName, Type = fi.Extension });
                    fileLIstView.Items.Add(new Filess() { Name = fi.Name, FilePath = fi.FullName, Type = fi.Extension });
                }
            }
        }

        #region search

        private int _currentSearchOccurence;

        private void ResetSearch()
        {
            lstSearchResults.Items.Clear();
            FileViewer.RemoveAllRegions();
            _currentSearchOccurence = 0;
        }
        private void SearchCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = !string.IsNullOrWhiteSpace(tbSearch.Text); ;
        }

        private delegate void UpdateProgressBarDelegate(
        System.Windows.DependencyProperty dp, Object value);

        private void SearchNextCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            Search(true);
        }

        private void SearchPreviousCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            Search(false);
        }

        private void SearchCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(tbSearch.Text))
            {
                return;
            }

            ResetSearch();
            int page = 0;
            bool found = false;
            if (rbAllPages.IsChecked == true)
            {
                page = 1;
            }
            else
            {
                page = FileViewer.CurrentPage;
            }
            bool finish = false;
            int countResults = 0;
            Cursor = System.Windows.Input.Cursors.Wait;
            searchProgressBar.Value = 1;
            searchProgressBar.Maximum = FileViewer.PageCount;
            stSearchProgress.Visibility = System.Windows.Visibility.Visible;

            UpdateProgressBarDelegate updatePbDelegate =
       new UpdateProgressBarDelegate(searchProgressBar.SetValue);
            double value = searchProgressBar.Value;
            while (!finish)
            {
                lblSearchResults.Text = "Find results for page " + page + " of " + FileViewer.PageCount;
                value++;
                Dispatcher.Invoke(updatePbDelegate,
           System.Windows.Threading.DispatcherPriority.Background,
           new object[] { System.Windows.Controls.ProgressBar.ValueProperty, value });
                ;


                int count = GetSearchResultCount(page, tbSearch.Text);
                if (count > 0)
                {
                    found = true;
                    String content = "Page " + page + ", ";
                    content += count + " occurence(s) found";
                    System.Windows.Controls.ListViewItem item = new System.Windows.Controls.ListViewItem();
                    item.Content = content;
                    item.Name = "Page" + page;
                    item.Tag = page;
                    lstSearchResults.Items.Add(item);
                }
                countResults += count;
                page++;
                finish = rbCurrentPage.IsChecked == true || page > FileViewer.PageCount;
            }

            Cursor = System.Windows.Input.Cursors.Arrow;

            lblSearchResults.Text = "";
            stSearchProgress.Visibility = System.Windows.Visibility.Hidden;

            if (!found)
            {
                System.Windows.MessageBox.Show("No match found", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private bool Search(bool ascending)
        {
            bool go = true;
            int page = FileViewer.CurrentPage;
            bool found = false;
            int newOccurence = 0;
            double occurenceLeft = 0;
            double occurenceTop = 0;
            double occurenceWidth = 0;
            double occurenceHeight = 0;
            if (ascending)
            {
                newOccurence = _currentSearchOccurence + 1;
            }
            else
            {
                newOccurence = _currentSearchOccurence - 1;
            }
            while (go)
            {
                if (FileViewer.SearchText(page, tbSearch.Text, newOccurence, chkCaseSensitive.IsChecked == true, chkWholeWord.IsChecked == true, ref occurenceLeft, ref occurenceTop, ref occurenceWidth, ref occurenceHeight))
                {
                    if (page != FileViewer.CurrentPage)
                    {
                        FileViewer.DisplayPage(page);
                    }
                    FileViewer.RemoveAllRegions();
                    AddSearchRegion(newOccurence, occurenceLeft, occurenceTop, occurenceWidth, occurenceHeight, true);
                    _currentSearchOccurence = newOccurence;
                    found = true;
                    go = false;
                }
                else
                {
                    if (rbAllPages.IsChecked == true)
                    {
                        if (ascending)
                        {
                            page++;
                            newOccurence = 1;
                        }
                        else
                        {
                            page--;
                            newOccurence = GetSearchResultCount(page, tbSearch.Text);
                        }
                        if (page == 0 | page > FileViewer.PageCount)
                        {
                            go = false;
                        }
                    }
                    else
                    {
                        go = false;
                    }
                }
            }
            if (!found)
            {
                System.Windows.MessageBox.Show(this, "No match found !", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            return found;
        }

        private void AddSearchRegion(int occurence, double leftCoordinate, double topCoordinate, double regionWidth, double regionheight, bool ensureVisibility)
        {
            int searchRegion = FileViewer.AddRegion("SearchResult" + Convert.ToString(occurence), leftCoordinate, topCoordinate, regionWidth, regionheight, Colors.Yellow, GdPicture14.WPF.GdViewer.RegionFillMode.Multiply);
            FileViewer.SetRegionEditable(searchRegion, false);
            if (ensureVisibility)
            {
                FileViewer.EnsureRegionVisibility(searchRegion);
            }
        }

        private void lstSearchResults_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (lstSearchResults.SelectedItems.Count != 0)
            {
                FileViewer.RemoveAllRegions();
                System.Windows.Controls.ListViewItem item = (System.Windows.Controls.ListViewItem)lstSearchResults.SelectedItems[0];
                int page = Convert.ToInt32(item.Tag);
                if (FileViewer.CurrentPage != page)
                {
                    FileViewer.DisplayPage(page);
                }
                int occurence = 1;
                double occurenceLeft = 0;
                double occurenceTop = 0;
                double occurenceWidth = 0;
                double occurenceHeight = 0;
                while (FileViewer.SearchText(page, tbSearch.Text, occurence, chkCaseSensitive.IsChecked == true, chkWholeWord.IsChecked == true, ref occurenceLeft, ref occurenceTop, ref occurenceWidth, ref occurenceHeight))
                {
                    AddSearchRegion(occurence, occurenceLeft, occurenceTop, occurenceWidth, occurenceHeight, occurence == 1);
                    occurence++;
                }
            }
        }

        private int GetSearchResultCount(int page, string searchedTerm)
        {
            return FileViewer.GetTextOccurrenceCount(page, searchedTerm, chkCaseSensitive.IsChecked == true, chkWholeWord.IsChecked == true);
        }

        //private void openPDF_Click(object sender, RoutedEventArgs e)
        //{
        //    OpenFileDialog ofd = new OpenFileDialog();
        //    ofd.Filter = "PDF Files|*.pdf";

        //    if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        //    {

        //        FileViewer.DisplayFromFile(ofd.FileName);


        //    }
        //}

        #endregion

        private void searchText()
        {
            var text = searchText_tb.Text;
            List<Filess> FileSearch = new List<Filess>();
            int page = 1;
            bool found = false;
            bool finish = false;
            int countResults = 0;

            foreach (var file in myFile)
            {
                FileInfo fi = new FileInfo(file.FilePath);
                FileViewer.DisplayFromFile(file.FilePath);
                while (!finish)
                {

                    int count = GetSearchResultCount(page, text);
                    if (count > 0)
                    {
                        found = true;
                        FileSearch.Add(new Filess() { Name = fi.Name, FilePath = fi.FullName, Type = fi.Extension });
                        finish = true;
                    }
                    countResults += count;
                    page++;
                    finish = page > FileViewer.PageCount;
                }


                if (!found)
                {
                    System.Windows.MessageBox.Show("No match found for " + file.Name, "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }

            foreach (var x in FileSearch)
            {
                FileInfo info = new FileInfo(x.FilePath);
                fileLIstView.Items.Clear();
                fileLIstView.Items.Add(new Filess() { Name = info.Name, FilePath = info.FullName, Type = info.Extension });
            }

        }

        private void searchWord_Click(object sender, RoutedEventArgs e)
        {
            if (!searchText_tb.Text.Equals(""))
            {
                searchText();
                tbSearch.Text = searchText_tb.Text;
                FileViewer.Clear();
            }
            else
            {
                System.Windows.MessageBox.Show("Please enter word to search!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }


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
