#define PROGRAM_RUNNING

using PSeatTester.Properties;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;

public class InputBox : IDisposable
{
    public static DialogResult Show(string title, string promptText, ref string value)
    {
        return Show(title, promptText, ref value, null);
    }
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
        return;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {

        }
        return;
    }

    public static DialogResult Show(string title, string promptText, ref string value,
                                    InputBoxValidation validation)
    {
        Form form = new Form();
        Label label = new Label();
        TextBox textBox = new TextBox();
        Button buttonOk = new Button();
        Button buttonCancel = new Button();

        form.Text = title;
        label.Text = promptText;
        textBox.Text = value;

        buttonOk.Text = "OK";
        buttonCancel.Text = "Cancel";
        buttonOk.DialogResult = DialogResult.OK;
        buttonCancel.DialogResult = DialogResult.Cancel;

        label.SetBounds(9, 20, 372, 13);
        textBox.SetBounds(12, 36, 372, 20);
        buttonOk.SetBounds(228, 72, 75, 23);
        buttonCancel.SetBounds(309, 72, 75, 23);

        label.AutoSize = true;
        textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
        buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

        form.ClientSize = new Size(396, 107);
        form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
        form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
        form.FormBorderStyle = FormBorderStyle.FixedDialog;
        form.StartPosition = FormStartPosition.CenterScreen;
        form.MinimizeBox = false;
        form.MaximizeBox = false;
        form.TopMost = true;
        form.AcceptButton = buttonOk;
        form.CancelButton = buttonCancel;
        if (validation != null)
        {
            form.FormClosing += delegate (object sender, FormClosingEventArgs e)
            {
                if (form.DialogResult == DialogResult.OK)
                {
                    string errorText = validation(textBox.Text);
                    if (e.Cancel = (errorText != ""))
                    {
                        MessageBox.Show(form, errorText, "Validation Error",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                        textBox.Focus();
                    }
                }
            };
        }
        DialogResult dialogResult = form.ShowDialog();
        value = textBox.Text;
        return dialogResult;
    }
}
public delegate string InputBoxValidation(string errorMessage);


public class uMessageBox
{
    public static void Show(string title = "경고", string promptText = null)
    {
        xShow(title, promptText);
        return;
    }

    public static void xShow(string title, string promptText)
    {
        MessageBox.Show(caption: title, text: promptText, buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Warning,defaultButton:MessageBoxDefaultButton.Button1,options:MessageBoxOptions.ServiceNotification);
    }
}


public class uxMessageBox : IDisposable
{
    static Button[] Buttonx;
    static Form form = new Form();
    
    public static DialogResult Show(string title, string promptText, string[] sButton, DialogResult[] Result)
    {
        DialogResult sResult = Show(title, promptText, sButton, Result, null);
        return sResult;
    }

    public static DialogResult Show(string title = "경고", string promptText = null)
    {
        DialogResult sResult = Show(title, promptText, null);
        return sResult;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
        return;
    }

    
    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {

        }
        return;
    }

    public static DialogResult Show(string title, string promptText, string[] sButton, DialogResult[] Result,
                                    uxMessageBoxValidation validation)
    {
        Label label = new Label();
        Buttonx = new Button[sButton.Length];
        PictureBox picture = new PictureBox()
        {
            Parent = form,
            Image = PSeatTester.Properties.Resources.Danger_Shield,

        };
        picture.SetBounds(9, 10, 32, 32);
                

        int i = 0;
        int bSize;

        for (i = 0; i < Buttonx.Length; i++)
        {
            Buttonx[i] = new Button()
            {
                AutoSize = false
            };
        }

        i = 0;
        foreach (string s in sButton)
        {
            Buttonx[i].Text = s;
            //Buttonx[i].AutoSize = true;
            i++;
        }
        int Max = 0;
        foreach (string s in sButton)
        {
            label.Text = s;
            if (Max < label.Width) Max = label.Right;
        }

        bSize = Max + 10;

        form.Text = title;
        label.Text = promptText;


        int Size = 0;
        for (i = 0; i < sButton.Length; i++)
        {
            Size += bSize + 8;
        }

        //label.SetBounds(9, 20, 372, 13);
        label.SetBounds(50, 20, 372, 13);
        label.Parent = form;
        label.AutoSize = true;

        Max = Math.Max(300, label.Right + 10);
        Max = Math.Max(Size, label.Right + 10);


        form.ClientSize = new Size(396 + 50, 107);
        form.ClientSize = new Size(Math.Max(300, Max) + 50, form.ClientSize.Height);

        int Widgth = (form.Width / 2) - (Size / 2);
        int X = 0;

        X = Widgth;
        for (i = 0; i < sButton.Length; i++)
        {
            Buttonx[i].SetBounds(X, 72, bSize, 23);
            X += bSize + 8;
            Buttonx[i].Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            Buttonx[i].Parent = form;
            Buttonx[i].Visible = true;
            Buttonx[i].DialogResult = Result[i];
            Buttonx[i].Click += new EventHandler(ButtonClick);
        }

        form.FormBorderStyle = FormBorderStyle.FixedDialog;
        form.StartPosition = FormStartPosition.CenterScreen;
        form.MinimizeBox = false;
        form.MaximizeBox = false;
        form.ControlBox = false;
        form.TopMost = true;

        if (validation != null)
        {
            form.FormClosing += delegate (object sender, FormClosingEventArgs e)
            {
                if (form.DialogResult == DialogResult.OK)
                {
                }
                else if (form.DialogResult == DialogResult.No)
                {
                }
                else if (form.DialogResult == DialogResult.Cancel)
                {
                }
                form.Dispose();
                form = null;
            };
        }
        DialogResult dialogResult = form.ShowDialog();
        return dialogResult;
    }

    public static DialogResult Show(string title, string promptText, uxMessageBoxValidation validation)
    {
        Label label = new Label();
        Buttonx = new Button[1];
        Buttonx[0] = new Button() { AutoSize = false };
        PictureBox picture = new PictureBox()
        {
            Parent = form,
            Image = PSeatTester.Properties.Resources.Danger_Shield

        };
        picture.SetBounds(9, 10, 32, 32);

        System.Drawing.Graphics formGraphics = form.CreateGraphics();

        float FWidth = (form.Font.SizeInPoints / 72) * formGraphics.DpiX;
        int FHeight = form.Font.Height;

        Buttonx[0].Text = "확인";

        int bSize = (int)("확인".Length * FWidth);
        int Size = bSize + 50;

        form.Text = title;
        label.Text = promptText;

        //label.SetBounds(9, 20, 372, 13);
        label.SetBounds(50, 20, 372, 13);
        label.Parent = form;
        label.AutoSize = true;

        int Max = 0;
        Max = Math.Max(300, label.Right + 10);
        Max = Math.Max(Size, label.Right + 10);


        form.ClientSize = new Size(396 + 50, 107);
        form.ClientSize = new Size(Math.Max(300, Max) + 50, form.ClientSize.Height);

        int Widgth = (form.Width / 2) - (Size / 2);
        int X = 0;

        X = Widgth;
        //for (i = 0; i < sButton.Length; i++)
        //{
        //    Buttonx[i].SetBounds(X, 72, bSize, 23);
        //    X += bSize + 8;
        //    Buttonx[i].Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        //    Buttonx[i].Parent = form;
        //    Buttonx[i].Visible = true;
        //    Buttonx[i].DialogResult = Result[i];
        //    Buttonx[i].Click += new EventHandler(ButtonClick);
        //}

        Buttonx[0].SetBounds(X, 72, Size, 23);
        X += bSize + 8;
        Buttonx[0].Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        Buttonx[0].Parent = form;
        Buttonx[0].Visible = true;
        Buttonx[0].DialogResult = DialogResult.OK;
        Buttonx[0].Click += new EventHandler(ButtonClick);

        form.FormBorderStyle = FormBorderStyle.FixedDialog;
        form.StartPosition = FormStartPosition.CenterScreen;
        form.MinimizeBox = false;
        form.MaximizeBox = false;
        form.ControlBox = false;
        form.TopMost = true;

        if (validation != null)
        {
            form.FormClosing += delegate (object sender, FormClosingEventArgs e)
            {
                if (form.DialogResult == DialogResult.OK)
                {
                }
                else if (form.DialogResult == DialogResult.No)
                {
                }
                else if (form.DialogResult == DialogResult.Cancel)
                {
                }
                form.Dispose();
            };
        }
        DialogResult dialogResult = form.ShowDialog();
        return dialogResult;
    }

    ~uxMessageBox()
    {
        
    }

    private static void ButtonClick(object sender, EventArgs e)
    {
        form.Close();        
        return;
    }

    private void FormClosing(object sender, FormClosingEventArgs e)
    {
        form.Dispose();
        form = null;
    }
}
public delegate string uxMessageBoxValidation(string errorMessage);




namespace PSeatTester
{
    //public class InputBox : IDisposable
    //{
    //    public static DialogResult Show(string title, string promptText, ref string value)
    //    {
    //        return Show(title, promptText, ref value, null);
    //    }

    //    public static DialogResult Show(string title, string promptText, ref string value,
    //                                    InputBoxValidation validation)
    //    {
    //        Form form = new Form();
    //        Label label = new Label();
    //        TextBox textBox = new TextBox();
    //        Button buttonOk = new Button();
    //        Button buttonCancel = new Button();

    //        form.Text = title;
    //        label.Text = promptText;
    //        textBox.Text = value;

    //        buttonOk.Text = "OK";
    //        buttonCancel.Text = "Cancel";
    //        buttonOk.DialogResult = DialogResult.OK;
    //        buttonCancel.DialogResult = DialogResult.Cancel;

    //        label.SetBounds(9, 20, 372, 13);
    //        textBox.SetBounds(12, 36, 372, 20);
    //        buttonOk.SetBounds(228, 72, 75, 23);
    //        buttonCancel.SetBounds(309, 72, 75, 23);

    //        label.AutoSize = true;
    //        textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
    //        buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
    //        buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

    //        form.ClientSize = new Size(396, 107);
    //        form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
    //        form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
    //        form.FormBorderStyle = FormBorderStyle.FixedDialog;
    //        form.StartPosition = FormStartPosition.CenterScreen;
    //        form.MinimizeBox = false;
    //        form.MaximizeBox = false;
    //        form.TopMost = true;
    //        form.AcceptButton = buttonOk;
    //        form.CancelButton = buttonCancel;
    //        if (validation != null)
    //        {
    //            form.FormClosing += delegate (object sender, FormClosingEventArgs e)
    //            {
    //                if (form.DialogResult == DialogResult.OK)
    //                {
    //                    string errorText = validation(textBox.Text);
    //                    if (e.Cancel = (errorText != ""))
    //                    {
    //                        MessageBox.Show(form, errorText, "Validation Error",
    //                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
    //                        textBox.Focus();
    //                    }
    //                }
    //            };
    //        }
    //        DialogResult dialogResult = form.ShowDialog();
    //        value = textBox.Text;
    //        return dialogResult;
    //    }

    //    public void Dispose()
    //    {
    //        Dispose(true);
    //        GC.SuppressFinalize(this);
    //        return;
    //    }

    //    protected virtual void Dispose(bool disposing)
    //    {
    //        if (disposing)
    //        {

    //        }
    //        return;
    //    }
    //}
    //public delegate string InputBoxValidation(string errorMessage);


    //public class uMessageBox : IDisposable
    //{
    //    static Button[] Buttonx;
    //    static Form form = new Form();

    //    public static DialogResult Show(string title, string promptText, string[] sButton, DialogResult[] Result)
    //    {
    //        DialogResult sResult = Show(title, promptText, sButton, Result, null);
    //        return sResult;
    //    }

    //    public static DialogResult Show(string title, string promptText, string[] sButton, DialogResult[] Result,
    //                                    uMessageBoxValidation validation)
    //    {
    //        Label label = new Label();
    //        Buttonx = new Button[sButton.Length];
    //        PictureBox picture = new PictureBox()
    //        {
    //            Parent = form,
    //            Image = PSeatTester.Properties.Resources.Browser,

    //        };
    //        picture.SetBounds(9, 10, 32, 32);

    //        int i = 0;
    //        int bSize;

    //        for (i = 0; i < Buttonx.Length; i++)
    //        {
    //            Buttonx[i] = new Button()
    //            {
    //                AutoSize = false
    //            };
    //        }

    //        i = 0;
    //        foreach (string s in sButton)
    //        {
    //            Buttonx[i].Text = s;
    //            //Buttonx[i].AutoSize = true;
    //            i++;
    //        }
    //        int Max = 0;
    //        foreach (string s in sButton)
    //        {
    //            label.Text = s;
    //            if (Max < label.Width) Max = label.Right;
    //        }

    //        bSize = Max + 10;

    //        form.Text = title;
    //        label.Text = promptText;


    //        int Size = 0;
    //        for (i = 0; i < sButton.Length; i++)
    //        {
    //            Size += bSize + 8;
    //        }

    //        //label.SetBounds(9, 20, 372, 13);
    //        label.SetBounds(50, 20, 372, 13);
    //        label.Parent = form;
    //        label.AutoSize = true;

    //        Max = Math.Max(300, label.Right + 10);
    //        Max = Math.Max(Size, label.Right + 10);


    //        form.ClientSize = new Size(396 + 50, 107);
    //        form.ClientSize = new Size(Math.Max(300, Max) + 50, form.ClientSize.Height);

    //        int Widgth = (form.Width / 2) - (Size / 2);
    //        int X = 0;

    //        X = Widgth;
    //        for (i = 0; i < sButton.Length; i++)
    //        {
    //            Buttonx[i].SetBounds(X, 72, bSize, 23);
    //            X += bSize + 8;
    //            Buttonx[i].Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
    //            Buttonx[i].Parent = form;
    //            Buttonx[i].Visible = true;
    //            Buttonx[i].DialogResult = Result[i];
    //            Buttonx[i].Click += new EventHandler(ButtonClick);
    //        }

    //        form.FormBorderStyle = FormBorderStyle.FixedDialog;
    //        form.StartPosition = FormStartPosition.CenterScreen;
    //        form.MinimizeBox = false;
    //        form.MaximizeBox = false;
    //        form.ControlBox = false;
    //        form.TopMost = true;

    //        if (validation != null)
    //        {
    //            form.FormClosing += delegate (object sender, FormClosingEventArgs e)
    //            {
    //                if (form.DialogResult == DialogResult.OK)
    //                {
    //                }
    //                else if (form.DialogResult == DialogResult.No)
    //                {
    //                }
    //                else if (form.DialogResult == DialogResult.Cancel)
    //                {
    //                }
    //            };
    //        }
    //        DialogResult dialogResult = form.ShowDialog();
    //        return dialogResult;
    //    }

    //    ~uMessageBox()
    //    {
    //    }

    //    private static void ButtonClick(object sender, EventArgs e)
    //    {
    //        form.Close();
    //        return;
    //    }

    //    public void Dispose()
    //    {
    //        Dispose(true);
    //        GC.SuppressFinalize(this);
    //        return;
    //    }

    //    protected virtual void Dispose(bool disposing)
    //    {
    //        if (disposing)
    //        {

    //        }
    //        return;
    //    }
    //}
    //public delegate string uMessageBoxValidation(string errorMessage);

    public class COMMON_FUCTION : IDisposable
    {
        public Stopwatch STOP_WATCH = new Stopwatch();
        private List<string> FileList = new List<string>();
        private List<string> DirList = new List<string>();


        //        private IntPtr hWnd;
        //uint vncstyles;


        public COMMON_FUCTION()
        {
            STOP_WATCH.Start();
        }
        ~COMMON_FUCTION()
        {
            STOP_WATCH.Stop();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
            return;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {

            }
            return;
        }

        /// <summary>
        /// 화면을 갱신할 때 잔상이 남지 않도록 한다 
        /// 프로그램 스타트가 되면 화면을 갱신할때 잔상이 남을 만한 컴포너트를 이곳으로 보내 속성을 바꾼다.
        /// </summary>
        /// <param name="control"></param>
        public void SetDoubleBuffered(Control control)
        {
            // set instance non-public property with name "DoubleBuffered" to true
            typeof(Control).InvokeMember("DoubleBuffered", BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.NonPublic, null, control, new object[] { true });
            return;
        }

        /*
         컴퓨터를 종료/재부팅/로그오프를 하는방법은 다음과 같습니다.

        기본적으로 Using System.Diagnostics을 사용하며,

        예제코드는 다음과 같습니다.

        종료

        Process.Start("shutdown.exe", "-s"); // 기본적으로 30초 후 종료

        Process.Start("shutdown.exe", "-s -t xx") // xx 초 후 종료



        재부팅

        Process.Start("shutdown.exe","-r"); // 종료과 유사하며 커멘드만 "-r"을 사용

        Process.Start("shutdown.exe","-r -t xx");



        로그오프

        Process.Start("shutdown.exe","-l"); // 위 코드와 유사하며 커멘드만 "-l"(숫자 1이 아닌 소문자'l')을 사용

        원하는 시간에 종료하길 원하신다면, 현재시간을 비교해 종료하는 방법이 있음.


        */

        // 모니터 On/Off 관련
        const int WM_SYSCOMMAND = 0x0112;
        const int SC_MONITORPOWER = 0xF170;
        const int MONITOR_ON = -1;
        const int MONITOR_OFF = 2;
        const int MONITOR_STANBY = 1;
        [System.Runtime.InteropServices.DllImport("User32")]
        private static extern int SendMessage(int hWnd, int hMsg, int wParam, int lParam);
        /// <summary>
        /// this.Handle 을 상수로 주어야 한다.;
        /// </summary>
        /// <param name="Handle"></param>
        public void MonitorOff(IntPtr Handle)
        {
            SendMessage(Handle.ToInt32(), WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_OFF);
            return;
        }
        /// <summary>
        /// this.Handle 을 상수로 주어야 한다.;
        /// </summary>
        /// <param name="Handle"></param>
        public void MonitorOn(IntPtr Handle)
        {
            SendMessage(Handle.ToInt32(), WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_ON);
            return;
        }

        public void ExternalProgramRun(string name) //외부 프로그램 실행
        {
            Process.Start(name);
            return;
        }

        public void ExternalProgramExit(string name) //외부 프로그램 종료
        {
            string exitname;

            if (name.IndexOf(".exe") < 0)
            {
                exitname = name;
            }
            else
            {
                exitname = name.Substring(0, name.IndexOf(".exe"));
            }
            foreach (Process process in Process.GetProcesses())
            {
                if (process.ProcessName.StartsWith(exitname))
                {
                    process.Kill();
                }
            }
        }

        /// <summary>
        /// 윈도우를 로그 오프를 한다.
        /// </summary>
        public void LogOff()
        {
            Process.Start("shutdown.exe", " -l"); // 위 코드와 유사하며 커멘드만 "-l"(숫자 1이 아닌 소문자'l')을 사용
            return;
        }
        /// <summary>
        /// 윈도우를 설정 시간후 재 부팅을 하다.
        /// </summary>
        /// <param name="Sec"></param>
        public void WindowRestartToDelay(int Sec)
        {
            string s;

            s = "shutdown.exe" + " -r -t " + Sec.ToString();
            Process.Start(s);
            return;
        }

        /// <summary>
        /// 윈도우를 30초 후 재부팅 한다.
        /// </summary>
        public void WindowRestartTo30Secconds()
        {
            Process.Start("shutdown.exe", " -r");
            return;
        }

        /// <summary>
        /// 윈도우를 바로 재부팅 한다.
        /// </summary>
        public void WindowRestartToRightNow()
        {
            Process.Start("shutdown.exe", " -r -f");
            return;
        }

        /// <summary>
        /// 윈도우 종료를 취소한다.
        /// </summary>
        public void WindowShutdownCancel()
        {
            Process.Start("shutdown.exe", " -a");
            return;
        }



        /// <summary>
        /// Sec 시간 후 종료한다.
        /// </summary>
        /// <param name="Sec"></param>
        public void WindowExitToDelay(int Sec)
        {
            string s;

            s = "shutdown.exe" + " -s -t " + Sec.ToString();
            Process.Start(s);
            return;
        }

        /// <summary>
        /// PC 종료(프로그램 정리후 종료)
        /// </summary>
        public void WindowExitTo30Secconds()
        {
            Process.Start("shutdown.exe", " -s");
            return;
        }

        /// <summary>
        /// PC 바로 종료
        /// </summary>
        public void WindowExit()
        {
            Process.Start("shutdown.exe", " -s -f");
            return;
        }


        /// <summary>
        /// 시작 프로그램을 레지스터리에 등록한다.
        /// </summary>
        public void StartProgramRegistrySet()
        {
            var rkApp1 = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            var sName = rkApp1.GetValue("name");

            if (sName != null)
            {
                if (sName.ToString() != System.IO.Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName))
                {
                    var rkApp2 = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
                    //rkApp2.SetValue("name", System.IO.Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName));
                    rkApp2.SetValue("name", Process.GetCurrentProcess().MainModule.FileName);
                }
            }
            else
            {
                var rkApp2 = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
                //rkApp2.SetValue("name", System.IO.Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName));
                rkApp2.SetValue("name", Process.GetCurrentProcess().MainModule.FileName);
            }
            return;
        }

        /// <summary>
        /// 레지스터리에 등록된 시작 프로그래을 삭제한다.
        /// </summary>
        public void StartProgramRegistryDel()
        {
            var rkApp = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            rkApp.DeleteValue("name");
            return;
        }

        /// <summary>
        /// Property 에 변수를 만든다.
        /// Type 은 typeof(string), typeof(int) 와 같이 형을 만들수 있다.
        /// </summary>
        /// <param name="Name"></param>
        /// <param name="Defalut"></param>
        /// <param name="sType"></param>
        //public void SetProperty(string Name, Type sType)
        //{
        //    System.Configuration.SettingsProperty property = new System.Configuration.SettingsProperty(Name);

        //    //if(sType == typeof(string))
        //    //    property.DefaultValue = "";
        //    //else if (sType == typeof(int))
        //    //    property.DefaultValue = 0;
        //    //else if (sType == typeof(short))
        //    //    property.DefaultValue = 0;
        //    //else if (sType == typeof(long))
        //    //    property.DefaultValue = 0;

        //    property.IsReadOnly = false;
        //    property.PropertyType = sType;
        //    property.Provider = Settings.Default.Providers["LocalFileSettingsProvider"];
        //    property.Attributes.Add(typeof(System.Configuration.UserScopedSettingAttribute), new System.Configuration.UserScopedSettingAttribute());
        //    try 
        //    {
        //        Properties.Settings.Default.Properties.Add(property);
        //    }
        //    catch
        //    {
        //        Properties.Settings.Default.Properties.Remove(property.Name);
        //        Properties.Settings.Default.Properties.Add(property);
        //    }

        //    return;
        //}


        public List<string> GetFileList
        {
            get
            {
                return FileList;
            }
        }

        public List<string> GetDirList
        {
            get
            {
                return DirList;
            }
        }


        [StructLayout(LayoutKind.Explicit)]
        private struct union_r
        {
            [FieldOffset(0)]
            public ushort i;
            [FieldOffset(0)]
            public byte c1;
            [FieldOffset(1)]
            public byte c2;
        };

        public class FileSortMode
        {
            public const short LASTWRITE_ODERBY = 0; //마지막 저장 날짜 기준 내림차순
            public const short LASTWRITE_ORDERBYDESCENDING = 1;//마지막 저장 날짜 기준 오름차순
            public const short FILENAME_ODERBY = 2; //파일명 기준으로 내림차순
            public const short FILENAME_ORDERBYDESCENDIN = 3; //파일명 기준으로 오름 차순
            public const short CREATETIME_ODERBY = 2; //생성날짜 기준으로 내림차순
            public const short CREATETIME_ORDERBYDESCENDIN = 3; //생성날짜 기준으로 오름 차순
            public const short NORMAL = 4;
        }

        public void ReadFileList(string Path, string Ext, short Mode = FileSortMode.NORMAL)
        {
            try
            {
                DirectoryInfo DirInf = new DirectoryInfo(Path);
                if (DirInf.Exists == true) //폴더가 존제하면
                {
                    //FileInfo[] FileInf = DirInf.GetFiles("*.*");

                    if (Mode == FileSortMode.NORMAL)
                    {
                        FileInfo[] FileInf = DirInf.GetFiles(Ext);

                        if (FileInf.Length == 0)
                        {
                            //MessageBox.Show("파일이 존재하지 않습니다.");
                            uMessageBox.Show(title: "경고", promptText:"파일이 존재하지 않습니다.");
                        }
                        else
                        {
                            string s = "";
                            for (int i = 0; i < FileInf.Length; i++)
                            {
                                s += FileInf[i].Name.ToString() + Environment.NewLine;

                                FileList.Add(s);
                                s = "";
                            }
                        }
                    }
                    else
                    {
                        //

                        //방식 1
                        //var Files = new DirectoryInfo(FileDirectory)
                        //.GetFiles()
                        //.OrderBy(f => f.Name.Substring(f.Name.Length - 23, 19)
                        //.ToList();

                        //방식 2
                        //생성 날짜를 기준으로 정렬한다.
                        //내림차순
                        //var FileInf = DirInf.GetFiles(Ext).OrderBy(f => f.LastWriteTime).ToList();
                        //오름차순

                        FileList.Clear();
                        string s = "";
                        if (Mode == FileSortMode.LASTWRITE_ODERBY)//마지막 저장 날짜 기준 오름차순
                        {
                            var FileInf = DirInf.GetFiles(Ext).OrderBy(f => f.LastWriteTime).ToList();
                            foreach (FileInfo fs in FileInf)
                            {
                                s = "";
                                s = fs.Name.ToString() + Environment.NewLine;

                                FileList.Add(s);
                            }
                        }
                        else if (Mode == FileSortMode.LASTWRITE_ORDERBYDESCENDING)//마지막 저장 날짜 기준 내림차순
                        {
                            var FileInf = DirInf.GetFiles(Ext).OrderByDescending(f => f.LastWriteTime).ToList();
                            foreach (FileInfo fs in FileInf)
                            {
                                s = "";
                                s = fs.Name.ToString() + Environment.NewLine;

                                FileList.Add(s);
                            }
                        }
                        else if (Mode == FileSortMode.FILENAME_ODERBY) //파일명을 기준으로 내림차순
                        {
                            var FileInf = DirInf.GetFiles(Ext).OrderBy(f => f.Name).ToList();
                            foreach (FileInfo fs in FileInf)
                            {
                                s = "";
                                s = fs.Name.ToString() + Environment.NewLine;

                                FileList.Add(s);
                            }
                        }
                        else if (Mode == FileSortMode.FILENAME_ORDERBYDESCENDIN)//파일명을 기준으로 오름 차순
                        {
                            var FileInf = DirInf.GetFiles(Ext).OrderByDescending(f => f.Name).ToList();
                            foreach (FileInfo fs in FileInf)
                            {
                                s = "";
                                s = fs.Name.ToString() + Environment.NewLine;

                                FileList.Add(s);
                            }
                        }
                        else if (Mode == FileSortMode.FILENAME_ODERBY) //파일명을 기준으로 내림차순
                        {
                            var FileInf = DirInf.GetFiles(Ext).OrderBy(f => f.CreationTime).ToList();
                            foreach (FileInfo fs in FileInf)
                            {
                                s = "";
                                s = fs.Name.ToString() + Environment.NewLine;

                                FileList.Add(s);
                            }
                        }
                        else if (Mode == FileSortMode.FILENAME_ORDERBYDESCENDIN)//파일명을 기준으로 오름 차순
                        {
                            var FileInf = DirInf.GetFiles(Ext).OrderByDescending(f => f.CreationTime).ToList();
                            foreach (FileInfo fs in FileInf)
                            {
                                s = "";
                                s = fs.Name.ToString() + Environment.NewLine;

                                FileList.Add(s);
                            }
                        }
                    }
                }
                else
                {
                    FileList.Clear();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
                //uMessageBox.Show(promptText: ex.Message + "\n" + ex.StackTrace);
            }
        }
        public void ReadFileListNotExt(string Path, string Ext, short Mode = FileSortMode.NORMAL)
        {
            try
            {
                DirectoryInfo DirInf = new DirectoryInfo(Path);
                if (DirInf.Exists == true) //폴더가 존제하면
                {
                    //FileInfo[] FileInf = DirInf.GetFiles("*.*");

                    if (Mode == FileSortMode.NORMAL)
                    {
                        FileInfo[] FileInf = DirInf.GetFiles(Ext);

                        FileList.Clear();
                        if (FileInf.Length == 0)
                        {
                            //MessageBox.Show("파일이 존재하지 않습니다.");
                            uMessageBox.Show(promptText: "파일이 존재하지 않습니다.");
                        }
                        else
                        {
                            string s = "";
                            for (int i = 0; i < FileInf.Length; i++)
                            {
                                s += FileInf[i].Name.ToString() + Environment.NewLine;
                                string[] t = s.Split('.');
                                FileList.Add(t[0]);
                                s = "";
                            }
                        }
                    }
                    else
                    {
                        //

                        //방식 1
                        //var Files = new DirectoryInfo(FileDirectory)
                        //.GetFiles()
                        //.OrderBy(f => f.Name.Substring(f.Name.Length - 23, 19)
                        //.ToList();

                        //방식 2
                        //생성 날짜를 기준으로 정렬한다.
                        //내림차순
                        //var FileInf = DirInf.GetFiles(Ext).OrderBy(f => f.LastWriteTime).ToList();
                        //오름차순

                        FileList.Clear();
                        string s = "";
                        if (Mode == FileSortMode.LASTWRITE_ODERBY)//마지막 저장 날짜 기준 오름차순
                        {
                            var FileInf = DirInf.GetFiles(Ext).OrderBy(f => f.LastWriteTime).ToList();
                            foreach (FileInfo fs in FileInf)
                            {
                                s = "";
                                s = fs.Name.ToString() + Environment.NewLine;

                                string[] t = s.Split('.');
                                FileList.Add(t[0]);
                            }
                        }
                        else if (Mode == FileSortMode.LASTWRITE_ORDERBYDESCENDING)//마지막 저장 날짜 기준 내림차순
                        {
                            var FileInf = DirInf.GetFiles(Ext).OrderByDescending(f => f.LastWriteTime).ToList();
                            foreach (FileInfo fs in FileInf)
                            {
                                s = "";
                                s = fs.Name.ToString() + Environment.NewLine;

                                string[] t = s.Split('.');
                                FileList.Add(t[0]);
                            }
                        }
                        else if (Mode == FileSortMode.FILENAME_ODERBY) //파일명을 기준으로 내림차순
                        {
                            var FileInf = DirInf.GetFiles(Ext).OrderBy(f => f.Name).ToList();
                            foreach (FileInfo fs in FileInf)
                            {
                                s = "";
                                s = fs.Name.ToString() + Environment.NewLine;

                                string[] t = s.Split('.');
                                FileList.Add(t[0]);
                            }
                        }
                        else if (Mode == FileSortMode.FILENAME_ORDERBYDESCENDIN)//파일명을 기준으로 오름 차순
                        {
                            var FileInf = DirInf.GetFiles(Ext).OrderByDescending(f => f.Name).ToList();
                            foreach (FileInfo fs in FileInf)
                            {
                                s = "";
                                s = fs.Name.ToString() + Environment.NewLine;

                                string[] t = s.Split('.');
                                FileList.Add(t[0]);
                            }
                        }
                        else if (Mode == FileSortMode.FILENAME_ODERBY) //파일명을 기준으로 내림차순
                        {
                            var FileInf = DirInf.GetFiles(Ext).OrderBy(f => f.CreationTime).ToList();
                            foreach (FileInfo fs in FileInf)
                            {
                                s = "";
                                s = fs.Name.ToString() + Environment.NewLine;

                                string[] t = s.Split('.');
                                FileList.Add(t[0]);
                            }
                        }
                        else if (Mode == FileSortMode.FILENAME_ORDERBYDESCENDIN)//파일명을 기준으로 오름 차순
                        {
                            var FileInf = DirInf.GetFiles(Ext).OrderByDescending(f => f.CreationTime).ToList();
                            foreach (FileInfo fs in FileInf)
                            {
                                s = "";
                                s = fs.Name.ToString() + Environment.NewLine;

                                string[] t = s.Split('.');
                                FileList.Add(t[0]);
                            }
                        }
                    }
                }
                else
                {
                    FileList.Clear();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
                //uMessageBox.Show(promptText: ex.Message + "\n" + ex.StackTrace);
            }
        }

        public void ReadDirList(string Path)
        {
            try
            {
                DirectoryInfo DirInf = new DirectoryInfo(Path + "\\");
                if (DirInf.Exists == true) //폴더가 존제하면
                {
                    DirectoryInfo[] DirListData = DirInf.GetDirectories();

                    DirList.Clear();
                    if (DirListData.Length == 0)
                    {
                        //MessageBox.Show("폴더가 존재하지 않습니다.");
                        uMessageBox.Show(promptText: "파일이 존재하지 않습니다.");
                    }
                    else
                    {
                        string s = "";
                        for (int i = 0; i < DirListData.Length; i++)
                        {
                            s += DirListData[i].Name.ToString();
                            //string[] t = s.Split('.');
                            //DirList.Add(t[0]);
                            string[] t = s.Split('\r');
                            DirList.Add(t[0]);
                            s = "";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
                //uMessageBox.Show(promptText: ex.Message + "\n" + ex.StackTrace);
            }
        }

        

        public __Infor__ OpenInfor(string Name,ref __Infor__ Inf)
        {
            if (File.Exists(Name) == true)
            {
                TIniFile Ini = new TIniFile(Name);

            }
            else
            {
                
            }

            return Inf;
        }

        public void SaveSpec(string Name, __Spec__ Spec)
        {
            TIniFile Ini = new TIniFile(Name);

            Ini.WriteString("MODEL", "VALUE", Spec.ModelName);

            Ini.WriteFloat("CURRENT", "IMS MIN", Spec.Current.IMS.Min);
            Ini.WriteFloat("CURRENT", "IMS MAX", Spec.Current.IMS.Max);

            Ini.WriteFloat("CURRENT", "SLIDE MIN", Spec.Current.Slide.Min);
            Ini.WriteFloat("CURRENT", "SLIDE MAX", Spec.Current.Slide.Max);

            Ini.WriteFloat("CURRENT", "RECLINE MIN", Spec.Current.Recliner.Min);
            Ini.WriteFloat("CURRENT", "RECLINE MAX", Spec.Current.Recliner.Max);

            Ini.WriteFloat("CURRENT", "TILT MIN", Spec.Current.Tilt.Min);
            Ini.WriteFloat("CURRENT", "TILT MAX", Spec.Current.Tilt.Max);

            Ini.WriteFloat("CURRENT", "HEIGHT MIN", Spec.Current.Height.Min);
            Ini.WriteFloat("CURRENT", "HEIGHT MAX", Spec.Current.Height.Max);
                        
            Ini.WriteFloat("CURRENT", "LUMBER_FWDBWD MIN", Spec.Current.LumberFwdBwd.Min);
            Ini.WriteFloat("CURRENT", "LUMBER_FWDBWD MAX", Spec.Current.LumberFwdBwd.Max);
            
            Ini.WriteFloat("CURRENT", "LUMBER_UPDN MIN", Spec.Current.LumberUpDn.Min);
            Ini.WriteFloat("CURRENT", "LUMBER_UPDN MAX", Spec.Current.LumberUpDn.Max);


            Ini.WriteFloat("MOVE SPEED", "SLIDE FWD MIN", Spec.MovingSpeed.SlideFwd.Min);
            Ini.WriteFloat("MOVE SPEED", "SLIDE FWD MAX", Spec.MovingSpeed.SlideFwd.Max);
            Ini.WriteFloat("MOVE SPEED", "SLIDE BWD MIN", Spec.MovingSpeed.SlideBwd.Min);
            Ini.WriteFloat("MOVE SPEED", "SLIDE BWD MAX", Spec.MovingSpeed.SlideBwd.Max);
             
            Ini.WriteFloat("MOVE SPEED", "RECLINE FWD MIN", Spec.MovingSpeed.ReclinerFwd.Min);
            Ini.WriteFloat("MOVE SPEED", "RECLINE FWD MAX", Spec.MovingSpeed.ReclinerFwd.Max);
            Ini.WriteFloat("MOVE SPEED", "RECLINE BWD MIN", Spec.MovingSpeed.ReclinerBwd.Min);
            Ini.WriteFloat("MOVE SPEED", "RECLINE BWD MAX", Spec.MovingSpeed.ReclinerBwd.Max);

            Ini.WriteFloat("MOVE SPEED", "TILT UP MIN", Spec.MovingSpeed.TiltUp.Min);
            Ini.WriteFloat("MOVE SPEED", "TILT UP MAX", Spec.MovingSpeed.TiltUp.Max);
            Ini.WriteFloat("MOVE SPEED", "TILT DN MIN", Spec.MovingSpeed.TiltDn.Min);
            Ini.WriteFloat("MOVE SPEED", "TILT DN MAX", Spec.MovingSpeed.TiltDn.Max);

            Ini.WriteFloat("MOVE SPEED", "HEIGHT UP MIN", Spec.MovingSpeed.HeightUp.Min);
            Ini.WriteFloat("MOVE SPEED", "HEIGHT UP MAX", Spec.MovingSpeed.HeightUp.Max);
            Ini.WriteFloat("MOVE SPEED", "HEIGHT DN MIN", Spec.MovingSpeed.HeightDn.Min);
            Ini.WriteFloat("MOVE SPEED", "HEIGHT DN MAX", Spec.MovingSpeed.HeightDn.Max);

            Ini.WriteFloat("SOUND", "SPEC", Spec.Sound.RunMax);
            Ini.WriteFloat("SOUND", "START SPEC", Spec.Sound.StartMax);
            Ini.WriteBool("SOUND", "ARV MODE", Spec.Sound.RMSMode);
            Ini.WriteFloat("SOUND", "구분시간", Spec.Sound.기동음구동음구분사이시간);
            Ini.WriteFloat("SOUND", "OFFSET", Spec.Sound.Offset);

            Ini.WriteFloat("LIMIT TIME", "SLIDE", Spec.SlideLimitTime);
            Ini.WriteFloat("LIMIT TIME", "RECLINE", Spec.ReclinerLimitTime);
            Ini.WriteFloat("LIMIT TIME", "TILT", Spec.TiltLimitTime);
            Ini.WriteFloat("LIMIT TIME", "HEIGHT", Spec.HeightLimitTime);
            Ini.WriteFloat("LIMIT TIME", "LUMBER", Spec.LumberLimitTime);

            //Ini.WriteBool("CAN MODE", "VALUE", Spec.Can);

            Ini.WriteFloat("MOVE STROKE", "SLIDE", Spec.MovingStroke.Slide);
            Ini.WriteFloat("MOVE STROKE", "RECLINE", Spec.MovingStroke.Recliner);
            Ini.WriteFloat("MOVE STROKE", "TILT", Spec.MovingStroke.Tilt);
            Ini.WriteFloat("MOVE STROKE", "HEIGHT", Spec.MovingStroke.Height);

            Ini.WriteInteger("납입 위치", "SLIDE", Spec.DeliveryPos.Slide);
            Ini.WriteInteger("납입 위치", "RECLINE", Spec.DeliveryPos.Recliner);
            Ini.WriteInteger("납입 위치", "TILT", Spec.DeliveryPos.Tilt);
            Ini.WriteInteger("납입 위치", "HEIGHT", Spec.DeliveryPos.Height);

            Ini.WriteFloat("SOUND CHECK", "START CHECK TIME", Spec.Sound.StartTime);
            Ini.WriteFloat("SOUND CHECK", "SLIDE TIME", Spec.SlideSoundCheckTime);
            Ini.WriteFloat("SOUND CHECK", "RECLINE TIME", Spec.ReclinerSoundCheckTime);
            Ini.WriteFloat("SOUND CHECK", "TILT TIME", Spec.TiltSoundCheckTime);
            Ini.WriteFloat("SOUND CHECK", "HEIGHT TIME", Spec.HeightSoundCheckTime);
            Ini.WriteFloat("SOUND CHECK", "LUMBER TIME", Spec.LumberSoundCheckTime);
            //Ini.WriteFloat("SOUND CHECK", "SOUNDCHECKTIMERANGE", Spec.Sound.SoundCheckRange);

            Ini.WriteFloat("MOVE TIME", "SLIDE", Spec.SlideTestTime);
            Ini.WriteFloat("MOVE TIME", "RECLINE", Spec.ReclinerTestTime);
            Ini.WriteFloat("MOVE TIME", "TILT", Spec.TiltTestTime);
            Ini.WriteFloat("MOVE TIME", "HEIGHT", Spec.HeightTestTime);
            Ini.WriteFloat("MOVE TIME", "LUMBERFWDBWD", Spec.LumberFwdBwdTestTime);
            Ini.WriteFloat("MOVE TIME", "LUMBERUPDN", Spec.LumberUpDnTestTime);

            Ini.WriteFloat("LIMIT CURR", "SLIDE", Spec.SlideLimitCurr);
            Ini.WriteFloat("LIMIT CURR", "RECLINE", Spec.ReclinerLimitCurr);
            Ini.WriteFloat("LIMIT CURR", "TILT", Spec.TiltLimitCurr);
            Ini.WriteFloat("LIMIT CURR", "HEIGHT", Spec.HeightLimitCurr);

            Ini.WriteFloat("TEST VOLT", "VALUE", Spec.TestVolt);


            //---------------------------------------------

            Ini.WriteInteger("PIN MAP", "SLIDE FWD MODE", Spec.PinMap.SlideFWD.Mode);
            Ini.WriteInteger("PIN MAP", "SLIDE FWD PIN", Spec.PinMap.SlideFWD.PinNo);

            Ini.WriteInteger("PIN MAP", "SLIDE BWD MODE", Spec.PinMap.SlideBWD.Mode);
            Ini.WriteInteger("PIN MAP", "SLIDE BWD PIN", Spec.PinMap.SlideBWD.PinNo);

            Ini.WriteInteger("PIN MAP", "RECLINE FWD MODE", Spec.PinMap.ReclineFWD.Mode);
            Ini.WriteInteger("PIN MAP", "RECLINE FWD PIN", Spec.PinMap.ReclineFWD.PinNo);

            Ini.WriteInteger("PIN MAP", "RECLINE BWD MODE", Spec.PinMap.ReclineBWD.Mode);
            Ini.WriteInteger("PIN MAP", "RECLINE BWD PIN", Spec.PinMap.ReclineBWD.PinNo);

            Ini.WriteInteger("PIN MAP", "TILT UP MODE", Spec.PinMap.TiltUp.Mode);
            Ini.WriteInteger("PIN MAP", "TILT UP PIN", Spec.PinMap.TiltUp.PinNo);

            Ini.WriteInteger("PIN MAP", "TILT DN MODE", Spec.PinMap.TiltDn.Mode);
            Ini.WriteInteger("PIN MAP", "TILT DN PIN", Spec.PinMap.TiltDn.PinNo);

            Ini.WriteInteger("PIN MAP", "HEIGHT UP MODE", Spec.PinMap.HeightUp.Mode);
            Ini.WriteInteger("PIN MAP", "HEIGHT UP PIN", Spec.PinMap.HeightUp.PinNo);

            Ini.WriteInteger("PIN MAP", "HEIGHT DN MODE", Spec.PinMap.HeightDn.Mode);
            Ini.WriteInteger("PIN MAP", "HEIGHT DN PIN", Spec.PinMap.HeightDn.PinNo);

            Ini.WriteInteger("PIN MAP", "LUMBER FWD BATT", Spec.PinMap.LumberFwdBwd.Batt);
            Ini.WriteInteger("PIN MAP", "LUMBER FWD GND", Spec.PinMap.LumberFwdBwd.Gnd);

            Ini.WriteInteger("PIN MAP", "LUMBER UP BATT", Spec.PinMap.LumberUpDn.Batt);
            Ini.WriteInteger("PIN MAP", "LUMBER UP GND", Spec.PinMap.LumberUpDn.Gnd);

            Ini.WriteInteger("PIN MAP", "POWER1 BATT", Spec.PinMap.Power.Batt1.Batt);
            Ini.WriteInteger("PIN MAP", "POWER1 GND", Spec.PinMap.Power.Batt1.Gnd);

            Ini.WriteInteger("PIN MAP", "POWER2 BATT", Spec.PinMap.Power.Batt2.Batt);
            Ini.WriteInteger("PIN MAP", "POWER2 GND", Spec.PinMap.Power.Batt2.Gnd);

            Ini.WriteBool("PIN MAP", "직구동", Spec.PinMap.PSeat_직구동);
            Ini.WriteBool("PIN MAP", "WALKIN", Spec.PinMap.WalkIn);
            return;
        }

        public void OpenSpec(string Name, ref __Spec__ Spec)
        {
            TIniFile Ini = new TIniFile(Name);

            if (Ini.ReadString("MODEL", "VALUE", ref Spec.ModelName) == false) Spec.ModelName = string.Empty;

            if (Ini.ReadDouble("CURRENT", "IMS MIN", ref Spec.Current.IMS.Min) == false) Spec.Current.IMS.Min = 0;
            if (Ini.ReadDouble("CURRENT", "IMS MAX", ref Spec.Current.IMS.Max) == false) Spec.Current.IMS.Max = 0;

            if (Ini.ReadDouble("CURRENT", "SLIDE MIN", ref Spec.Current.Slide.Min) == false) Spec.Current.Slide.Min = 0;
            if (Ini.ReadDouble("CURRENT", "SLIDE MAX", ref Spec.Current.Slide.Max) == false) Spec.Current.Slide.Max = 0;
            if (Ini.ReadDouble("CURRENT", "RECLINE MIN", ref Spec.Current.Recliner.Min) == false) Spec.Current.Recliner.Min = 0;
            if (Ini.ReadDouble("CURRENT", "RECLINE MAX", ref Spec.Current.Recliner.Max) == false) Spec.Current.Recliner.Max = 0;

            if (Ini.ReadDouble("CURRENT", "TILT MIN", ref Spec.Current.Tilt.Min) == false) Spec.Current.Tilt.Min = 0;
            if (Ini.ReadDouble("CURRENT", "TILT MAX", ref Spec.Current.Tilt.Max) == false) Spec.Current.Tilt.Max = 0;

            if (Ini.ReadDouble("CURRENT", "HEIGHT MIN", ref Spec.Current.Height.Min) == false) Spec.Current.Height.Min = 0;
            if (Ini.ReadDouble("CURRENT", "HEIGHT MAX", ref Spec.Current.Height.Max) == false) Spec.Current.Height.Max = 0;

            if (Ini.ReadDouble("CURRENT", "LUMBER_FWDBWD MIN", ref Spec.Current.LumberFwdBwd.Min) == false) Spec.Current.LumberFwdBwd.Min = 0;
            if (Ini.ReadDouble("CURRENT", "LUMBER_FWDBWD MAX", ref Spec.Current.LumberFwdBwd.Max) == false) Spec.Current.LumberFwdBwd.Max = 0;

            if (Ini.ReadDouble("CURRENT", "LUMBER_UPDN MIN", ref Spec.Current.LumberUpDn.Min) == false) Spec.Current.LumberUpDn.Min = 0;
            if (Ini.ReadDouble("CURRENT", "LUMBER_UPDN MAX", ref Spec.Current.LumberUpDn.Max) == false) Spec.Current.LumberUpDn.Max = 0;

            if (Ini.ReadDouble("MOVE SPEED", "SLIDE FWD MIN", ref Spec.MovingSpeed.SlideFwd.Min) == false) Spec.MovingSpeed.SlideFwd.Min = 0;
            if (Ini.ReadDouble("MOVE SPEED", "SLIDE FWD MAX", ref Spec.MovingSpeed.SlideFwd.Max) == false) Spec.MovingSpeed.SlideFwd.Max = 0;
            if (Ini.ReadDouble("MOVE SPEED", "SLIDE BWD MIN", ref Spec.MovingSpeed.SlideBwd.Min) == false) Spec.MovingSpeed.SlideBwd.Min = 0;
            if (Ini.ReadDouble("MOVE SPEED", "SLIDE BWD MAX", ref Spec.MovingSpeed.SlideBwd.Max) == false) Spec.MovingSpeed.SlideBwd.Max = 0;

            if (Ini.ReadDouble("MOVE SPEED", "RECLINE FWD MIN", ref Spec.MovingSpeed.ReclinerFwd.Min) == false) Spec.MovingSpeed.ReclinerFwd.Min = 0;
            if (Ini.ReadDouble("MOVE SPEED", "RECLINE FWD MAX", ref Spec.MovingSpeed.ReclinerFwd.Max) == false) Spec.MovingSpeed.ReclinerFwd.Max = 0;
            if (Ini.ReadDouble("MOVE SPEED", "RECLINE BWD MIN", ref Spec.MovingSpeed.ReclinerBwd.Min) == false) Spec.MovingSpeed.ReclinerBwd.Min = 0;
            if (Ini.ReadDouble("MOVE SPEED", "RECLINE BWD MAX", ref Spec.MovingSpeed.ReclinerBwd.Max) == false) Spec.MovingSpeed.ReclinerBwd.Max = 0;
             
            if (Ini.ReadDouble("MOVE SPEED", "TILT UP MIN", ref Spec.MovingSpeed.TiltUp.Min) == false) Spec.MovingSpeed.TiltUp.Min = 0;
            if (Ini.ReadDouble("MOVE SPEED", "TILT UP MAX", ref Spec.MovingSpeed.TiltUp.Max) == false) Spec.MovingSpeed.TiltUp.Max = 0;
            if (Ini.ReadDouble("MOVE SPEED", "TILT DN MIN", ref Spec.MovingSpeed.TiltDn.Min) == false) Spec.MovingSpeed.TiltDn.Min = 0;
            if (Ini.ReadDouble("MOVE SPEED", "TILT DN MAX", ref Spec.MovingSpeed.TiltDn.Max) == false) Spec.MovingSpeed.TiltDn.Max = 0;

            if (Ini.ReadDouble("MOVE SPEED", "HEIGHT UP MIN", ref Spec.MovingSpeed.HeightUp.Min) == false) Spec.MovingSpeed.HeightUp.Min = 0;
            if (Ini.ReadDouble("MOVE SPEED", "HEIGHT UP MAX", ref Spec.MovingSpeed.HeightUp.Max) == false) Spec.MovingSpeed.HeightUp.Max = 0;
            if (Ini.ReadDouble("MOVE SPEED", "HEIGHT DN MIN", ref Spec.MovingSpeed.HeightDn.Min) == false) Spec.MovingSpeed.HeightDn.Min = 0;
            if (Ini.ReadDouble("MOVE SPEED", "HEIGHT DN MAX", ref Spec.MovingSpeed.HeightDn.Max) == false) Spec.MovingSpeed.HeightDn.Max = 0;

            if (Ini.ReadFloat("SOUND", "SPEC", ref Spec.Sound.RunMax) == false) Spec.Sound.RunMax = 0;
            if (Ini.ReadFloat("SOUND", "START SPEC", ref Spec.Sound.StartMax) == false) Spec.Sound.StartMax = 0;
            if (Ini.ReadBool("SOUND", "ARV MODE", ref Spec.Sound.RMSMode) == false) Spec.Sound.RMSMode = false;
            if (Ini.ReadFloat("SOUND", "구분시간", ref Spec.Sound.기동음구동음구분사이시간) == false) Spec.Sound.기동음구동음구분사이시간 = 0;
            if (Ini.ReadFloat("SOUND", "OFFSET", ref Spec.Sound.Offset) == false) Spec.Sound.Offset = 0;

            if (Ini.ReadFloat("LIMIT TIME", "SLIDE", ref Spec.SlideLimitTime) == false) Spec.SlideLimitTime = 0;
            if (Ini.ReadFloat("LIMIT TIME", "RECLINE", ref Spec.ReclinerLimitTime) == false) Spec.ReclinerLimitTime = 0;
            if (Ini.ReadFloat("LIMIT TIME", "TILT", ref Spec.TiltLimitTime) == false) Spec.TiltLimitTime = 0;
            if (Ini.ReadFloat("LIMIT TIME", "HEIGHT", ref Spec.HeightLimitTime) == false) Spec.HeightLimitTime = 0;
            if (Ini.ReadFloat("LIMIT TIME", "LUMBER", ref Spec.LumberLimitTime) == false) Spec.LumberLimitTime = 0;
            
            //if (Ini.ReadBool("CAN MODE", "VALUE", ref Spec.Can) == false) Spec.Can = false;

            if (Ini.ReadFloat("MOVE STROKE", "SLIDE", ref Spec.MovingStroke.Slide) == false) Spec.MovingStroke.Slide = 0;
            if (Ini.ReadFloat("MOVE STROKE", "RECLINE", ref Spec.MovingStroke.Recliner) == false) Spec.MovingStroke.Recliner = 0;
            if (Ini.ReadFloat("MOVE STROKE", "TILT", ref Spec.MovingStroke.Tilt) == false) Spec.MovingStroke.Tilt = 0;
            if (Ini.ReadFloat("MOVE STROKE", "HEIGHT", ref Spec.MovingStroke.Height) == false) Spec.MovingStroke.Height = 0;

            if (Ini.ReadShort("납입 위치", "SLIDE", ref Spec.DeliveryPos.Slide) == false) Spec.DeliveryPos.Slide = 0;
            if (Ini.ReadShort("납입 위치", "RECLINE", ref Spec.DeliveryPos.Recliner) == false) Spec.DeliveryPos.Recliner = 0;
            if (Ini.ReadShort("납입 위치", "TILT", ref Spec.DeliveryPos.Tilt) == false) Spec.DeliveryPos.Tilt = 0;
            if (Ini.ReadShort("납입 위치", "HEIGHT", ref Spec.DeliveryPos.Height) == false) Spec.DeliveryPos.Height = 0;

            if (Ini.ReadFloat("SOUND CHECK", "START CHECK TIME", ref Spec.Sound.StartTime) == false) Spec.Sound.StartTime = 0;            
            if (Ini.ReadFloat("SOUND CHECK", "SLIDE TIME", ref Spec.SlideSoundCheckTime) == false) Spec.SlideSoundCheckTime = 0;
            if (Ini.ReadFloat("SOUND CHECK", "RECLINE TIME", ref Spec.ReclinerSoundCheckTime) == false) Spec.ReclinerSoundCheckTime = 0;
            if (Ini.ReadFloat("SOUND CHECK", "TILT TIME", ref Spec.TiltSoundCheckTime) == false) Spec.TiltSoundCheckTime = 0;
            if (Ini.ReadFloat("SOUND CHECK", "HEIGHT TIME", ref Spec.HeightSoundCheckTime) == false) Spec.HeightSoundCheckTime = 0;
            if (Ini.ReadFloat("SOUND CHECK", "LUMBER TIME", ref Spec.LumberSoundCheckTime) == false) Spec.LumberSoundCheckTime = 0;
            //if (Ini.ReadFloat("SOUND CHECK", "SOUNDCHECKTIMERANGE", ref Spec.Sound.SoundCheckRange) == false) Spec.Sound.SoundCheckRange = 0;

            if (Ini.ReadFloat("MOVE TIME", "SLIDE", ref Spec.SlideTestTime) == false) Spec.SlideTestTime = 0;
            if (Ini.ReadFloat("MOVE TIME", "RECLINE", ref Spec.ReclinerTestTime) == false) Spec.ReclinerTestTime = 0;
            if (Ini.ReadFloat("MOVE TIME", "TILT", ref Spec.TiltTestTime) == false) Spec.TiltTestTime = 0;
            if (Ini.ReadFloat("MOVE TIME", "HEIGHT", ref Spec.HeightTestTime) == false) Spec.HeightTestTime = 0;
            if (Ini.ReadFloat("MOVE TIME", "LUMBERFWDBWD", ref Spec.LumberFwdBwdTestTime) == false) Spec.LumberFwdBwdTestTime = 0;
            if (Ini.ReadFloat("MOVE TIME", "LUMBERUPDN", ref Spec.LumberUpDnTestTime) == false) Spec.LumberUpDnTestTime = 0;

            if (Ini.ReadFloat("LIMIT CURR", "SLIDE", ref Spec.SlideLimitCurr) == false) Spec.SlideLimitCurr = 0;
            if (Ini.ReadFloat("LIMIT CURR", "RECLINE", ref Spec.ReclinerLimitCurr) == false) Spec.ReclinerLimitCurr = 0;
            if (Ini.ReadFloat("LIMIT CURR", "TILT", ref Spec.TiltLimitCurr) == false) Spec.TiltLimitCurr = 0;
            if (Ini.ReadFloat("LIMIT CURR", "HEIGHT", ref Spec.HeightLimitCurr) == false) Spec.HeightLimitCurr = 0;

            if (Ini.ReadFloat("TEST VOLT", "VALUE", ref Spec.TestVolt) == false) Spec.TestVolt = 0;

            //---------------------------------------------

            if (Ini.ReadInteger("PIN MAP", "SLIDE FWD MODE", ref Spec.PinMap.SlideFWD.Mode) == false) Spec.PinMap.SlideFWD.Mode = -1;
            if (Ini.ReadInteger("PIN MAP", "SLIDE FWD PIN", ref Spec.PinMap.SlideFWD.PinNo) == false) Spec.PinMap.SlideFWD.PinNo = -1;

            if (Ini.ReadInteger("PIN MAP", "SLIDE BWD MODE", ref Spec.PinMap.SlideBWD.Mode) == false) Spec.PinMap.SlideBWD.Mode = -1;
            if (Ini.ReadInteger("PIN MAP", "SLIDE BWD PIN", ref Spec.PinMap.SlideBWD.PinNo) == false) Spec.PinMap.SlideBWD.PinNo = -1;

            if (Ini.ReadInteger("PIN MAP", "RECLINE FWD MODE", ref Spec.PinMap.ReclineFWD.Mode) == false) Spec.PinMap.ReclineFWD.Mode = -1;
            if (Ini.ReadInteger("PIN MAP", "RECLINE FWD PIN", ref Spec.PinMap.ReclineFWD.PinNo) == false) Spec.PinMap.ReclineFWD.PinNo = -1;

            if (Ini.ReadInteger("PIN MAP", "RECLINE BWD MODE", ref Spec.PinMap.ReclineBWD.Mode) == false) Spec.PinMap.ReclineBWD.Mode = -1;
            if (Ini.ReadInteger("PIN MAP", "RECLINE BWD PIN", ref Spec.PinMap.ReclineBWD.PinNo) == false) Spec.PinMap.ReclineBWD.PinNo = -1;

            if (Ini.ReadInteger("PIN MAP", "TILT UP MODE", ref Spec.PinMap.TiltUp.Mode) == false) Spec.PinMap.TiltUp.Mode = -1;
            if (Ini.ReadInteger("PIN MAP", "TILT UP PIN", ref Spec.PinMap.TiltUp.PinNo) == false) Spec.PinMap.TiltUp.PinNo = -1;

            if (Ini.ReadInteger("PIN MAP", "TILT DN MODE", ref Spec.PinMap.TiltDn.Mode) == false) Spec.PinMap.TiltDn.Mode = -1;
            if (Ini.ReadInteger("PIN MAP", "TILT DN PIN", ref Spec.PinMap.TiltDn.PinNo) == false) Spec.PinMap.TiltDn.PinNo = -1;

            if (Ini.ReadInteger("PIN MAP", "HEIGHT UP MODE", ref Spec.PinMap.HeightUp.Mode) == false) Spec.PinMap.HeightUp.Mode = -1;
            if (Ini.ReadInteger("PIN MAP", "HEIGHT UP PIN", ref Spec.PinMap.HeightUp.PinNo) == false) Spec.PinMap.HeightUp.PinNo = -1;

            if (Ini.ReadInteger("PIN MAP", "HEIGHT DN MODE", ref Spec.PinMap.HeightDn.Mode) == false) Spec.PinMap.HeightDn.Mode = -1;
            if (Ini.ReadInteger("PIN MAP", "HEIGHT DN PIN", ref Spec.PinMap.HeightDn.PinNo) == false) Spec.PinMap.HeightDn.PinNo = -1;

            if (Ini.ReadInteger("PIN MAP", "LUMBER FWD BATT", ref Spec.PinMap.LumberFwdBwd.Batt) == false) Spec.PinMap.LumberFwdBwd.Batt = -1;
            if (Ini.ReadInteger("PIN MAP", "LUMBER FWD GND", ref Spec.PinMap.LumberFwdBwd.Gnd) == false) Spec.PinMap.LumberFwdBwd.Gnd = -1;

            if (Ini.ReadInteger("PIN MAP", "LUMBER UP BATT", ref Spec.PinMap.LumberUpDn.Batt) == false) Spec.PinMap.LumberUpDn.Batt = -1;
            if (Ini.ReadInteger("PIN MAP", "LUMBER UP GND", ref Spec.PinMap.LumberUpDn.Gnd) == false) Spec.PinMap.LumberUpDn.Gnd = -1;

            if (Ini.ReadInteger("PIN MAP", "POWER1 BATT", ref Spec.PinMap.Power.Batt1.Batt) == false) Spec.PinMap.Power.Batt1.Batt = -1;
            if (Ini.ReadInteger("PIN MAP", "POWER1 GND", ref Spec.PinMap.Power.Batt1.Gnd) == false) Spec.PinMap.Power.Batt1.Gnd = -1;

            if (Ini.ReadInteger("PIN MAP", "POWER2 BATT", ref Spec.PinMap.Power.Batt2.Batt) == false) Spec.PinMap.Power.Batt2.Batt = -1;
            if (Ini.ReadInteger("PIN MAP", "POWER2 GND", ref Spec.PinMap.Power.Batt2.Gnd) == false) Spec.PinMap.Power.Batt2.Gnd = -1;

            if (Ini.ReadBool("PIN MAP", "직구동", ref Spec.PinMap.PSeat_직구동) == false) Spec.PinMap.PSeat_직구동 = false;

            if (Ini.ReadBool("PIN MAP", "WALKIN", ref Spec.PinMap.WalkIn) == false) Spec.PinMap.WalkIn = false;
            return;
        }


        public void SaveInfor(string Name, __Infor__ Inf,bool Flag = false)
        {
            TIniFile Ini = new TIniFile(Name);
            
            return;
        }

        
        public void SaveRunTime<__Struct__>(short Ch, __Struct__ Time)
        {
            try
            {
                string Name = null;

                FileStream DataStream = new FileStream(Name, FileMode.Append, FileAccess.Write);

                byte[] buffer = new byte[Marshal.SizeOf(typeof(__Struct__))];
                GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
                Marshal.StructureToPtr(Time, handle.AddrOfPinnedObject(), false);
                DataStream.Write(buffer, 0, Marshal.SizeOf(typeof(__Struct__)));
                handle.Free();
                DataStream.Close();
            }
            catch //(Exception ex)
            {
                //MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }


        

        public string[] GetComName(string[] List)
        {
            string[] Name = new string[List.Length];
            int x;

            for (int i = 0; i < Name.Length; i++) Name[i] = "";
            x = 0;
            foreach (string s in List)
            {
                Name[x] = "COM";
                Name[x] = Name[x] + GetNumber(s);
                x++;
            }

            return Name;
        }

        public string GetNumber(string Data)
        {
            string s = "";

            foreach (char c in Data)
            {
                if (isnumeric(c) == true) s = s + c;
            }
            return s;
        }

        public int GetToNumInt(string s, bool ReturnToChar = true, bool ToHex = true) //ReturnToChar == true 일때 숫자가 아니면 바로 리턴하고 false 이면 숫자값을 모두 읽어 온다.
        {
            int value;
            string sb;

            sb = GetString(s, ReturnToChar, ToHex);

            if (ToHex == true)
            {
                if (int.TryParse(sb, System.Globalization.NumberStyles.HexNumber, null, out value) == false)
                    return 0;
                else return value;
            }
            else
            {
                if (int.TryParse(sb, System.Globalization.NumberStyles.Number, null, out value) == false)
                    return 0;
                else return value;
            }
        }
                
        public string GetString(string data = "0", bool ReturnToChar = true, bool ToHexString = true)
        {
            int i;
            StringBuilder sb = new StringBuilder();

            i = 0;
            foreach (char c in data)
            {
                i++;
                // Check for numeric characters (hex in this case).  Add "." and "e" if float,
                // and remove letters.  Include initial space because it is harmless.
                if (ToHexString == true)
                {
                    if (((c >= '0') && (c <= '9')) || ((c >= 'A') && (c <= 'F')) || ((c >= 'a') && (c <= 'f')) || (c == ' ') || (c == '+') || (c == '-'))
                    {
                        sb.Append(c);
                    }
                    else
                    {
                        if (ReturnToChar == true)
                        {
                            if ((c == 'x') || (c == 'X'))
                            {
                                if (2 < i) break;
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                }
                else
                {
                    if (((c >= '0') && (c <= '9')) || (c == '.') || (c == ' ') || (c == '+') || (c == '-'))
                    {
                        sb.Append(c);
                    }
                    else
                    {
                        break;
                    }
                }
            }
            return sb.ToString();
        }

        public void strcpy(out string target, char[] source)
        {
            target = "";

            foreach (char c in source)
            {
                if (c != 0x00)
                    target = target + c;
                else break;
            }
            //target = target1;
            return;
        }

        public void strcpy(out char[] target, char[] source)
        {
            int i;

            target = new char[source.Length];
            i = 0;
            foreach (char c in source)
            {
                //if (c != 0x00)
                target[i++] = c;
                //else break;
            }
            return;
        }

        public void strcpy(out char[] target, string source)
        {
            int i;

            target = new char[source.Length];
            i = 0;
            foreach (char c in source)
            {
                //if (c != 0x00)
                target[i++] = c;
                //else break;
            }
            return;
        }

        public void strcpy(out byte[] target, string source)
        {
            int i;

            target = new byte[source.Length];
            i = 0;
            foreach (char c in source)
            {
                if (c != 0x00)
                    target[i++] = (byte)c;
                else break;
            }
            return;
        }

        public void strncpy(out char[] target, string source, int Length)
        {
            int i;

            i = 0;
            target = new char[Length];
            foreach (char c in source)
            {
                if (c != 0x00)
                    target[i++] = (char)c;
                else break;
            }
            return;
        }

        public void strncpy(out char[] target, byte[] source, int Length)
        {
            int i;

            i = 0;
            target = new char[Length];
            foreach (char c in source)
            {
                if (c != 0x00)
                    target[i++] = (char)c;
                else break;
            }
            return;
        }


        public float GetToNumFloat(string s, bool ReturnToChar = true, bool ToHex = true) //ReturnToChar == true 일때 숫자가 아니면 바로 리턴하고 false 이면 숫자값을 모두 읽어 온다.
        {
            float value;
            string sb;

            sb = GetString(s, ReturnToChar, ToHex);

            if (ToHex == true)
            {
                if (float.TryParse(sb, System.Globalization.NumberStyles.HexNumber, null, out value) == false)
                    return 0;
                else return value;
            }
            else
            {
                if (float.TryParse(sb, System.Globalization.NumberStyles.Number, null, out value) == false)
                    return 0;
                else return value;
            }
        }

        public short GetToNumShort(string s, bool ReturnToChar = true, bool ToHex = true) //ReturnToChar == true 일때 숫자가 아니면 바로 리턴하고 false 이면 숫자값을 모두 읽어 온다.
        {
            short value;
            string sb;

            sb = GetString(s, ReturnToChar, ToHex);

            if (ToHex == true)
            {
                if (short.TryParse(sb, System.Globalization.NumberStyles.HexNumber, null, out value) == false)
                    return 0;
                else return value;
            }
            else
            {
                if (short.TryParse(sb, System.Globalization.NumberStyles.Number, null, out value) == false)
                    return 0;
                else return value;
            }
        }

        public long GetToNumLong(string s, bool ReturnToChar = true, bool ToHex = true) //ReturnToChar == true 일때 숫자가 아니면 바로 리턴하고 false 이면 숫자값을 모두 읽어 온다.
        {
            long value;
            string sb;

            sb = GetString(s, ReturnToChar, ToHex);

            if (ToHex == true)
            {
                if (long.TryParse(sb, System.Globalization.NumberStyles.HexNumber, null, out value) == false)
                    return 0;
                else return value;
            }
            else
            {
                if (long.TryParse(sb, System.Globalization.NumberStyles.Number, null, out value) == false)
                    return 0;
                else return value;
            }
        }

        public double GetToNumDouble(string s, bool ReturnToChar = true, bool ToHex = true) //ReturnToChar == true 일때 숫자가 아니면 바로 리턴하고 false 이면 숫자값을 모두 읽어 온다.
        {
            double value;
            string sb;

            sb = GetString(s, ReturnToChar, ToHex);

            if (ToHex == true)
            {
                if (double.TryParse(sb, System.Globalization.NumberStyles.HexNumber, null, out value) == false)
                    return 0;
                else return value;
            }
            else
            {
                if (double.TryParse(sb, System.Globalization.NumberStyles.Number, null, out value) == false)
                    return 0;
                else return value;
            }
        }

        public byte GetToNumByte(string s, bool ReturnToChar = true, bool ToHex = true) //ReturnToChar == true 일때 숫자가 아니면 바로 리턴하고 false 이면 숫자값을 모두 읽어 온다.
        {
            byte value;
            string sb;

            sb = GetString(s, ReturnToChar, ToHex);

            if (ToHex == true)
            {
                if (byte.TryParse(sb, System.Globalization.NumberStyles.HexNumber, null, out value) == false)
                    return 0x00;
                else return value;
            }
            else
            {
                if (byte.TryParse(sb, System.Globalization.NumberStyles.Number, null, out value) == false)
                    return 0x00;
                else return value;
            }
        }

        public void GetSystemDate(out string outTime)
        {
            outTime = DateTime.Now.ToString("yyyy.MM.dd");

            return;
        }

        //---------------------------------------------------------------------------


        public void GetSystemTime(out string outTime)
        {
            outTime = DateTime.Now.ToString("HH:mm:ss");
            return;
        }

        //---------------------------------------------------------------------------
        public void GetSystemDateTime(out string outTime)
        {
            outTime = "[" + DateTime.Now.ToString("yyyy.MM.dd") + "-" + DateTime.Now.ToString("HH:mm:ss") + "]";
            return;
        }

        //---------------------------------------------------------------------------
        public bool isnumeric(char s)
        {
            switch (s)
            {
                case '0':
                case '1':
                case '2':
                case '3':
                case '4':
                case '5':
                case '6':
                case '7':
                case '8':
                case '9': return true;
            }
            return false;
        }
        //---------------------------------------------------------------------------
        public int StringToHex(char[] Data)
        {
            int Value;
            //short i;
            //short Length;

            //Value = 0;

            //Length = (short)Data.Length;

            //for (i = 0; i < Length; i++)
            //{
            //    Value |= CharToHex(Data[i]) << (4 * ((Length - 1) - i));
            //}

            byte[] s = Encoding.UTF8.GetBytes(Data);
            string s2 = Encoding.UTF8.GetString(s).ToUpper();

            var sx = new String(s2.Where(Char.IsLetterOrDigit).ToArray());

            s2 = sx;
            //if (0 <= s2.IndexOf("0x")) s2 = s2.Remove(0, 2);
            if (0 <= s2.IndexOf("0X")) s2 = s2.Remove(0, 2);


            if (int.TryParse(s2, System.Globalization.NumberStyles.HexNumber, null, out Value) == false) Value = 0;
            return Value;
        }

        public int StringToHex(byte[] Data)
        {
            int Value;
            //short i;
            //short Length;

            //Value = 0;

            //Length = (short)Data.Length;

            //for (i = 0; i < Length; i++)
            //{
            //    Value |= CharToHex(Data[i]) << (4 * ((Length - 1) - i));
            //}

            string s2 = Encoding.UTF8.GetString(Data).ToUpper();

            //if (0 <= s2.IndexOf("0x")) s2 = s2.Remove(0, 2);

            var sx = new String(s2.Where(Char.IsLetterOrDigit).ToArray());
            s2 = sx;

            if (0 <= s2.IndexOf("0X")) s2 = s2.Remove(0, 2);
            
            if (int.TryParse(s2, System.Globalization.NumberStyles.HexNumber, null, out Value) == false) Value = 0;
            return Value;
        }
        //---------------------------------------------------------------------------
        public int StringToHex(string Data)
        {
            int Value;
            //short i;
            //short Length;

            //if(0 <= Data.IndexOf("0x"))
            //{
            //    Data = Data.Substring(2);
            //}
            //else if (0 <= Data.IndexOf("0X"))
            //{
            //    Data = Data.Substring(2);
            //}

            //Value = 0;

            //Length = (short)Data.Length;

            //i = 0;
            //foreach (char c in Data)
            //{
            //    Value |= CharToHex(c) << (4 * ((Length - 1) - i));
            //    i++;
            //}

            string s2 = Data.ToUpper();
            var sx = new String(s2.Where(Char.IsLetterOrDigit).ToArray());
            s2 = sx;
            //if (0 <= s2.IndexOf("0x")) s2 = s2.Remove(0, 2);
            if (0 <= s2.IndexOf("0X")) s2 = s2.Remove(0, 2);

            if (int.TryParse(s2, System.Globalization.NumberStyles.HexNumber, null, out Value) == false) Value = 0;
            return Value;
        }
        //---------------------------------------------------------------------------
        public short CharToHex(char c)
        {
            //    switch(c)
            //    {
            //        case '0' : return 0x00;
            //        case '1' : return 0x01;
            //        case '2' : return 0x02;
            //        case '3' : return 0x03;
            //        case '4' : return 0x04;
            //        case '5' : return 0x05;
            //        case '6' : return 0x06;
            //        case '7' : return 0x07;
            //        case '8' : return 0x08;
            //        case '9' : return 0x09;
            //        case 'A' : return 0x0a;
            //        case 'B' : return 0x0b;
            //        case 'C' : return 0x0c;
            //        case 'D' : return 0x0d;
            //        case 'E' : return 0x0e;
            //        case 'F' : return 0x0f;
            //        default  : return 0x00;
            //    }
            short Value = 0x00;

            if (('0' <= c) && (c <= '9'))
            {
                Value = (short)(c - '0');
            }
            else if (('A' <= c) && (c <= 'F'))
            {
                Value = (short)(c - 'A');
                Value += 10;
            }

            else if (('a' <= c) && (c <= 'f'))
            {
                Value = (short)(c - 'a');
                Value += 10;
            }
            return Value;
        }
        //------------------------------------------------------------------------------
        public string ToString(char[] Data)
        {
            string s;
            //short i;

            if (Data == null) return "";
            if (Data.Length == 0) return "";
            if (Data[0] == 0x00) return "";

            /*
            for (i = 0; i < Data.Length; i++)
            {
                if (Data[i] == 0x0000) break;
            }

            s = new string(Data).Substring(0, i);
            */
            s = "";
            foreach (char c in Data)
            {
                if (c != 0x00)
                    s = s + c;
                else break;
            }
            return s;
        }
        //------------------------------------------------------------------------------
        public string ReturnNumData(string Data)
        {
            string s;

            //if (Data.Length == 0) return "0";
            //s = "";
            //foreach (char c in Data)
            //{
            //    if (char.IsNumber(c))
            //    {
            //        s = s + c;
            //    }
            //    else
            //    {
            //        s = "0";
            //        return s;
            //    }
            //}
            //s = new string(Data).Substring(0, i);
            //s = Data.ToString().Substring(0,i);


            //using System.Text.RegularExpressions;

            //위와 같이 사용하시면 됩니다. "\D" 는 숫자가 아닌 문자열을 뜻하므로 숫자를 제외하고 다 없애라는 뜻이 됩니다.간단하지만 요긴한것 같습니다.
            string strTmp = Regex.Replace(Data, @"\D", "");
            int nTmp = int.Parse(strTmp);

            s = Convert.ToString(nTmp);

            //숫자만 가져오기
            //var onlyLetters = new String(Data.Where(Char.IsNumber).ToArray());
            return s;
        }
        //------------------------------------------------------------------------------
        public string ReturnStringData(string Data)
        {
            string s;

            //if (Data.Length == 0) return "0";
            //s = "";
            //foreach (char c in Data)
            //{
            //    if (char.IsNumber(c))
            //    {
            //        s = s + c;
            //    }
            //    else
            //    {
            //        s = "0";
            //        return s;
            //    }
            //}
            //s = new string(Data).Substring(0, i);
            //s = Data.ToString().Substring(0,i);


            //using System.Text.RegularExpressions;

            //위와 같이 사용하시면 됩니다. "\d" 는 문자가 아닌 숫자열을 뜻하므로 문자를 제외하고 다 없애라는 뜻이 됩니다.간단하지만 요긴한것 같습니다.
            string strTmp = Regex.Replace(Data, @"\d", "");
            int nTmp = int.Parse(strTmp);

            s = Convert.ToString(nTmp);

            //문자열만 축출
            //var onlyLetters = new String(Data.Where(Char.IsLetter).ToArray());
            //대문자만 축출
            //var onlyLetters = new String(Data.Where(c => Char.IsLetter(c) && Char.IsUpper(c)).ToArray());
                       

            return s;
        }
        //------------------------------------------------------------------------------

        /*
        public void SaveIOMap<__IOMap__>(string Name, __IOMap__ Data)
        {
            //string Name;            

            //Path = Application.StartupPath + "\\SYSTEM";
            //Name = Path + "\\IO.dat";
            //File.Delete(Name);

            FileStream DataStream = new FileStream(Name, FileMode.Create, FileAccess.Write);

            byte[] buffer = new byte[Marshal.SizeOf(typeof(__IOMap__))];
            GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
            //Marshal.StructureToPtr((object)Data, handle.AddrOfPinnedObject(), false);
            Marshal.StructureToPtr(Data, handle.AddrOfPinnedObject(), false);
            DataStream.Write(buffer, 0, Marshal.SizeOf(typeof(__IOMap__)));
            handle.Free();
            DataStream.Close();
            return;
        }
        
        //------------------------------------------------------------------------------
        public __IOMap__ ReadIOMap<__IOMap__>(string Name, __IOMap__ Data)
        {
            //string Name;            
            int Length;
            FileStream DataStream;

            //Path = Application.StartupPath + "\\SYSTEM";
            //Name = Path + "\\IO.dat";

            byte[] buffer = new byte[Marshal.SizeOf(typeof(__IOMap__))];

            if (File.Exists(Name))
            {
                DataStream = new FileStream(Name, FileMode.Open, FileAccess.Read);
                Length = DataStream.Read(buffer, 0, Marshal.SizeOf(typeof(__IOMap__)));
                DataStream.Close();

                //if (Marshal.SizeOf(typeof(__IOMap__)) == Length) //나중에 함수가 추가 되더라도 데이타를 읽어 오기 위해 이 비교문을 삭제한다.
                {
                    GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
                    Data = (__IOMap__)Marshal.PtrToStructure(handle.AddrOfPinnedObject(), typeof(__IOMap__));
                    handle.Free();
                }
            }
            return Data;
        }
        */
        //------------------------------------------------------------------------------
        public void SaveData<__Struct__>(string Name, __Struct__ Data)
        {
            try
            {
                using (File.Open(Name, FileMode.Append)) { } //파일 엑세스가 가능한지 확인한다.

                if (File.Exists(Name) == false)
                {
                    using (var DataStream = new FileStream(Name, FileMode.Create, FileAccess.Write))
                    {
                        byte[] buffer = new byte[Marshal.SizeOf(typeof(__Struct__))];
                        GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
                        //Marshal.StructureToPtr((object)Data, handle.AddrOfPinnedObject(), false);
                        Marshal.StructureToPtr(Data, handle.AddrOfPinnedObject(), false);
                        DataStream.Write(buffer, 0, Marshal.SizeOf(typeof(__Struct__)));
                        handle.Free();
                        DataStream.Close();
                    }
                }
                else
                {
                    using (var DataStream = new FileStream(Name, FileMode.Append, FileAccess.Write))
                    {
                        byte[] buffer = new byte[Marshal.SizeOf(typeof(__Struct__))];
                        GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
                        //Marshal.StructureToPtr((object)Data, handle.AddrOfPinnedObject(), false);
                        Marshal.StructureToPtr(Data, handle.AddrOfPinnedObject(), false);
                        DataStream.Write(buffer, 0, Marshal.SizeOf(typeof(__Struct__)));
                        handle.Free();
                        DataStream.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }
        //------------------------------------------------------------------------------
        public void SaveBuffer<__Struct__>(string Name, __Struct__ Data)
        {
            try
            {
                //using (File.Open(Name, FileMode.Append)) { } //파일 엑세스가 가능한지 확인한다.

                using (var DataStream = new FileStream(Name, FileMode.Create, FileAccess.Write))
                {
                    //FileStream DataStream = new FileStream(Name, FileMode.Create, FileAccess.Write);

                    byte[] buffer = new byte[Marshal.SizeOf(typeof(__Struct__))];
                    GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
                    //Marshal.StructureToPtr((object)Data, handle.AddrOfPinnedObject(), false);
                    Marshal.StructureToPtr(Data, handle.AddrOfPinnedObject(), false);
                    DataStream.Write(buffer, 0, Marshal.SizeOf(typeof(__Struct__)));
                    handle.Free();
                    DataStream.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
                //uMessageBox.Show(promptText: ex.Message + "\n" + ex.StackTrace);
            }
        }
        //------------------------------------------------------------------------------
        public __Struct__ ReadBuffer<__Struct__>(string Name, __Struct__ Data)
        {
            int Length;
            //FileStream DataStream;

            //using (File.Open(Name, FileMode.Open)) { } //파일 엑세스가 가능한지 확인한다.

            byte[] buffer = new byte[Marshal.SizeOf(typeof(__Struct__))];

            if (File.Exists(Name))
            {
                using (var DataStream = new FileStream(Name, FileMode.Open, FileAccess.Read))
                {
                    //DataStream = new FileStream(Name, FileMode.Open, FileAccess.Read);
                    Length = DataStream.Read(buffer, 0, Marshal.SizeOf(typeof(__Struct__)));
                    DataStream.Close();

                    //if (Marshal.SizeOf(typeof(__Spec__)) == Length) //나중에 함수가 추가 되더라도 데이타를 읽어 오기 위해 이 비교문을 삭제한다.
                    {
                        GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
                        Data = (__Struct__)Marshal.PtrToStructure(handle.AddrOfPinnedObject(), typeof(__Struct__));
                        handle.Free();
                    }
                }
            }
            return Data;
        }


        //------------------------------------------------------------------------------
        /*
        public __Can__ OpenCan<__Can__>(string Name, __Can__ Data)
        {
            //string Path;
            int Length;
            FileStream DataStream;

            //Path = Application.StartupPath + "\\SYSTEM";

            //Name = Path + "\\Can.dat";

            byte[] buffer = new byte[Marshal.SizeOf(typeof(__Can__))];

            if (File.Exists(Name))
            {
                DataStream = new FileStream(Name, FileMode.Open, FileAccess.Read);
                Length = DataStream.Read(buffer, 0, Marshal.SizeOf(typeof(__Can__)));
                DataStream.Close();

                //if (Marshal.SizeOf(typeof(__Can__)) == Length) //나중에 함수가 추가 되더라도 데이타를 읽어 오기 위해 이 비교문을 삭제한다.
                {
                    GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
                    Data = (__Can__)Marshal.PtrToStructure(handle.AddrOfPinnedObject(), typeof(__Can__));
                    handle.Free();
                }
            }

            return Data;
        }
        //---------------------------------------------------------------------------
        public void SaveCan<__Can__>(string Name, __Can__ Data)
        {
            //string Name;            

            //Path = Application.StartupPath + "\\SYSTEM";
            //Name = Path + "\\Can.dat";
            //            File.Delete(Name);

            FileStream DataStream = new FileStream(Name, FileMode.Create, FileAccess.Write);

            byte[] buffer = new byte[Marshal.SizeOf(typeof(__Can__))];
            GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
            //Marshal.StructureToPtr((object)Data, handle.AddrOfPinnedObject(), false);
            Marshal.StructureToPtr(Data, handle.AddrOfPinnedObject(), false);
            DataStream.Write(buffer, 0, Marshal.SizeOf(typeof(__Can__)));
            handle.Free();
            DataStream.Close();
            
            return;
        }



        public void SaveCount<__NgCount__>(string Name, __NgCount__ Data)
        {
            try
            {
                FileStream DataStream;

                if (File.Exists(Name) == false)
                    DataStream = new FileStream(Name, FileMode.Create, FileAccess.Write);
                else DataStream = new FileStream(Name, FileMode.Append, FileAccess.Write);
                byte[] buffer = new byte[Marshal.SizeOf(typeof(__NgCount__))];
                GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
                //Marshal.StructureToPtr((object)Data, handle.AddrOfPinnedObject(), false);
                Marshal.StructureToPtr(Data, handle.AddrOfPinnedObject(), false);
                DataStream.Write(buffer, 0, Marshal.SizeOf(typeof(__NgCount__)));
                handle.Free();
                DataStream.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        //------------------------------------------------------------------------------
        public __NgCount__ ReadCount<__NgCount__>(string Name, __NgCount__ Data)
        {
            int Length;
            FileStream DataStream;


            byte[] buffer = new byte[Marshal.SizeOf(typeof(__Infor__))];

            if (File.Exists(Name))
            {
                DataStream = new FileStream(Name, FileMode.Open, FileAccess.Read);
                Length = DataStream.Read(buffer, 0, Marshal.SizeOf(typeof(__NgCount__)));
                DataStream.Close();

                //if (Marshal.SizeOf(typeof(__NgCount__)) == Length) //나중에 함수가 추가 되더라도 데이타를 읽어 오기 위해 이 비교문을 삭제한다.
                {
                    GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
                    Data = (__NgCount__)Marshal.PtrToStructure(handle.AddrOfPinnedObject(), typeof(__NgCount__));
                    handle.Free();
                }
            }
            return Data;
        }
        //------------------------------------------------------------------------------

        public void SaveData<__TestData__>(string Name, __TestData__ Data)
        {
            try
            {
                FileStream DataStream;

                if (File.Exists(Name) == false)
                    DataStream = new FileStream(Name, FileMode.Create, FileAccess.Write);
                else DataStream = new FileStream(Name, FileMode.Append, FileAccess.Write);
                byte[] buffer = new byte[Marshal.SizeOf(typeof(__TestData__))];
                GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
                //Marshal.StructureToPtr((object)Data, handle.AddrOfPinnedObject(), false);
                Marshal.StructureToPtr(Data, handle.AddrOfPinnedObject(), false);
                DataStream.Write(buffer, 0, Marshal.SizeOf(typeof(__TestData__)));
                handle.Free();
                DataStream.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
            return;
        }
        //------------------------------------------------------------------------------
        public __TestData__ ReadData<__TestData__>(string Name, __TestData__ Data)
        {
            int Length;
            FileStream DataStream;


            byte[] buffer = new byte[Marshal.SizeOf(typeof(__TestData__))];

            if (File.Exists(Name))
            {
                DataStream = new FileStream(Name, FileMode.Open, FileAccess.Read);
                Length = DataStream.Read(buffer, 0, Marshal.SizeOf(typeof(__TestData__)));
                DataStream.Close();

                //if (Marshal.SizeOf(typeof(__TestData__)) == Length) //나중에 함수가 추가 되더라도 데이타를 읽어 오기 위해 이 비교문을 삭제한다.
                {
                    GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
                    Data = (__TestData__)Marshal.PtrToStructure(handle.AddrOfPinnedObject(), typeof(__TestData__));
                    handle.Free();
                }
            }
            return Data;
        }
        //------------------------------------------------------------------------------
        */
        //public __IOPort CheckLowInPos(short Pos)
        //{
        //    __IOPort Data = new __IOPort()
        //    {
        //        Pos = 0,
        //        Data = 0x00
        //    };

        //    //unsafe
        //    //{
        //    if (Pos == -1) return Data;
        //    //}

        //    if (Pos < 200)
        //    {
        //        //Data.Pos = (short)(MainForm.IOMap.In[Pos].Pos / 8);
        //        //Data.Data = (byte)(0x01 << (MainForm.IOMap.In[Pos].Pos % 8));
        //        Data.Pos = (short)(Pos / 8);
        //        Data.Data = (byte)(0x01 << (Pos % 8));
        //    }
        //    else
        //    {
        //        //Data.Pos = (short)(MainForm.IOMap.ExtIn[Pos - 200].Pos / 8);
        //        //Data.Data = (byte)(0x01 << (MainForm.IOMap.ExtIn[Pos - 200].Pos % 8));
        //        Data.Pos = (short)((Pos - 200) / 8);
        //        Data.Data = (byte)(0x01 << ((Pos - 200) % 8));
        //    }
        //    return Data;
        //}


        //public __IOPort CheckHighInPos(short Pos)
        //{
        //    __IOPort Data = new __IOPort()
        //    {
        //        Pos = 0,
        //        Data = 0x00
        //    };

        //    if (Pos == -1) return Data;

        //    //Data.Pos = (short)(MainForm.IOMap.HiIn[Pos].Pos / 8);
        //    //Data.Data = (byte)(0x01 << (MainForm.IOMap.HiIn[Pos].Pos % 8));
        //    Data.Pos = (short)(Pos / 8);
        //    Data.Data = (byte)(0x01 << (Pos % 8));
        //    return Data;
        //}




        //double ticks = sw.ElapsedTicks;
        //double seconds = ticks / Stopwatch.Frequency;
        //double milliseconds = (ticks / Stopwatch.Frequency) * 1000;
        //double nanoseconds = (ticks / Stopwatch.Frequency) * 1000000000;

        public long timeGetTimems()
        {
            double ticks = STOP_WATCH.ElapsedTicks;

            //return STOP_WATCH.ElapsedMilliseconds;
            return (long)((ticks / Stopwatch.Frequency) * 1000);
        }

        public long timeGetNanoTimes()
        {
            //1ticks == 100 nanosecond
            //return STOP_WATCH.Elapsed.Ticks;// ElapsedTicks;
            double ticks = STOP_WATCH.ElapsedTicks;
            return (long)((ticks / Stopwatch.Frequency) * 1000000000);
        }


        public void delay(long time)
        {
            long first;
            long last;

            first = timeGetTimems();
            last = timeGetTimems();
            do
            {
                last = timeGetTimems();
            } while ((last - first) < time);
            return;
        }

        //private DelayForm delayForm = null;
        public void timedelay2(long time)
        {
            //long first;
            //long last;

            //first = timeGetTimems();
            //last = timeGetTimems();
            //do
            //{
            //    Application.DoEvents();
            //    last = timeGetTimems();
            //} while ((last - first) < time);

            using (DelayForm delayForm = new DelayForm((int)time) { WindowState = FormWindowState.Minimized })
            {
                delayForm.FormClosing += delegate (object sender, FormClosingEventArgs e)
                {
                    e.Cancel = false;
                    delayForm.Dispose();
                    //delayForm = null;
                };

                delayForm.ShowDialog();

                //delayForm.Dispose();
                //delayForm = null;
            }
            return;
        }

        public void timedelay(long time)
        {
            long first;
            long last;

            first = timeGetTimems();
            last = timeGetTimems();
            do
            {
                Application.DoEvents();
                last = timeGetTimems();
            } while ((last - first) < time);

            return;
        }

        //private SecurityUnlockForm SecurityForm = null;


        public void nanotimedelay(long time)
        {
            long first;
            long last;

            first = timeGetNanoTimes();
            last = timeGetNanoTimes();
            do
            {
                Application.DoEvents();
                last = timeGetNanoTimes();
            } while ((last - first) < time);
            return;
        }
    }

    //------------------------------------------------------------------------------
    

    public class ConfigSetting
    {
        __Config__ xConfig = new __Config__();
        //{
        //    Lin = new __LinDevice__
        //    {
        //        Device = 0,
        //        Speed = 0
        //    },
        //    Can = new __CanDevice__
        //    {
        //        Channel = 0,
        //        Device = 0,
        //        ID = 0,
        //        Speed = 0
        //    },
        //    Client = new __TcpIP__
        //    {
        //        IP = string.Empty,
        //        Port = 0
        //    },
        //    NoiseMeter = new __Port__
        //    {
        //        Port = string.Empty,
        //        Speed = 2
        //    }
        //};

        //TIniFile Ini;// = new TIniFile(Name);                

        public ConfigSetting()
        {
            return;
        }

        public __Config__ ReadWriteConfig
        {
            set { xConfig = value; SaveConfig(); }

            get { OpenConfig(); return xConfig; }
        }

        private void OpenConfig()
        {
            //string Path;

            ////Settings 는 using PSeatTester.Properties 즉 Settings 를 사용

            //Path = Program.SYSTEM_PATH.ToString() + "\\Config.ini";

            //if (File.Exists(Path) == false) return;
            //Ini = new TIniFile(Path);

            //xConfig.Lin.Device = Settings.Default.LIN_DEVICE;
            //xConfig.Lin.Speed = Settings.Default.LIN_SPEED;

            xConfig.Can.Channel = Settings.Default.CAN_CHANNEL;
            xConfig.Can.Device = Settings.Default.CAN_DEVICE;
            xConfig.Can.ID = Settings.Default.CAN_ID;
            xConfig.Can.Speed = Settings.Default.CAN_SPEED;

            xConfig.Client.IP = Settings.Default.CLIENT_IP;
            xConfig.Client.Port = Settings.Default.CLIENT_PORT;
            xConfig.Server.IP = Settings.Default.SERVER_IP;
            xConfig.Server.Port = Settings.Default.SERVER_PORT;

            xConfig.PC.IP = Settings.Default.PC_IP;
            xConfig.PC.Port = Settings.Default.PC_PORT;
            xConfig.Board.IP = Settings.Default.BOARD_IP;
            xConfig.Board.Port = Settings.Default.BOARD_PORT;

            xConfig.NoiseMeter.Port = Settings.Default.NOISE_METER_PORT;
            xConfig.NoiseMeter.Speed = Settings.Default.NOISE_METER_SPEED;

            xConfig.Panel.Port = Settings.Default.PANEL_METER_PORT;
            xConfig.Panel.Speed = Settings.Default.PANEL_METER_SPEED;

            xConfig.Power.Port = Settings.Default.POWER_PORT;
            
            xConfig.Power.Speed = Settings.Default.POWER_SPEED;

            xConfig.BattID = Settings.Default.BATT_ID;
            xConfig.CurrID = Settings.Default.CURR_ID;

            xConfig.AutoConnection = Settings.Default.AUTO_CONNECION;
            xConfig.PinConnectionDelay = Settings.Default.CONNECTION_DELAY;
            xConfig.IMSSeatToCanControl = Settings.Default.IMS_SEAT_CONTROL;
            xConfig.KalmanSpeed = Settings.Default.KALMAN_SPEED;

            xConfig.SmartIOPort.Port = Settings.Default.SWIO_PORT;
            xConfig.SmartIOPort.Speed = Settings.Default.SWIO_SPEED;
            xConfig.UseSmartIO = Settings.Default.USE_SMARTIO;
            return;
        }

        private void SaveConfig()
        {
            //string Path = Program.SYSTEM_PATH.ToString() + "\\Config.ini";
            //TIniFile Ini = new TIniFile(Path);


            //Settings.Default.LIN_DEVICE = xConfig.Lin.Device;
            //Settings.Default.LIN_SPEED = xConfig.Lin.Speed;

            Settings.Default.CAN_CHANNEL = xConfig.Can.Channel;
            Settings.Default.CAN_DEVICE = xConfig.Can.Device;
            Settings.Default.CAN_ID = xConfig.Can.ID;
            Settings.Default.CAN_SPEED = xConfig.Can.Speed; 

            Settings.Default.CLIENT_IP = xConfig.Client.IP;
            Settings.Default.CLIENT_PORT = xConfig.Client.Port;
            
            Settings.Default.SERVER_IP = xConfig.Server.IP;
            Settings.Default.SERVER_PORT = xConfig.Server.Port;

            Settings.Default.PC_IP = xConfig.PC.IP;
            Settings.Default.PC_PORT = xConfig.PC.Port;
            Settings.Default.BOARD_IP = xConfig.Board.IP;
            Settings.Default.BOARD_PORT = xConfig.Board.Port;

            Settings.Default.NOISE_METER_PORT = xConfig.NoiseMeter.Port;
            Settings.Default.NOISE_METER_SPEED = xConfig.NoiseMeter.Speed;

            Settings.Default.PANEL_METER_PORT = xConfig.Panel.Port;
            Settings.Default.PANEL_METER_SPEED = xConfig.Panel.Speed;

            Settings.Default.POWER_PORT = xConfig.Power.Port;
            Settings.Default.POWER_SPEED = xConfig.Power.Speed;

            Settings.Default.BATT_ID = xConfig.BattID;
            Settings.Default.CURR_ID = xConfig.CurrID;

            Settings.Default.AUTO_CONNECION = xConfig.AutoConnection;
            Settings.Default.CONNECTION_DELAY = xConfig.PinConnectionDelay;

            Settings.Default.IMS_SEAT_CONTROL = xConfig.IMSSeatToCanControl;
            Settings.Default.KALMAN_SPEED = xConfig.KalmanSpeed;

            Settings.Default.SWIO_PORT = xConfig.SmartIOPort.Port;
            Settings.Default.SWIO_SPEED = xConfig.SmartIOPort.Speed;
            Settings.Default.USE_SMARTIO = xConfig.UseSmartIO;
            Settings.Default.Save();
            return;
        }

        ~ConfigSetting()
        {
        }
    }

    /*
    public class ELoadControl
    {
        private string SendString;
        private char[] SendBuf = new char[1000];
        //private bool ELoadCheckFlag; //extern 

        private char[] ELoad_R_Data = new char[200];//extern 
        //private string rxData;
        private short ELoad_R_Length;//extern 
        private bool ELoad_R_Flag;//extern 
        private const byte STX = 0x13;
        private const byte ETX = 0x11;
        private const byte END = 0x0A;


        private float Power = 0;
        private float Current = 0;
        private bool InputLoadOnOffFlag = false;
        private bool OutputLoadOnOffFlag = false;
        private short ReadMode = 0;
        private float LoadVoltValue = 0;
        private float LoadCurrentValue = 0;

        private SerialPort ELoadPort;

        public ELoadControl(SerialPort Ps)
        {
            //rxData = "";
            ELoadPort = Ps;
            //ELoadCheckFlag = false;
        }

        public SerialPort ELoadPortOpen
        {
            set
            {
                ELoadPort = value;
                //ELoadCheckFlag = false;
                //rxData = "";
            }
        }

        ~ELoadControl()
        {
        }

        public bool IsOpen
        {
            get
            {
                return ELoadPort.IsOpen;
            }
        }

        
        public void ELoadCheckMessage()
        {
            if (ELoadPort.IsOpen == true)
            {
                Array.Clear(SendBuf, 0, 1000);

                SendString = "*IDN?\n";
                Array.Copy(SendString.ToCharArray(), SendBuf, SendString.Length);
#if PROGRAM_RUNNING
                ELoadPort.Write(SendBuf, 0, SendString.Length);
#endif
            }

            return;
        }

        public char[] SetReceiveData
        {
            set
            {
                Array.Copy(value, 0, ELoad_R_Data, 0, value.Length);
            }        
        }
        public short SetReceiveLength
        {
            set
            {
                ELoad_R_Length = value;
            }
        }

        public bool SetReceiveFlag
        {
            set
            {
                ELoad_R_Flag = value;
            }
            get
            {
                return ELoad_R_Flag;
            }
        }

        

        public void ReceiveMsgCheck()
        {
            if (ELoad_R_Flag == false) return;

            string s = MainForm.ComF.ToString(ELoad_R_Data);

            if (0 <= s.IndexOf("PLZ664WA"))
            {
                //ELoadCheckFlag = true;
            }
            else
            {
                if (ELoad_R_Data[0] == STX)
                {
                    if ((ELoad_R_Data[ELoad_R_Length - 2] == END) && (ELoad_R_Data[ELoad_R_Length - 1] == ETX))
                    {
                        string[] sx = s.Split((char)STX);
                        if (0 < sx.Length)
                        {
                            string[] sx1 = sx[0].Split((char)END);
                            if (0 < sx1.Length)
                            {
                                
                                switch (ReadMode)
                                {
                                    case 1:
                                        if (float.TryParse(sx1[0], out Power) == false) Power = 0;
                                        break;
                                    case 2:
                                        if (float.TryParse(sx1[0], out Current) == false) Current = 0;
                                        break;
                                }
                            }
                        }
                    }
                }
            }
            ReadMode = 0;
            Array.Clear(ELoad_R_Data, 0, ELoad_R_Data.Length);
            ELoad_R_Flag = false;
            ELoad_R_Length = 0;
            return;
        }

        private void SetCurrent(float Current)
        {
            if (ELoadPort.IsOpen == true)
            {
                Array.Clear(SendBuf, 0, 1000);

                SendString = string.Format("SOURce:CURRent:LEVel:IMMediate:AMPLitude {0:0.00}\n", Current);
                Array.Copy(SendString.ToCharArray(), SendBuf, SendString.Length);
#if PROGRAM_RUNNING
                ELoadPort.Write(SendBuf, 0, SendString.Length);
#endif
            }
            return;
        }

        private void SetPower(float Volt)
        {
            if (ELoadPort.IsOpen == true)
            {
                Array.Clear(SendBuf, 0, 1000);

                SendString = string.Format("SOURce:POWer:LEVel:IMMediate:AMPLitude {0:0.00}\n", Volt);
                Array.Copy(SendString.ToCharArray(), SendBuf, SendString.Length);
#if PROGRAM_RUNNING
                ELoadPort.Write(SendBuf, 0, SendString.Length);
#endif
            }
            return;
        }
        
        public float LoadVoltage
        {
            set
            {
                LoadVoltValue = value;
                SetPower(value);
            }
            get
            {
                return LoadVoltValue;
            }
        }

        public float LoadCurrent
        {
            set
            {
                LoadCurrentValue = value;
                SetCurrent(value);
            }
            get
            {
                return LoadCurrentValue;
            }
        }


        public bool InputLoadOnOff
        {
            set
            {
                InputLoadOnOffFlag = value;
                SetInputLoadOn(value);
            }
            get
            {
                return InputLoadOnOffFlag;
            }
        }

        public bool OutputLoadOnOff
        {
            set
            {
                OutputLoadOnOffFlag = value;
                SetOutputLoadOn(value);
            }
            get
            {
                return OutputLoadOnOffFlag;
            }
        }

        private void SetInputLoadOn(bool Flag)
        {
            if (ELoadPort.IsOpen == true)
            {
                Array.Clear(SendBuf, 0, 1000);

                if (Flag == false)
                    SendString = "INPut:STATe:IMMEdiate 0\n";
                else SendString = "INPut:STATe:IMMEdiate 1\n";
                Array.Copy(SendString.ToCharArray(), SendBuf, SendString.Length);
#if PROGRAM_RUNNING
                ELoadPort.Write(SendBuf, 0, SendString.Length);
#endif
            }
            return;
        }

        private void SetOutputLoadOn(bool Flag)
        {
            if (ELoadPort.IsOpen == true)
            {
                Array.Clear(SendBuf, 0, 1000);

                if (Flag == false)
                    SendString = "OUTPut:STATe:IMMEdiate 0\n";
                else SendString = "OUTPut:STATe:IMMEdiate 1\n";
                Array.Copy(SendString.ToCharArray(), SendBuf, SendString.Length);
#if PROGRAM_RUNNING
                ELoadPort.Write(SendBuf, 0, SendString.Length);
#endif
            }
            return;
        }
        
        public void GetPower()
        {
            if (ELoadPort.IsOpen == true)
            {
                Array.Clear(SendBuf, 0, 1000);

                SendString = "SOURce:POWer:LEVel:IMMediate:AMPLitude?\n";
                
                Array.Copy(SendString.ToCharArray(), SendBuf, SendString.Length);
                ReadMode = 1;
#if PROGRAM_RUNNING
                ELoadPort.Write(SendBuf, 0, SendString.Length);
#endif
                return;
            }
        }

        public void GetCurrent()
        {
            if (ELoadPort.IsOpen == true)
            {
                Array.Clear(SendBuf, 0, 1000);

                SendString = "SOURce:CURRent:LEVel:IMMediate:AMPLitude?\n";

                Array.Copy(SendString.ToCharArray(), SendBuf, SendString.Length);
                ReadMode = 2;
#if PROGRAM_RUNNING
                ELoadPort.Write(SendBuf, 0, SendString.Length);
#endif
                return;
            }
        }
    }

    */
    public class PowerControl
    {
        //public string SendString;
        //public char[] SendBuf = new char[1000];
        public bool DCPowerCheckFlag; //extern 

        public char[] POWER_R_Data = new char[200];//extern 
        //private string rxData;
        public short Power_R_Length;//extern 
        private bool Power_R_Flag;//extern 
        private bool ExitFlag = false;
        private Timer pwTimer = new Timer()
        {
            Interval = 50
        };

        //long DCPowerCheckFirst = 0;
        //long DCPowerCheckLast = 0;


        public Stopwatch STOP_WATCH = new Stopwatch();

        private SerialPort PowerPort;

        public PowerControl(__Port__ Ps)
        {
            int Speed;

            Speed = 9600;
            //rxData = "";



            bool Flag = true;
            if (Ps.Port == "") Flag = false;
            if (Ps.Port == null) Flag = false;
            if (Ps.Port == string.Empty) Flag = false;

            if(Flag == true)
                PowerPort = new SerialPort(Ps.Port, Speed, Parity.None, 8, StopBits.One);
            else PowerPort = new SerialPort();

            if (Flag == false)
            {
                PowerPort.Parity = Parity.None;
                PowerPort.DataBits = 8;
                PowerPort.StopBits = StopBits.One;
            }
            PowerPort.ReadTimeout = 500;
            PowerPort.WriteTimeout = 500;
            PowerPort.ReadBufferSize = 4096;
            PowerPort.WriteBufferSize = 2048;
            PowerPort.DataReceived += new SerialDataReceivedEventHandler(PowerDataCatch);
            
            if (Flag == true) PowerPort.Open();
            STOP_WATCH.Start();
            pwTimer.Tick += new EventHandler(pwTimer_Tick);
            pwTimer.Enabled = true;
        }

        public SerialPort PowerPortOpen
        {
            set
            {
                PowerPort = value;
                pwTimer.Tick += new EventHandler(pwTimer_Tick);
                pwTimer.Enabled = true;
                //rxData = "";
                STOP_WATCH.Start();
            }
        }

        ~PowerControl()
        {
            ExitFlag = true;
            STOP_WATCH.Stop();
        }

        public void Close()
        {
            if (PowerPort != null)
            {
                if (PowerPort.IsOpen == true) PowerPort.Close();
            }
            return;
        }

        private List<string> MsgList = new List<string>();

        private bool Question = false;
        private long QuestionFirst = 0;
        private long QuestionLast = 0;
        private bool MsgSendFlag = false;

        private void pwTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                pwTimer.Enabled = false;

                if (Question == false)
                {
                    if (0 < MsgList.Count)
                    {
                        string SendString = MsgList[0] + "\n";
                        char[] SendBuf = new char[1000];

                        if (0 <= SendString.IndexOf("?"))
                        {
                            Question = true;
                            QuestionFirst = timeGetTimems();
                            QuestionLast = timeGetTimems();
                        }
                        Array.Clear(SendBuf, 0, 1000);
                        Array.Copy(SendString.ToCharArray(), SendBuf, SendString.Length);
                        PowerPort.Write(SendBuf, 0, SendString.Length);
                        MsgList.RemoveAt(0);
                        QuestionFirst = timeGetTimems();
                        QuestionLast = timeGetTimems();
                    }
                    else
                    {
                        QuestionLast = timeGetTimems();
                        if (500 <= (QuestionLast - QuestionFirst))
                        {
                            if (MsgSendFlag == false)
                                PowerReadMessage();
                            else PowerOnOffReadMessage();
                            MsgSendFlag = !MsgSendFlag;
                        }
                    }
                }
                else
                {
                    QuestionLast = timeGetTimems();
                    if (1500 <= (QuestionLast - QuestionFirst))
                    {
                        //1.5초 동안 응답이 없으면 타임 종료로 다시 통신을 제게한다.
                        Question = false;
                    }
                }
            }
            catch
            {

            }
            finally
            {
                pwTimer.Enabled = !ExitFlag;
            }
            return;
        }

        public bool IsOpen
        {
            get
            {
                return PowerPort.IsOpen;
            }
        }

        public long timeGetTimems()
        {
            return STOP_WATCH.ElapsedMilliseconds;
        }

        public void POWER_PWSetting(float Volt)   // Power Setting
        {
            if (PowerPort.IsOpen == false) return;

#if PROGRAM_RUNNING
            if (PowerPort.IsOpen == true) MsgList.Add(string.Format("VOLT {0:00.00}", Volt));
#endif
            return;
        }
        //---------------------------------------------------------------------------
        public void POWER_CURRENTSetting(float Current)   // Current Setting
        {
            if (PowerPort.IsOpen == false) return;
#if PROGRAM_RUNNING
            if (PowerPort.IsOpen == true) MsgList.Add(string.Format("CURR {0:00.00}", Current));
#endif
            return;
        }
        //---------------------------------------------------------------------------

        public void POWER_PWON()   // Power Output ON
        {
            if (PowerPort.IsOpen == false) return;

#if PROGRAM_RUNNING
            if (PowerPort.IsOpen == true) MsgList.Add("Outp ON");
#endif
            PowerOnOff = true;
            return;
        }
        //---------------------------------------------------------------------------
        public void POWER_PWOFF()  // Power Output OFF
        {
            if (PowerPort.IsOpen == false) return;

#if PROGRAM_RUNNING
            if (PowerPort.IsOpen == true) MsgList.Add("Outp OFF");
#endif
            PowerOnOff = false;
            return;
        }
        //---------------------------------------------------------------------------
        public void CHANNELSELECT()   // Power Output ON
        {
            if (PowerPort.IsOpen == false) return;

#if PROGRAM_RUNNING
            if (PowerPort.IsOpen == true) MsgList.Add("INST OUT1");
#endif
            return;
        }
        //---------------------------------------------------------------------------
        public void timedelay(long time)
        {
            long first;
            long last;

            first = timeGetTimems();
            last = timeGetTimems();
            do
            {
                Application.DoEvents();
                last = timeGetTimems();
            } while ((last - first) < time);
            return;
        }
        //---------------------------------------------------------------------------

        public void DCPowerReset()
        {
#if PROGRAM_RUNNING
            DCPowerCheckFlag = false;
            if (PowerPort.IsOpen == true) MsgList.Add("*RST");
#endif
            return;
        }
        //---------------------------------------------------------------------------

        public string CheckPowerSerialNo()
        {
#if PROGRAM_RUNNING
            string s;

            DCPowerCheckFlag = false;
            if (PowerPort.IsOpen == true) MsgList.Add("*SN?");
            s = ReadMessage();
            return s;
#else
            return "";
#endif
        }

        private string ToString(char[] Data)
        {
            string s;
            short i;

            for (i = 0; i < 50; i++)
            {
                if (Data[i] == 0x0000) break;
            }

            s = new string(Data).Substring(0, i);

            return s;
        }
        //---------------------------------------------------------------------------

        public enum SendMsgMode
        {
            VOLT,
            ONOFF
        }

        private bool SendFlag = false;
        private SendMsgMode SendMessageMode = SendMsgMode.VOLT;

        public SendMsgMode GetSendMsg
        {
            get
            {
                return SendMessageMode;
            }
        }

        private float SetVolt = 0;
        private bool PowerOnOff = false;

        public float ReadVolt
        {
            get
            {
                if (PowerOnOff == true)
                    return SetVolt;
                else return 0;
            }
        }

        public bool isPowerOn
        {
            get { return PowerOnOff; }
        }

        public bool GetReadEnd
        {
            get { return !SendFlag; }
            set { SendFlag = !value; }
        }

        public void PowerReadMessage()
        {

#if PROGRAM_RUNNING
            if (PowerPort.IsOpen == true) MsgList.Add("VOLT?");
#endif
            SendFlag = true;
            SendMessageMode = SendMsgMode.VOLT;
            return;
        }

        public void PowerOnOffReadMessage()
        {


#if PROGRAM_RUNNING
            if (PowerPort.IsOpen == true) MsgList.Add("OUTP?");
#endif
            SendFlag = true;
            SendMessageMode = SendMsgMode.ONOFF;
            return;
        }



        void POWER_Read()
        {
            /*
                        try
                        {
#if PROGRAM_RUNNING
                        try
                        {
                            s = ReadMessage();                    
                            return;
                        }
                        else{
                            Data->Flag = false;
                            Data->ReadFlag = false;
                        }

                        catch(...)
                        {
                        }
#endif
                        }
                        finally
                        {
                        }
            */
            return;
        }

        public string ReadMessage()
        {
            //#if PROGRAM_RUNNING	        
            long First;
            long Last;
            bool Flag;
            string s;

            if (PowerPort.IsOpen == false) return "";
            s = "";
            try
            {
                try
                {
                    First = timeGetTimems();
                    Flag = PowerPort.IsOpen;

                    do
                    {
                        if (Flag == true)
                        {
                            if (Power_R_Flag == true)
                            {
                                s = POWER_R_Data.ToString();
                                Power_R_Flag = false;
                                Power_R_Length = 0;
                                Array.Clear(POWER_R_Data, 0, 200);
                                return s;
                            }
                        }
                        else break;
                        Application.DoEvents();
                        Last = timeGetTimems();
                    } while ((Last - First) < 1000);
                }
                catch (Exception Msg)
                {
                    MessageBox.Show(Msg.Message + "\n" + Msg.StackTrace);
                    //uMessageBox.Show(promptText: ex.Message + "\n" + ex.StackTrace);
                }
            }
            finally
            {

            }
            return s;
            //#else
            //            return "";
            //#endif
        }
        //---------------------------------------------------------------------------

        private void PowerDataCatch(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                timedelay(30);
                int Length = PowerPort.BytesToRead;
                byte[] buffer = new byte[Length + 10];

                PowerPort.Read(buffer, 0, Length);

                string[] s = Encoding.Default.GetString(buffer).Split('\n');

                if (Question == true) Question = false;
                if (SendMessageMode == SendMsgMode.VOLT)
                {
                    if (float.TryParse(s[0], out SetVolt) == false) { }
                }
                else
                {
                    //PowerOnOff = s[0] == "1" ? true : false;
                }
            }
            catch
            {

            }
            finally
            {
                SendFlag = false;
            }
            return;
        }
    }
}

