#define PROGRAM_RUNNING
#define IMS_USE
#undef IMS_USE

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using MES;

namespace PSeatTester
{
    public interface MyInterface
    {
        /// <summary>
        /// 공용 함수를 호출한다.
        /// </summary>
        COMMON_FUCTION 공용함수 { get; }
        /// <summary>
        /// 환경 변수 설정을 진해안다.
        /// </summary>
        __Config__ GetConfig { get; set; }
        /// <summary>
        /// LIN 통신 함수를 호출한다.
        /// </summary>
        //LinControl GetLin { get; }
        /// <summary>
        /// CAN 통신 함수를 호출한다.
        /// </summary>
        __CanControl GetCan { get; }
        /// <summary>
        /// 현재 검사를 진행하는지 정보를 읽어 오거나 설정한다.
        /// </summary>
        bool isRunning { get; set; }
        /// <summary>
        /// 프로그램 종료 여부를 같는다.
        /// </summary>
        bool isExit { get; }
        short GetCanChannel { get; }
        //short GetLinChannel { get; }
        CanMap GetCanReWrite { get; }
        PinMapStruct GetPinMap { get; set; }
        PanelMeter GetPMeter { get; }
        /// <summary>
        /// 파워 제어 함수를 호출한다.
        /// </summary>
        PowerControl GetPower { get; }
        IOControl GetIO { get; }
        SoundMeter GetSound { get; }
    }
    public partial class MainForm : Form, MyInterface
    {
        private COMMON_FUCTION ComF = new COMMON_FUCTION();
        private __Config__ Config;
        private IOControl IOPort = null;
        private MES_Control PopCtrl = null;
        //private LinControl LinCtrl = null;
        private __CanControl CanCtrl = null;
        private PinMapStruct PinMap = new PinMapStruct();
        private __CheckItem__ CheckItem = new __CheckItem__();
        private __Spec__ mSpec = new __Spec__();
        private PowerControl PwCtrl = null;
        private PanelMeter pMeter = null;
        private Color NotSelectColor;
        private Color SelectColor;
        private SoundMeter Sound = null;
        private __TestData__ mData = new __TestData__();
        private KALMAN_FILETER CurrFilter = null;
        
        private CanMap CanReWrite = null;
        private __Infor__ Infor = new __Infor__();

        public const float PSEAT_OFF_CURRENT = 0.5F;
        private MES_Control.__ReadMesData__ SendBarcodeData = new MES_Control.__ReadMesData__();
        private bool SlideMotorMoveEndFlag { get; set; }
        private bool ReclineMotorMoveEndFlag { get; set; }
        private bool TiltMotorMoveEndFlag { get; set; }
        private bool HeightMotorMoveEndFlag { get; set; }

        public MainForm()
        {
            InitializeComponent();
        }

        public COMMON_FUCTION 공용함수
        {
            get
            {
                return ComF;
            }
        }

        public __Config__ GetConfig
        {
            get
            {
                return Config;
            }
            set
            {
                Config = value;
                ConfigSetting ReadConfig = new ConfigSetting();
                ReadConfig.ReadWriteConfig = Config;
            }
        }


        private void MainForm_Load(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;

            ConfigSetting ReadConfig = new ConfigSetting();
            Config = ReadConfig.ReadWriteConfig;

            CurrFilter = new KALMAN_FILETER(this);

            NotSelectColor = fpSpread1.ActiveSheet.Cells[0, 0].BackColor;
            SelectColor = Color.FromArgb(172, 227, 175);

            OpenInfor();

            PopCtrl = new MES_Control(ClientIp: Config.Client, ServerIp: Config.Server, mControl: this);
            PopCtrl.Open();
            IOPort = new IOControl(Board: Config.Board, PC: Config.PC, mControl: this);

            if (Config.UseSmartIO == true)
                IOPort.Open(Config.SmartIOPort.Port, Config.SmartIOPort.Speed, Config.UseSmartIO);
            else IOPort.Open();

            if (IOPort.isOpen == false) MessageBox.Show("I/O Card IP 를 확인해 주십시오.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            if (IOPort.isConnection == false) MessageBox.Show("I/O Card 와 접속되지 않습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            if (PopCtrl.isClientConnection == false)
            {
                //MessageBox.Show("서버와 접속되지 않습니다.", "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            label2.Text = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();

            테스ToolStripMenuItem_Click(테스ToolStripMenuItem, new EventArgs());
#if PROGRAM_RUNNING

            comboBox4.SelectedItem = null;
            comboBox3.SelectedItem = null;
            //comboBox5.SelectedItem = "12 WAY";            

            if (SerialOpen() == false)
            {

            }

            CanCtrl = new __CanControl(this);
            //LinCtrl = new LinControl(false, this);

            CanReWrite = new CanMap(this);
            CanLinDefaultSetting();
            if (0 <= Config.Can.Device)
            {
                CanCtrl.OpenCan(0, CanPosition(), (short)Config.Can.Speed, false);
                if (CanCtrl.isOpen(0) == false) { }

            }

            //if (0 <= Config.Lin.Device)
            //{
            //    if (LinCtrl.isOpen(LinChannel) == false)
            //    {
            //        LinCtrl.LinOpen(LinPosition(Config.Lin.Device), LinControl.HW_MODE.SLAVE, Config.Lin.Speed);
            //        if (LinCtrl.isOpen(LinChannel) == false) { }
            //    }
            //}

            ComF.ReadFileListNotExt(Program.SPEC_PATH.ToString(), "*.Spc", COMMON_FUCTION.FileSortMode.FILENAME_ODERBY);
            List<string> FList = ComF.GetFileList;

            if (0 < FList.Count)
            {
                ModelBoxChangeFlag = true;
                comboBox1.Items.Clear();
                foreach (string s in FList) comboBox1.Items.Add(s);

                if (0 < comboBox1.Items.Count)
                {
                    comboBox1.SelectedIndex = 0;
                }

                if (comboBox1.SelectedItem != null)
                {
                    string sName = Program.SPEC_PATH.ToString() + "\\" + comboBox1.SelectedItem.ToString() + ".Spc";
                    if (File.Exists(sName) == true) ComF.OpenSpec(sName, ref mSpec);

                    SpecName = sName;
                    mName = comboBox1.SelectedItem.ToString();
                }
                ModelBoxChangeFlag = false;
            }
#endif
            CheckItem.LhdRhd = LHD_RHD.LHD;
            CheckItem.LhRh = LH_RH.LH;
            CheckItem.ProductTestRunFlag = false;
            CheckItem.PSeat = PSEAT_TYPE.MANUAL;
            CheckItem.PSeat12Way = false;
            CheckItem.Can = false;
            FormLoadFirstSpecDisplayTime = ComF.timeGetTimems();
            comboBox4.SelectedItem = "IMS";
            timer1.Enabled = true;
            timer2.Enabled = true;
            DisplaySpec();
            return;
        }

        private bool FormLoadFirstSpecDisplay = true;
        private long FormLoadFirstSpecDisplayTime = 0;

        private string mName = null;
        private string SpecName = null;

        private bool SerialOpen()
        {
            bool Flag = true;

            Sound = new SoundMeter(Config.NoiseMeter.Port, true);
            if ((Config.NoiseMeter.Port != null) && (Config.NoiseMeter.Port != string.Empty))
            {
                if (Sound.Open() == false) Flag = false;
            }

            PwCtrl = new PowerControl(Config.Power);
            if (PwCtrl.IsOpen == false)
            {
                MessageBox.Show("파워 제어용 포트를 열수 없습니다.", "경고", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Flag = false;
            }

            pMeter = new PanelMeter(this);
            pMeter.Open(Config.Panel);
            if (pMeter.isOpen == false)
            {
                MessageBox.Show("판넬메타 통신 포트를 열수 없습니다.", "경고", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Flag = false;
            }

            return Flag;
        }
        public SoundMeter GetSound
        {
            get { return Sound; }
        }
        public PowerControl GetPower
        {
            get { return PwCtrl; }
        }

        //public LinControl GetLin
        //{
        //    get { return LinCtrl; }
        //}
        public __CanControl GetCan
        {
            get { return CanCtrl; }
        }

        public CanMap GetCanReWrite
        {
            get { return CanReWrite; }
        }
        //public short GetLinChannel
        //{
        //    get { return LinChannel; }
        //}
        public short GetCanChannel
        {
            get { return CanChannel; }
        }
        private void CanLinDefaultSetting()
        {
            CanReWrite.CanLinDefaultSetting();
            return;
        }

        public PanelMeter GetPMeter
        {
            get
            {
                return pMeter;
            }
        }

        private bool ExitFlag { get; set; }
        private bool RunningFlag { get; set; }
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (RunningFlag == true)
            {
                e.Cancel = true;
                return;
            }


            ExitFlag = true;

            CloseForm xClose = new CloseForm()
            {
                Owner = this,
                WindowState = FormWindowState.Normal,
                StartPosition = FormStartPosition.CenterScreen,
                //Location = new Point(1, 1),
                FormBorderStyle = FormBorderStyle.FixedSingle,
                MinimizeBox = false,
                MaximizeBox = false,
                TopMost = true
            };


            xClose.FormClosing += delegate (object sender1, FormClosingEventArgs e1)
            {
                e1.Cancel = false;
                xClose.Dispose();
#if PROGRAM_RUNNING
                if (Sound != null)
                {
                    if (Sound.Connection == true) Sound.StartStopMesurement(false);
                    ComF.timedelay(1000);
                }
                if (CanCtrl.isOpen(0) == true) CanCtrl.CanClose(0);
                //if (LinCtrl != null)
                //{
                //    if (LinCtrl.isOpen(LinChannel) == true) LinCtrl.LinClose();
                //}

                if (PwCtrl.IsOpen == true) PwCtrl.Close();
                if (pMeter.isOpen == true) pMeter.Close();
#endif
                System.Diagnostics.Process[] mProcess = System.Diagnostics.Process.GetProcessesByName(Application.ProductName);
                foreach (System.Diagnostics.Process p in mProcess) p.Kill();
            };

            e.Cancel = true;
            xClose.Show();
            return;
        }
        public bool isRunning
        {
            get { return RunningFlag; }
            set { RunningFlag = value; }
        }
        public bool isExit
        {
            get { return ExitFlag; }
        }


        //        private short LinChannel = 0;
        //        public short LinPosition(short Pos)
        //        {
        //#if PROGRAM_RUNNING
        //            short ID = 0;

        //            string[] Device = LinCtrl.GetDevice;

        //            for (short i = 0; i < Device.Length; i++)
        //            {
        //                string s = Device[i];
        //                string s1 = Pos.ToString() + " - ID";

        //                if (0 <= s.IndexOf(s1))
        //                {
        //                    ID = i;
        //                    break;
        //                }
        //            }

        //            LinChannel = ID;
        //            return ID;
        //#else
        //            return 0;
        //#endif
        //        }

        private short CanChannel = 0;
        public short CanPosition()
        {
#if PROGRAM_RUNNING
            //short ID = 0;

            string[] Device = CanCtrl.GetDevice;

            //for (short i = 0; i < Device.Length; i++)
            //{
            //    string s = Device[i];
            //    string s1 = "0x" + Pos.ToString("X2");

            //    if (0 <= s.IndexOf(s1))
            //    {
            //        ID = i;
            //        break;
            //    }
            //}
            short Pos = -1;
            string s1 = "Device=" + Config.Can.Device.ToString();
            string s2 = "Channel=" + Config.Can.Channel.ToString() + "h";

            foreach (string s in Device)
            {

                if (0 <= s.IndexOf(s1))
                {
                    if (0 <= s.IndexOf(s2))
                    {
                        string ss = s.Substring(s.IndexOf("ID=") + "ID=".Length);
                        string[] ss1 = ss.Split(',');
                        if (1 < ss1.Length)
                        {
                            string ss2 = ss1[0].Replace("(", null);

                            ss2 = ss2.Replace(")", null);
                            Pos = (short)ComF.StringToHex(ss2);
                        }
                    }
                }
            }

            if (Pos == -1)
            {
                Pos = Config.Can.ID;
            }
            CanChannel = Pos;
            return Pos;
#else
            return 0;
#endif
        }

        //private void toolStripButton4_Click(object sender, EventArgs e)
        //{
        //    //Master Lin 설정
        //    if (toolStripButton6.Text == "로그인")
        //    {
        //        //MessageBox.Show("로그인을 먼저 진행하십시오.", "경고");
        //        //uMessageBox.Show(promptText: "로그인을 먼저 진행하십시오.", title: "경고");
        //        MessageBox.Show(this, "로그인을 먼저 진행하십시오.", "경고", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return;
        //    }
        //    if (RunningFlag == false)
        //    {
        //        if (0 <= Config.Lin.Device)
        //        {
        //            if (LinCtrl.isOpen(LinChannel) == false)
        //            {
        //                LinCtrl.LinOpen(LinPosition(Config.Lin.Device), LinControl.HW_MODE.MASTER, Config.Lin.Speed);
        //                if (LinCtrl.isOpen(LinChannel) == false) { }
        //            }
        //        }
        //        LinCtrl.LinConfigSetting(LinChannel);
        //    }
        //    return;
        //}

        private SelfTest SelfForm = null;

        private void OpenFormClose()
        {
            if (SelfForm != null) SelfForm.Close();
            if (SpecForm != null) SpecForm.Close();
            return;
        }

        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            if (IOPort.GetAuto == false)
            {
                if (SelfForm == null)
                {
                    OpenFormClose();
                    panel1.SendToBack();
                    panel1.Visible = false;

                    panel2.Visible = true;
                    if (panel2.Parent != this) panel2.Parent = this;
                    panel2.BringToFront();
                    SelfForm = new SelfTest(this)
                    {
                        MaximizeBox = false,
                        MinimizeBox = false,
                        ControlBox = false,
                        ShowIcon = false,
                        StartPosition = FormStartPosition.CenterScreen,
                        WindowState = FormWindowState.Maximized,
                        TopMost = false,
                        TopLevel = false,
                        FormBorderStyle = FormBorderStyle.None,
                        Location = new Point(0, 0),
                        Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
                    };

                    SelfForm.FormClosing += delegate (object sender1, FormClosingEventArgs e1)
                    {
                        e1.Cancel = false;
                        SelfForm.Parent = null;
                        SelfForm.Dispose();
                        SelfForm = null;
                    };

                    SelfForm.Parent = panel2;
                    SelfForm.Show();
                }
            }
            return;
        }

        private void 테스ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripButton1.Visible = true; //테스트
            toolStripButton4.Visible = IOPort.GetAuto == false ? true : false;
            //설정
            toolStripButton2.Visible = false;
            toolStripButton3.Visible = false;
            //toolStripButton4.Visible = false;
            //toolStripButton5.Visible = false;            
            toolStripSeparator1.Visible = false;
            //toolStripSeparator2.Visible = false;
            //toolStripSeparator3.Visible = false;            
            //로그인
            toolStripButton6.Visible = false;
            toolStripButton10.Visible = false;

            //보기
            toolStripButton7.Visible = false;
            //종료
            toolStripButton8.Visible = false;
            toolStripButton9.Visible = false;
            toolStripSeparator4.Visible = false;
            return;
        }

        private void 설정ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripButton1.Visible = false; //테스트
            toolStripButton4.Visible = false;
            //설정
            toolStripButton2.Visible = true;
            toolStripButton3.Visible = true;
            //toolStripButton4.Visible = true;
            //toolStripButton5.Visible = true;
            toolStripSeparator1.Visible = true;
            //toolStripSeparator2.Visible = true;
            //toolStripSeparator3.Visible = true;
            //로그인
            toolStripButton6.Visible = false;
            toolStripButton10.Visible = false;
            //toolStripSeparator4.Visible = false;
            //보기
            toolStripButton7.Visible = false;
            //종료
            toolStripButton8.Visible = false;
            toolStripButton9.Visible = false;
            toolStripSeparator4.Visible = false;
            return;
        }

        private void 로그인ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripButton1.Visible = false; //테스트
            toolStripButton4.Visible = false;
            //설정
            toolStripButton2.Visible = false;
            toolStripButton3.Visible = false;
            //toolStripButton4.Visible = false;
            //toolStripButton5.Visible = false;
            toolStripSeparator1.Visible = false;
            //toolStripSeparator2.Visible = false;
            //toolStripSeparator3.Visible = false;
            //로그인
            toolStripButton6.Visible = true;
            toolStripButton10.Visible = true;
            toolStripSeparator4.Visible = true;
            //보기
            toolStripButton7.Visible = false;
            //종료
            toolStripButton8.Visible = false;
            toolStripButton9.Visible = false;
            //toolStripSeparator4.Visible = false;
            return;
        }

        private void 보기ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripButton1.Visible = false; //테스트
            toolStripButton4.Visible = false;
            //설정
            toolStripButton2.Visible = false;
            toolStripButton3.Visible = false;
            //toolStripButton4.Visible = false;
            //toolStripButton5.Visible = false;
            toolStripSeparator1.Visible = false;
            //toolStripSeparator2.Visible = false;
            //toolStripSeparator3.Visible = false;
            //로그인
            toolStripButton6.Visible = false;
            toolStripButton10.Visible = false;
            toolStripSeparator4.Visible = false;
            //보기
            toolStripButton7.Visible = true;
            //종료
            toolStripButton8.Visible = false;
            toolStripButton9.Visible = false;
            //toolStripSeparator4.Visible = false;
            return;
        }

        private void 종료ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripButton1.Visible = false; //테스트
            toolStripButton4.Visible = false;
            //설정
            toolStripButton2.Visible = false;
            toolStripButton3.Visible = false;
            //toolStripButton4.Visible = false;
            //toolStripButton5.Visible = false;
            toolStripSeparator1.Visible = false;
            //toolStripSeparator2.Visible = false;
            //toolStripSeparator3.Visible = false;
            //로그인
            toolStripButton6.Visible = false;
            toolStripButton10.Visible = false;
            //toolStripSeparator4.Visible = false;
            //보기
            toolStripButton7.Visible = false;
            //종료
            toolStripButton8.Visible = true;
            toolStripButton9.Visible = true;
            toolStripSeparator4.Visible = true;
            return;
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            if (toolStripButton6.Text == "로그인")
            {
                PasswordCheckForm pass = new PasswordCheckForm();
                //아래와 같이 해 주면 폼을 닫을때 Dialog로 오픈을 하지 않아도 {} 안을 실행하게 된다. 동시에 해당 폼의 FormClosing이 동시에 실행되므로 Dialog로 오픈한것 같은 효과를 얻는다.
                pass.FormClosing += delegate (object sender1, FormClosingEventArgs e1)
                {
                    if (pass.result == true)
                    {
                        toolStripButton6.Text = "로그아웃";
                        로그인ToolStripMenuItem.Text = "로그아웃";
                        toolStripButton6.Image = Properties.Resources.Pad_Unlock_36Pixel1;

                        toolStripButton2.Enabled = true;
                        toolStripButton3.Enabled = true;
                        //toolStripButton4.Enabled = true;
                        //toolStripButton5.Enabled = true;
                    }
                };
                pass.Show();
            }
            else
            {
                toolStripButton6.Text = "로그인";
                로그인ToolStripMenuItem.Text = "로그인";
                toolStripButton6.Image = Properties.Resources.Pad_Lock3;
                toolStripButton2.Enabled = false;
                toolStripButton3.Enabled = false;
                //toolStripButton4.Enabled = false;
                //toolStripButton5.Enabled = false;
            }
            return;
        }

        private void toolStripButton10_Click(object sender, EventArgs e)
        {
            if (toolStripButton6.Text == "로그인")
            {
                //MessageBox.Show("로그인을 먼저 진행하십시오.", "경고");
                //uMessageBox.Show(promptText: "로그인을 먼저 진행하십시오.", title: "경고");
                MessageBox.Show(this, "로그인을 먼저 진행하십시오.", "경고", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            PasswordSetForm set = new PasswordSetForm();
            set.MinimizeBox = false;
            set.MaximizeBox = false;
            set.FormBorderStyle = FormBorderStyle.FixedSingle;
            set.Show();
            return;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (toolStrip1.Enabled == IOPort.GetAuto) toolStrip1.Enabled = !IOPort.GetAuto;
            label2.Text = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();

            if (PopCtrl.isClientConnection != led7.Value.AsBoolean) led7.Value.AsBoolean = PopCtrl.isClientConnection;
            if (IOPort.isConnection != led1.Value.AsBoolean) led1.Value.AsBoolean = IOPort.isConnection;

            if (Infor.Date != DateTime.Now.ToString("yyyyMMdd"))
            {
                string Path = Program.DATA_PATH.ToString() + "\\" + DateTime.Now.ToString("yyyyMM") + ".xls";
                Infor.Date = DateTime.Now.ToString("yyyyMMdd");
                Infor.DataName = Path;
                Infor.TotalCount = 0;
                
                Infor.OkCount = 0;
                Infor.NgCount = 0;
                SaveInfor();
            }
            return;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (panel1.Visible == false)
            {
                OpenFormClose();
                panel2.SendToBack();
                panel2.Visible = false;

                panel1.Visible = true;
                panel1.BringToFront();
                if (panel1.Parent != this) panel1.Parent = this;
            }
            return;
        }

        private void toolStripButton9_Click(object sender, EventArgs e)
        {
            if (IOPort.GetAuto == false)
            {
                if (RunningFlag == false)
                {
                    this.Close();
                }
            }
            return;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ModelBoxChangeFlag == true) return;
            string s = comboBox1.SelectedItem.ToString();

            string sName = Program.SPEC_PATH.ToString() + "\\" + s + ".Spc";

            if (File.Exists(sName) == true)
            {
                if(SpecName != sName) ComF.OpenSpec(sName, ref mSpec);
                SpecName = sName;
            }
            DisplaySpec();
            mName = comboBox1.SelectedItem.ToString();
            return;
        }

        private MES_Control.__ReadMesData__ PopReadData = new MES_Control.__ReadMesData__()
        {
            Barcode = null,
            Check = null,
            Date = null,
            LineCode = null,
            MachineNo = null
        };
        private void PopDataCheck()
        {
            if (PopReadData.Check == null) return;

            //RunningFlag = false;
            StartSwOnFlag = false;
            RetestFlag = false;
            ProductOutputFlag = false;
            ScreenInit();
            if (IOPort.TestOKOnOff == true) IOPort.TestOKOnOff = false;
            if (IOPort.YellowLampOnOff == true) IOPort.YellowLampOnOff = false;
            if (IOPort.RedLampOnOff == true) IOPort.RedLampOnOff = false;
            if (IOPort.GreenLampOnOff == true) IOPort.GreenLampOnOff = false;

            //CheckItem.Height = false;
            //CheckItem.Slide = false;
            //CheckItem.Tilt = false;
            //CheckItem.Recline = false;
            CheckItem.PSeat10Way = false;
            CheckItem.PSeat12Way = false;
            CheckItem.PSeat4Way = false;
            CheckItem.PSeat8Way = false;
            
            ProductOutputFlag = false;
            textBox1.Text = PopReadData.Barcode;
            label18.Text = label18.Text;
            label17.Text = PopReadData.Barcode;
            label23.Text = PopReadData.Check;

            SendBarcodeData.Barcode = PopReadData.Barcode;
            SendBarcodeData.Check = PopReadData.Check;
            SendBarcodeData.Date = PopReadData.Date;
            SendBarcodeData.LineCode = PopReadData.LineCode;
            SendBarcodeData.MachineNo = PopReadData.MachineNo;

            if (IOPort.GetAuto == true)
            {
                PopReadOk = true;

                bool Flag = false;
                //IMS
                if (PopReadData.Check.Substring(0, 1) == "0")
                {
                    if (CheckItem.PSeat == PSEAT_TYPE.IMS) Flag = false;
                    CheckItem.PSeat = PSEAT_TYPE.POWER;
                }
                else
                {
                    if (CheckItem.PSeat != PSEAT_TYPE.IMS) Flag = false;
                    if (comboBox4.SelectedItem == null) comboBox4.SelectedItem = "IMS";
                    CheckItem.PSeat = PSEAT_TYPE.IMS;
                    comboBox4.SelectedItem = "IMS";
                    CheckItem.Can = true;
                    CheckItem.PSeat12Way = true;
                }
                //Slide            
                if (PopReadData.Check.Substring(1, 1) == "0")
                {
                    CheckItem.Slide = false;
                }
                else
                {
                    CheckItem.Slide = true;
                }

                //Tilt
                if (PopReadData.Check.Substring(2, 1) == "0")
                {
                    CheckItem.Tilt = false;
                }
                else
                {
                    CheckItem.Tilt = true;
                }
                //Height
                if (PopReadData.Check.Substring(3, 1) == "0")
                {
                    CheckItem.Height = false;
                }
                else
                {
                    CheckItem.Height = true;
                }

                //Recline
                if (PopReadData.Check.Substring(4, 1) == "0")
                {
                    CheckItem.Recline = false;
                }
                else
                {
                    CheckItem.Recline = true;
                }

                if (PopReadData.Check.Substring(15, 1) == "1")
                {
                    if (comboBox2.SelectedItem == null) comboBox2.SelectedItem = "RHD";
                    if (comboBox2.SelectedItem.ToString() != "RHD")
                    {
                        comboBox2.SelectedItem = "RHD";
                        CheckItem.LhdRhd = LHD_RHD.RHD;
                        Flag = true;
                    }
                }
                else
                {
                    if (comboBox2.SelectedItem == null) comboBox2.SelectedItem = "LHD";
                    if (comboBox2.SelectedItem.ToString() != "LHD")
                    {
                        comboBox2.SelectedItem = "LHD";
                        CheckItem.LhdRhd = LHD_RHD.LHD;
                        Flag = true;
                    }
                }

                if (PopReadData.Check.Substring(16, 1) == "1")
                {
                    if (comboBox3.SelectedItem == null) comboBox3.SelectedItem = "RH";
                    if (comboBox3.SelectedItem.ToString() != "RH")
                    {
                        comboBox3.SelectedItem = "RH";
                        CheckItem.LhRh = LH_RH.RH;
                        Flag = true;
                    }
                }
                else
                {
                    if (comboBox3.SelectedItem == null) comboBox3.SelectedItem = "LH";
                    if (comboBox3.SelectedItem.ToString() != "LH")
                    {
                        comboBox3.SelectedItem = "LH";
                        CheckItem.LhRh = LH_RH.LH;
                        Flag = true;
                    }
                }


                //Lumber 2way
                if (PopReadData.Check.Substring(5, 1) == "0")
                {
                }
                else
                {
                    if(((comboBox3.SelectedItem.ToString() == "LH") && comboBox2.SelectedItem.ToString() == "LHD") || ((comboBox3.SelectedItem.ToString() == "RH") && (comboBox2.SelectedItem.ToString() == "RHD")))
                    {
                        CheckItem.PSeat = PSEAT_TYPE.POWER;
                        comboBox4.SelectedItem = "POWER";
                        CheckItem.Can = false;
                        CheckItem.PSeat10Way = true;
                        CheckItem.PSeat12Way = false;
                        CheckItem.PSeat4Way = false;
                        Flag = true;
                    }
                    //else
                    //{
                    //    if(comboBox4.SelectedItem == null) comboBox4.SelectedItem = "WALK IN";
                    //    if (comboBox4.SelectedItem.ToString() != "WALK IN")
                    //    {
                    //        if (CheckItem.WalkIn == false) Flag = true;
                    //        CheckItem.PSeat = PSEAT_TYPE.POWER;
                    //        comboBox4.SelectedItem = "WALK IN";
                            
                    //        CheckItem.PSeat = PSEAT_TYPE.POWER;
                    //        comboBox4.SelectedItem = "WALK IN";
                    //        CheckItem.Can = false;
                    //        CheckItem.PSeat12Way = false;
                    //        CheckItem.PSeat10Way = false;
                    //        CheckItem.PSeat4Way = false;
                    //        CheckItem.WalkIn = true;
                    //        Flag = true;
                    //    }
                    //}
                }
                //Lumber 4way
                if (PopReadData.Check.Substring(6, 1) == "0")
                {
                }
                else
                {
                    if (PopReadData.Check.Substring(0, 1) == "0")
                    {
                        if (comboBox4.SelectedItem == null)
                        {
                            comboBox4.SelectedItem = "POWER";
                            CheckItem.PSeat = PSEAT_TYPE.POWER;
                            comboBox4.SelectedItem = "POWER";
                            CheckItem.Can = false;
                            CheckItem.PSeat12Way = true;
                            CheckItem.PSeat4Way = true;
                            CheckItem.PSeat10Way = false;
                            CheckItem.WalkIn = false;
                            Flag = true;
                        }
                        else
                        {
                            if (comboBox4.SelectedItem.ToString() != "POWER")
                            {
                                CheckItem.PSeat = PSEAT_TYPE.POWER;
                                comboBox4.SelectedItem = "POWER";
                                CheckItem.Can = false;
                                CheckItem.PSeat12Way = true;
                                CheckItem.PSeat4Way = true;
                                CheckItem.PSeat10Way = false;
                                CheckItem.WalkIn = false;
                                Flag = true;
                            }
                            else
                            {
                                if((CheckItem.PSeat12Way == false) || (CheckItem.PSeat4Way == false) || (CheckItem.PSeat10Way == true) || (CheckItem.WalkIn == true))
                                {
                                    CheckItem.PSeat = PSEAT_TYPE.POWER;
                                    comboBox4.SelectedItem = "POWER";
                                    CheckItem.Can = false;
                                    CheckItem.PSeat12Way = true;
                                    CheckItem.PSeat4Way = true;
                                    CheckItem.PSeat10Way = false;
                                    CheckItem.WalkIn = false;
                                    Flag = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        CheckItem.PSeat12Way = true;
                        Flag = true;
                    }
                }

                if ((PopReadData.Check.Substring(5, 1) == "0") && (PopReadData.Check.Substring(6, 1) == "0"))
                {
                    if (((comboBox3.SelectedItem.ToString() == "RH") && comboBox2.SelectedItem.ToString() == "LHD") || ((comboBox3.SelectedItem.ToString() == "LH") && (comboBox2.SelectedItem.ToString() == "RHD")))
                    {
                        if (comboBox4.SelectedItem == null) comboBox4.SelectedItem = "WALK IN";
                        if (comboBox4.SelectedItem.ToString() != "WALK IN")
                        {
                            if (CheckItem.WalkIn == false) Flag = true;
                            CheckItem.PSeat = PSEAT_TYPE.POWER;
                            comboBox4.SelectedItem = "WALK IN";

                            CheckItem.PSeat = PSEAT_TYPE.POWER;
                            comboBox4.SelectedItem = "WALK IN";
                            CheckItem.Can = false;
                            CheckItem.PSeat12Way = false;
                            CheckItem.PSeat10Way = false;
                            CheckItem.PSeat4Way = false;
                            CheckItem.WalkIn = true;
                            Flag = true;
                        }
                    }
                }
                //Sound      
                if (PopReadData.Check.Substring(7, 1) == "0")
                {
                    CheckItem.Sound = false;
                }
                else
                {
                    CheckItem.Sound = true;
                }

                Flag = OpenSpec(true);

                //----------------------------------------------------------------------
                //if (Flag == true) 
                DisplaySpec();

                if (isCheckItem == false) label5.Text = "검사 사양이 아닙니다.";
            }
            return;
        }

        private bool ProductOutputFlag { get; set; }

        private bool BuzzerRunFlag = false;
        private bool BuzerOnOff = false;
        private long BuzzerLast = 0;
        private long BuzzerFirst = 0;
        private short BuzzerOnCount = 0;
        //private bool JigUpFlag = false;
        //private short JigUpCount = 0;
        private bool StartSwOnFlag = false;
        private bool RetestFlag = false;
        private void timer2_Tick(object sender, EventArgs e)
        {
            try
            {
                timer2.Enabled = false;

                if ((DateTime.Now.Hour == 5) && (DateTime.Now.Minute == 1))
                {
                    if (Infor.ReBootingFlag == false)
                    {
                        Infor.ReBootingFlag = true;
                        SaveInfor();
                        //if (Sound != null)
                        //{
                            //if (Sound.Connection == true) Sound.StartStopMesurement(false);
                            //Sound.Close();
                            //ComF.timedelay(1000);
                        //}

                        ComF.WindowRestartToDelay(0);
                        this.Close();
                        //System.Diagnostics.Process[] mProcess = System.Diagnostics.Process.GetProcessesByName(Application.ProductName);
                        //foreach (System.Diagnostics.Process p in mProcess) p.Kill();
                    }
                }


                //this.Text = IOPort.GetInData[0].ToString("X");
                if (BuzzerRunFlag == true)
                {
                    BuzzerLast = ComF.timeGetTimems();
                    if (BuzerOnOff == true)
                    {
                        if (700 <= (BuzzerLast - BuzzerFirst))
                        {
                            BuzerOnOff = false;
                            IOPort.BuzzerOnOff = BuzerOnOff;
                            BuzzerFirst = ComF.timeGetTimems();
                            BuzzerOnCount++;
                        }
                    }
                    else
                    {
                        if (500 <= (BuzzerLast - BuzzerFirst))
                        {
                            if (BuzzerOnCount < 3)
                            {
                                BuzerOnOff = true;
                                IOPort.BuzzerOnOff = BuzerOnOff;
                                BuzzerFirst = ComF.timeGetTimems();
                            }
                            else
                            {
                                BuzerOnOff = false;
                                IOPort.BuzzerOnOff = BuzerOnOff;
                                BuzzerOnCount = 0;
                                BuzzerRunFlag = false;
                            }
                        }
                    }
                }


                if (MemSetButton == true)
                {
                    if (2000 <= (ComF.timeGetTimems() - MemSetButtonToTime))
                    {
                        MemSetButton = false;
                        ledBulb1.On = false;
                    }
                }

                if (panel2.Visible == false)
                {
                    if ((IOPort.GetJigUp == false) && (IOPort.SetDoorOpen == true)) IOPort.SetDoorOpen = false;

                    if ((IOPort.SetPinConnection == true) || (IOPort.GetPinConnectionFwd == true) || (RunningFlag == true))
                    {
                        //if (IOPort.SetDoorOpen == false) IOPort.SetDoorOpen = true;
                        if (IOPort.TestINGOnOff == false) IOPort.TestINGOnOff = true;
                    }
                    else if ((IOPort.SetPinConnection == false) && (IOPort.GetPinConnectionBwd == true) && (RunningFlag == false))
                    {
                        //if (IOPort.SetJigDown == true) IOPort.SetJigDown = false;
                        if (IOPort.TestINGOnOff == true) IOPort.TestINGOnOff = false;
                    }

                    if (IOPort.ProductInOut != IOPort.GetProductIn) IOPort.ProductInOut = IOPort.GetProductIn;
                    if (led4.Value.AsBoolean != IOPort.GetProductIn) led4.Value.AsBoolean = IOPort.GetProductIn;

                    if ((IOPort.GetProductIn == false) && (RunningFlag == false))
                    {
                        if (PowerOnFlag == true) PowerOnOff = false;
                        if (SlideDeliveryPosEndFlag == true) SlideDeliveryPosEndFlag = false;
                        if ((IOPort.SetPinConnection == false) && (IOPort.GetPinConnectionFwd == false))
                        {
                            if (IOPort.TestINGOnOff == true) IOPort.TestINGOnOff = false;
                        }

                        if (RetestFlag == true) RetestFlag = false;
                        if (panel7.Visible == true) panel7.Visible = false;
                        if (label5.ForeColor != Color.White) label5.ForeColor = Color.White;
                        if (IOPort.TestOKOnOff == true) IOPort.TestOKOnOff = false;
                        if (IOPort.SetDoorOpen == true) IOPort.SetDoorOpen = false;
                        if (IOPort.YellowLampOnOff == true) IOPort.YellowLampOnOff = false;
                        if (IOPort.RedLampOnOff == true) IOPort.RedLampOnOff = false;
                        if (IOPort.GreenLampOnOff == true) IOPort.GreenLampOnOff = false;

                        //제품이 없을 경우 배출 제품이 있다고 설정된 값을 초기화 한다.
                        if (ProductOutputFlag == true)
                        {
                            ProductOutputFlag = false;
                            if (IOPort.GetAuto == false) ScreenInit();
                            //PopReadOk = false;
                            PopReadOkOld = false;
                        }

                        if (label16.Text != "") label16.Text = "";
                        if (label5.Text != "대기중입니다.") label5.Text = "대기중입니다.";
                    }

                    if (IOPort.GetAuto == false)
                    {
                        if (IOPort.SetPinConnection == IOPort.GetPinConnectionSw) IOPort.SetPinConnection = IOPort.GetPinConnectionSw;
                    }
                }


                if (led9.Value.AsBoolean != IOPort.GetJigUp) led9.Value.AsBoolean = IOPort.GetJigUp;
                if (led6.Value.AsBoolean != IOPort.TestOKOnOff) led6.Value.AsBoolean = IOPort.TestOKOnOff;
                if (led5.Value.AsBoolean != IOPort.TestINGOnOff) led5.Value.AsBoolean = IOPort.TestINGOnOff;

                if ((IOPort.GetSlideMidSensor == true) || (SlideDeliveryPosEndFlag == true))
                {
                    if (led10.Value.AsBoolean == false) led10.Value.AsBoolean = true;
                }
                else
                {
                    if (led10.Value.AsBoolean == true) led10.Value.AsBoolean = false;
                }
                if (led8.Value.AsBoolean != IOPort.GetReclineMidSensor) led8.Value.AsBoolean = IOPort.GetReclineMidSensor;

                if (PopCtrl.isReading == true)
                {
                    if (RunningFlag == false)
                    {
                        PopReadData = PopCtrl.GetReadData;
                        PopDataCheck();                        

                        SaveLogData = "RECEIVE - " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                        SaveLogData = PopCtrl.SourceData;                        
                    }
                    PopCtrl.isReading = false;
                }

                if (IOPort.GetAuto == false)
                {
                    if (led2.Indicator.Text != "수동")
                    {
                        led2.Indicator.Text = "수동";
                        if (panel1.Visible == true)
                            toolStripButton4.Visible = true;
                        else toolStripButton4.Visible = false;

                        label17.Visible = false;
                        textBox1.Visible = true;
                    }
                    if (toolStrip1.Enabled == false) toolStrip1.Enabled = true;
                    if (RunningFlag == false)
                    {
                        //자동 모드에서는 수동 조작 버튼을 인식하지 않는다.
                        //다만 검사를 진행할 경우에만 자동 컨넥팅 모드에 따라 동작할 뿐이다.
                        if (panel2.Visible == false) //세프 테스트 모드일때 스위치 입력으로 버튼이 안먹는것을 막기 위한 조치
                        {
                            if (IOPort.GetPinConnectionSw != IOPort.SetPinConnection)
                            {
                                IOPort.SetPinConnection = IOPort.GetPinConnectionSw;

                                PowerOnOff = IOPort.GetPinConnectionSw;
                                BattOnOff = IOPort.GetPinConnectionSw;
                            }
                        }
                    }
                }
                else
                {                    
                    if (led2.Indicator.Text != "자동")
                    {
                        led2.Indicator.Text = "자동";
                        toolStripSeparator4.Visible = false;
                        IOPort.SetPinConnection = false;
                        label17.Visible = true;
                        textBox1.Visible = false;
                        PopDataCheck();
                        PopReadOk = false;
                        if (panel4.Visible == true) panel4.Visible = false;
                    }

                    if (panel1.Visible == true)
                    {
                        if (toolStrip1.Enabled == true) toolStrip1.Enabled = false;
                    }
                }

                if (RunningFlag == false)
                {
                    if (IOPort.GetAuto == true)
                    {
                        //자동 일때
                        bool Flag = false;
                        if (Config.AutoConnection == true)
                        {
                            //제품이 있고 + 지그 업 + POP 데이타 수신 + 배출될 제품이 없고
                            if ((RetestFlag == false) && (IOPort.GetJigUp == true) && (PopReadOk == true) && (ProductOutputFlag == false) && (IOPort.GetProductIn == true) && (label16.Text == "")) Flag = true;
                            if ((IOPort.GetStartSw == true) && (IOPort.GetJigUp == true) && (PopReadOk == true) && (ProductOutputFlag == false) && (IOPort.GetProductIn == true) && (label16.Text == "")) Flag = true;

                            if (Flag == true)
                            {
                                if (isCheckItem == false)
                                {
                                    Flag = false;
                                    IOPort.TestINGOnOff = false;
                                    IOPort.TestOKOnOff = true;
                                    PopReadOk = false;
                                }
                            }
                            else
                            {
                                //if ((RetestFlag == false) && (IOPort.GetJigUp == true) && (PopReadOk == false) && (ProductOutputFlag == false) && (IOPort.GetProductIn == true) && (label16.Text == ""))
                                //{
                                //    PopReadOk = PopReadOkOld;
                                //}
                                //else 
                                if ((IOPort.GetStartSw == true) && (IOPort.GetJigUp == true) && (PopReadOk == false) && (ProductOutputFlag == false) && (IOPort.GetProductIn == true) && (label16.Text == ""))
                                {
                                    PopReadOk = PopReadOkOld;
                                }
                            }
                        }
                        else
                        {
                            //Pass Sw On + 지그 업 + POP Data 수신 + 배출될 제품 없고 + 제품이 있고
                            if ((IOPort.GetStartSw == true) && (IOPort.GetJigUp == true) && (PopReadOk == true) && (ProductOutputFlag == false) && (IOPort.GetProductIn == true) && (label16.Text == "")) Flag = true;

                            if (isCheckItem == false)
                            {
                                Flag = false;
                                IOPort.TestINGOnOff = false;
                                IOPort.TestOKOnOff = true;
                            }
                        }

                        if (StartSwOnFlag == false)
                        {
                            if (Flag == true)
                            {
                                Infor.TotalCount++;
                                DisplayCount();
                                StartSetting();
                            }
                            else if ((IOPort.GetStartSw == true) && (ProductOutputFlag == true) && (label16.Text == "OK"))
                            {
                                if (IOPort.TestOKOnOff == false) IOPort.TestOKOnOff = true;
                                if (IOPort.SetDoorOpen == false) IOPort.SetDoorOpen = true;
                                if (panel7.Visible == false) panel7.Visible = false;
                                if (PowerOnFlag == true) PowerOnOff = false;
                            }
                            else if ((IOPort.GetStartSw == true) && (ProductOutputFlag == true) && (label16.Text != ""))
                            {
                                if (IOPort.TestOKOnOff == true) IOPort.TestOKOnOff = false;
                                if (IOPort.SetDoorOpen == false) IOPort.SetDoorOpen = true;
                                if (panel7.Visible == false) panel7.Visible = false;
                                if (PowerOnFlag == true) PowerOnOff = false;
                            }
                        }
                        if (IOPort.GetStartSw == true)
                            StartSwOnFlag = true;
                        else StartSwOnFlag = false;
                    }
                    else
                    {
                        //수동 일때
                        bool Flag = false;
                        //if (Config.AutoConnection == true)
                        //{
                        //    //제품이 있고 + 지그 업 + 배출될 제품이 없고
                        //    if ((IOPort.GetStartSw == true) && (RetestFlag == false) && (JigUpFlag == true) && (ProductOutputFlag == false) && (IOPort.GetProductIn == true) && (label16.Text == "")) Flag = true;
                        //    if ((IOPort.GetStartSw == true) && (JigUpFlag == true) && (ProductOutputFlag == false) && (IOPort.GetProductIn == true) && (label16.Text == "")) Flag = true;
                        //}
                        //else
                        //{
                        //Pass Sw On + 지그 업 + 배출될 제품 없고 + 제품이 있고
                        if ((IOPort.GetStartSw == true) && (IOPort.GetJigUp == true) && (ProductOutputFlag == false) && (IOPort.GetProductIn == true) && (label16.Text == "")) Flag = true;

                        //}
                        if (StartSwOnFlag == false)
                        {
                            if (Flag == true)
                            {
                                Infor.TotalCount++;
                                DisplayCount();
                                StartSetting();
                            }
                            else if ((IOPort.GetStartSw == true) && (ProductOutputFlag == true) && /*(label16.Text != "")*/(label16.Text == "OK")) //수동 모드에서 양품일때만 패스 버튼으로 배출 할 수 있도록
                            {
                                if (IOPort.TestOKOnOff == false) IOPort.TestOKOnOff = true;
                                if (IOPort.SetDoorOpen == false) IOPort.SetDoorOpen = true;
                                if (PowerOnFlag == true) PowerOnOff = false;
                            }
                            else if ((IOPort.GetStartSw == true) && (ProductOutputFlag == true) && /*(label16.Text != "")*/(label16.Text != "")) //수동 모드에서 양품일때만 패스 버튼으로 배출 할 수 있도록
                            {
                                if (IOPort.TestOKOnOff == true) IOPort.TestOKOnOff = false;
                                if (IOPort.SetDoorOpen == false) IOPort.SetDoorOpen = true;
                                if (PowerOnFlag == true) PowerOnOff = false;
                            }
                        }
                        if (IOPort.GetStartSw == true)
                            StartSwOnFlag = true;
                        else StartSwOnFlag = false;

                        if (RunningFlag == false) IOInCheck();
                    }

                    if (IOPort.GetResetSw == true)
                    {
                        if (ResetOnButton == false)
                        {
                            if (IOPort.SetDoorOpen == true) IOPort.SetDoorOpen = false;
                            if (IOPort.TestOKOnOff == true) IOPort.TestOKOnOff = false;
                            RetestFlag = true;
                            ResetOnButton = true;
                            //PopReadOk = PopReadOkOld;
                            label16.Text = "";
                            ProductOutputFlag = false;
                            if (panel7.Visible == true) panel7.Visible = false;
                            IOPort.FunctionIOInit();
                            IOPort.IOInit();
                        }
                    }
                    else
                    {
                        if (ResetOnButton == true) ResetOnButton = false;
                    }
                }
                else
                {
                    if (IOPort.GetResetSw == true)
                    {
                        if (ResetOnButton == false)
                        {
                            ResetOnButton = true;
                            StopSetting();
                            RetestFlag = true;
                            ResetOnButton = true;
                            label16.Text = "";
                            ProductOutputFlag = false;
                            if (panel7.Visible == true) panel7.Visible = false;
                            IOPort.FunctionIOInit();
                            IOPort.IOInit();
                        }
                    }
                    else
                    {
                        if (ResetOnButton == true) ResetOnButton = false;
                    }
                }

                if (FormLoadFirstSpecDisplay == true)
                {
                    if (1000 <= (ComF.timeGetTimems() - FormLoadFirstSpecDisplayTime))
                    {
                        FormLoadFirstSpecDisplay = false;
                        DisplaySpec();
                    }
                }
                if (sevenSegmentAnalog2.Value.AsDouble != pMeter.GetBatt) sevenSegmentAnalog2.Value.AsDouble = pMeter.GetBatt;
                if (sevenSegmentAnalog5.Value.AsDouble != pMeter.GetPSeat) sevenSegmentAnalog5.Value.AsDouble = pMeter.GetPSeat;
                //if (sevenSegmentAnalog1.Value.AsDouble != (Sound.GetSound + mSpec.Sound.Offset)) sevenSegmentAnalog1.Value.AsDouble = (Sound.GetSound + mSpec.Sound.Offset);
                if (sevenSegmentAnalog1.Value.AsDouble != Sound.GetSound) sevenSegmentAnalog1.Value.AsDouble = Sound.GetSound;
                DisplayAD();
            }
            catch
            {

            }
            finally
            {
                timer2.Enabled = !ExitFlag;
            }
        }

        private bool ResetOnButton = false;

        private void DisplayCount()
        {
            return;
        }

        //private string NAutoBarcode = "";
        private void StartSetting()
        {
            if (panel7.Visible == true) panel7.Visible = false;
            if (SlideDeliveryPosEndFlag == true) SlideDeliveryPosEndFlag = false;
            label5.ForeColor = Color.White;
            
            if ((Config.AutoConnection == true) && (IOPort.GetAuto == true))
            {
                StepTimeFirst = ComF.timeGetTimems();
                StepTimeLast = ComF.timeGetTimems();
                IOPort.SetPinConnection = true;
                do
                {
                    StepTimeLast = ComF.timeGetTimems();
                    if ((long)(Config.PinConnectionDelay * 1000) <= (StepTimeLast - StepTimeFirst)) break;
                    Application.DoEvents();
                } while (IOPort.GetPinConnectionFwd == false);
            }

            if (mName != comboBox1.SelectedItem.ToString())
            {
                string s = comboBox1.SelectedItem.ToString();

                string sName = Program.SPEC_PATH.ToString() + "\\" + s + ".Spc";

                if (File.Exists(sName) == true)
                {
                    if (SpecName != sName) ComF.OpenSpec(sName, ref mSpec);
                    SpecName = sName;
                    DisplaySpec();
                    mName = comboBox1.SelectedItem.ToString();
                }
            }
                if (IOPort.TestOKOnOff == true) IOPort.TestOKOnOff = false;

            RunningFlag = true;
            ScreenInit();
            PopReadOkOld = PopReadOk;
            PopReadOk = false;
            ProductOutputFlag = true;
            Step = 0;
            SpecOutputFlag = false;
            PowerOnOff = true;
            label16.Text = "검사중";
            label16.ForeColor = Color.Yellow;
            IOPort.YellowLampOnOff = true;
            IOPort.GreenLampOnOff = false;
            IOPort.RedLampOnOff = false;
            //IOPort.TestINGOnOff = true;

            //plot1.Channels[0].Clear();
            plot2.Channels[0].Clear();
            plot2.Channels[1].Clear();
            IOPort.FunctionIOInit();

            TotalTimeToFirst = ComF.timeGetTimems();
            TotalTimeToLast = ComF.timeGetTimems();

            ThreadSetting();
            
            return;
        }

        private void ScreenInit()
        {
            for (int i = 2; i < fpSpread1.ActiveSheet.RowCount; i++)
            {
                fpSpread1.ActiveSheet.Cells[i, 6].Text = "";
                fpSpread1.ActiveSheet.Cells[i, 7].Text = "";
                fpSpread1.ActiveSheet.Cells[i, 7].ForeColor = Color.White;
            }
            label16.Text = "";


            mData.Height.Data = 0;
            mData.Height.Result = RESULT.CLEAR;
            mData.Height.Test = false;

            mData.HeightDnTime.Test = false;
            mData.HeightDnTime.Data = 0;
            mData.HeightDnTime.Result = RESULT.CLEAR;

            mData.Ims.Test = false;
            mData.Ims.Data = 0;
            mData.Ims.Result = RESULT.CLEAR;

            mData.LumberFwdBwd.Test = false;
            mData.LumberFwdBwd.Data = 0;
            mData.LumberFwdBwd.Result = RESULT.CLEAR;

            mData.LumberUpDn.Data = 0;
            mData.LumberUpDn.Test = false;
            mData.LumberUpDn.Result = RESULT.CLEAR;

            mData.Recline.Data = 0;
            mData.Recline.Test = false;
            mData.Recline.Result = RESULT.CLEAR;

            mData.ReclineBwdTime.Data = 0;
            mData.ReclineBwdTime.Test = false;
            mData.ReclineBwdTime.Result = RESULT.CLEAR;
            mData.ReclineFwdTime.Data = 0;
            mData.ReclineFwdTime.Test = false;
            mData.ReclineFwdTime.Result = RESULT.CLEAR;

            mData.Result = RESULT.CLEAR;

            mData.Slide.Test = false;
            mData.Slide.Result = RESULT.CLEAR;
            mData.Slide.Data = 0;

            mData.SlideBwdTime.Test = false;
            mData.SlideBwdTime.Result = RESULT.CLEAR;
            mData.SlideBwdTime.Data = 0;
            mData.SlideFwdTime.Test = false;
            mData.SlideFwdTime.Result = RESULT.CLEAR;
            mData.SlideFwdTime.Data = 0;

            mData.SoundHeight.RunData = 0;
            mData.SoundHeight.RunResult = RESULT.CLEAR;
            mData.SoundHeight.StartData = 0;
            mData.SoundHeight.StartResult = RESULT.CLEAR;
            mData.SoundHeight.Test = false;

            mData.SoundLumberFwdBwd.RunData = 0;
            mData.SoundLumberFwdBwd.RunResult = RESULT.CLEAR;
            mData.SoundLumberFwdBwd.StartData = 0;
            mData.SoundLumberFwdBwd.StartResult = RESULT.CLEAR;
            mData.SoundLumberFwdBwd.Test = false;

            mData.SoundLumberUpDn.RunData = 0;
            mData.SoundLumberUpDn.RunResult = RESULT.CLEAR;
            mData.SoundLumberUpDn.StartData = 0;
            mData.SoundLumberUpDn.StartResult = RESULT.CLEAR;
            mData.SoundLumberUpDn.Test = false;

            mData.SoundRecline.RunData = 0;
            mData.SoundRecline.RunResult = RESULT.CLEAR;
            mData.SoundRecline.StartData = 0;
            mData.SoundRecline.StartResult = RESULT.CLEAR;
            mData.SoundRecline.Test = false;

            mData.SoundSlide.RunData = 0;
            mData.SoundSlide.RunResult = RESULT.CLEAR;
            mData.SoundSlide.StartData = 0;
            mData.SoundSlide.StartResult = RESULT.CLEAR;
            mData.SoundSlide.Test = false;

            mData.SoundTilt.RunData = 0;
            mData.SoundTilt.RunResult = RESULT.CLEAR;
            mData.SoundTilt.StartData = 0;
            mData.SoundTilt.StartResult = RESULT.CLEAR;
            mData.SoundTilt.Test = false;

            mData.Tilt.Test = false;
            mData.Tilt.Result = RESULT.CLEAR;
            mData.Tilt.Data = 0;

            mData.TiltDnTime.Data = 0;
            mData.TiltDnTime.Result = RESULT.CLEAR;
            mData.TiltDnTime.Test = false;
            mData.TiltUpTime.Data = 0;
            mData.TiltUpTime.Result = RESULT.CLEAR;
            mData.TiltUpTime.Test = false;

            mData.Time = "";
            return;
        }

        private void StopSetting()
        {
            RunningFlag = false;
            PowerOnOff = false;
            BattOnOff = false;
            label16.Text = "NG";
            label16.ForeColor = Color.Red;
            IOPort.YellowLampOnOff = false;
            IOPort.GreenLampOnOff = false;
            IOPort.RedLampOnOff = true;
            //IOPort.TestINGOnOff = true; //검사중 리셋이 눌리면 다른 신호는 다 끄더라도 검사중 신호는 끄면 안됨.
            if (ProductOutputFlag == true) ProductOutputFlag = false;
            Infor.TotalCount--;
            DisplayCount();
            if (IOPort.GetAuto == true)
            {
                if (Config.AutoConnection == true) IOPort.SetPinConnection = false;
            }
            IOPort.FunctionIOInit();
            IOPort.IOInit();
            return;
        }

        private bool PowerOnFlag = false;
        private bool PowerOnOff
        {
            set
            {
                if (value == true)
                {
                    PwCtrl.POWER_PWSetting(mSpec.TestVolt);
                    PwCtrl.POWER_CURRENTSetting(20);
                    PwCtrl.POWER_PWON();
                }
                else
                {
                    PwCtrl.POWER_PWSetting(0);
                    PwCtrl.POWER_PWOFF();
                }

                PowerOnFlag = value;
            }
        }

        private bool BattOnOff
        {
            set
            {
                //if (value == true)
                //{
                //    IOPort.PinSelectToGndOnOff((short)mSpec.PinMap.Power.Batt1.Gnd, value);
                //    IOPort.PinSelectToGndOnOff((short)mSpec.PinMap.Power.Batt2.Gnd, value);
                //    ComF.timedelay(100);
                //    IOPort.PinSelectToBattOnOff((short)mSpec.PinMap.Power.Batt1.Batt, value);
                //    IOPort.PinSelectToBattOnOff((short)mSpec.PinMap.Power.Batt2.Batt, value);
                //}
                //else
                //{
                //    IOPort.PinSelectToBattOnOff((short)mSpec.PinMap.Power.Batt1.Batt, value);
                //    IOPort.PinSelectToBattOnOff((short)mSpec.PinMap.Power.Batt2.Batt, value);
                //    ComF.timedelay(100);
                //    IOPort.PinSelectToGndOnOff((short)mSpec.PinMap.Power.Batt1.Gnd, value);
                //    IOPort.PinSelectToGndOnOff((short)mSpec.PinMap.Power.Batt2.Gnd, value);
                //}
            }
        }

        private bool PopReadOk { get; set; }
        private bool PopReadOkOld { get; set; }

        public PinMapStruct GetPinMap
        {
            set { PinMap = value; }
            get { return PinMap; }
        }

        private void IMSSetButton(short Pos)
        {
            if (Pos == 0)
                CanReWrite.CanDataOutput(OUT_CAN_LIST.C_AVNIMSButtonCmd, (byte)C_AVNIMSButtonCmd.Data.Memory_P1_CMD);
            else CanReWrite.CanDataOutput(OUT_CAN_LIST.C_AVNIMSButtonCmd, (byte)C_AVNIMSButtonCmd.Data.Memory_P2_CMD);
            ComF.timedelay(300);
            CanReWrite.CanDataOutput(OUT_CAN_LIST.C_AVNIMSButtonCmd, (byte)C_AVNIMSButtonCmd.Data.Default);

            return;
        }

        private void IMSM1Button()
        {
            CanReWrite.CanDataOutput(OUT_CAN_LIST.C_AVNIMSButtonCmd, (byte)C_AVNIMSButtonCmd.Data.PBack_P1_CMD);
            ComF.timedelay(300);
            CanReWrite.CanDataOutput(OUT_CAN_LIST.C_AVNIMSButtonCmd, (byte)C_AVNIMSButtonCmd.Data.Default);
            return;
        }
        
        private void IMSM1Button(bool Flag)
        {
            if(Flag == true)
                CanReWrite.CanDataOutput(OUT_CAN_LIST.C_AVNIMSButtonCmd, (byte)C_AVNIMSButtonCmd.Data.PBack_P1_CMD);
            else CanReWrite.CanDataOutput(OUT_CAN_LIST.C_AVNIMSButtonCmd, (byte)C_AVNIMSButtonCmd.Data.Default);
            return;
        }

        private void IMSM2Button()
        {
            CanReWrite.CanDataOutput(OUT_CAN_LIST.C_AVNIMSButtonCmd, (byte)C_AVNIMSButtonCmd.Data.PBack_P2_CMD);
            ComF.timedelay(300);
            CanReWrite.CanDataOutput(OUT_CAN_LIST.C_AVNIMSButtonCmd, (byte)C_AVNIMSButtonCmd.Data.Default);

            return;
        }

        private void IMS1Set()
        {
            if (comboBox4.SelectedItem.ToString() == "IMS")
            {
                IMSSetButton(0);
            }
            return;
        }

        private void IMS2Set()
        {
            if (comboBox4.SelectedItem.ToString() == "IMS")
            {
                IMSSetButton(1);
            }
            return;
        }


        private bool SlideFwdOnOff
        {
            set
            {
                bool Can = false;
                if (((comboBox2.SelectedItem.ToString() == "LHD") && (comboBox3.SelectedItem.ToString() == "LH")) || ((comboBox2.SelectedItem.ToString() == "RHD") && (comboBox3.SelectedItem.ToString() == "RH"))) Can = true;
                if (comboBox4.SelectedItem.ToString() != "IMS") Can = false;

                if (Config.IMSSeatToCanControl == false) Can = false;

                if (Can == true)
                {
                    //제어가 통신 타입이라면 
                    CanReWrite.CanDataOutput(OUT_CAN_LIST.PSeat_Slide_Fwd, value == true ? (byte)0x01 : (byte)0x00);
                }
                else
                {
                    //제어가 통신 타입이 아니라면
                    if (mSpec.PinMap.PSeat_직구동 == true)
                    {
                        if (mSpec.PinMap.WalkIn == false)
                        {
                            if (value == true)
                            {
                                IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.SlideBWD.PinNo);
                                ComF.timedelay(50);
                                IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.SlideFWD.PinNo);
                            }
                            else
                            {
                                IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.SlideBWD.PinNo);
                                IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.SlideFWD.PinNo);
                            }
                        }
                        else
                        {
                            if (mSpec.PinMap.SlideFWD.Mode == PSeatRuNMode.Gnd)
                                IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.SlideFWD.PinNo);
                            else IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.SlideFWD.PinNo);
                        }
                    }
                    else
                    {
                        if (mSpec.PinMap.SlideFWD.Mode == PSeatRuNMode.Gnd)
                            IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.SlideFWD.PinNo);
                        else IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.SlideFWD.PinNo);
                    }
                }
            }
        }


        private bool SlideBwdOnOff
        {
            set
            {
                bool Can = false;
                if (((comboBox2.SelectedItem.ToString() == "LHD") && (comboBox3.SelectedItem.ToString() == "LH")) || ((comboBox2.SelectedItem.ToString() == "RHD") && (comboBox3.SelectedItem.ToString() == "RH"))) Can = true;
                if (comboBox4.SelectedItem.ToString() != "IMS") Can = false;
                if (Config.IMSSeatToCanControl == false) Can = false;
                if (Can == true)
                {
                    //제어가 통신 타입이라면 
                    CanReWrite.CanDataOutput(OUT_CAN_LIST.PSeat_Slide_Bwd, value == true ? (byte)0x01 : (byte)0x00);
                }
                else
                {
                    //제어가 통신 타입이 아니라면
                    if (mSpec.PinMap.PSeat_직구동 == true)
                    {
                        if (mSpec.PinMap.WalkIn == false)
                        {
                            if (value == true)
                            {
                                IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.SlideFWD.PinNo);
                                ComF.timedelay(50);
                                IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.SlideBWD.PinNo);
                            }
                            else
                            {
                                IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.SlideFWD.PinNo);
                                IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.SlideBWD.PinNo);
                            }
                        }
                        else
                        {
                            if (mSpec.PinMap.SlideBWD.Mode == PSeatRuNMode.Gnd)
                                IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.SlideBWD.PinNo);
                            else IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.SlideBWD.PinNo);
                        }
                    }
                    else
                    {
                        if (mSpec.PinMap.SlideBWD.Mode == PSeatRuNMode.Gnd)
                            IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.SlideBWD.PinNo);
                        else IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.SlideBWD.PinNo);
                    }
                }
            }
        }
        private bool ReclineFwdOnOff
        {
            set
            {
                bool Can = false;
                if (((comboBox2.SelectedItem.ToString() == "LHD") && (comboBox3.SelectedItem.ToString() == "LH")) || ((comboBox2.SelectedItem.ToString() == "RHD") && (comboBox3.SelectedItem.ToString() == "RH"))) Can = true;
                if (comboBox4.SelectedItem.ToString() != "IMS") Can = false;
                if (Config.IMSSeatToCanControl == false) Can = false;
                if (Can == true)
                {
                    //제어가 통신 타입이라면 
                    CanReWrite.CanDataOutput(OUT_CAN_LIST.PSeat_Recline_Fwd, value == true ? (byte)0x01 : (byte)0x00);
                }
                else
                {
                    //제어가 통신 타입이 아니라면
                    if (mSpec.PinMap.PSeat_직구동 == true)
                    {
                        if (mSpec.PinMap.WalkIn == false)
                        {
                            if (value == true)
                            {
                                IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.ReclineBWD.PinNo);
                                ComF.timedelay(50);
                                IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.ReclineFWD.PinNo);
                            }
                            else
                            {
                                IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.ReclineBWD.PinNo);
                                IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.ReclineFWD.PinNo);
                            }
                        }
                        else
                        {
                            if (mSpec.PinMap.ReclineFWD.Mode == PSeatRuNMode.Gnd)
                                IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.ReclineFWD.PinNo);
                            else IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.ReclineFWD.PinNo);
                        }
                    }
                    else
                    {
                        if (mSpec.PinMap.ReclineFWD.Mode == PSeatRuNMode.Gnd)
                            IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.ReclineFWD.PinNo);
                        else IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.ReclineFWD.PinNo);
                    }
                }
            }
        }
        private bool ReclineBwdOnOff
        {
            set
            {
                bool Can = false;
                if (((comboBox2.SelectedItem.ToString() == "LHD") && (comboBox3.SelectedItem.ToString() == "LH")) || ((comboBox2.SelectedItem.ToString() == "RHD") && (comboBox3.SelectedItem.ToString() == "RH"))) Can = true;
                if (comboBox4.SelectedItem.ToString() != "IMS") Can = false;
                if (Config.IMSSeatToCanControl == false) Can = false;
                if (Can == true)
                {
                    //제어가 통신 타입이라면 
                    CanReWrite.CanDataOutput(OUT_CAN_LIST.PSeat_Recline_Bwd, value == true ? (byte)0x01 : (byte)0x00);
                }
                else
                {
                    //제어가 통신 타입이 아니라면
                    if (mSpec.PinMap.PSeat_직구동 == true)
                    {
                        if (mSpec.PinMap.WalkIn == false)
                        {
                            if (value == true)
                            {
                                IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.ReclineFWD.PinNo);
                                ComF.timedelay(50);
                                IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.ReclineBWD.PinNo);
                            }
                            else
                            {
                                IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.ReclineFWD.PinNo);
                                IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.ReclineBWD.PinNo);
                            }
                        }
                        else
                        {
                            if (mSpec.PinMap.ReclineBWD.Mode == PSeatRuNMode.Gnd)
                                IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.ReclineBWD.PinNo);
                            else IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.ReclineBWD.PinNo);
                        }
                    }
                    else
                    {
                        if (mSpec.PinMap.ReclineBWD.Mode == PSeatRuNMode.Gnd)
                            IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.ReclineBWD.PinNo);
                        else IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.ReclineBWD.PinNo);
                    }
                }
            }
        }

        private bool TiltUpOnOff
        {
            set
            {
                bool Can = false;
                if (((comboBox2.SelectedItem.ToString() == "LHD") && (comboBox3.SelectedItem.ToString() == "LH")) || ((comboBox2.SelectedItem.ToString() == "RHD") && (comboBox3.SelectedItem.ToString() == "RH"))) Can = true;
                if (comboBox4.SelectedItem.ToString() != "IMS") Can = false;
                if (Config.IMSSeatToCanControl == false) Can = false;
                if (Can == true)
                {
                    //제어가 통신 타입이라면 
                    CanReWrite.CanDataOutput(OUT_CAN_LIST.PSeat_Tilt_Up, value == true ? (byte)0x01 : (byte)0x00);
                }
                else
                {
                    //제어가 통신 타입이 아니라면
                    if (mSpec.PinMap.PSeat_직구동 == true)
                    {
                        if (value == true)
                        {
                            IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.TiltDn.PinNo);
                            ComF.timedelay(50);
                            IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.TiltUp.PinNo);
                        }
                        else
                        {
                            IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.TiltDn.PinNo);
                            IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.TiltUp.PinNo);
                        }
                    }
                    else
                    {
                        if (mSpec.PinMap.SlideBWD.Mode == PSeatRuNMode.Gnd)
                            IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.TiltUp.PinNo);
                        else IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.TiltUp.PinNo);
                    }
                }
            }
        }

        private bool TiltDnOnOff
        {
            set
            {
                bool Can = false;
                if (((comboBox2.SelectedItem.ToString() == "LHD") && (comboBox3.SelectedItem.ToString() == "LH")) || ((comboBox2.SelectedItem.ToString() == "RHD") && (comboBox3.SelectedItem.ToString() == "RH"))) Can = true;
                if (comboBox4.SelectedItem.ToString() != "IMS") Can = false;
                if (Config.IMSSeatToCanControl == false) Can = false;
                if (Can == true)
                {
                    //제어가 통신 타입이라면 
                    CanReWrite.CanDataOutput(OUT_CAN_LIST.PSeat_Tilt_Down, value == true ? (byte)0x01 : (byte)0x00);
                }
                else
                {
                    //제어가 통신 타입이 아니라면
                    if (mSpec.PinMap.PSeat_직구동 == true)
                    {
                        if (value == true)
                        {
                            IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.TiltUp.PinNo);
                            ComF.timedelay(50);
                            IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.TiltDn.PinNo);
                        }
                        else
                        {
                            IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.TiltUp.PinNo);
                            IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.TiltDn.PinNo);
                        }
                    }
                    else
                    {
                        if (mSpec.PinMap.SlideBWD.Mode == PSeatRuNMode.Gnd)
                            IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.TiltDn.PinNo);
                        else IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.TiltDn.PinNo);
                    }
                }
            }
        }

        private bool HeightUpOnOff
        {
            set
            {
                bool Can = false;
                if (((comboBox2.SelectedItem.ToString() == "LHD") && (comboBox3.SelectedItem.ToString() == "LH")) || ((comboBox2.SelectedItem.ToString() == "RHD") && (comboBox3.SelectedItem.ToString() == "RH"))) Can = true;
                if (comboBox4.SelectedItem.ToString() != "IMS") Can = false;
                if (Config.IMSSeatToCanControl == false) Can = false;
                if (Can == true)
                {
                    //제어가 통신 타입이라면 
                    CanReWrite.CanDataOutput(OUT_CAN_LIST.PSeat_Height_Up, value == true ? (byte)0x01 : (byte)0x00);
                }
                else
                {
                    //제어가 통신 타입이 아니라면
                    if (mSpec.PinMap.PSeat_직구동 == true)
                    {
                        if (value == true)
                        {
                            IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.HeightDn.PinNo);
                            ComF.timedelay(50);
                            IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.HeightUp.PinNo);
                        }
                        else
                        {
                            IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.HeightDn.PinNo);
                            IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.HeightUp.PinNo);
                        }
                    }
                    else
                    {
                        if (mSpec.PinMap.SlideBWD.Mode == PSeatRuNMode.Gnd)
                            IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.HeightUp.PinNo);
                        else IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.HeightUp.PinNo);
                    }
                }
            }
        }

        private bool HeightDnOnOff
        {
            set
            {
                bool Can = false;
                if (((comboBox2.SelectedItem.ToString() == "LHD") && (comboBox3.SelectedItem.ToString() == "LH")) || ((comboBox2.SelectedItem.ToString() == "RHD") && (comboBox3.SelectedItem.ToString() == "RH"))) Can = true;
                if (comboBox4.SelectedItem.ToString() != "IMS") Can = false;
                if (Config.IMSSeatToCanControl == false) Can = false;
                if (Can == true)
                {
                    //제어가 통신 타입이라면 
                    CanReWrite.CanDataOutput(OUT_CAN_LIST.PSeat_Height_Down, value == true ? (byte)0x01 : (byte)0x00);
                }
                else
                {
                    //제어가 통신 타입이 아니라면
                    if (mSpec.PinMap.PSeat_직구동 == true)
                    {
                        if (value == true)
                        {
                            IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.HeightUp.PinNo);
                            ComF.timedelay(50);
                            IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.HeightDn.PinNo);
                        }
                        else
                        {
                            IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.HeightUp.PinNo);
                            IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.HeightDn.PinNo);
                        }
                    }
                    else
                    {
                        if (mSpec.PinMap.SlideBWD.Mode == PSeatRuNMode.Gnd)
                            IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.HeightDn.PinNo);
                        else IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.HeightDn.PinNo);
                    }
                }
            }
        }

        private bool LumberFwd
        {
            set
            {
                if (value == true)
                {
                    IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.LumberFwdBwd.Gnd);
                    ComF.timedelay(50);
                    IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.LumberFwdBwd.Batt);
                }
                else
                {
                    IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.LumberFwdBwd.Gnd);
                    IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.LumberFwdBwd.Batt);
                }
            }
        }
        private bool LumberBwd
        {
            set
            {
                if (value == true)
                {
                    IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.LumberFwdBwd.Batt);
                    ComF.timedelay(50);
                    IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.LumberFwdBwd.Gnd);
                }
                else
                {
                    IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.LumberFwdBwd.Batt);
                    IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.LumberFwdBwd.Gnd);
                }
            }
        }

        private bool LumberUp
        {
            set
            {
                if (value == true)
                {
                    IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.LumberUpDn.Gnd);
                    ComF.timedelay(50);
                    IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.LumberUpDn.Batt);
                }
                else
                {
                    IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.LumberUpDn.Gnd);
                    IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.LumberUpDn.Batt);
                }
            }
        }

        private bool LumberDn
        {
            set
            {
                if (value == true)
                {
                    IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.LumberUpDn.Batt);
                    ComF.timedelay(50);
                    IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.LumberUpDn.Gnd);
                }
                else
                {
                    IOPort.PinSelectToGndOnOff(OnOff: value, Port: (short)mSpec.PinMap.LumberUpDn.Batt);
                    IOPort.PinSelectToBattOnOff(OnOff: value, Port: (short)mSpec.PinMap.LumberUpDn.Gnd);
                }
            }
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            if (toolStripButton6.Text == "로그인")
            {
                MessageBox.Show(this, "로그인을 먼저 진행하십시오.", "경고", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (RunningFlag == false)
            {
                SystemSet SysSet = new SystemSet(this)
                {
                    Owner = this,
                    MaximizeBox = false,
                    MinimizeBox = false,
                    ControlBox = false,
                    FormBorderStyle = FormBorderStyle.FixedSingle,
                    WindowState = FormWindowState.Normal,
                    StartPosition = FormStartPosition.CenterParent
                };

                SysSet.FormClosing += delegate (object sender1, FormClosingEventArgs e1)
                {
                    e1.Cancel = false;
                    SysSet.Dispose();
                    SysSet = null;
                };

                SysSet.Show();
            }
            return;
        }
        private SpecSetting SpecForm = null;
        private bool ModelBoxChangeFlag = false;
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (toolStripButton6.Text == "로그인")
            {
                MessageBox.Show(this, "로그인을 먼저 진행하십시오.", "경고", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (IOPort.GetAuto == false)
            {
                if (SpecForm == null)
                {
                    OpenFormClose();
                    panel1.SendToBack();
                    panel1.Visible = false;

                    panel2.Visible = true;
                    if (panel2.Parent != this) panel2.Parent = this;
                    panel2.BringToFront();
                    SpecForm = new SpecSetting(this, mSpec.ModelName)
                    {
                        MaximizeBox = false,
                        MinimizeBox = false,
                        ControlBox = false,
                        ShowIcon = false,
                        StartPosition = FormStartPosition.CenterScreen,
                        WindowState = FormWindowState.Maximized,
                        TopMost = false,
                        TopLevel = false,
                        FormBorderStyle = FormBorderStyle.None,
                        Location = new Point(0, 0),
                        Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
                    };

                    SpecForm.FormClosing += delegate (object sender1, FormClosingEventArgs e1)
                    {
                        e1.Cancel = false;
                        SpecForm.Parent = null;
                        SpecForm.Dispose();
                        SpecForm = null;

                        ComF.ReadFileListNotExt(Program.SPEC_PATH.ToString(), "*.Spc", COMMON_FUCTION.FileSortMode.FILENAME_ODERBY);
                        List<string> FList = ComF.GetFileList;

                        if (0 < FList.Count)
                        {
                            ModelBoxChangeFlag = true;
                            comboBox1.Items.Clear();
                            foreach (string s in FList) comboBox1.Items.Add(s);

                            if (0 < comboBox1.Items.Count)
                            {
                                if ((mSpec.ModelName != null) && (mSpec.ModelName != "") && (mSpec.ModelName != string.Empty))
                                {
                                    if (comboBox1.Items.Contains(mSpec.ModelName) == true) comboBox1.SelectedItem = mSpec.ModelName;
                                }
                            }
                            if (mSpec.ModelName != null)
                            {
                                string sName = Program.SPEC_PATH.ToString() + "\\" + mSpec.ModelName + ".Spc";
                                ComF.OpenSpec(sName, ref mSpec);
                                SpecName = sName;
                                DisplaySpec();
                            }
                            ModelBoxChangeFlag = false;
                        }
                    };

                    SpecForm.Parent = panel2;
                    SpecForm.Show();
                }
            }
            return;
        }

        private void DisplaySpec()
        {
            //IMS 
            if (CheckItem.PSeat == PSEAT_TYPE.IMS)
            {
                fpSpread1.ActiveSheet.Cells[2, 1].BackColor = SelectColor;
                fpSpread1.ActiveSheet.Cells[2, 3].ForeColor = Color.Black;
                fpSpread1.ActiveSheet.Cells[2, 4].ForeColor = Color.Black;
            }
            else
            {
                fpSpread1.ActiveSheet.Cells[2, 1].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[2, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[2, 4].ForeColor = Color.Silver;
            }

            //Power Seat 이동거리
            if (CheckItem.PSeat != PSEAT_TYPE.MANUAL)
            {
                fpSpread1.ActiveSheet.Cells[3, 1].BackColor = SelectColor;

                if (CheckItem.Slide == true)
                {
                    fpSpread1.ActiveSheet.Cells[3, 2].BackColor = SelectColor; //Slide F
                    fpSpread1.ActiveSheet.Cells[4, 2].BackColor = SelectColor; //Slide B

                    //Slide Fwd
                    fpSpread1.ActiveSheet.Cells[3, 3].ForeColor = Color.Black;
                    fpSpread1.ActiveSheet.Cells[3, 4].ForeColor = Color.Black;
                    //Slide Bwd
                    fpSpread1.ActiveSheet.Cells[4, 3].ForeColor = Color.Black;
                    fpSpread1.ActiveSheet.Cells[4, 4].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread1.ActiveSheet.Cells[3, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[4, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[3, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[3, 4].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[4, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[4, 4].ForeColor = Color.Silver;
                }

                if (CheckItem.Recline == true)
                {
                    fpSpread1.ActiveSheet.Cells[5, 2].BackColor = SelectColor; //Reckine F
                    fpSpread1.ActiveSheet.Cells[6, 2].BackColor = SelectColor; //Recline B

                    //Recline Fwd
                    fpSpread1.ActiveSheet.Cells[5, 3].ForeColor = Color.Black;
                    fpSpread1.ActiveSheet.Cells[5, 4].ForeColor = Color.Black;
                    //Recline Bwd
                    fpSpread1.ActiveSheet.Cells[6, 3].ForeColor = Color.Black;
                    fpSpread1.ActiveSheet.Cells[6, 4].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread1.ActiveSheet.Cells[5, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[6, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[5, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[5, 4].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[6, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[6, 4].ForeColor = Color.Silver;
                }

                if (CheckItem.Tilt == true)
                {
                    fpSpread1.ActiveSheet.Cells[7, 2].BackColor = SelectColor; //Tilt Up
                    fpSpread1.ActiveSheet.Cells[8, 2].BackColor = SelectColor; //Tilt Dn

                    //Tilt Up
                    fpSpread1.ActiveSheet.Cells[7, 3].ForeColor = Color.Black;
                    fpSpread1.ActiveSheet.Cells[7, 4].ForeColor = Color.Black;
                    //Tilt Down
                    fpSpread1.ActiveSheet.Cells[8, 3].ForeColor = Color.Black;
                    fpSpread1.ActiveSheet.Cells[8, 4].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread1.ActiveSheet.Cells[7, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[8, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[7, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[7, 4].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[8, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[8, 4].ForeColor = Color.Silver;
                }

                if (CheckItem.Height == true)
                {
                    fpSpread1.ActiveSheet.Cells[9, 2].BackColor = SelectColor; //Height Up
                    fpSpread1.ActiveSheet.Cells[10, 2].BackColor = SelectColor; //Height Dn

                    //Height Up
                    fpSpread1.ActiveSheet.Cells[9, 3].ForeColor = Color.Black;
                    fpSpread1.ActiveSheet.Cells[9, 4].ForeColor = Color.Black;
                    //Height Down
                    fpSpread1.ActiveSheet.Cells[10, 3].ForeColor = Color.Black;
                    fpSpread1.ActiveSheet.Cells[10, 4].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread1.ActiveSheet.Cells[9, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[10, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[9, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[9, 4].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[10, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[10, 4].ForeColor = Color.Silver;
                }
            }
            else
            {
                fpSpread1.ActiveSheet.Cells[3, 1].BackColor = NotSelectColor;

                fpSpread1.ActiveSheet.Cells[3, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[4, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[5, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[6, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[7, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[8, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[9, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[10, 2].BackColor = NotSelectColor;

                fpSpread1.ActiveSheet.Cells[3, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[3, 4].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[4, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[4, 4].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[5, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[5, 4].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[6, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[6, 4].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[7, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[7, 4].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[8, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[8, 4].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[9, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[9, 4].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[10, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[10, 4].ForeColor = Color.Silver;
            }

            //Power Seat 전류
            if (CheckItem.PSeat != PSEAT_TYPE.MANUAL)
            {
                fpSpread1.ActiveSheet.Cells[11, 1].BackColor = SelectColor;

                if (CheckItem.Slide == true)
                {
                    fpSpread1.ActiveSheet.Cells[11, 2].BackColor = SelectColor;
                    fpSpread1.ActiveSheet.Cells[11, 3].ForeColor = Color.Black;
                    fpSpread1.ActiveSheet.Cells[11, 4].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread1.ActiveSheet.Cells[11, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[11, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[11, 4].ForeColor = Color.Silver;
                }

                if (CheckItem.Recline == true)
                {
                    fpSpread1.ActiveSheet.Cells[12, 2].BackColor = SelectColor;
                    fpSpread1.ActiveSheet.Cells[12, 3].ForeColor = Color.Black;
                    fpSpread1.ActiveSheet.Cells[12, 4].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread1.ActiveSheet.Cells[12, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[12, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[12, 4].ForeColor = Color.Silver;
                }

                if (CheckItem.Tilt == true)
                {
                    fpSpread1.ActiveSheet.Cells[13, 2].BackColor = SelectColor;
                    fpSpread1.ActiveSheet.Cells[13, 3].ForeColor = Color.Black;
                    fpSpread1.ActiveSheet.Cells[13, 4].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread1.ActiveSheet.Cells[13, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[13, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[13, 4].ForeColor = Color.Silver;
                }

                if (CheckItem.Height == true)
                {
                    fpSpread1.ActiveSheet.Cells[14, 2].BackColor = SelectColor;
                    fpSpread1.ActiveSheet.Cells[14, 3].ForeColor = Color.Black;
                    fpSpread1.ActiveSheet.Cells[14, 4].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread1.ActiveSheet.Cells[14, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[14, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[14, 4].ForeColor = Color.Silver;
                }
            }
            else
            {
                fpSpread1.ActiveSheet.Cells[11, 1].BackColor = NotSelectColor;

                fpSpread1.ActiveSheet.Cells[11, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[12, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[13, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[14, 2].BackColor = NotSelectColor;

                fpSpread1.ActiveSheet.Cells[11, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[11, 4].ForeColor = Color.Silver;

                fpSpread1.ActiveSheet.Cells[12, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[12, 4].ForeColor = Color.Silver;

                fpSpread1.ActiveSheet.Cells[13, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[13, 4].ForeColor = Color.Silver;

                fpSpread1.ActiveSheet.Cells[14, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[14, 4].ForeColor = Color.Silver;
            }
            //Lumber
            if (CheckItem.PSeat12Way == true)
            {
                fpSpread1.ActiveSheet.Cells[15, 1].BackColor = SelectColor;

                //Lumber Fwd/Bwd
                fpSpread1.ActiveSheet.Cells[15, 2].BackColor = SelectColor;
                fpSpread1.ActiveSheet.Cells[15, 3].ForeColor = Color.Black;
                fpSpread1.ActiveSheet.Cells[15, 4].ForeColor = Color.Black;
                //Lumber Up/Dn
                fpSpread1.ActiveSheet.Cells[16, 2].BackColor = SelectColor;
                fpSpread1.ActiveSheet.Cells[16, 3].ForeColor = Color.Black;
                fpSpread1.ActiveSheet.Cells[16, 4].ForeColor = Color.Black;

            }
            else if (CheckItem.PSeat10Way == true)
            {
                fpSpread1.ActiveSheet.Cells[15, 1].BackColor = SelectColor;

                //Lumber Fwd/Bwd
                fpSpread1.ActiveSheet.Cells[15, 2].BackColor = SelectColor;
                fpSpread1.ActiveSheet.Cells[15, 3].ForeColor = Color.Black;
                fpSpread1.ActiveSheet.Cells[15, 4].ForeColor = Color.Black;
                //Lumber Up/Dn
                fpSpread1.ActiveSheet.Cells[16, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[16, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[16, 4].ForeColor = Color.Silver;

            }
            else
            {
                fpSpread1.ActiveSheet.Cells[15, 1].BackColor = NotSelectColor;
                //Lumber Fwd/Bwd
                fpSpread1.ActiveSheet.Cells[15, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[15, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[15, 4].ForeColor = Color.Silver;
                //Lumber Up/Dn
                fpSpread1.ActiveSheet.Cells[16, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[16, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[16, 4].ForeColor = Color.Silver;
            }

            if (CheckItem.Sound == true)
            {
                if ((CheckItem.PSeat != PSEAT_TYPE.MANUAL) || (CheckItem.PSeat == PSEAT_TYPE.IMS))
                {
                    if (CheckItem.Sound == true)
                    {
                        if (CheckItem.Slide == true)
                            fpSpread1.ActiveSheet.Cells[17, 1].BackColor = SelectColor; //Slide
                        else fpSpread1.ActiveSheet.Cells[17, 1].BackColor = NotSelectColor; //Slide

                        if (CheckItem.Recline == true)
                            fpSpread1.ActiveSheet.Cells[19, 1].BackColor = SelectColor; //Recline
                        else fpSpread1.ActiveSheet.Cells[19, 1].BackColor = NotSelectColor; //Recline

                        if (CheckItem.Tilt == true)
                            fpSpread1.ActiveSheet.Cells[21, 1].BackColor = SelectColor; //Tilt
                        else fpSpread1.ActiveSheet.Cells[21, 1].BackColor = NotSelectColor; //Tilt


                        if (CheckItem.Height == true)
                            fpSpread1.ActiveSheet.Cells[23, 1].BackColor = SelectColor; //Height
                        else fpSpread1.ActiveSheet.Cells[23, 1].BackColor = NotSelectColor; //Height

                        if ((CheckItem.PSeat == PSEAT_TYPE.IMS) || (CheckItem.PSeat12Way == true)) //12Way
                        {
                            fpSpread1.ActiveSheet.Cells[25, 1].BackColor = SelectColor; //L/Fwd Bwd
                            fpSpread1.ActiveSheet.Cells[27, 1].BackColor = SelectColor; //L/Up Dn
                        }
                        else if (CheckItem.PSeat10Way == true) //운전석 2Way
                        {
                            fpSpread1.ActiveSheet.Cells[25, 1].BackColor = SelectColor; //L/Fwd Bwd
                            fpSpread1.ActiveSheet.Cells[27, 1].BackColor = NotSelectColor; //L/Up Dn
                        }
                        else //조수석
                        {
                            fpSpread1.ActiveSheet.Cells[25, 1].BackColor = NotSelectColor; //L/Fwd Bwd
                            fpSpread1.ActiveSheet.Cells[27, 1].BackColor = NotSelectColor; //L/Up Dn
                        }
                    }
                    else
                    {
                        fpSpread1.ActiveSheet.Cells[17, 1].BackColor = NotSelectColor;
                        fpSpread1.ActiveSheet.Cells[17, 1].BackColor = NotSelectColor; //Slide
                        fpSpread1.ActiveSheet.Cells[19, 1].BackColor = NotSelectColor; //Recline
                        fpSpread1.ActiveSheet.Cells[21, 1].BackColor = NotSelectColor; //Height
                        fpSpread1.ActiveSheet.Cells[23, 1].BackColor = NotSelectColor; //Tilt
                        fpSpread1.ActiveSheet.Cells[25, 1].BackColor = NotSelectColor; //L/Fwd Bwd
                        fpSpread1.ActiveSheet.Cells[27, 1].BackColor = NotSelectColor; //L/Up Dn
                    }
                }


                if ((CheckItem.PSeat != PSEAT_TYPE.MANUAL) || (CheckItem.PSeat == PSEAT_TYPE.IMS))
                {

                    if ((CheckItem.PSeat == PSEAT_TYPE.IMS) || (CheckItem.PSeat12Way == true))
                    {
                        //--------Slide
                        if (CheckItem.Slide == true)
                        {
                            fpSpread1.ActiveSheet.Cells[17, 2].BackColor = SelectColor;
                            fpSpread1.ActiveSheet.Cells[18, 2].BackColor = SelectColor;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[17, 2].BackColor = NotSelectColor;
                            fpSpread1.ActiveSheet.Cells[18, 2].BackColor = NotSelectColor;
                        }
                        //--------Recline
                        if (CheckItem.Recline == true)
                        {
                            fpSpread1.ActiveSheet.Cells[19, 2].BackColor = SelectColor;
                            fpSpread1.ActiveSheet.Cells[20, 2].BackColor = SelectColor;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[19, 2].BackColor = NotSelectColor;
                            fpSpread1.ActiveSheet.Cells[20, 2].BackColor = NotSelectColor;
                        }
                        //--------Tilt
                        if (CheckItem.Tilt == true)
                        {
                            fpSpread1.ActiveSheet.Cells[21, 2].BackColor = SelectColor;
                            fpSpread1.ActiveSheet.Cells[22, 2].BackColor = SelectColor;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[21, 2].BackColor = NotSelectColor;
                            fpSpread1.ActiveSheet.Cells[22, 2].BackColor = NotSelectColor;
                        }
                        //--------Height
                        if (CheckItem.Height == true)
                        {
                            fpSpread1.ActiveSheet.Cells[23, 2].BackColor = SelectColor;
                            fpSpread1.ActiveSheet.Cells[24, 2].BackColor = SelectColor;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[23, 2].BackColor = NotSelectColor;
                            fpSpread1.ActiveSheet.Cells[24, 2].BackColor = NotSelectColor;
                        }
                        //--------Lumber Fwd/Bwd
                        fpSpread1.ActiveSheet.Cells[25, 2].BackColor = SelectColor;
                        fpSpread1.ActiveSheet.Cells[26, 2].BackColor = SelectColor;
                        //--------Lumber Up/Dn
                        fpSpread1.ActiveSheet.Cells[27, 2].BackColor = SelectColor;
                        fpSpread1.ActiveSheet.Cells[28, 2].BackColor = SelectColor;
                    }
                    else if (CheckItem.PSeat10Way == true) //운전석 2Way
                    {
                        //--------Slide
                        if (CheckItem.Slide == true)
                        {
                            fpSpread1.ActiveSheet.Cells[17, 2].BackColor = SelectColor;
                            fpSpread1.ActiveSheet.Cells[18, 2].BackColor = SelectColor;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[17, 2].BackColor = NotSelectColor;
                            fpSpread1.ActiveSheet.Cells[18, 2].BackColor = NotSelectColor;
                        }
                        //--------Recline
                        if (CheckItem.Recline == true)
                        {
                            fpSpread1.ActiveSheet.Cells[19, 2].BackColor = SelectColor;
                            fpSpread1.ActiveSheet.Cells[20, 2].BackColor = SelectColor;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[19, 2].BackColor = NotSelectColor;
                            fpSpread1.ActiveSheet.Cells[20, 2].BackColor = NotSelectColor;
                        }
                        //--------Tilt
                        if (CheckItem.Tilt == true)
                        {
                            fpSpread1.ActiveSheet.Cells[21, 2].BackColor = SelectColor;
                            fpSpread1.ActiveSheet.Cells[22, 2].BackColor = SelectColor;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[21, 2].BackColor = NotSelectColor;
                            fpSpread1.ActiveSheet.Cells[22, 2].BackColor = NotSelectColor;
                        }
                        //--------Height
                        if (CheckItem.Height == true)
                        {
                            fpSpread1.ActiveSheet.Cells[23, 2].BackColor = SelectColor;
                            fpSpread1.ActiveSheet.Cells[24, 2].BackColor = SelectColor;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[23, 2].BackColor = NotSelectColor;
                            fpSpread1.ActiveSheet.Cells[24, 2].BackColor = NotSelectColor;
                        }
                        //--------Lumber Fwd/Bwd
                        fpSpread1.ActiveSheet.Cells[25, 2].BackColor = SelectColor;
                        fpSpread1.ActiveSheet.Cells[26, 2].BackColor = SelectColor;
                        //--------Lumber Up/Dn
                        fpSpread1.ActiveSheet.Cells[27, 2].BackColor = NotSelectColor;
                        fpSpread1.ActiveSheet.Cells[28, 2].BackColor = NotSelectColor;
                    }
                    else //조수석
                    {
                        //--------Slide
                        if (CheckItem.Slide == true)
                        {
                            fpSpread1.ActiveSheet.Cells[17, 2].BackColor = SelectColor;
                            fpSpread1.ActiveSheet.Cells[18, 2].BackColor = SelectColor;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[17, 2].BackColor = NotSelectColor;
                            fpSpread1.ActiveSheet.Cells[18, 2].BackColor = NotSelectColor;
                        }
                        //--------Recline
                        if (CheckItem.Recline == true)
                        {
                            fpSpread1.ActiveSheet.Cells[19, 2].BackColor = SelectColor;
                            fpSpread1.ActiveSheet.Cells[20, 2].BackColor = SelectColor;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[19, 2].BackColor = NotSelectColor;
                            fpSpread1.ActiveSheet.Cells[20, 2].BackColor = NotSelectColor;
                        }
                        //--------Tilt
                        if (CheckItem.Tilt == true)
                        {
                            fpSpread1.ActiveSheet.Cells[21, 2].BackColor = SelectColor;
                            fpSpread1.ActiveSheet.Cells[22, 2].BackColor = SelectColor;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[21, 2].BackColor = NotSelectColor;
                            fpSpread1.ActiveSheet.Cells[22, 2].BackColor = NotSelectColor;
                        }
                        //--------Height
                        if (CheckItem.Height == true)
                        {
                            fpSpread1.ActiveSheet.Cells[23, 2].BackColor = SelectColor;
                            fpSpread1.ActiveSheet.Cells[24, 2].BackColor = SelectColor;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[23, 2].BackColor = NotSelectColor;
                            fpSpread1.ActiveSheet.Cells[24, 2].BackColor = NotSelectColor;
                        }
                        //--------Lumber Fwd/Bwd
                        fpSpread1.ActiveSheet.Cells[25, 2].BackColor = NotSelectColor;
                        fpSpread1.ActiveSheet.Cells[26, 2].BackColor = NotSelectColor;
                        //--------Lumber Up/Dn
                        fpSpread1.ActiveSheet.Cells[27, 2].BackColor = NotSelectColor;
                        fpSpread1.ActiveSheet.Cells[28, 2].BackColor = NotSelectColor;
                    }

                    if ((CheckItem.PSeat == PSEAT_TYPE.IMS) || (CheckItem.PSeat12Way == true))
                    {
                        if (CheckItem.Slide == true)
                        {
                            fpSpread1.ActiveSheet.Cells[17, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[17, 4].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[18, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[18, 4].ForeColor = Color.Black;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[17, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[17, 4].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[18, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[18, 4].ForeColor = Color.Silver;
                        }

                        if (CheckItem.Recline == true)
                        {
                            fpSpread1.ActiveSheet.Cells[19, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[19, 4].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[20, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[20, 4].ForeColor = Color.Black;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[19, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[19, 4].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[20, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[20, 4].ForeColor = Color.Silver;
                        }

                        if (CheckItem.Tilt == true)
                        {
                            fpSpread1.ActiveSheet.Cells[21, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[21, 4].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[22, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[22, 4].ForeColor = Color.Black;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[21, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[21, 4].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[22, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[22, 4].ForeColor = Color.Silver;
                        }
                        if (CheckItem.Height == true)
                        {
                            fpSpread1.ActiveSheet.Cells[23, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[23, 4].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[24, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[24, 4].ForeColor = Color.Black;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[23, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[23, 4].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[24, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[24, 4].ForeColor = Color.Silver;
                        }
                        fpSpread1.ActiveSheet.Cells[25, 3].ForeColor = Color.Black;
                        fpSpread1.ActiveSheet.Cells[25, 4].ForeColor = Color.Black;
                        fpSpread1.ActiveSheet.Cells[26, 3].ForeColor = Color.Black;
                        fpSpread1.ActiveSheet.Cells[26, 4].ForeColor = Color.Black;

                        fpSpread1.ActiveSheet.Cells[27, 3].ForeColor = Color.Black;
                        fpSpread1.ActiveSheet.Cells[27, 4].ForeColor = Color.Black;
                        fpSpread1.ActiveSheet.Cells[28, 3].ForeColor = Color.Black;
                        fpSpread1.ActiveSheet.Cells[28, 4].ForeColor = Color.Black;
                    }
                    else if (CheckItem.PSeat10Way == true)
                    {
                        if (CheckItem.Slide == true)
                        {
                            fpSpread1.ActiveSheet.Cells[17, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[17, 4].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[18, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[18, 4].ForeColor = Color.Black;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[17, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[17, 4].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[18, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[18, 4].ForeColor = Color.Silver;
                        }

                        if (CheckItem.Recline == true)
                        {
                            fpSpread1.ActiveSheet.Cells[19, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[19, 4].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[20, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[20, 4].ForeColor = Color.Black;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[19, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[19, 4].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[20, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[20, 4].ForeColor = Color.Silver;
                        }

                        if (CheckItem.Tilt == true)
                        {
                            fpSpread1.ActiveSheet.Cells[21, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[21, 4].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[22, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[22, 4].ForeColor = Color.Black;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[21, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[21, 4].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[22, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[22, 4].ForeColor = Color.Silver;
                        }
                        if (CheckItem.Height == true)
                        {
                            fpSpread1.ActiveSheet.Cells[23, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[23, 4].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[24, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[24, 4].ForeColor = Color.Black;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[23, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[23, 4].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[24, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[24, 4].ForeColor = Color.Silver;
                        }

                        fpSpread1.ActiveSheet.Cells[25, 3].ForeColor = Color.Black;
                        fpSpread1.ActiveSheet.Cells[25, 4].ForeColor = Color.Black;
                        fpSpread1.ActiveSheet.Cells[26, 3].ForeColor = Color.Black;
                        fpSpread1.ActiveSheet.Cells[26, 4].ForeColor = Color.Black;

                        fpSpread1.ActiveSheet.Cells[27, 3].ForeColor = Color.Silver;
                        fpSpread1.ActiveSheet.Cells[27, 4].ForeColor = Color.Silver;
                        fpSpread1.ActiveSheet.Cells[28, 3].ForeColor = Color.Silver;
                        fpSpread1.ActiveSheet.Cells[28, 4].ForeColor = Color.Silver;
                    }
                    else
                    {
                        if (CheckItem.Slide == true)
                        {
                            fpSpread1.ActiveSheet.Cells[17, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[17, 4].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[18, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[18, 4].ForeColor = Color.Black;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[17, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[17, 4].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[18, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[18, 4].ForeColor = Color.Silver;
                        }

                        if (CheckItem.Recline == true)
                        {
                            fpSpread1.ActiveSheet.Cells[19, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[19, 4].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[20, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[20, 4].ForeColor = Color.Black;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[19, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[19, 4].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[20, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[20, 4].ForeColor = Color.Silver;
                        }

                        if (CheckItem.Tilt == true)
                        {
                            fpSpread1.ActiveSheet.Cells[21, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[21, 4].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[22, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[22, 4].ForeColor = Color.Black;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[21, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[21, 4].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[22, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[22, 4].ForeColor = Color.Silver;
                        }
                        if (CheckItem.Height == true)
                        {
                            fpSpread1.ActiveSheet.Cells[23, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[23, 4].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[24, 3].ForeColor = Color.Black;
                            fpSpread1.ActiveSheet.Cells[24, 4].ForeColor = Color.Black;
                        }
                        else
                        {
                            fpSpread1.ActiveSheet.Cells[23, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[23, 4].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[24, 3].ForeColor = Color.Silver;
                            fpSpread1.ActiveSheet.Cells[24, 4].ForeColor = Color.Silver;
                        }

                        fpSpread1.ActiveSheet.Cells[25, 3].ForeColor = Color.Silver;
                        fpSpread1.ActiveSheet.Cells[25, 4].ForeColor = Color.Silver;
                        fpSpread1.ActiveSheet.Cells[26, 3].ForeColor = Color.Silver;
                        fpSpread1.ActiveSheet.Cells[26, 4].ForeColor = Color.Silver;

                        fpSpread1.ActiveSheet.Cells[27, 3].ForeColor = Color.Silver;
                        fpSpread1.ActiveSheet.Cells[27, 4].ForeColor = Color.Silver;
                        fpSpread1.ActiveSheet.Cells[28, 3].ForeColor = Color.Silver;
                        fpSpread1.ActiveSheet.Cells[28, 4].ForeColor = Color.Silver;
                    }

                }
                else
                {
                    fpSpread1.ActiveSheet.Cells[17, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[18, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[19, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[20, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[21, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[22, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[23, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[24, 2].BackColor = NotSelectColor;
                    //--------Lumber Fwd/Bwd
                    fpSpread1.ActiveSheet.Cells[25, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[26, 2].BackColor = NotSelectColor;
                    //--------Lumber Up/Dn
                    fpSpread1.ActiveSheet.Cells[27, 2].BackColor = NotSelectColor;
                    fpSpread1.ActiveSheet.Cells[28, 2].BackColor = NotSelectColor;

                    fpSpread1.ActiveSheet.Cells[17, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[17, 4].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[18, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[18, 4].ForeColor = Color.Silver;

                    fpSpread1.ActiveSheet.Cells[19, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[19, 4].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[20, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[20, 4].ForeColor = Color.Silver;

                    fpSpread1.ActiveSheet.Cells[21, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[21, 4].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[22, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[22, 4].ForeColor = Color.Silver;

                    fpSpread1.ActiveSheet.Cells[23, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[23, 4].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[24, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[24, 4].ForeColor = Color.Silver;

                    fpSpread1.ActiveSheet.Cells[25, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[25, 4].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[26, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[26, 4].ForeColor = Color.Silver;

                    fpSpread1.ActiveSheet.Cells[27, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[27, 4].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[28, 3].ForeColor = Color.Silver;
                    fpSpread1.ActiveSheet.Cells[28, 4].ForeColor = Color.Silver;
                }
            }
            else
            {
                fpSpread1.ActiveSheet.Cells[17, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[18, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[19, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[20, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[21, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[22, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[23, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[24, 2].BackColor = NotSelectColor;
                //--------Lumber Fwd/Bwd
                fpSpread1.ActiveSheet.Cells[25, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[26, 2].BackColor = NotSelectColor;
                //--------Lumber Up/Dn
                fpSpread1.ActiveSheet.Cells[27, 2].BackColor = NotSelectColor;
                fpSpread1.ActiveSheet.Cells[28, 2].BackColor = NotSelectColor;

                fpSpread1.ActiveSheet.Cells[17, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[17, 4].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[18, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[18, 4].ForeColor = Color.Silver;

                fpSpread1.ActiveSheet.Cells[19, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[19, 4].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[20, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[20, 4].ForeColor = Color.Silver;

                fpSpread1.ActiveSheet.Cells[21, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[21, 4].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[22, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[22, 4].ForeColor = Color.Silver;

                fpSpread1.ActiveSheet.Cells[23, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[23, 4].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[24, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[24, 4].ForeColor = Color.Silver;

                fpSpread1.ActiveSheet.Cells[25, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[25, 4].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[26, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[26, 4].ForeColor = Color.Silver;

                fpSpread1.ActiveSheet.Cells[27, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[27, 4].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[28, 3].ForeColor = Color.Silver;
                fpSpread1.ActiveSheet.Cells[28, 4].ForeColor = Color.Silver;
            }

            fpSpread1.ActiveSheet.Cells[2, 3].Text = "";
            fpSpread1.ActiveSheet.Cells[2, 4].Text = "";

            fpSpread1.ActiveSheet.Cells[3, 3].Text = mSpec.MovingSpeed.SlideFwd.Min.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[3, 4].Text = mSpec.MovingSpeed.SlideFwd.Max.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[4, 3].Text = mSpec.MovingSpeed.SlideBwd.Min.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[4, 4].Text = mSpec.MovingSpeed.SlideBwd.Max.ToString("0.00");

            fpSpread1.ActiveSheet.Cells[5, 3].Text = mSpec.MovingSpeed.ReclinerFwd.Min.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[5, 4].Text = mSpec.MovingSpeed.ReclinerFwd.Max.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[6, 3].Text = mSpec.MovingSpeed.ReclinerBwd.Min.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[6, 4].Text = mSpec.MovingSpeed.ReclinerBwd.Max.ToString("0.00");

            fpSpread1.ActiveSheet.Cells[7, 3].Text = mSpec.MovingSpeed.TiltUp.Min.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[7, 4].Text = mSpec.MovingSpeed.TiltUp.Max.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[8, 3].Text = mSpec.MovingSpeed.TiltDn.Min.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[8, 4].Text = mSpec.MovingSpeed.TiltDn.Max.ToString("0.00");

            fpSpread1.ActiveSheet.Cells[9, 3].Text = mSpec.MovingSpeed.HeightUp.Min.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[9, 4].Text = mSpec.MovingSpeed.HeightUp.Max.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[10, 3].Text = mSpec.MovingSpeed.HeightDn.Min.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[10, 4].Text = mSpec.MovingSpeed.HeightDn.Max.ToString("0.00");


            fpSpread1.ActiveSheet.Cells[11, 3].Text = mSpec.Current.Slide.Min.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[11, 4].Text = mSpec.Current.Slide.Max.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[12, 3].Text = mSpec.Current.Recliner.Min.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[12, 4].Text = mSpec.Current.Recliner.Max.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[13, 3].Text = mSpec.Current.Tilt.Min.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[13, 4].Text = mSpec.Current.Tilt.Max.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[14, 3].Text = mSpec.Current.Height.Min.ToString("0.00");
            fpSpread1.ActiveSheet.Cells[14, 4].Text = mSpec.Current.Height.Max.ToString("0.00");

            fpSpread1.ActiveSheet.Cells[15, 3].Text = mSpec.Current.LumberFwdBwd.Min.ToString("0.0");
            fpSpread1.ActiveSheet.Cells[15, 4].Text = mSpec.Current.LumberFwdBwd.Max.ToString("0.0");
            fpSpread1.ActiveSheet.Cells[16, 3].Text = mSpec.Current.LumberUpDn.Min.ToString("0.0");
            fpSpread1.ActiveSheet.Cells[16, 4].Text = mSpec.Current.LumberUpDn.Max.ToString("0.0");

            fpSpread1.ActiveSheet.Cells[17, 4].Text = Math.Truncate(mSpec.Sound.StartMax).ToString("0");
            fpSpread1.ActiveSheet.Cells[18, 4].Text = Math.Truncate(mSpec.Sound.RunMax).ToString("0");

            fpSpread1.ActiveSheet.Cells[19, 4].Text = Math.Truncate(mSpec.Sound.StartMax).ToString("0");
            fpSpread1.ActiveSheet.Cells[20, 4].Text = Math.Truncate(mSpec.Sound.RunMax).ToString("0");

            fpSpread1.ActiveSheet.Cells[21, 4].Text = Math.Truncate(mSpec.Sound.StartMax).ToString("0");
            fpSpread1.ActiveSheet.Cells[22, 4].Text = Math.Truncate(mSpec.Sound.RunMax).ToString("0");

            fpSpread1.ActiveSheet.Cells[23, 4].Text = Math.Truncate(mSpec.Sound.StartMax).ToString("0");
            fpSpread1.ActiveSheet.Cells[24, 4].Text = Math.Truncate(mSpec.Sound.RunMax).ToString("0");

            fpSpread1.ActiveSheet.Cells[25, 4].Text = Math.Truncate(mSpec.Sound.StartMax).ToString("0");
            fpSpread1.ActiveSheet.Cells[26, 4].Text = Math.Truncate(mSpec.Sound.RunMax).ToString("0");

            fpSpread1.ActiveSheet.Cells[27, 4].Text = Math.Truncate(mSpec.Sound.StartMax).ToString("0");
            fpSpread1.ActiveSheet.Cells[28, 4].Text = Math.Truncate(mSpec.Sound.RunMax).ToString("0");

            return;
        }

        public IOControl GetIO
        {
            get { return IOPort; }
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            this.Location = new Point(1, 1);
            //DisplaySpec();
            return;
        }


        public bool ImsButtonOnFlag { get; set; }
        public bool M1ButtonOnFlag { get; set; }
        public bool M2ButtonOnFlag { get; set; }
        //public long ImsButtonOnFirstTime { get; set; }
        //public long ImsButtonOnLastTime { get; set; }

        private void IOInCheck()
        {
            bool Flag = false;

            if (IOPort.GetRHDSelect == true)
            {
                if (comboBox2.SelectedItem == null) comboBox2.SelectedItem = "RHD";
                if (comboBox2.SelectedItem.ToString() != "RHD")
                {
                    comboBox2.SelectedItem = "RHD";
                    CheckItem.LhdRhd = LHD_RHD.RHD;
                    Flag = true;
                }
            }
            else
            {
                if (comboBox2.SelectedItem == null) comboBox2.SelectedItem = "LHD";
                if (comboBox2.SelectedItem.ToString() != "LHD")
                {
                    comboBox2.SelectedItem = "LHD";
                    CheckItem.LhdRhd = LHD_RHD.LHD;
                    Flag = true;
                }
            }

            if (IOPort.GetRHSelect == true)
            {
                if (comboBox3.SelectedItem == null) comboBox3.SelectedItem = "RH";
                if (comboBox3.SelectedItem.ToString() != "RH")
                {
                    comboBox3.SelectedItem = "RH";
                    CheckItem.LhRh = LH_RH.RH;
                    Flag = true;
                }
            }
            else
            {
                if (comboBox3.SelectedItem == null) comboBox3.SelectedItem = "LH";
                if (comboBox3.SelectedItem.ToString() != "LH")
                {
                    comboBox3.SelectedItem = "LH";
                    CheckItem.LhRh = LH_RH.LH;
                    Flag = true;
                }
            }

            if (IOPort.GetIMS == true)
            {
                if (comboBox4.SelectedItem == null) comboBox4.SelectedItem = "IMS";

                if (comboBox4.SelectedItem.ToString() != "IMS")
                {
                    CheckItem.PSeat = PSEAT_TYPE.IMS;
                    comboBox4.SelectedItem = "IMS";
                    CheckItem.Can = true;
                    CheckItem.PSeat12Way = true;
                    CheckItem.Slide = true;
                    CheckItem.Tilt = true;
                    CheckItem.Recline = true;
                    CheckItem.Height = true;
                    Flag = true;
                }
            }
            else if (IOPort.GetSeatPower == true)
            {
                if (comboBox4.SelectedItem == null) comboBox4.SelectedItem = "POWER";
                if (comboBox4.SelectedItem.ToString() != "POWER")
                {
                    CheckItem.PSeat = PSEAT_TYPE.POWER;
                    comboBox4.SelectedItem = "POWER";
                    CheckItem.Can = false;
                    CheckItem.PSeat12Way = true;
                    CheckItem.Slide = true;
                    CheckItem.Tilt = true;
                    CheckItem.Recline = true;
                    CheckItem.Height = true;
                    Flag = true;
                }
            }
            else if (IOPort.GetWalkIn == true)
            {
                if (comboBox4.SelectedItem == null) comboBox4.SelectedItem = "WALK IN";
                if (comboBox4.SelectedItem.ToString() != "WALK IN")
                {
                    CheckItem.PSeat = PSEAT_TYPE.POWER;
                    comboBox4.SelectedItem = "WALK IN";
                    CheckItem.Can = false;
                    CheckItem.PSeat12Way = false;
                    CheckItem.PSeat10Way = false;
                    CheckItem.WalkIn = true;
                    CheckItem.Slide = true;
                    CheckItem.Tilt = true;
                    CheckItem.Recline = true;
                    CheckItem.Height = true;
                    Flag = true;
                }
            }

            if (ledArrow1.Value.AsBoolean != IOPort.GetSlideFwd) ledArrow1.Value.AsBoolean = IOPort.GetSlideFwd;
            if (ledArrow2.Value.AsBoolean != IOPort.GetSlideBwd) ledArrow2.Value.AsBoolean = IOPort.GetSlideBwd;

            if (ledArrow4.Value.AsBoolean != IOPort.GetReclinerFwd) ledArrow4.Value.AsBoolean = IOPort.GetReclinerFwd;
            if (ledArrow3.Value.AsBoolean != IOPort.GetReclinerBwd) ledArrow3.Value.AsBoolean = IOPort.GetReclinerBwd;

            if (ledArrow7.Value.AsBoolean != IOPort.GetTiltUp) ledArrow7.Value.AsBoolean = IOPort.GetTiltUp;
            if (ledArrow8.Value.AsBoolean != IOPort.GetTiltDn) ledArrow8.Value.AsBoolean = IOPort.GetTiltDn;

            if (ledArrow5.Value.AsBoolean != IOPort.GetHeightUp) ledArrow5.Value.AsBoolean = IOPort.GetHeightUp;
            if (ledArrow6.Value.AsBoolean != IOPort.GetHeightDn) ledArrow6.Value.AsBoolean = IOPort.GetHeightDn;

            if (ledArrow11.Value.AsBoolean != IOPort.GetLumberFwd) ledArrow11.Value.AsBoolean = IOPort.GetLumberFwd;
            if (ledArrow12.Value.AsBoolean != IOPort.GetLumberBwd) ledArrow12.Value.AsBoolean = IOPort.GetLumberBwd;

            if (ledArrow9.Value.AsBoolean != IOPort.GetLumberUp) ledArrow9.Value.AsBoolean = IOPort.GetLumberUp;
            if (ledArrow10.Value.AsBoolean != IOPort.GetLumberDn) ledArrow10.Value.AsBoolean = IOPort.GetLumberDn;

            if (IOPort.GetIMS == true)
            {
                if (ImsButtonOnFlag == false)
                {
                    if (MemSetButton == false)
                    {
                        MemSetButton = true;
                        ledBulb1.On = true;
                        MemSetButtonToTime = ComF.timeGetTimems();
                    }
                    ImsButtonOnFlag = true;
                }
            }
            else
            {
                ImsButtonOnFlag = false;
            }

            if (IOPort.GetM1Sw == true)
            {
                if (M1ButtonOnFlag == false)
                {
                    ledBulb2.On = true;
                    if (MemSetButton == true)
                        IMSSetButton(0);
                    else IMSM1Button();
                    ledBulb1.On = false;
                    ledBulb2.On = false;
                    MemSetButton = false;
                    M1ButtonOnFlag = true;
                }
            }
            else
            {
                M1ButtonOnFlag = false;
            }

            if (IOPort.GetM2Sw == true)
            {
                if (M2ButtonOnFlag == true)
                {
                    ledBulb3.On = true;
                    if (MemSetButton == true)
                        IMSSetButton(1);
                    else IMSM2Button();
                    ledBulb1.On = false;
                    ledBulb3.On = false;
                    MemSetButton = false;
                    M2ButtonOnFlag = true;
                }
            }
            else
            {
                M2ButtonOnFlag = false;
            }

            //----------------------------------------------------------------------
            //사양에 따라 모델 선택을 진행한다.

            if (IOPort.GetLumber2Way4WaySelect == true)
            {
                if (CheckItem.PSeat4Way == false) Flag = true;
                CheckItem.PSeat4Way = true;

                CheckItem.Slide = true;
                CheckItem.Tilt = true;
                CheckItem.Recline = true;
                CheckItem.Height = true;
            }
            else
            {
                if (CheckItem.PSeat4Way == true) Flag = true;
                CheckItem.PSeat4Way = false;
            }

            if (IOPort.GetWalkIn == true)
            {
                if (CheckItem.WalkIn == false)
                {
                    CheckItem.WalkIn = true;
                    CheckItem.Slide = true;
                    CheckItem.Tilt = true;
                    CheckItem.Recline = true;
                    CheckItem.Height = true;
                }
            }
            else
            {
                if (CheckItem.WalkIn == true) CheckItem.WalkIn = false;
            }
            //if (Flag == true)
            //{
            Flag = OpenSpec();
            if (CheckItem.Sound == false) CheckItem.Sound = true;
            //}

            //----------------------------------------------------------------------
            if (Flag == true) DisplaySpec();
            return;
        }

        private BackgroundWorker backgroundWorker1 = null;
        private void ThreadSetting()
        {
            backgroundWorker1 = new BackgroundWorker();

            //ReportProgress메소드를 호출하기 위해서 반드시 true로 설정, false일 경우 ReportProgress메소드를 호출하면 exception 발생
            backgroundWorker1.WorkerReportsProgress = true;
            //스레드에서 취소 지원 여부
            backgroundWorker1.WorkerSupportsCancellation = true;
            //스레드가 run시에 호출되는 핸들러 등록
            backgroundWorker1.DoWork += new DoWorkEventHandler(BackgroundWorker1_DoWork);
            // ReportProgress메소드 호출시 호출되는 핸들러 등록
            backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(BackgroundWorker1_ProgressChanged);
            // 스레드 완료(종료)시 호출되는 핸들러 동록
            backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BackgroundWorker1_RunWorkerCompleted);


            // 스레드가 Busy(즉, run)가 아니라면
            if (backgroundWorker1.IsBusy != true)
            {
                // 스레드 작동!! 아래 함수 호출 시 위에서 bw.DoWork += new DoWorkEventHandler(bw_DoWork); 에 등록한 핸들러가
                // 호출 됩니다.

                backgroundWorker1.RunWorkerAsync();
            }
            return;
        }

        private void BackgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //바로 위에서 worker.ReportProgress((i * 10));호출 시 
            // bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged); 등록한 핸들러가 호출 된다고
            // 하였는데요.. 이 부분에서는 기존 Thread에서 처럼 Dispatcher를 이용하지 않아도 됩니다. 
            // 즉 아래처럼!!사용이 가능합니다.
            //this.tbProgress.Text = (e.ProgressPercentage.ToString() + "%");

            // 기존의 Thread클래스에서 아래와 같이 UI 엘리먼트를 갱신하려면
            // Dispatcher.BeginInvoke(delegate() 
            // {
            //        this.tbProgress.Text = (e.ProgressPercentage.ToString() + "%");
            // )};
            //처럼 처리해야 할 것입니다. 그러나 바로 UI 엘리먼트를 업데이트 하고 있죠??
        }


        //스레드의 run함수가 종료될 경우 해당 핸들러가 호출됩니다.
        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            //스레드가 종료한 이유(사용자 취소, 완료, 에러)에 맞쳐 처리하면 됩니다.
            if ((e.Cancelled == true))
            {
            }
            else if (!(e.Error == null))
            {

            }
            else
            {

            }
        }


        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            do
            {
                //CancellationPending 속성이 true로 set되었다면(위에서 CancelAsync 메소드 호출 시 true로 set된다고 하였죠?
                if ((worker.CancellationPending == true))
                {
                    //루프를 break한다.(즉 스레드 run 핸들러를 벗어나겠죠)
                    e.Cancel = true;
                    break;
                }
                else
                {
                    // 이곳에는 스레드에서 처리할 연산을 넣으시면 됩니다.

                    this.Invoke(new EventHandler(Processing));

                    Thread.Sleep(1);
                    // 스레드 진행상태 보고 - 이 메소드를 호출 시 위에서 
                    // bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged); 등록한 핸들러가 호출 됩니다.
                    worker.ReportProgress(5);
                }
                if ((ExitFlag == true) || (RunningFlag == false))
                {
                    worker.CancelAsync();
                }
            } while (true);
            //while (ExitFlag == false);
        }

        private void SpreadPositionSelect(int sRow, int sCol)
        {
            if (fpSpread1.Sheets[0].ShowRowSelector == false) fpSpread1.Sheets[0].ShowRowSelector = true; //선택된 셀의 헤더에 삼각형 표시가 나타난다.

            fpSpread1.Sheets[0].SetActiveCell(sRow, sCol);
            //fpSpread1.Sheets[0].ActiveRowIndex = sRow;
            //fpSpread1.Sheets[0].ActiveColumnIndex = sCol;


            //Bottom 이나 Nearest 등과 같은 옵션을 설정하여 앞으로 가운데로 또는 끝으로 표시할 수 있다.

            //방법 1 

            //fpSpread1.ShowRow(0, sRow, FarPoint.Win.Spread.VerticalPosition.Nearest);     //자동스크롤
            //fpSpread1.ShowColumn(0, sCol, FarPoint.Win.Spread.HorizontalPosition.Nearest);

            //방법 2              
            fpSpread1.ShowActiveCell(FarPoint.Win.Spread.VerticalPosition.Bottom, FarPoint.Win.Spread.HorizontalPosition.Right);

            //FarPoint.Win.Spread.CellType.NumberCellType CType = new FarPoint.Win.Spread.CellType.NumberCellType() { WordWrap = true };
            //fpSpread1.Sheets[0].Cells[sRow, 0, sRow, 1].CellType = CType;
            //fpSpread1.Sheets[0].Cells[sRow, 3, sRow, 5].CellType = CType;
            //fpSpread1.Sheets[0].Cells[sRow, 6, sRow, 13].CellType = CType;

            return;
        }


        private bool SpecOutputFlag { get; set; }
        private short Step { get; set; }
        private short SubStep { get; set; }

        
        private long StepTimeFirst = 0;
        private long StepTimeLast = 0;
        private long TotalTimeToFirst = 0;
        private long TotalTimeToLast = 0;
        //private void Processing()
        private void Processing(object sender, EventArgs e)
        {
            TotalTimeToLast = ComF.timeGetTimems();

            int TotalTime = (int)((TotalTimeToLast - TotalTimeToFirst) / 1000);
            if (sevenSegmentInteger4.Value.AsInteger != TotalTime) sevenSegmentInteger4.Value.AsInteger = TotalTime;

            switch (Step)
            {
                case 0:
                    //Jig On
                    if (SpecOutputFlag == false)
                    {
                        mData.Time = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                        SubStep = 0;
                        //if (Config.AutoConnection == true) IOPort.SetPinConnection = true;
                        SpecOutputFlag = true;
                        CurrFilter.InitAll();
                        StepTimeFirst = ComF.timeGetTimems();
                        StepTimeLast = ComF.timeGetTimems();
                        
                        DeliveryPosStep = 0;
                        SpreadPositionSelect(11, 7);
                    }
                    else
                    {
                        //if (Config.AutoConnection == true)
                        //{
                        //    StepTimeLast = ComF.timeGetTimems();
                        //    if ((long)(Config.PinConnectionDelay * 1000) <= (StepTimeLast - StepTimeFirst))
                        //    {
                        //        SpecOutputFlag = false;
                        //        Step++;
                        //    }
                        //    else
                        //    {
                        //        if (IOPort.GetPinConnectionFwd == true)
                        //        {
                        //            SpecOutputFlag = false;
                        //            Step++;
                        //        }
                        //    }
                        //}
                        //else
                        //{
                        SpecOutputFlag = false;
                        Step++;
                        //}
                    }
                    break;
                case 1:
                    //Batt
                    if (SpecOutputFlag == false)
                    {
                        BattOnOff = true;
                        StepTimeFirst = ComF.timeGetTimems();
                        StepTimeLast = ComF.timeGetTimems();
                        SpecOutputFlag = true;
                    }
                    else
                    {
                        StepTimeLast = ComF.timeGetTimems();
                        if (1500 <= (StepTimeLast - StepTimeFirst))
                        {
                            Step++;
                            SpecOutputFlag = false;
                        }
                    }
                    break;
                case 2:
                    //Ims1 Set
                    // 가상리미트에서 메모리 설정을 하는 관계로 진행을 하지 않는다.
                    //IMS1Set();
                    //ComF.timedelay(300);
                    Step++;
                    SpecOutputFlag = false;
                    break;
                case 3:
                    if (CheckItem.Slide == true)
                    {
                        SildeCheck();
                    }
                    else
                    {
                        Step++;
                        SpecOutputFlag = false;
                        ComF.timedelay(500);
                    }
                    break;
                case 4:                    
                    if (SpecOutputFlag == false)
                    {
                        StepTimeFirst = ComF.timeGetTimems();
                        StepTimeLast = ComF.timeGetTimems();
                        SpecOutputFlag = true;
                    }
                    else
                    {
                        StepTimeLast = ComF.timeGetTimems();
                        if (300 <= (StepTimeLast - StepTimeFirst))
                        {
                            Step++;
                            SpecOutputFlag = false;
                        }
                    }
                    break;
                case 5:
                    if (CheckItem.Recline == true)
                    {
                        ReclineCheck();
                    }
                    else
                    {
                        Step++;
                        SpecOutputFlag = false;
                        ComF.timedelay(500);
                    }
                    break;
                case 6:
                    if (SpecOutputFlag == false)
                    {
                        StepTimeFirst = ComF.timeGetTimems();
                        StepTimeLast = ComF.timeGetTimems();
                        SpecOutputFlag = true;
                    }
                    else
                    {
                        StepTimeLast = ComF.timeGetTimems();
                        if (300 <= (StepTimeLast - StepTimeFirst))
                        {
                            Step++;
                            SpecOutputFlag = false;
                        }
                    }
                    break;
                case 7:
                    if (CheckItem.Tilt == true)
                    {
                        TiltCheck();
                    }
                    else
                    {
                        Step++;
                        SpecOutputFlag = false;
                        ComF.timedelay(500);
                    }
                    break;
                case 8:
                    if (SpecOutputFlag == false)
                    {
                        StepTimeFirst = ComF.timeGetTimems();
                        StepTimeLast = ComF.timeGetTimems();
                        SpecOutputFlag = true;
                    }
                    else
                    {
                        StepTimeLast = ComF.timeGetTimems();
                        if (300 <= (StepTimeLast - StepTimeFirst))
                        {
                            Step++;
                            SpecOutputFlag = false;
                        }
                    }
                    break;
                case 9:
                    if (CheckItem.Height == true)
                    {
                        HeightCheck();
                    }
                    else
                    {
                        Step++;
                        SpecOutputFlag = false;
                        ComF.timedelay(500);
                    }
                    break;
                case 10:
                    if (SpecOutputFlag == false)
                    {
                        StepTimeFirst = ComF.timeGetTimems();
                        StepTimeLast = ComF.timeGetTimems();
                        SpecOutputFlag = true;
                    }
                    else
                    {
                        StepTimeLast = ComF.timeGetTimems();
                        if (300 <= (StepTimeLast - StepTimeFirst))
                        {
                            Step++;
                            SpecOutputFlag = false;
                        }
                    }
                    break;
                case 11:
                    if ((CheckItem.PSeat12Way == true) || (CheckItem.PSeat10Way == true))
                    {
                        LumberFwdBwdCheck();
                    }
                    else
                    {
                        Step++;
                        SpecOutputFlag = false;
                    }
                    break;
                case 12:
                    if (SpecOutputFlag == false)
                    {
                        StepTimeFirst = ComF.timeGetTimems();
                        StepTimeLast = ComF.timeGetTimems();
                        SpecOutputFlag = true;
                    }
                    else
                    {
                        StepTimeLast = ComF.timeGetTimems();
                        if (300 <= (StepTimeLast - StepTimeFirst))
                        {
                            Step++;
                            SpecOutputFlag = false;
                        }
                    }
                    break;
                case 13:
                    if (CheckItem.PSeat12Way == true)
                    {
                        LumberUpDnCheck();
                    }
                    else
                    {
                        Step++;
                        SpecOutputFlag = false;
                    }
                    break;
                case 14:
                    //SoundCheckInitPositionMove();
                    Step++;
                    SpecOutputFlag = false;
                    break;
                case 15:
                    //if (comboBox4.SelectedItem.ToString() == "IMS")
                    //{
                    //    CheckSpeedToIMS(SubStep);
                    //    if (7 < SubStep) Step++;
                    //}
                    //else
                    //{
                    //    CheckSpeedNotIMS();
                    //}
                    Step++;
                    SpecOutputFlag = false;
                    break;
                case 16:
                    //SoundCheck();
                    if (comboBox4.SelectedItem.ToString() != "IMS") 
                        Step = 18;
                    else Step++;
                    SpecOutputFlag = false;
                    MemorySetStep = 0;
                    break;
                case 17:
                    if (comboBox4.SelectedItem.ToString() == "IMS")
                    {
                        MemoryPlayCheck();
                    }
                    else
                    {
                        SpreadPositionSelect(28, 7);
                        SpecOutputFlag = false;
                        Step++;
                    }
                    break;
                case 18:
                    ResultCheck();
                    if (mData.Result != RESULT.REJECT)
                    {
                        //양품이고 납입 불량이 아니면
                        //if(panel7.Visible == false) 
                        IOPort.SetDoorOpen = true;
                    }
                    StepTimeFirst = ComF.timeGetTimems();
                    StepTimeLast = ComF.timeGetTimems();
                    SpecOutputFlag = false;
                    Step++;
                    break;                
                case 19:
                    ////딜레이로 사용
                    //if (SpecOutputFlag == false)
                    //{
                    //    label5.Text = "납입 위치 이동 준비작업 중...";
                    //    PwCtrl.POWER_PWOFF();
                    //    CheckTimeToFirst = ComF.timeGetTimems();
                    //    CheckTimeToLast = ComF.timeGetTimems();
                    //    SpecOutputFlag = true;
                    //}
                    //else
                    //{
                    //    CheckTimeToLast = ComF.timeGetTimems();
                    //    if (1000 <= (CheckTimeToLast - CheckTimeToFirst))
                    //    {
                    //        PwCtrl.POWER_PWON();                            

                    //        SpecOutputFlag = false;
                    //        Step++;
                    //    }
                    //}
                    SpecOutputFlag = false;
                    Step++;
                    break;
                case 20:
                    //딜레이로 사용
                    if (SpecOutputFlag == false)
                    {
                        CheckTimeToFirst = ComF.timeGetTimems();
                        CheckTimeToLast = ComF.timeGetTimems();
                        SpecOutputFlag = true;
                    }
                    else
                    {
                        CheckTimeToLast = ComF.timeGetTimems();

                        if (500 <= (CheckTimeToLast - CheckTimeToFirst))
                        {
                            SpecOutputFlag = false;
                            Step++;
                        }
                    }
                    break;                    
                case 21:
                    DeliveryPosMoveing();
                    break;
                default:
                    InComingResult = RESULT.CLEAR;

                    if ((Config.AutoConnection == true) && (IOPort.GetAuto == true))
                    {
                        if (IOPort.SetPinConnection == true) IOPort.SetPinConnection = false;

                        StepTimeLast = ComF.timeGetTimems();
                        if ((int)(Config.PinConnectionDelay * 1000) <= (StepTimeLast - StepTimeFirst))
                        {
                            IMS1Set();
                            if (mData.Result == RESULT.REJECT)
                            {
                                label16.Text = "NG";
                                label16.ForeColor = Color.Red;
                                //label5.Text = "불량 제품 입니다.";
                            }
                            else
                            {
                                label16.Text = "OK";
                                label16.ForeColor = Color.Lime;
                                //label5.Text = "양품 제품 입니다.";
                                IOPort.TestOKOnOff = true;
                            }

                            //if (panel7.Visible == false) IOPort.TestINGOnOff = false;
                            
                            DeliveryResultDisplay();

                            SendTestData();
                            RunningFlag = false;
                            SaveData();
                        }
                        else
                        {
                            if (IOPort.GetPinConnectionBwd == true)
                            {
                                IMS1Set();

                                if (mData.Result == RESULT.REJECT)
                                {
                                    label16.Text = "NG";
                                    label16.ForeColor = Color.Red;
                                    //label5.Text = "불량 제품 입니다.";
                                }
                                else
                                {
                                    label16.Text = "OK";
                                    label16.ForeColor = Color.Lime;
                                    //label5.Text = "양품 제품 입니다.";
                                    IOPort.TestOKOnOff = true;
                                }
                                DeliveryResultDisplay();
                                //if (panel7.Visible == false) IOPort.TestINGOnOff = false;
                                //IOPort.TestINGOnOff = false;
                                //if (mData.Result != RESULT.REJECT)
                                //{
                                //    if (Config.AutoConnection == true) IOPort.TestOKOnOff = true;
                                //}
                                SendTestData();
                                RunningFlag = false;
                                SaveData();
                            }
                        }
                    }
                    else
                    {
                        //수동 , 양품일때 Test Ing 를 끄면 제품이 배출 되므로 IGN를 끄지 않는다.
                        //IOPort.TestINGOnOff = false;
                        //if (mData.Result != RESULT.REJECT)
                        //{
                        //    if (Config.AutoConnection == true) IOPort.TestOKOnOff = true;
                        //}

                        IMS1Set();
                        if (mData.Result == RESULT.REJECT)
                        {
                            label16.Text = "NG";
                            label16.ForeColor = Color.Red;
                            //label5.Text = "불량 제품 입니다.";
                        }
                        else
                        {

                            label16.Text = "OK";
                            label16.ForeColor = Color.Lime;
                            //label5.Text = "양품 제품 입니다.";
                        }
                        DeliveryResultDisplay();
                        SendTestData();
                        RunningFlag = false;
                        SaveData();
                    }

                    break;
            }
            return;
        }

        private short InComingResult = RESULT.CLEAR;

        private void DeliveryResultDisplay()
        {
            bool xFlag = false;
            if (mSpec.DeliveryPos.Recliner == 0)//Fwd
            {
                if (IOPort.GetReclineMidSensor == false)
                {
                    xFlag = true;
                    label5.Text = "RECLINE 납입 위치 불량 입니다.";
                    label5.ForeColor = Color.Red;
                    InComingResult = RESULT.REJECT;
                }
            }
            if (mSpec.DeliveryPos.Slide == 1)//Mid
            {
                if ((IOPort.GetSlideMidSensor == false) && (SlideDeliveryPosEndFlag == false))
                {
                    xFlag = true;
                    label5.Text = "SLIDE 납입 위치 불량 입니다.";
                    label5.ForeColor = Color.Red;
                    InComingResult = RESULT.REJECT;
                }
            }
            if (xFlag == false)
            {
                label5.Text = "납입 위치 정상 입니다.";
                label5.ForeColor = Color.Lime;
                InComingResult = RESULT.PASS;
            }
            return;
        }

        private short MemorySetStep { set; get; }
        private bool ImsPlayButtonOnFlag = false;
        private bool MemoryPlay;
        private int MemoryCount = 0;
        private long CheckTimeToFirst;
        private long CheckTimeToLast;
        private void MemoryPlayCheck()
        {
            if (SpecOutputFlag == false)
            {
                label5.Text = "메모리 검사 중 입니다.";
                MemorySetStep = 0;
                CheckTimeToFirst = ComF.timeGetTimems();
                CheckTimeToLast = ComF.timeGetTimems();
                SpecOutputFlag = true;
            }
            else
            {
                CheckTimeToLast = ComF.timeGetTimems();

                switch (MemorySetStep)
                {
                    case 0:
                        if (300 <= (CheckTimeToLast - CheckTimeToFirst))
                        {
                            IMS1Set();
                            MemorySetStep = 1;
                            CheckTimeToFirst = ComF.timeGetTimems();
                            CheckTimeToLast = ComF.timeGetTimems();
                        }
                        break;
                    case 1:
                        if (500 <= (CheckTimeToLast - CheckTimeToFirst))
                        {
                            IMSM1Button(true);
                            MemorySetStep = 2;
                            MemoryPlay = false;
                            ImsPlayButtonOnFlag = true;
                            MemoryCount = 0;
                            SpreadPositionSelect(2, 7);
                            CheckTimeToFirst = ComF.timeGetTimems();
                            CheckTimeToLast = ComF.timeGetTimems();
                        }
                        break;
                    default:
                        if (MemoryPlay == false)
                        {
                            byte Data = CanReWrite.CheckCanInMessage(IN_CAN_LIST.DrvIMS_PlyBckP1Req);
                            if ((byte)DrvIMS_PlyBckP1Req.Data.P1 == Data) MemoryPlay = true;

                            if (PSEAT_OFF_CURRENT <= pMeter.GetPSeat) MemoryPlay = true;

                            if (ImsPlayButtonOnFlag == true)
                            {
                                if (400 <= (CheckTimeToLast - CheckTimeToFirst))
                                {
                                    ImsPlayButtonOnFlag = false;
                                    IMSM1Button(false);
                                    CheckTimeToFirst = ComF.timeGetTimems();
                                    CheckTimeToLast = ComF.timeGetTimems();
                                }
                            }
                            else
                            {
                                if (400 <= (CheckTimeToLast - CheckTimeToFirst))
                                {
                                    ImsPlayButtonOnFlag = true;
                                    IMSM1Button(true);
                                    CheckTimeToFirst = ComF.timeGetTimems();
                                    CheckTimeToLast = ComF.timeGetTimems();
                                }
                            }

                            //3초
                            if (3000 <= (CheckTimeToLast - CheckTimeToFirst)) //메모리 재생 안됨
                            {
                                mData.Ims.Data = mData.Ims.Data / (float)MemoryCount;
                                //fpSpread1.ActiveSheet.Cells[2, 6].Text = mData.Ims.Data.ToString("0.0");
                                fpSpread1.ActiveSheet.Cells[2, 6].Text = "NG";
                                fpSpread1.ActiveSheet.Cells[2, 7].Text = "NG";
                                fpSpread1.ActiveSheet.Cells[2, 7].ForeColor = Color.Red;
                                mData.Ims.Test = true;
                                mData.Ims.Result = RESULT.REJECT;
                                SpecOutputFlag = false;
                                Step++;
                            }
                        }
                        else
                        {
                            if (1000 <= (CheckTimeToLast - CheckTimeToFirst))
                            {
                                if (mData.Ims.Test == false) mData.Ims.Test = true;
                                mData.Ims.Data += pMeter.GetPSeat;//Kalman.CheckData(pMeter.GetPSeat);
                                MemoryCount++;

                                if (pMeter.GetPSeat < PSEAT_OFF_CURRENT)
                                {
                                    mData.Ims.Data = mData.Ims.Data / (float)MemoryCount;

                                    if ((1 <= mData.Ims.Data) || (MemoryPlay == true))
                                        mData.Ims.Result = RESULT.PASS;
                                    else mData.Ims.Result = RESULT.REJECT;
                                    SpecOutputFlag = false;
                                    Step++;

                                    //fpSpread1.ActiveSheet.Cells[2, 6].Text = mData.Ims.Data.ToString("0.0");
                                    if (mData.Ims.Result == RESULT.PASS)
                                    {
                                        fpSpread1.ActiveSheet.Cells[2, 6].Text = "OK";
                                        fpSpread1.ActiveSheet.Cells[2, 7].Text = "OK";
                                        fpSpread1.ActiveSheet.Cells[2, 7].ForeColor = Color.Lime;
                                    }
                                    else
                                    {
                                        fpSpread1.ActiveSheet.Cells[2, 6].Text = "NG";
                                        fpSpread1.ActiveSheet.Cells[2, 7].Text = "NG";
                                        fpSpread1.ActiveSheet.Cells[2, 7].ForeColor = Color.Red;
                                    }
                                }
                            }
                        }

                        if (CanReWrite.GetCanFail == true) //통신 불량
                        {
                            mData.Ims.Data = mData.Ims.Data / (float)MemoryCount;
                            //fpSpread1.ActiveSheet.Cells[2, 6].Text = mData.Ims.Data.ToString("0.0");
                            fpSpread1.ActiveSheet.Cells[2, 6].Text = "NG";
                            fpSpread1.ActiveSheet.Cells[2, 7].Text = "NG";
                            fpSpread1.ActiveSheet.Cells[2, 7].ForeColor = Color.Red;
                            mData.Ims.Test = true;
                            mData.Ims.Result = RESULT.REJECT;
                            SpecOutputFlag = false;
                            Step++;
                        }

                        //25초
                        if ((25F * 1000F) <= (CheckTimeToLast - CheckTimeToFirst)) //Time Over
                        {
                            mData.Ims.Data = mData.Ims.Data / (float)MemoryCount;
                            //fpSpread1.ActiveSheet.Cells[2, 6].Text = mData.Ims.Data.ToString("0.0");
                            fpSpread1.ActiveSheet.Cells[2, 6].Text = "NG";
                            fpSpread1.ActiveSheet.Cells[2, 7].Text = "NG";
                            fpSpread1.ActiveSheet.Cells[2, 7].ForeColor = Color.Red;
                            mData.Ims.Test = true;
                            mData.Ims.Result = RESULT.REJECT;
                            SpecOutputFlag = false;
                            Step++;
                        }
                        break;
                }
            }
            return;
        }

        //private void LumberDeliveryPos(short Pos)
        //{
        //    if (SpecOutputFlag == false)
        //    {
        //        ComF.timedelay(1000);
        //        if (Pos == 0)
        //            LumberBwd = true;
        //        else LumberDn = true;
        //        SpecOutputFlag = true;
        //        CheckTimeToFirst = ComF.timeGetTimems();
        //        CheckTimeToLast = ComF.timeGetTimems();
        //        label5.Text = "럼버 납입위치 이동 중 입니다.";
        //    }
        //    else
        //    {
        //        CheckTimeToLast = ComF.timeGetTimems();

        //        if (1.5 <= pMeter.GetPSeat)
        //        {
        //            if (Pos == 0)
        //                LumberBwd = false;
        //            else LumberDn = false;
        //            SpecOutputFlag = false;
        //            Step++;
        //        }

        //        //60초
        //        if ((25F * 1000F) <= (CheckTimeToLast - CheckTimeToFirst))
        //        {
        //            if (Pos == 0)
        //                LumberBwd = false;
        //            else LumberDn = false;
        //            SpecOutputFlag = false;
        //            Step++;
        //        }
        //        if (comboBox4.SelectedItem.ToString() == "IMS")
        //        {
        //            if (CanReWrite.GetCanFail == true)
        //            {
        //                if (Pos == 0)
        //                    LumberBwd = false;
        //                else LumberDn = false;
        //                SpecOutputFlag = false;
        //                Step++;
        //            }
        //        }
        //    }
        //    return;
        //}

        //public int[] SoundCheckTimeToStart = new int[12];
        //public int[] SoundCheckTimeToEnd = new int[12];

        private float MinData = 0;
        private float MaxData = 0;
        private double PlotTime = 0;

        private void SildeCheck()
        {
            if (SpecOutputFlag == false)
            {
                label5.Text = "슬라이드 전류 측정 중 입니다.";
                StepTimeFirst = ComF.timeGetTimems();
                StepTimeLast = ComF.timeGetTimems();
                SpecOutputFlag = true;

                //Slide Fwd Sw On
                SlideFwdOnOff = true;

                plot2.Channels[0].Clear();
                SpreadPositionSelect(11, 7);
                mData.Slide.Data = 0;
                MinData = 99999;
                MaxData = 0;
                CurrFilter.InitAll();

                if (0 < plot2.Channels[1].Count)
                    PlotTime = plot2.Channels[1].GetX(plot2.Channels[1].Count - 1);
                else PlotTime = 0;
            }
            else
            {
                float AdData;
                double RunTime = 0;

                StepTimeLast = ComF.timeGetTimems();

                AdData = pMeter.GetPSeat;

                RunTime = StepTimeLast - StepTimeFirst;

                if (mSpec.Sound.RMSMode ==  true)
                    plot2.Channels[0].AddXY(RunTime, CurrFilter.CheckData(0, (Sound.GetSound + mSpec.Sound.Offset)));
                else plot2.Channels[0].AddXY(RunTime, (Sound.GetSound + mSpec.Sound.Offset));

                if (mSpec.Sound.RMSMode == true)
                    plot2.Channels[1].AddXY(PlotTime + RunTime, CurrFilter.CheckData(0, (Sound.GetSound + mSpec.Sound.Offset)));
                else plot2.Channels[1].AddXY(PlotTime + RunTime, (Sound.GetSound + mSpec.Sound.Offset));

                plot2.XAxes[0].Tracking.ZoomToFitAll();
                plot2.YAxes[0].Tracking.ZoomToFitAll();
                               

                if ((mSpec.SlideTestTime * 1000) <= RunTime)
                {
                    SlideFwdOnOff = false;

                    mData.Slide.Test = true;

                    if ((mSpec.Current.Slide.Min <= MaxData) && (MaxData <= mSpec.Current.Slide.Max))
                    {
                        mData.Slide.Data = MaxData;
                        mData.Slide.Result = RESULT.PASS;
                    }
                    else
                    {
                        mData.Slide.Result = RESULT.REJECT;
                        mData.Slide.Data = MaxData;
                    }
                    fpSpread1.ActiveSheet.Cells[11, 6].Text = mData.Slide.Data.ToString("0.00");

                    if (mData.Slide.Result == RESULT.REJECT) mData.Result = RESULT.REJECT;

                    fpSpread1.ActiveSheet.Cells[11, 7].Text = mData.Slide.Result == RESULT.PASS ? "O.K" : "N.G";
                    fpSpread1.ActiveSheet.Cells[11, 7].ForeColor = mData.Slide.Result == RESULT.PASS ? Color.Lime : Color.Red;
                    SoundLavelCheck(SoundLevePos.SLIDE_FWD);
                    SoundCheck(SoundLevePos.SLIDE_FWD);
                    Step++;
                    SpecOutputFlag = false;
                    ComF.timedelay(500);
                }
                else
                {
                    if ((500 <= RunTime) && (RunTime < ((mSpec.SlideTestTime - 0.5) * 1000)))
                    {
                        //mData.Slide.Data = Math.Max(mData.Slide.Data, AdData);
                        if (MinData == 99999)
                        {
                            if (mSpec.Current.Slide.Min < AdData) MinData = AdData;
                        }
                        else
                        {
                            MinData = Math.Min(MinData, AdData);
                        }
                        MaxData = Math.Max(MaxData, AdData);
                        fpSpread1.ActiveSheet.Cells[11, 6].Text = AdData.ToString("0.00");
                    }
                    //fpSpread1.ActiveSheet.Cells[11, 6].Text = mData.Slide.Data.ToString("0.0");

                }
            }
            return;
        }

        private void ReclineCheck()
        {
            if (SpecOutputFlag == false)
            {
                label5.Text = "리클라인 전류 측정 중 입니다.";
                StepTimeFirst = ComF.timeGetTimems();
                StepTimeLast = ComF.timeGetTimems();
                SpecOutputFlag = true;

                //Recline Fwd Sw On
                ReclineBwdOnOff = true;
                mData.Recline.Data = 0;
                plot2.Channels[0].Clear();
                SpreadPositionSelect(12, 7);
                MinData = 99999;
                MaxData = 0;
                CurrFilter.InitAll();
                if (0 < plot2.Channels[1].Count)
                    PlotTime = plot2.Channels[1].GetX(plot2.Channels[1].Count - 1);
                else PlotTime = 0;
            }
            else
            {
                float AdData;
                double RunTime = 0;

                StepTimeLast = ComF.timeGetTimems();

                AdData = pMeter.GetPSeat;

                RunTime = StepTimeLast - StepTimeFirst;

                if (mSpec.Sound.RMSMode == true)
                    plot2.Channels[0].AddXY(RunTime, CurrFilter.CheckData(0, (Sound.GetSound + mSpec.Sound.Offset)));
                else plot2.Channels[0].AddXY(RunTime, (Sound.GetSound + mSpec.Sound.Offset));

                if (mSpec.Sound.RMSMode == true)
                    plot2.Channels[1].AddXY(PlotTime + RunTime, CurrFilter.CheckData(0, (Sound.GetSound + mSpec.Sound.Offset)));
                else plot2.Channels[1].AddXY(PlotTime + RunTime, (Sound.GetSound + mSpec.Sound.Offset));

                plot2.XAxes[0].Tracking.ZoomToFitAll();
                plot2.YAxes[0].Tracking.ZoomToFitAll();

                //plot1.Channels[0].AddXY(plot1.Channels[0].Count, AdData);
                //plot1.YAxes[0].Tracking.ZoomToFitAll();
                //plot1.XAxes[0].Tracking.ZoomToFitAll();


                if ((mSpec.ReclinerTestTime * 1000) <= RunTime)
                {
                    //Recline Fwd Sw Off

                    //mData.Recline.Data = Math.Max(AdData, mData.Recline.Data);
                    //fpSpread1.ActiveSheet.Cells[12, 6].Text = AdData.ToString("0.0");
                    //fpSpread1.ActiveSheet.Cells[12, 6].Text = mData.Recline.Data.ToString("0.0");

                    //float CheckData;

                    //if (float.TryParse(fpSpread1.ActiveSheet.Cells[12, 6].Text, out CheckData) == false) CheckData = 0;

                    ReclineBwdOnOff = false;

                    //mData.Recline.Data = CheckData;
                    mData.Recline.Test = true;
                    //if ((mSpec.Current.Recliner.Min <= CheckData) && (CheckData <= mSpec.Current.Recliner.Max))
                    //    mData.Recline.Result = RESULT.PASS;
                    //else mData.Recline.Result = RESULT.REJECT;

                    //if (MinData < mSpec.Current.Recliner.Min)
                    //{
                    //    mData.Recline.Result = RESULT.REJECT;
                    //    mData.Recline.Data = MinData;
                    //}
                    //else 
                    //if (mSpec.Current.Recliner.Max < MaxData)
                    if ((mSpec.Current.Recliner.Min <= MaxData) && (MaxData <= mSpec.Current.Recliner.Max))
                    {
                        mData.Recline.Data = MaxData;
                        mData.Recline.Result = RESULT.PASS;
                    }
                    else
                    {
                        mData.Recline.Result = RESULT.REJECT;
                        mData.Recline.Data = MaxData;
                    }

                    if (mData.Recline.Result == RESULT.REJECT) mData.Result = RESULT.REJECT;
                    fpSpread1.ActiveSheet.Cells[12, 6].Text = mData.Recline.Data.ToString("0.00");
                    fpSpread1.ActiveSheet.Cells[12, 7].Text = mData.Recline.Result == RESULT.PASS ? "O.K" : "N.G";
                    fpSpread1.ActiveSheet.Cells[12, 7].ForeColor = mData.Recline.Result == RESULT.PASS ? Color.Lime : Color.Red;

                    SoundLavelCheck(SoundLevePos.RECLINE_BWD);
                    SoundCheck(SoundLevePos.RECLINE_BWD);

                    Step++;
                    SpecOutputFlag = false;
                    ComF.timedelay(500);
                }
                else
                {
                    if ((500 <= RunTime) && (RunTime < ((mSpec.ReclinerTestTime - 0.5) * 1000)))
                    {
                        //mData.Recline.Data = Math.Max(AdData, mData.Recline.Data);
                        fpSpread1.ActiveSheet.Cells[12, 6].Text = AdData.ToString("0.00");
                        //fpSpread1.ActiveSheet.Cells[12, 6].Text = mData.Recline.Data.ToString("0.0");
                        if (MinData == 99999)
                        {
                            if (mSpec.Current.Recliner.Min < AdData) MinData = AdData;
                        }
                        else
                        {
                            MinData = Math.Min(MinData, AdData);
                        }
                        MaxData = Math.Max(MaxData, AdData);
                    }
                }
            }
            return;
        }

        private void TiltCheck()
        {
            if (SpecOutputFlag == false)
            {
                label5.Text = "틸트 전류 측정 중 입니다.";
                StepTimeFirst = ComF.timeGetTimems();
                StepTimeLast = ComF.timeGetTimems();
                SpecOutputFlag = true;

                //Tilt Up Sw
#if IMS_USE
                if (comboBox4.SelectedItem.ToString() == "IMS")
                    TiltUpOnOff = true;
                else TiltDnOnOff = true;
#else
                TiltDnOnOff = true;
#endif

                mData.Tilt.Data = 0;

                plot2.Channels[0].Clear();
                SpreadPositionSelect(13, 7);
                MinData = 99999;
                MaxData = 0;
                CurrFilter.InitAll();
                if (0 < plot2.Channels[1].Count)
                    PlotTime = plot2.Channels[1].GetX(plot2.Channels[1].Count - 1);
                else PlotTime = 0;
            }
            else
            {
                float AdData;
                double RunTime = 0;

                StepTimeLast = ComF.timeGetTimems();

                AdData = pMeter.GetPSeat;

                RunTime = StepTimeLast - StepTimeFirst;

                if (mSpec.Sound.RMSMode == true)
                    plot2.Channels[0].AddXY(RunTime, CurrFilter.CheckData(0, (Sound.GetSound + mSpec.Sound.Offset)));
                else plot2.Channels[0].AddXY(RunTime, (Sound.GetSound + mSpec.Sound.Offset));

                if (mSpec.Sound.RMSMode == true)
                    plot2.Channels[1].AddXY(PlotTime + RunTime, CurrFilter.CheckData(0, (Sound.GetSound + mSpec.Sound.Offset)));
                else plot2.Channels[1].AddXY(PlotTime + RunTime, (Sound.GetSound + mSpec.Sound.Offset));

                plot2.XAxes[0].Tracking.ZoomToFitAll();
                plot2.YAxes[0].Tracking.ZoomToFitAll();

                //plot1.Channels[0].AddXY(plot1.Channels[0].Count, AdData);
                //plot1.YAxes[0].Tracking.ZoomToFitAll();
                //plot1.XAxes[0].Tracking.ZoomToFitAll();

                if ((mSpec.TiltTestTime * 1000) <= RunTime)
                {
                    //Slide Fwd Sw Off
#if IMS_USE
                    if (comboBox4.SelectedItem.ToString() == "IMS")
                        TiltUpOnOff = false;
                    else TiltDnOnOff = false;
#else
                    TiltDnOnOff = false;
#endif
                    //mData.Tilt.Data = CheckData;
                    mData.Tilt.Test = true;
                    //if ((mSpec.Current.Tilt.Min <= CheckData) && (CheckData <= mSpec.Current.Tilt.Max))
                    //    mData.Tilt.Result = RESULT.PASS;
                    //else mData.Tilt.Result = RESULT.REJECT;

                    //if (MinData < mSpec.Current.Tilt.Min)
                    //{
                    //    mData.Tilt.Result = RESULT.REJECT;
                    //    mData.Tilt.Data = MinData;
                    //}
                    //else 
                    //if (mSpec.Current.Tilt.Max < MaxData)
                    if ((mSpec.Current.Tilt.Min <= MaxData) && (MaxData <= mSpec.Current.Tilt.Max))
                    {
                        mData.Tilt.Data = MaxData;
                        mData.Tilt.Result = RESULT.PASS;
                    }
                    else
                    {
                        mData.Tilt.Result = RESULT.REJECT;
                        mData.Tilt.Data = MaxData;
                    }
                    if (mData.Tilt.Result == RESULT.REJECT) mData.Result = RESULT.REJECT;

                    fpSpread1.ActiveSheet.Cells[13, 6].Text = mData.Tilt.Data.ToString("0.00");
                    fpSpread1.ActiveSheet.Cells[13, 7].Text = mData.Tilt.Result == RESULT.PASS ? "O.K" : "N.G";
                    fpSpread1.ActiveSheet.Cells[13, 7].ForeColor = mData.Tilt.Result == RESULT.PASS ? Color.Lime : Color.Red;
                    SoundLavelCheck(SoundLevePos.TILT_UP);
                    SoundCheck(SoundLevePos.TILT_UP);
                    Step++;
                    SpecOutputFlag = false;
                    ComF.timedelay(500);
                }
                else
                {
                    if ((500 <= RunTime) && (RunTime < ((mSpec.TiltTestTime - 0.5) * 1000)))
                    {
                        //mData.Tilt.Data = Math.Max(AdData, mData.Tilt.Data);
                        if (MinData == 99999)
                        {
                            if (mSpec.Current.Tilt.Min < AdData) MinData = AdData;
                        }
                        else
                        {
                            MinData = Math.Min(MinData, AdData);
                        }
                        MaxData = Math.Max(MaxData, AdData);
                        fpSpread1.ActiveSheet.Cells[13, 6].Text = AdData.ToString("0.00");
                    }
                    //fpSpread1.ActiveSheet.Cells[13, 6].Text = mData.Tilt.Data.ToString("0.0");
                }
            }
            return;
        }

        private void HeightCheck()
        {
            if (SpecOutputFlag == false)
            {
                label5.Text = "하이트 전류 측정 중 입니다.";
                StepTimeFirst = ComF.timeGetTimems();
                StepTimeLast = ComF.timeGetTimems();
                SpecOutputFlag = true;

                //Slide Fwd Sw On
#if IMS_USE
                if (comboBox4.SelectedItem.ToString() == "IMS")
                    HeightUpOnOff = true;
                else HeightDnOnOff = true;
#else
                HeightDnOnOff = true;
#endif
                mData.Height.Data = 0;
                plot2.Channels[0].Clear();
                SpreadPositionSelect(14, 7);
                MinData = 99999;
                MaxData = 0;
                CurrFilter.InitAll();
                if (0 < plot2.Channels[1].Count)
                    PlotTime = plot2.Channels[1].GetX(plot2.Channels[1].Count - 1);
                else PlotTime = 0;
            }
            else
            {
                float AdData;
                double RunTime = 0;

                StepTimeLast = ComF.timeGetTimems();

                AdData = pMeter.GetPSeat;

                RunTime = StepTimeLast - StepTimeFirst;

                if (mSpec.Sound.RMSMode == true)
                    plot2.Channels[0].AddXY(RunTime, CurrFilter.CheckData(0, (Sound.GetSound + mSpec.Sound.Offset)));
                else plot2.Channels[0].AddXY(RunTime, (Sound.GetSound + mSpec.Sound.Offset));

                if (mSpec.Sound.RMSMode == true)
                    plot2.Channels[1].AddXY(PlotTime + RunTime, CurrFilter.CheckData(0, (Sound.GetSound + mSpec.Sound.Offset)));
                else plot2.Channels[1].AddXY(PlotTime + RunTime, (Sound.GetSound + mSpec.Sound.Offset));

                plot2.XAxes[0].Tracking.ZoomToFitAll();
                plot2.YAxes[0].Tracking.ZoomToFitAll();

                //plot1.Channels[0].AddXY(plot1.Channels[0].Count, AdData);
                //plot1.YAxes[0].Tracking.ZoomToFitAll();
                //plot1.XAxes[0].Tracking.ZoomToFitAll();

                if ((mSpec.HeightTestTime * 1000) <= RunTime)
                {
                    //Height

                    //mData.Height.Data = Math.Max(AdData, mData.Height.Data);
                    //fpSpread1.ActiveSheet.Cells[14, 6].Text = AdData.ToString("0.0");
                    //fpSpread1.ActiveSheet.Cells[14, 6].Text = mData.Height.Data.ToString("0.0");

                    //float CheckData;

                    //if (float.TryParse(fpSpread1.ActiveSheet.Cells[14, 6].Text, out CheckData) == false) CheckData = 0;
#if IMS_USE
                    if (comboBox4.SelectedItem.ToString() == "IMS")
                        HeightUpOnOff = false;
                    else HeightDnOnOff = false;
#else
                    HeightDnOnOff = false;
#endif
                    //mData.Height.Data = CheckData;
                    mData.Height.Test = true;
                    //if ((mSpec.Current.Height.Min <= CheckData) && (CheckData <= mSpec.Current.Height.Max))
                    //    mData.Height.Result = RESULT.PASS;
                    //else mData.Height.Result = RESULT.REJECT;

                    //if (MinData < mSpec.Current.Height.Min)
                    //{
                    //    mData.Height.Result = RESULT.REJECT;
                    //    mData.Height.Data = MinData;
                    //}
                    //else 
                    //if (mSpec.Current.Height.Max < MaxData)
                    if ((mSpec.Current.Height.Min <= MaxData) && (MaxData <= mSpec.Current.Height.Max))
                    {
                        mData.Height.Data = MaxData;
                        mData.Height.Result = RESULT.PASS;
                    }
                    else
                    {
                        mData.Height.Result = RESULT.REJECT;
                        mData.Height.Data = MaxData;
                    }

                    if (mData.Height.Result == RESULT.REJECT) mData.Result = RESULT.REJECT;

                    fpSpread1.ActiveSheet.Cells[14, 6].Text = mData.Height.Data.ToString("0.00");
                    fpSpread1.ActiveSheet.Cells[14, 7].Text = mData.Height.Result == RESULT.PASS ? "O.K" : "N.G";
                    fpSpread1.ActiveSheet.Cells[14, 7].ForeColor = mData.Height.Result == RESULT.PASS ? Color.Lime : Color.Red;

                    SoundLavelCheck(SoundLevePos.HEIGHT_UP);
                    SoundCheck(SoundLevePos.HEIGHT_UP);
                    Step++;
                    SpecOutputFlag = false;
                    ComF.timedelay(500);
                }
                else
                {
                    if ((500 <= RunTime) && (RunTime < ((mSpec.HeightTestTime - 0.5) * 1000)))
                    {
                        //mData.Height.Data = Math.Max(AdData, mData.Height.Data);
                        fpSpread1.ActiveSheet.Cells[14, 6].Text = AdData.ToString("0.00");
                        //fpSpread1.ActiveSheet.Cells[14, 6].Text = mData.Height.Data.ToString("0.0");
                        if (MinData == 99999)
                        {
                            if (mSpec.Current.Height.Min < AdData) MinData = AdData;
                        }
                        else
                        {
                            MinData = Math.Min(MinData, AdData);
                        }
                        MaxData = Math.Max(MaxData, AdData);
                    }
                }
            }
            return;
        }

        private void LumberFwdBwdCheck()
        {
            if (SpecOutputFlag == false)
            {
                label5.Text = "럼버 전진/후진 전류 측정 중 입니다.";
                StepTimeFirst = ComF.timeGetTimems();
                StepTimeLast = ComF.timeGetTimems();
                SpecOutputFlag = true;

                LumberFwd = true;
                MinData = 99999;
                MaxData = 0;
                plot2.Channels[0].Clear();
                SpreadPositionSelect(15, 7);
                CurrFilter.InitAll();
                if (0 < plot2.Channels[1].Count)
                    PlotTime = plot2.Channels[1].GetX(plot2.Channels[1].Count - 1);
                else PlotTime = 0;
            }
            else
            {
                float AdData;
                double RunTime = 0;

                StepTimeLast = ComF.timeGetTimems();

                AdData = pMeter.GetPSeat;

                RunTime = StepTimeLast - StepTimeFirst;

                if (mSpec.Sound.RMSMode == true)
                    plot2.Channels[0].AddXY(RunTime, CurrFilter.CheckData(0, (Sound.GetSound + mSpec.Sound.Offset)));
                else plot2.Channels[0].AddXY(RunTime, (Sound.GetSound + mSpec.Sound.Offset));

                if (mSpec.Sound.RMSMode == true)
                    plot2.Channels[1].AddXY(PlotTime + RunTime, CurrFilter.CheckData(0, (Sound.GetSound + mSpec.Sound.Offset)));
                else plot2.Channels[1].AddXY(PlotTime + RunTime, (Sound.GetSound + mSpec.Sound.Offset));
                plot2.XAxes[0].Tracking.ZoomToFitAll();
                plot2.YAxes[0].Tracking.ZoomToFitAll();

                //plot1.Channels[0].AddXY(plot1.Channels[0].Count, AdData);
                //plot1.YAxes[0].Tracking.ZoomToFitAll();
                //plot1.XAxes[0].Tracking.ZoomToFitAll();

                if ((mSpec.LumberFwdBwdTestTime * 1000) <= RunTime)
                {
                    //Slide Fwd Sw Off
                    //mData.LumberFwdBwd.Data = Math.Max(AdData, mData.LumberFwdBwd.Data);
                    //fpSpread1.ActiveSheet.Cells[15, 6].Text = mData.LumberFwdBwd.Data.ToString("0.0");
                    //float CheckData;

                    //if (float.TryParse(fpSpread1.ActiveSheet.Cells[15, 6].Text, out CheckData) == false) CheckData = 0;

                    LumberFwd = false;

                    //mData.LumberFwdBwd.Data = CheckData;
                    mData.LumberFwdBwd.Test = true;
                    //if ((mSpec.Current.LumberFwdBwd.Min <= CheckData) && (CheckData <= mSpec.Current.LumberFwdBwd.Max))
                    //    mData.LumberFwdBwd.Result = RESULT.PASS;
                    //else mData.LumberFwdBwd.Result = RESULT.REJECT;

                    //if (MinData < mSpec.Current.LumberFwdBwd.Min)
                    //{
                    //    mData.LumberFwdBwd.Result = RESULT.REJECT;
                    //    mData.LumberFwdBwd.Data = MinData;
                    //}
                    //else 
                    //if (mSpec.Current.LumberFwdBwd.Max < MaxData)
                    if ((mSpec.Current.LumberFwdBwd.Min <= MaxData) && (MaxData <= mSpec.Current.LumberFwdBwd.Max))
                    {
                        mData.LumberFwdBwd.Data = MaxData;
                        mData.LumberFwdBwd.Result = RESULT.PASS;
                    }
                    else
                    {
                        mData.LumberFwdBwd.Result = RESULT.REJECT;
                        mData.LumberFwdBwd.Data = MaxData;
                    }

                    if (mData.LumberFwdBwd.Result == RESULT.REJECT) mData.Result = RESULT.REJECT;
                    fpSpread1.ActiveSheet.Cells[15, 6].Text = mData.LumberFwdBwd.Data.ToString("0.0");
                    fpSpread1.ActiveSheet.Cells[15, 7].Text = mData.LumberFwdBwd.Result == RESULT.PASS ? "O.K" : "N.G";
                    fpSpread1.ActiveSheet.Cells[15, 7].ForeColor = mData.LumberFwdBwd.Result == RESULT.PASS ? Color.Lime : Color.Red;

                    SoundLavelCheck(SoundLevePos.LUMBER_FWD);
                    SoundCheck(SoundLevePos.LUMBER_FWD);
                    Step++;
                    SpecOutputFlag = false;
                    ComF.timedelay(500);
                }
                else
                {
                    if ((500 <= RunTime) && (RunTime < ((mSpec.LumberFwdBwdTestTime - 0.5) * 1000)))
                    {
                        if (MinData == 99999)
                        {
                            if (mSpec.Current.LumberFwdBwd.Min < AdData) MinData = AdData;
                        }
                        else
                        {
                            MinData = Math.Min(MinData, AdData);
                        }
                        MaxData = Math.Max(MaxData, AdData);
                        //mData.LumberFwdBwd.Data = Math.Max(AdData, mData.LumberFwdBwd.Data);
                        fpSpread1.ActiveSheet.Cells[15, 6].Text = AdData.ToString("0.0");
                    }
                }
            }
            return;
        }
        private void LumberUpDnCheck()
        {
            if (SpecOutputFlag == false)
            {
                label5.Text = "럼버 상승/하강 전류 측정 중 입니다.";
                StepTimeFirst = ComF.timeGetTimems();
                StepTimeLast = ComF.timeGetTimems();
                SpecOutputFlag = true;
                MinData = 99999;
                MaxData = 0;
                //Lumber Fwd Sw On
                LumberUp = true;

                plot2.Channels[0].Clear();
                SpreadPositionSelect(16, 7);
                CurrFilter.InitAll();
                if (0 < plot2.Channels[1].Count)
                    PlotTime = plot2.Channels[1].GetX(plot2.Channels[1].Count - 1);
                else PlotTime = 0;
            }
            else
            {
                float AdData;
                double RunTime = 0;

                StepTimeLast = ComF.timeGetTimems();

                AdData = pMeter.GetPSeat;

                RunTime = StepTimeLast - StepTimeFirst;

                if (mSpec.Sound.RMSMode == true)
                    plot2.Channels[0].AddXY(RunTime, CurrFilter.CheckData(0, (Sound.GetSound + mSpec.Sound.Offset)));
                else plot2.Channels[0].AddXY(RunTime, (Sound.GetSound + mSpec.Sound.Offset));

                if (mSpec.Sound.RMSMode == true)
                    plot2.Channels[1].AddXY(PlotTime + RunTime, CurrFilter.CheckData(0, (Sound.GetSound + mSpec.Sound.Offset)));
                else plot2.Channels[1].AddXY(PlotTime + RunTime, (Sound.GetSound + mSpec.Sound.Offset));

                plot2.XAxes[0].Tracking.ZoomToFitAll();
                plot2.YAxes[0].Tracking.ZoomToFitAll();

                //plot1.Channels[0].AddXY(plot1.Channels[0].Count, AdData);
                //plot1.YAxes[0].Tracking.ZoomToFitAll();
                //plot1.XAxes[0].Tracking.ZoomToFitAll();

                if ((mSpec.LumberUpDnTestTime * 1000) <= RunTime)
                {
                    //Slide Fwd Sw Off
                    //mData.LumberUpDn.Data = Math.Max(mData.LumberUpDn.Data, AdData);
                    //fpSpread1.ActiveSheet.Cells[16, 6].Text = mData.LumberUpDn.Data.ToString("0.0");

                    //float CheckData;

                    //if (float.TryParse(fpSpread1.ActiveSheet.Cells[16, 6].Text, out CheckData) == false) CheckData = 0;

                    LumberUp = false;

                    //mData.LumberUpDn.Data = CheckData;
                    mData.LumberUpDn.Test = true;
                    //if ((mSpec.Current.LumberUpDn.Min <= CheckData) && (CheckData <= mSpec.Current.LumberUpDn.Max))
                    //    mData.LumberUpDn.Result = RESULT.PASS;
                    //else mData.LumberUpDn.Result = RESULT.REJECT;

                    //if (MinData < mSpec.Current.LumberUpDn.Min)
                    //{
                    //    mData.LumberUpDn.Result = RESULT.REJECT;
                    //    mData.LumberUpDn.Data = MinData;
                    //}
                    //else 
                    if ((mSpec.Current.LumberUpDn.Min <= MaxData) && (MaxData <= mSpec.Current.LumberUpDn.Max))
                    {
                        mData.LumberUpDn.Data = MaxData;
                        mData.LumberUpDn.Result = RESULT.PASS;
                    }
                    else
                    {
                        mData.LumberUpDn.Result = RESULT.REJECT;
                        mData.LumberUpDn.Data = MaxData;
                    }

                    if (mData.LumberUpDn.Result == RESULT.REJECT) mData.Result = RESULT.REJECT;
                    fpSpread1.ActiveSheet.Cells[16, 6].Text = mData.LumberUpDn.Data.ToString("0.0");
                    fpSpread1.ActiveSheet.Cells[16, 7].Text = mData.LumberUpDn.Result == RESULT.PASS ? "O.K" : "N.G";
                    fpSpread1.ActiveSheet.Cells[16, 7].ForeColor = mData.LumberUpDn.Result == RESULT.PASS ? Color.Lime : Color.Red;

                    SoundLavelCheck(SoundLevePos.LUMBER_UP);
                    SoundCheck(SoundLevePos.LUMBER_UP);
                    Step++;
                    SpecOutputFlag = false;
                    ComF.timedelay(500);
                }
                else
                {
                    if ((500 <= RunTime) && (RunTime < ((mSpec.LumberUpDnTestTime - 0.5) * 1000)))
                    {
                        if (MinData == 99999)
                        {
                            if (mSpec.Current.LumberUpDn.Min < AdData) MinData = AdData;
                        }
                        else
                        {
                            MinData = Math.Min(MinData, AdData);
                        }
                        MaxData = Math.Max(MaxData, AdData);                    //mData.LumberUpDn.Data = Math.Max(mData.LumberUpDn.Data, AdData);
                        fpSpread1.ActiveSheet.Cells[16, 6].Text = AdData.ToString("0.0");
                    }
                }
            }
            return;
        }

        private void SoundLavelCheck(short Pos)
        {
            float StartTime = 0;
            float RunTime = 0;
            float Time = 0;
            float Sound = 0;
            double StartSound = 0;
            int StartCount = 0;
            double RunSound = 0;
            int RunCount = 0;
            long DelayTime = (long)(mSpec.Sound.기동음구동음구분사이시간 * 1000F);

            if (CheckItem.Sound == false) return;
            StartTime = (float)(mSpec.Sound.StartTime * 1000F);

            switch (Pos)
            {
                case SoundLevePos.SLIDE_FWD:
                    RunTime = (float)(mSpec.SlideSoundCheckTime * 1000F);
                    break;
                case SoundLevePos.RECLINE_BWD:
                    RunTime = (float)(mSpec.ReclinerSoundCheckTime * 1000F);
                    break;
                case SoundLevePos.TILT_UP:
                    RunTime = (float)(mSpec.TiltSoundCheckTime * 1000F);
                    break;
                case SoundLevePos.HEIGHT_UP:
                    RunTime = (float)(mSpec.HeightSoundCheckTime * 1000F);
                    break;
                case SoundLevePos.LUMBER_FWD:
                case SoundLevePos.LUMBER_UP:
                    RunTime = (float)(mSpec.LumberSoundCheckTime * 1000F);
                    break;
            }
            //RunTime -= 300;

            for (int i = 0; i < plot2.Channels[0].Count; i++)
            {
                Time = (float)plot2.Channels[0].GetX(i);
                Sound = (float)plot2.Channels[0].GetY(i);

                if (Time < StartTime)
                {
                    if (300 <= Time)
                    {
                        if (StartCount == 0)
                        {
                            StartSound = Sound;
                        }
                        else
                        {
                            //if (mSpec.Sound.RMSMode == true)
                            //    StartSound += Sound;
                            //else StartSound = Math.Max(StartSound, Sound);
                            StartSound = Math.Max(StartSound, Sound);
                        }
                        StartCount++;
                    }
                }
                else if (((StartTime  + DelayTime ) <= Time) && (Time <= RunTime))
                {
                    if (RunCount == 0)
                    {
                        RunSound = Sound;
                    }
                    else
                    {
                        //if (mSpec.Sound.RMSMode == true)
                        //    RunSound += Sound;
                        //else RunSound = Math.Max(RunSound, Sound);
                        RunSound = Math.Max(RunSound, Sound);
                    }
                    RunCount++;
                }
            }

            //-------------------------------------------------------
            //소음기 속도가 느려 데이타가 늦게 올라오는 경향이 있다. 그래서 약간의 옵션을 줌
            StartSound = StartSound * 0.95;
            RunSound = RunSound * 0.95;
            if (RunSound <= 0) RunSound = (plot2.Channels[0].GetY(plot2.Channels[0].Count - 1) * 0.95) - 1F;

            if (StartSound < RunSound) StartSound = RunSound + 1F;
            //-------------------------------------------------------

            //if (mSpec.Sound.RMSMode == true) RunSound = RunSound / (float)RunCount;
            //if (mSpec.Sound.RMSMode == true) StartSound = StartSound / (float)StartCount;

            switch (Pos)
            {
                case SoundLevePos.SLIDE_FWD:
                    mData.SoundSlide.StartData = (float)Math.Truncate(StartSound);
                    mData.SoundSlide.RunData = (float)Math.Truncate(RunSound);
                    mData.SoundSlide.Test = true;
                    break;
                case SoundLevePos.RECLINE_BWD:
                    mData.SoundRecline.StartData = (float)Math.Truncate(StartSound);
                    mData.SoundRecline.RunData = (float)Math.Truncate(RunSound);
                    mData.SoundRecline.Test = true;
                    break;
                case SoundLevePos.TILT_UP:
                    mData.SoundTilt.StartData = (float)Math.Truncate(StartSound);
                    mData.SoundTilt.RunData = (float)Math.Truncate(RunSound);
                    mData.SoundTilt.Test = true;
                    break;
                case SoundLevePos.HEIGHT_UP:
                    mData.SoundHeight.StartData = (float)Math.Truncate(StartSound);
                    mData.SoundHeight.RunData = (float)Math.Truncate(RunSound);
                    mData.SoundHeight.Test = true;
                    break;
                case SoundLevePos.LUMBER_FWD:
                    mData.SoundLumberFwdBwd.StartData = (float)Math.Truncate(StartSound);
                    mData.SoundLumberFwdBwd.RunData = (float)Math.Truncate(RunSound);
                    mData.SoundLumberFwdBwd.Test = true;
                    break;
                default:
                    mData.SoundLumberUpDn.StartData = (float)Math.Truncate(StartSound);
                    mData.SoundLumberUpDn.RunData = (float)Math.Truncate(RunSound);
                    mData.SoundLumberUpDn.Test = true;
                    break;
            }
            return;
        }


        //void SoundCheckInitPositionMove()
        //{
        //    if (SpecOutputFlag == false)
        //    {
        //        SlideFwdOnOff = true;
        //        ReclineFwdOnOff = true;
        //        TiltUpOnOff = true;
        //        HeightUpOnOff = true;
        //        StepTimeFirst = ComF.timeGetTimems();
        //        StepTimeLast = ComF.timeGetTimems();
        //        SpecOutputFlag = true;
        //        SlideMotorMoveEndFlag = false;
        //        ReclineMotorMoveEndFlag = false;
        //        TiltMotorMoveEndFlag = false;
        //        HeightMotorMoveEndFlag = false;
        //        label5.Text = "속도 측정을 하기 위해 측정 위치로 이동 중 입니다.";
        //    }
        //    else
        //    {
        //        if(30 < pMeter.GetPSeat)
        //        {
        //            SlideFwdOnOff = false;
        //            ReclineFwdOnOff = false;
        //            TiltUpOnOff = false;
        //            HeightUpOnOff = false;
        //            SpecOutputFlag = false;
        //            SlideMotorMoveEndFlag = true;
        //            ReclineMotorMoveEndFlag = true;
        //            TiltMotorMoveEndFlag = true;
        //            HeightMotorMoveEndFlag = true;
        //        }

        //        StepTimeLast = ComF.timeGetTimems();

        //        if (2000 <= (StepTimeLast - StepTimeFirst))
        //        {
        //            if (SlideMotorMoveEndFlag == false)
        //            {
        //                float AdData;

        //                bool Flag = false;
        //                if (comboBox4.SelectedItem.ToString() == "IMS")
        //                    AdData = pMeter.GetPSeat;
        //                else AdData = IOPort.ADRead[mSpec.PinMap.SlideFWD.PinNo];
        //                if (comboBox4.SelectedItem.ToString() == "IMS")
        //                {
        //                    if (AdData < mSpec.SlideLimitCurr) Flag = true;
        //                }
        //                else
        //                {
        //                    if (mSpec.SlideLimitCurr < AdData) Flag = true;
        //                }
        //                if (Flag == true)
        //                {
        //                    SlideFwdOnOff = false;
        //                    SlideMotorMoveEndFlag = true;
        //                }
        //                else if ((mSpec.SlideLimitTime * 1000) <= (StepTimeLast - StepTimeFirst))
        //                {
        //                    SlideFwdOnOff = false;
        //                    SlideMotorMoveEndFlag = true;
        //                }
        //            }
        //            if (ReclineMotorMoveEndFlag == false)
        //            {
        //                float AdData;

        //                bool Flag = false;
        //                if (comboBox4.SelectedItem.ToString() == "IMS")
        //                    AdData = pMeter.GetPSeat;
        //                else AdData = IOPort.ADRead[mSpec.PinMap.SlideFWD.PinNo];
        //                if (comboBox4.SelectedItem.ToString() == "IMS")
        //                {
        //                    if (AdData < mSpec.ReclinerLimitCurr) Flag = true;
        //                }
        //                else
        //                {
        //                    if (mSpec.ReclinerLimitCurr < AdData) Flag = true;
        //                }
        //                if (Flag == true)
        //                {
        //                    ReclineFwdOnOff = false;
        //                    ReclineMotorMoveEndFlag = true;
        //                }
        //                else if ((mSpec.ReclinerLimitTime * 1000) <= (StepTimeLast - StepTimeFirst))
        //                {
        //                    ReclineFwdOnOff = false;
        //                    ReclineMotorMoveEndFlag = true;
        //                }
        //            }
        //            if (TiltMotorMoveEndFlag == false)
        //            {
        //                float AdData;
        //                bool Flag = false;

        //                if (comboBox4.SelectedItem.ToString() == "IMS")
        //                    AdData = pMeter.GetPSeat;
        //                else AdData = IOPort.ADRead[mSpec.PinMap.TiltUp.PinNo];

        //                if (comboBox4.SelectedItem.ToString() == "IMS")
        //                {
        //                    if (AdData < mSpec.TiltLimitCurr) Flag = true;
        //                }
        //                else
        //                {
        //                    if (mSpec.TiltLimitCurr < AdData) Flag = true;
        //                }
        //                if (Flag == true)
        //                {
        //                    TiltUpOnOff = false;
        //                    TiltMotorMoveEndFlag = true;
        //                }
        //                else if ((mSpec.TiltLimitTime * 1000) <= (StepTimeLast - StepTimeFirst))
        //                {
        //                    TiltUpOnOff = false;
        //                    TiltMotorMoveEndFlag = true;
        //                }
        //            }
        //            if (HeightMotorMoveEndFlag == false)
        //            {
        //                float AdData;
        //                bool Flag = false;

        //                if (comboBox4.SelectedItem.ToString() == "IMS")
        //                    AdData = pMeter.GetPSeat;
        //                else AdData = IOPort.ADRead[mSpec.PinMap.HeightUp.PinNo];
        //                if (comboBox4.SelectedItem.ToString() == "IMS")
        //                {
        //                    if (AdData < mSpec.HeightLimitCurr) Flag = true;
        //                }
        //                else
        //                {
        //                    if (mSpec.HeightLimitCurr < AdData) Flag = true;
        //                }
        //                if (Flag == true)
        //                {
        //                    HeightUpOnOff = false;
        //                    HeightMotorMoveEndFlag = true;
        //                }
        //                else if ((mSpec.HeightLimitTime * 1000) <= (StepTimeLast - StepTimeFirst))
        //                {
        //                    HeightUpOnOff = false;
        //                    HeightMotorMoveEndFlag = true;
        //                }
        //            }
        //        }

        //        if ((SlideMotorMoveEndFlag == true) && (TiltMotorMoveEndFlag == true) && (HeightMotorMoveEndFlag == true))
        //        {
        //            SpecOutputFlag = false;
        //            Step++;
        //            ComF.timedelay(500);
        //        }
        //    }
        //    return;
        //}
        //private bool[] SpeedCheckEnd = { false, false, false, false, false, false, false, false };
        //private void CheckSpeedToIMS(short SpeedCheckStep)
        //{
        //    if (SpecOutputFlag == false)
        //    {
        //        StepTimeFirst = ComF.timeGetTimems();
        //        StepTimeLast = ComF.timeGetTimems();
        //        SpecOutputFlag = true;

        //        //Slide Fwd Sw On

        //        switch (SpeedCheckStep)
        //        {
        //            case 0:
        //                //FirstPosCheckEnd[SpeedCheckStep] = false;
        //                plot2.Channels[0].Clear();
        //                SoundCheckTimeToStart[SpeedCheckStep] = plot2.Channels[0].Count;
        //                SlideBwdOnOff = true;
        //                label5.Text = "SLIDE BACKWORD 속도 측정 중 입니다.";
        //                break;
        //            case 1:
        //                //FirstPosCheckEnd[SpeedCheckStep] = false;
        //                plot2.Channels[0].Clear();
        //                SoundCheckTimeToStart[SpeedCheckStep] = plot2.Channels[0].Count;
        //                SlideFwdOnOff = true;
        //                label5.Text = "SLIDE FORWORD 속도 측정 중 입니다.";
        //                break;
        //            case 2:
        //                //FirstPosCheckEnd[SpeedCheckStep] = false;
        //                plot2.Channels[0].Clear();
        //                SoundCheckTimeToStart[SpeedCheckStep] = plot2.Channels[0].Count;
        //                ReclineBwdOnOff = true;
        //                label5.Text = "RECLINE BACKWORD 속도 측정 중 입니다.";
        //                break;
        //            case 3:
        //                //FirstPosCheckEnd[SpeedCheckStep] = false;
        //                plot2.Channels[0].Clear();
        //                SoundCheckTimeToStart[SpeedCheckStep] = plot2.Channels[0].Count;
        //                ReclineFwdOnOff = true;
        //                label5.Text = "RECLINE FORWORD 속도 측정 중 입니다.";
        //                break;
        //            case 4:
        //                //FirstPosCheckEnd[SpeedCheckStep] = false;
        //                plot2.Channels[0].Clear();
        //                SoundCheckTimeToStart[SpeedCheckStep] = plot2.Channels[0].Count;
        //                TiltDnOnOff = true;
        //                label5.Text = "TULT UP 속도 측정 중 입니다.";
        //                break;
        //            case 5:
        //                //FirstPosCheckEnd[SpeedCheckStep] = false;
        //                plot2.Channels[0].Clear();
        //                SoundCheckTimeToStart[SpeedCheckStep] = plot2.Channels[0].Count;
        //                TiltUpOnOff = true;
        //                label5.Text = "TULT DOWN 속도 측정 중 입니다.";
        //                break;
        //            case 6:
        //                //FirstPosCheckEnd[SpeedCheckStep] = false;
        //                plot2.Channels[0].Clear();
        //                SoundCheckTimeToStart[SpeedCheckStep] = plot2.Channels[0].Count;
        //                HeightDnOnOff = true;
        //                label5.Text = "HEIGHT UP 속도 측정 중 입니다.";
        //                break;
        //            case 7:
        //                //FirstPosCheckEnd[SpeedCheckStep] = false;
        //                plot2.Channels[0].Clear();
        //                SoundCheckTimeToStart[SpeedCheckStep] = plot2.Channels[0].Count;
        //                HeightUpOnOff = true;
        //                label5.Text = "HEIGHT DOWN 속도 측정 중 입니다.";
        //                break;
        //        }
        //        SpeedCheckStartTime[SpeedCheckStep] = ComF.timeGetTimems();
        //    }
        //    else
        //    {
        //        StepTimeLast = ComF.timeGetTimems();

        //        float AdData = pMeter.GetPSeat;

        //        if (500 <= (StepTimeLast - StepTimeFirst))
        //        {
        //            plot2.Channels[0].AddXY(plot2.Channels[0].Count, Math.Abs(Sound.GetSound - NormalSound));
        //            plot2.XAxes[0].Tracking.ZoomToFitAll();
        //            plot2.YAxes[0].Tracking.ZoomToFitAll();
        //            if (SpeedCheckStep == 0)
        //            {
        //                if (/*(FirstPosCheckEnd[SpeedCheckStep] == false) && (*/AdData < mSpec.SlideLimitCurr)//)
        //                {
        //                    SpeedCheckEnd[SpeedCheckStep] = true;
        //                    SlideBwdOnOff = false;
        //                    SoundCheckTimeToEnd[SpeedCheckStep] = plot2.Channels[0].Count;

        //                    float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[SpeedCheckStep]) / 1000F;
        //                    mData.SlideBwdTime.Data = (mSpec.MovingStroke.Slide / EndTime) * 1F;
        //                    mData.SlideBwdTime.Test = true;
        //                    if ((mSpec.MovingSpeed.SlideBwd.Min <= mData.SlideBwdTime.Data) && (mData.SlideBwdTime.Data <= mSpec.MovingSpeed.SlideBwd.Max))
        //                        mData.SlideBwdTime.Result = RESULT.PASS;
        //                    else mData.SlideBwdTime.Result = RESULT.REJECT;

        //                    fpSpread1.ActiveSheet.Cells[4, 6].Text = mData.SlideBwdTime.Data.ToString("0.0");
        //                    fpSpread1.ActiveSheet.Cells[4, 7].Text = mData.SlideBwdTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                    fpSpread1.ActiveSheet.Cells[4, 7].ForeColor = mData.SlideBwdTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                }
        //                else
        //                {
        //                    //if (FirstPosCheckEnd[SpeedCheckStep] == false)
        //                    //{
        //                    if ((mSpec.SlideLimitTime * 1000) <= (StepTimeLast - StepTimeFirst))
        //                    {
        //                        SlideBwdOnOff = false;
        //                        float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[SpeedCheckStep]) / 1000F;
        //                        mData.SlideBwdTime.Data = (mSpec.MovingStroke.Slide / EndTime) * 1F;
        //                        mData.SlideBwdTime.Test = true;
        //                        if ((mSpec.MovingSpeed.SlideBwd.Min <= mData.SlideBwdTime.Data) && (mData.SlideBwdTime.Data <= mSpec.MovingSpeed.SlideBwd.Max))
        //                            mData.SlideBwdTime.Result = RESULT.PASS;
        //                        else mData.SlideBwdTime.Result = RESULT.REJECT;

        //                        fpSpread1.ActiveSheet.Cells[4, 6].Text = mData.SlideBwdTime.Data.ToString("0.0");
        //                        fpSpread1.ActiveSheet.Cells[4, 7].Text = mData.SlideBwdTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                        fpSpread1.ActiveSheet.Cells[4, 7].ForeColor = mData.SlideBwdTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                    }
        //                    //}
        //                }
        //            }
        //            else if (SpeedCheckStep == 1)
        //            {
        //                if (/*(FirstPosCheckEnd[SpeedCheckStep] == false) && (*/AdData < mSpec.SlideLimitCurr)//)
        //                {
        //                    SpeedCheckEnd[SpeedCheckStep] = true;
        //                    SlideFwdOnOff = false;
        //                    SoundCheckTimeToEnd[SpeedCheckStep] = plot2.Channels[0].Count;

        //                    float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[SpeedCheckStep]) / 1000F;
        //                    mData.SlideFwdTime.Data = (mSpec.MovingStroke.Slide / EndTime) * 1F;
        //                    mData.SlideFwdTime.Test = true;
        //                    if ((mSpec.MovingSpeed.SlideFwd.Min <= mData.SlideFwdTime.Data) && (mData.SlideFwdTime.Data <= mSpec.MovingSpeed.SlideFwd.Max))
        //                        mData.SlideFwdTime.Result = RESULT.PASS;
        //                    else mData.SlideFwdTime.Result = RESULT.REJECT;

        //                    fpSpread1.ActiveSheet.Cells[3, 6].Text = mData.SlideFwdTime.Data.ToString("0.0");
        //                    fpSpread1.ActiveSheet.Cells[3, 7].Text = mData.SlideFwdTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                    fpSpread1.ActiveSheet.Cells[3, 7].ForeColor = mData.SlideFwdTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                }
        //                else
        //                {
        //                    //if (FirstPosCheckEnd[SpeedCheckStep] == false)
        //                    //{
        //                    if ((mSpec.SlideLimitTime * 1000) <= (StepTimeLast - StepTimeFirst))
        //                    {
        //                        SpeedCheckEnd[SpeedCheckStep] = true;
        //                        SlideFwdOnOff = false;
        //                        SoundCheckTimeToEnd[SpeedCheckStep] = plot2.Channels[0].Count;

        //                        float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[SpeedCheckStep]) / 1000F;
        //                        mData.SlideFwdTime.Data = (mSpec.MovingStroke.Slide / EndTime) * 1F;
        //                        mData.SlideFwdTime.Test = true;
        //                        if ((mSpec.MovingSpeed.SlideFwd.Min <= mData.SlideFwdTime.Data) && (mData.SlideFwdTime.Data <= mSpec.MovingSpeed.SlideFwd.Max))
        //                            mData.SlideFwdTime.Result = RESULT.PASS;
        //                        else mData.SlideFwdTime.Result = RESULT.REJECT;

        //                        fpSpread1.ActiveSheet.Cells[3, 6].Text = mData.SlideFwdTime.Data.ToString("0.0");
        //                        fpSpread1.ActiveSheet.Cells[3, 7].Text = mData.SlideFwdTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                        fpSpread1.ActiveSheet.Cells[3, 7].ForeColor = mData.SlideFwdTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                    }
        //                    //}
        //                }
        //            }
        //            else if (SpeedCheckStep == 2)
        //            {
        //                if (/*(FirstPosCheckEnd[SpeedCheckStep] == false) && (*/AdData < mSpec.SlideLimitCurr)//)
        //                {
        //                    SpeedCheckEnd[SpeedCheckStep] = true;
        //                    ReclineBwdOnOff = false;
        //                    SoundCheckTimeToEnd[SpeedCheckStep] = plot2.Channels[0].Count;

        //                    float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[SpeedCheckStep]) / 1000F;
        //                    mData.ReclineBwdTime.Data = (mSpec.MovingStroke.Recliner / EndTime) * 1F;
        //                    mData.ReclineBwdTime.Test = true;
        //                    if ((mSpec.MovingSpeed.ReclinerBwd.Min <= mData.ReclineBwdTime.Data) && (mData.ReclineBwdTime.Data <= mSpec.MovingSpeed.ReclinerBwd.Max))
        //                        mData.ReclineBwdTime.Result = RESULT.PASS;
        //                    else mData.ReclineBwdTime.Result = RESULT.REJECT;

        //                    fpSpread1.ActiveSheet.Cells[6, 6].Text = mData.ReclineBwdTime.Data.ToString("0.0");
        //                    fpSpread1.ActiveSheet.Cells[6, 7].Text = mData.ReclineBwdTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                    fpSpread1.ActiveSheet.Cells[6, 7].ForeColor = mData.ReclineBwdTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                }
        //                else
        //                {
        //                    //if (FirstPosCheckEnd[SpeedCheckStep] == false)
        //                    //{
        //                    if ((mSpec.ReclinerLimitTime * 1000) <= (StepTimeLast - StepTimeFirst))
        //                    {
        //                        ReclineBwdOnOff = false;
        //                        float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[SpeedCheckStep]) / 1000F;
        //                        mData.ReclineBwdTime.Data = (mSpec.MovingStroke.Recliner / EndTime) * 1F;
        //                        mData.ReclineBwdTime.Test = true;
        //                        if ((mSpec.MovingSpeed.ReclinerBwd.Min <= mData.ReclineBwdTime.Data) && (mData.ReclineBwdTime.Data <= mSpec.MovingSpeed.ReclinerBwd.Max))
        //                            mData.ReclineBwdTime.Result = RESULT.PASS;
        //                        else mData.ReclineBwdTime.Result = RESULT.REJECT;

        //                        fpSpread1.ActiveSheet.Cells[6, 6].Text = mData.ReclineBwdTime.Data.ToString("0.0");
        //                        fpSpread1.ActiveSheet.Cells[6, 7].Text = mData.ReclineBwdTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                        fpSpread1.ActiveSheet.Cells[6, 7].ForeColor = mData.ReclineBwdTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                    }
        //                    //}
        //                }
        //            }
        //            else if (SpeedCheckStep == 3)
        //            {
        //                if (/*(FirstPosCheckEnd[SpeedCheckStep] == false) && (*/AdData < mSpec.ReclinerLimitCurr)//)
        //                {
        //                    SpeedCheckEnd[SpeedCheckStep] = true;
        //                    ReclineFwdOnOff = false;
        //                    SoundCheckTimeToEnd[SpeedCheckStep] = plot2.Channels[0].Count;

        //                    float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[SpeedCheckStep]) / 1000F;
        //                    mData.ReclineFwdTime.Data = (mSpec.MovingStroke.Recliner / EndTime) * 1F;
        //                    mData.ReclineFwdTime.Test = true;
        //                    if ((mSpec.MovingSpeed.ReclinerFwd.Min <= mData.ReclineFwdTime.Data) && (mData.ReclineFwdTime.Data <= mSpec.MovingSpeed.ReclinerFwd.Max))
        //                        mData.ReclineFwdTime.Result = RESULT.PASS;
        //                    else mData.ReclineFwdTime.Result = RESULT.REJECT;

        //                    fpSpread1.ActiveSheet.Cells[5, 6].Text = mData.ReclineFwdTime.Data.ToString("0.0");
        //                    fpSpread1.ActiveSheet.Cells[5, 7].Text = mData.ReclineFwdTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                    fpSpread1.ActiveSheet.Cells[5, 7].ForeColor = mData.ReclineFwdTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                }
        //                else
        //                {
        //                    //if (FirstPosCheckEnd[SpeedCheckStep] == false)
        //                    //{
        //                    if ((mSpec.ReclinerLimitTime * 1000) <= (StepTimeLast - StepTimeFirst))
        //                    {
        //                        SpeedCheckEnd[SpeedCheckStep] = true;
        //                        ReclineFwdOnOff = false;
        //                        SoundCheckTimeToEnd[SpeedCheckStep] = plot2.Channels[0].Count;

        //                        float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[SpeedCheckStep]) / 1000F;
        //                        mData.ReclineFwdTime.Data = (mSpec.MovingStroke.Recliner / EndTime) * 1F;
        //                        mData.ReclineFwdTime.Test = true;
        //                        if ((mSpec.MovingSpeed.ReclinerFwd.Min <= mData.ReclineFwdTime.Data) && (mData.ReclineFwdTime.Data <= mSpec.MovingSpeed.ReclinerFwd.Max))
        //                            mData.ReclineFwdTime.Result = RESULT.PASS;
        //                        else mData.ReclineFwdTime.Result = RESULT.REJECT;

        //                        fpSpread1.ActiveSheet.Cells[5, 6].Text = mData.ReclineFwdTime.Data.ToString("0.0");
        //                        fpSpread1.ActiveSheet.Cells[5, 7].Text = mData.ReclineFwdTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                        fpSpread1.ActiveSheet.Cells[5, 7].ForeColor = mData.ReclineFwdTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                    }
        //                    //}
        //                }
        //            }
        //            else if (SpeedCheckStep == 4)
        //            {
        //                if (/*(FirstPosCheckEnd[SpeedCheckStep] == false) && (*/AdData < mSpec.TiltLimitCurr)//)
        //                {
        //                    SpeedCheckEnd[SpeedCheckStep] = true;
        //                    TiltDnOnOff = false;
        //                    SoundCheckTimeToEnd[SpeedCheckStep] = plot2.Channels[0].Count;
        //                    float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[SpeedCheckStep]) / 1000F;
        //                    mData.TiltDnTime.Data = (mSpec.MovingStroke.Tilt / EndTime) * 1F;
        //                    mData.TiltDnTime.Test = true;
        //                    if ((mSpec.MovingSpeed.TiltDn.Min <= mData.TiltDnTime.Data) && (mData.TiltDnTime.Data <= mSpec.MovingSpeed.TiltDn.Max))
        //                        mData.TiltDnTime.Result = RESULT.PASS;
        //                    else mData.TiltDnTime.Result = RESULT.REJECT;

        //                    fpSpread1.ActiveSheet.Cells[8, 6].Text = mData.TiltDnTime.Data.ToString("0.0");
        //                    fpSpread1.ActiveSheet.Cells[8, 7].Text = mData.TiltDnTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                    fpSpread1.ActiveSheet.Cells[8, 7].ForeColor = mData.TiltDnTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                }
        //                else
        //                {
        //                    //if (FirstPosCheckEnd[SpeedCheckStep] == false)
        //                    //{
        //                    if ((mSpec.TiltLimitTime * 1000) <= (StepTimeLast - StepTimeFirst))
        //                    {
        //                        SpeedCheckEnd[SpeedCheckStep] = true;
        //                        TiltDnOnOff = false;
        //                        SoundCheckTimeToEnd[SpeedCheckStep] = plot2.Channels[0].Count;
        //                        float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[SpeedCheckStep]) / 1000F;
        //                        mData.TiltDnTime.Data = (mSpec.MovingStroke.Tilt / EndTime) * 1F;
        //                        mData.TiltDnTime.Test = true;
        //                        if ((mSpec.MovingSpeed.TiltDn.Min <= mData.TiltDnTime.Data) && (mData.TiltDnTime.Data <= mSpec.MovingSpeed.TiltDn.Max))
        //                            mData.TiltDnTime.Result = RESULT.PASS;
        //                        else mData.TiltDnTime.Result = RESULT.REJECT;

        //                        fpSpread1.ActiveSheet.Cells[8, 6].Text = mData.TiltDnTime.Data.ToString("0.0");
        //                        fpSpread1.ActiveSheet.Cells[8, 7].Text = mData.TiltDnTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                        fpSpread1.ActiveSheet.Cells[8, 7].ForeColor = mData.TiltDnTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                    }
        //                    //}
        //                }
        //            }
        //            else if (SpeedCheckStep == 5)
        //            {
        //                if (/*(FirstPosCheckEnd[SpeedCheckStep] == false) && (*/AdData < mSpec.TiltLimitCurr)//)
        //                {
        //                    SpeedCheckEnd[SpeedCheckStep] = true;
        //                    TiltUpOnOff = false;
        //                    SoundCheckTimeToEnd[SpeedCheckStep] = plot2.Channels[0].Count;
        //                    float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[SpeedCheckStep]) / 1000F;
        //                    mData.TiltUpTime.Data = (mSpec.MovingStroke.Tilt / EndTime) * 1F;
        //                    mData.TiltUpTime.Test = true;
        //                    if ((mSpec.MovingSpeed.TiltUp.Min <= mData.TiltUpTime.Data) && (mData.TiltUpTime.Data <= mSpec.MovingSpeed.TiltUp.Max))
        //                        mData.TiltUpTime.Result = RESULT.PASS;
        //                    else mData.TiltUpTime.Result = RESULT.REJECT;

        //                    fpSpread1.ActiveSheet.Cells[7, 6].Text = mData.TiltUpTime.Data.ToString("0.0");
        //                    fpSpread1.ActiveSheet.Cells[7, 7].Text = mData.TiltUpTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                    fpSpread1.ActiveSheet.Cells[7, 7].ForeColor = mData.TiltUpTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                }
        //                else
        //                {
        //                    //if (FirstPosCheckEnd[SpeedCheckStep] == false)
        //                    //{
        //                    if ((mSpec.TiltLimitTime * 1000) <= (StepTimeLast - StepTimeFirst))
        //                    {
        //                        SpeedCheckEnd[SpeedCheckStep] = true;
        //                        TiltUpOnOff = false;
        //                        SoundCheckTimeToEnd[SpeedCheckStep] = plot2.Channels[0].Count;
        //                        float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[SpeedCheckStep]) / 1000F;
        //                        mData.TiltUpTime.Data = (mSpec.MovingStroke.Tilt / EndTime) * 1F;
        //                        mData.TiltUpTime.Test = true;
        //                        if ((mSpec.MovingSpeed.TiltUp.Min <= mData.TiltUpTime.Data) && (mData.TiltUpTime.Data <= mSpec.MovingSpeed.TiltUp.Max))
        //                            mData.TiltUpTime.Result = RESULT.PASS;
        //                        else mData.TiltUpTime.Result = RESULT.REJECT;

        //                        fpSpread1.ActiveSheet.Cells[7, 6].Text = mData.TiltUpTime.Data.ToString("0.0");
        //                        fpSpread1.ActiveSheet.Cells[7, 7].Text = mData.TiltUpTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                        fpSpread1.ActiveSheet.Cells[7, 7].ForeColor = mData.TiltUpTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                    }
        //                    //}
        //                }
        //            }
        //            else if (SpeedCheckStep == 6)
        //            {
        //                if (/*(FirstPosCheckEnd[SpeedCheckStep] == false) && (*/AdData < mSpec.HeightLimitCurr)//)
        //                {
        //                    SpeedCheckEnd[SpeedCheckStep] = true;
        //                    HeightDnOnOff = false;
        //                    SoundCheckTimeToEnd[SpeedCheckStep] = plot2.Channels[0].Count;
        //                    float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[SpeedCheckStep]) / 1000F;
        //                    mData.HeightDnTime.Data = (mSpec.MovingStroke.Height / EndTime) * 1F;
        //                    mData.HeightDnTime.Test = true;
        //                    if ((mSpec.MovingSpeed.HeightDn.Min <= mData.HeightDnTime.Data) && (mData.HeightDnTime.Data <= mSpec.MovingSpeed.HeightDn.Max))
        //                        mData.HeightDnTime.Result = RESULT.PASS;
        //                    else mData.HeightDnTime.Result = RESULT.REJECT;

        //                    fpSpread1.ActiveSheet.Cells[10, 6].Text = mData.HeightDnTime.Data.ToString("0.0");
        //                    fpSpread1.ActiveSheet.Cells[10, 7].Text = mData.HeightDnTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                    fpSpread1.ActiveSheet.Cells[10, 7].ForeColor = mData.HeightDnTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                }
        //                else
        //                {
        //                    //if (FirstPosCheckEnd[SpeedCheckStep] == false)
        //                    //{
        //                    if ((mSpec.HeightLimitTime * 1000) <= (StepTimeLast - StepTimeFirst))
        //                    {
        //                        SpeedCheckEnd[SpeedCheckStep] = true;
        //                        HeightDnOnOff = false;
        //                        SoundCheckTimeToEnd[SpeedCheckStep] = plot2.Channels[0].Count;
        //                        float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[SpeedCheckStep]) / 1000F;
        //                        mData.HeightDnTime.Data = (mSpec.MovingStroke.Height / EndTime) * 1F;
        //                        mData.HeightDnTime.Test = true;
        //                        if ((mSpec.MovingSpeed.HeightDn.Min <= mData.HeightDnTime.Data) && (mData.HeightDnTime.Data <= mSpec.MovingSpeed.HeightDn.Max))
        //                            mData.HeightDnTime.Result = RESULT.PASS;
        //                        else mData.HeightDnTime.Result = RESULT.REJECT;

        //                        fpSpread1.ActiveSheet.Cells[10, 6].Text = mData.HeightDnTime.Data.ToString("0.0");
        //                        fpSpread1.ActiveSheet.Cells[10, 7].Text = mData.HeightDnTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                        fpSpread1.ActiveSheet.Cells[10, 7].ForeColor = mData.HeightDnTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                    }
        //                    //}
        //                }
        //            }
        //            else
        //            {
        //                if (/*(FirstPosCheckEnd[SpeedCheckStep] == false) && (*/AdData < mSpec.HeightLimitCurr)//)
        //                {
        //                    SpeedCheckEnd[SpeedCheckStep] = true;
        //                    HeightUpOnOff = false;
        //                    SoundCheckTimeToEnd[SpeedCheckStep] = plot2.Channels[0].Count;
        //                    float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[SpeedCheckStep]) / 1000F;
        //                    mData.HeightUpTime.Data = (mSpec.MovingStroke.Height / EndTime) * 1F;
        //                    mData.HeightUpTime.Test = true;
        //                    if ((mSpec.MovingSpeed.HeightUp.Min <= mData.HeightUpTime.Data) && (mData.HeightUpTime.Data <= mSpec.MovingSpeed.HeightUp.Max))
        //                        mData.HeightUpTime.Result = RESULT.PASS;
        //                    else mData.HeightUpTime.Result = RESULT.REJECT;

        //                    fpSpread1.ActiveSheet.Cells[9, 6].Text = mData.HeightUpTime.Data.ToString("0.0");
        //                    fpSpread1.ActiveSheet.Cells[9, 7].Text = mData.HeightUpTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                    fpSpread1.ActiveSheet.Cells[9, 7].ForeColor = mData.HeightUpTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                }
        //                else
        //                {
        //                    //if (FirstPosCheckEnd[SpeedCheckStep] == false)
        //                    //{
        //                    if ((mSpec.HeightLimitTime * 1000) <= (StepTimeLast - StepTimeFirst))
        //                    {
        //                        SpeedCheckEnd[SpeedCheckStep] = true;
        //                        HeightDnOnOff = false;
        //                        SoundCheckTimeToEnd[SpeedCheckStep] = plot2.Channels[0].Count;
        //                        float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[SpeedCheckStep]) / 1000F;
        //                        mData.HeightUpTime.Data = (mSpec.MovingStroke.Height / EndTime) * 1F;
        //                        mData.HeightUpTime.Test = true;
        //                        if ((mSpec.MovingSpeed.HeightUp.Min <= mData.HeightUpTime.Data) && (mData.HeightUpTime.Data <= mSpec.MovingSpeed.HeightUp.Max))
        //                            mData.HeightUpTime.Result = RESULT.PASS;
        //                        else mData.HeightUpTime.Result = RESULT.REJECT;

        //                        fpSpread1.ActiveSheet.Cells[9, 6].Text = mData.HeightUpTime.Data.ToString("0.0");
        //                        fpSpread1.ActiveSheet.Cells[9, 7].Text = mData.HeightUpTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                        fpSpread1.ActiveSheet.Cells[9, 7].ForeColor = mData.HeightUpTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                    }
        //                    //}
        //                }
        //            }
        //        }
        //        if (SpeedCheckEnd[SpeedCheckStep] == true)
        //        {
        //            SoundLavelCheck((short)(SpeedCheckStep / 2));
        //            SubStep++;
        //            SpecOutputFlag = false;
        //        }
        //    }
        //    return;
        //}

        //private bool NotImsSlideSpeedCheckFlag { get; set; }
        //private bool NotImsReclineSpeedCheckFlag { get; set; }
        //private bool NotImsTiltSpeedCheckFlag { get; set; }
        //private bool NotImsHeightSpeedCheckFlag { get; set; }

        //private bool NotImsSlideSpeedCheckEnd { get; set; }
        //private bool NotImsReclineSpeedCheckEnd { get; set; }
        //private bool NotImsTiltSpeedCheckEnd { get; set; }
        //private bool NotImsHeightSpeedCheckEnd { get; set; }
        //private long[] SpeedCheckStartTime = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        //private void CheckSpeedNotIMS()
        //{
        //    if (SpecOutputFlag == false)
        //    {
        //        label5.Text = "P/SEAT 속도 측정 중 입니다.";
        //        //label5.Text = "TULT UP 속도 측정 중 입니다.";
        //        //label5.Text = "HEIGHT UP 속도 측정 중 입니다.";
        //        //label5.Text = "RECLINE UP 속도 측정 중 입니다.";
        //        SlideBwdOnOff = true;
        //        ReclineBwdOnOff = true;
        //        TiltDnOnOff = true;
        //        HeightDnOnOff = true;

        //        SpeedCheckStartTime[0] = ComF.timeGetTimems();
        //        SpeedCheckStartTime[1] = ComF.timeGetTimems();
        //        SpeedCheckStartTime[2] = ComF.timeGetTimems();
        //        SpeedCheckStartTime[3] = ComF.timeGetTimems();
        //        NotImsSlideSpeedCheckFlag = false;
        //        NotImsReclineSpeedCheckFlag = false;
        //        NotImsTiltSpeedCheckFlag = false;
        //        NotImsHeightSpeedCheckFlag = false;

        //        NotImsSlideSpeedCheckEnd = false;
        //        NotImsReclineSpeedCheckEnd = false;
        //        NotImsTiltSpeedCheckEnd = false;
        //        NotImsHeightSpeedCheckEnd = false;
        //        SpecOutputFlag = true;
        //    }
        //    else
        //    {
        //        float AdData1;
        //        float AdData2;
        //        float AdData3;
        //        float AdData4;

        //        if (NotImsSlideSpeedCheckEnd == false)
        //        {
        //            if (NotImsSlideSpeedCheckFlag == false)
        //                AdData1 = IOPort.ADRead[mSpec.PinMap.SlideBWD.PinNo];
        //            else AdData1 = IOPort.ADRead[mSpec.PinMap.SlideFWD.PinNo];

        //            if (1000 <= (ComF.timeGetTimems() - SpeedCheckStartTime[0]))
        //            {
        //                bool Flag = false;

        //                if (comboBox4.SelectedItem.ToString() == "IMS")
        //                {
        //                    if ((AdData1 <= mSpec.SlideLimitCurr) || ((mSpec.SlideLimitTime * 1000) <= (ComF.timeGetTimems() - SpeedCheckStartTime[0]))) Flag = true;
        //                }
        //                else
        //                {
        //                    if ((mSpec.SlideLimitCurr <= AdData1) || ((mSpec.SlideLimitTime * 1000) <= (ComF.timeGetTimems() - SpeedCheckStartTime[0]))) Flag = true;
        //                }
        //                if (Flag == true)
        //                {
        //                    if (NotImsSlideSpeedCheckFlag == false)
        //                        SlideBwdOnOff = false;
        //                    else SlideFwdOnOff = false;

        //                    float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[0]) / 1000F;

        //                    if (NotImsSlideSpeedCheckFlag == false)
        //                    {
        //                        mData.SlideBwdTime.Test = true;
        //                        mData.SlideBwdTime.Data = (mSpec.MovingStroke.Slide / EndTime) * 1F;
        //                        if ((mSpec.MovingSpeed.SlideBwd.Min <= mData.SlideBwdTime.Data) && (mData.SlideBwdTime.Data <= mSpec.MovingSpeed.SlideBwd.Max))
        //                            mData.SlideBwdTime.Result = RESULT.PASS;
        //                        else mData.SlideBwdTime.Result = RESULT.REJECT;

        //                        fpSpread1.Sheets[0].Cells[4, 6].Text = mData.SlideBwdTime.Data.ToString("0.0");
        //                        fpSpread1.Sheets[0].Cells[4, 7].Text = mData.SlideBwdTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                        fpSpread1.Sheets[0].Cells[4, 7].ForeColor = mData.SlideBwdTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                    }
        //                    else
        //                    {
        //                        mData.SlideFwdTime.Test = true;
        //                        mData.SlideFwdTime.Data = (mSpec.MovingStroke.Slide / EndTime) * 1F;
        //                        if ((mSpec.MovingSpeed.SlideFwd.Min <= mData.SlideFwdTime.Data) && (mData.SlideFwdTime.Data <= mSpec.MovingSpeed.SlideFwd.Max))
        //                            mData.SlideFwdTime.Result = RESULT.PASS;
        //                        else mData.SlideFwdTime.Result = RESULT.REJECT;

        //                        fpSpread1.Sheets[0].Cells[3, 6].Text = mData.SlideFwdTime.Data.ToString("0.0");
        //                        fpSpread1.Sheets[0].Cells[3, 7].Text = mData.SlideFwdTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                        fpSpread1.Sheets[0].Cells[3, 7].ForeColor = mData.SlideFwdTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                    }

        //                    if (NotImsSlideSpeedCheckFlag == false)
        //                    {
        //                        ComF.timedelay(200);
        //                        NotImsSlideSpeedCheckFlag = true;
        //                        SlideFwdOnOff = true;
        //                        SpeedCheckStartTime[0] = ComF.timeGetTimems();
        //                    }
        //                    else
        //                    {
        //                        NotImsSlideSpeedCheckEnd = true;
        //                    }
        //                }
        //            }
        //        }


        //        if (NotImsReclineSpeedCheckEnd == false)
        //        {
        //            if (NotImsReclineSpeedCheckFlag == false)
        //                AdData2 = IOPort.ADRead[mSpec.PinMap.ReclineBWD.PinNo];
        //            else AdData2 = IOPort.ADRead[mSpec.PinMap.ReclineFWD.PinNo];

        //            if (1000 <= (ComF.timeGetTimems() - SpeedCheckStartTime[0]))
        //            {
        //                bool Flag = false;

        //                if (comboBox4.SelectedItem.ToString() == "IMS")
        //                {
        //                    if ((AdData2 <= mSpec.ReclinerLimitCurr) || ((mSpec.ReclinerLimitTime * 1000) <= (ComF.timeGetTimems() - SpeedCheckStartTime[0]))) Flag = true;
        //                }
        //                else
        //                {
        //                    if ((mSpec.ReclinerLimitCurr <= AdData2) || ((mSpec.ReclinerLimitTime * 1000) <= (ComF.timeGetTimems() - SpeedCheckStartTime[0]))) Flag = true;
        //                }
        //                if (Flag == true)
        //                {
        //                    if (NotImsReclineSpeedCheckFlag == false)
        //                        ReclineBwdOnOff = false;
        //                    else ReclineFwdOnOff = false;

        //                    float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[0]) / 1000F;

        //                    if (NotImsReclineSpeedCheckFlag == false)
        //                    {
        //                        mData.ReclineBwdTime.Test = true;
        //                        mData.ReclineBwdTime.Data = (mSpec.MovingStroke.Recliner / EndTime) * 1F;
        //                        if ((mSpec.MovingSpeed.ReclinerBwd.Min <= mData.ReclineBwdTime.Data) && (mData.ReclineBwdTime.Data <= mSpec.MovingSpeed.ReclinerBwd.Max))
        //                            mData.ReclineBwdTime.Result = RESULT.PASS;
        //                        else mData.ReclineBwdTime.Result = RESULT.REJECT;

        //                        fpSpread1.Sheets[0].Cells[6, 6].Text = mData.ReclineBwdTime.Data.ToString("0.0");
        //                        fpSpread1.Sheets[0].Cells[6, 7].Text = mData.ReclineBwdTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                        fpSpread1.Sheets[0].Cells[6, 7].ForeColor = mData.ReclineBwdTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                    }
        //                    else
        //                    {
        //                        mData.ReclineFwdTime.Test = true;
        //                        mData.ReclineFwdTime.Data = (mSpec.MovingStroke.Recliner / EndTime) * 1F;
        //                        if ((mSpec.MovingSpeed.ReclinerFwd.Min <= mData.ReclineFwdTime.Data) && (mData.ReclineFwdTime.Data <= mSpec.MovingSpeed.ReclinerBwd.Max))
        //                            mData.ReclineFwdTime.Result = RESULT.PASS;
        //                        else mData.ReclineFwdTime.Result = RESULT.REJECT;

        //                        fpSpread1.Sheets[0].Cells[5, 6].Text = mData.ReclineFwdTime.Data.ToString("0.0");
        //                        fpSpread1.Sheets[0].Cells[5, 7].Text = mData.ReclineFwdTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                        fpSpread1.Sheets[0].Cells[5, 7].ForeColor = mData.ReclineFwdTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                    }

        //                    if (NotImsReclineSpeedCheckFlag == false)
        //                    {
        //                        ComF.timedelay(200);
        //                        NotImsReclineSpeedCheckFlag = true;
        //                        ReclineFwdOnOff = true;
        //                        SpeedCheckStartTime[1] = ComF.timeGetTimems();
        //                    }
        //                    else
        //                    {
        //                        NotImsReclineSpeedCheckEnd = true;
        //                    }
        //                }
        //            }
        //        }

        //        if (NotImsTiltSpeedCheckEnd == false)
        //        {
        //            if (NotImsTiltSpeedCheckFlag == false)
        //                AdData3 = IOPort.ADRead[mSpec.PinMap.TiltDn.PinNo];
        //            else AdData3 = IOPort.ADRead[mSpec.PinMap.TiltUp.PinNo];

        //            if (1000 <= (ComF.timeGetTimems() - SpeedCheckStartTime[1]))
        //            {
        //                bool Flag = false;

        //                if (comboBox4.SelectedItem.ToString() == "IMS")
        //                {
        //                    if ((AdData3 <= mSpec.TiltLimitCurr) || ((mSpec.TiltLimitTime * 1000) <= (ComF.timeGetTimems() - SpeedCheckStartTime[1]))) Flag = true;
        //                }
        //                else
        //                {
        //                    if ((mSpec.TiltLimitCurr <= AdData3) || ((mSpec.TiltLimitTime * 1000) <= (ComF.timeGetTimems() - SpeedCheckStartTime[1]))) Flag = true;
        //                }
        //                if (Flag == true)
        //                {
        //                    if (NotImsTiltSpeedCheckFlag == false)
        //                        TiltDnOnOff = false;
        //                    else TiltUpOnOff = false;

        //                    float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[1]) / 1000F;

        //                    if (NotImsTiltSpeedCheckFlag == false)
        //                    {
        //                        mData.TiltDnTime.Test = true;
        //                        mData.TiltDnTime.Data = (mSpec.MovingStroke.Tilt / EndTime) * 1F;
        //                        if ((mSpec.MovingSpeed.TiltDn.Min <= mData.TiltDnTime.Data) && (mData.TiltDnTime.Data <= mSpec.MovingSpeed.TiltDn.Max))
        //                            mData.TiltDnTime.Result = RESULT.PASS;
        //                        else mData.TiltDnTime.Result = RESULT.REJECT;

        //                        fpSpread1.Sheets[0].Cells[8, 6].Text = mData.TiltDnTime.Data.ToString("0.0");
        //                        fpSpread1.Sheets[0].Cells[8, 7].Text = mData.TiltDnTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                        fpSpread1.Sheets[0].Cells[8, 7].ForeColor = mData.TiltDnTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                    }
        //                    else
        //                    {
        //                        mData.TiltUpTime.Test = true;
        //                        mData.TiltUpTime.Data = (mSpec.MovingStroke.Tilt / EndTime) * 1F;
        //                        if ((mSpec.MovingSpeed.TiltUp.Min <= mData.TiltUpTime.Data) && (mData.TiltUpTime.Data <= mSpec.MovingSpeed.TiltUp.Max))
        //                            mData.TiltUpTime.Result = RESULT.PASS;
        //                        else mData.TiltUpTime.Result = RESULT.REJECT;

        //                        fpSpread1.Sheets[0].Cells[7, 6].Text = mData.TiltUpTime.Data.ToString("0.0");
        //                        fpSpread1.Sheets[0].Cells[7, 7].Text = mData.TiltUpTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                        fpSpread1.Sheets[0].Cells[7, 7].ForeColor = mData.TiltUpTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                    }

        //                    if (NotImsTiltSpeedCheckFlag == false)
        //                    {
        //                        ComF.timedelay(200);
        //                        NotImsTiltSpeedCheckFlag = true;
        //                        TiltUpOnOff = true;
        //                        SpeedCheckStartTime[2] = ComF.timeGetTimems();
        //                    }
        //                    else
        //                    {
        //                        NotImsTiltSpeedCheckEnd = true;
        //                    }
        //                }
        //            }
        //        }

        //        if (NotImsHeightSpeedCheckEnd == false)
        //        {
        //            if (NotImsHeightSpeedCheckFlag == false)
        //                AdData4 = IOPort.ADRead[mSpec.PinMap.HeightDn.PinNo];
        //            else AdData4 = IOPort.ADRead[mSpec.PinMap.HeightUp.PinNo];

        //            if (1000 <= (ComF.timeGetTimems() - SpeedCheckStartTime[2]))
        //            {
        //                bool Flag = false;

        //                if (comboBox4.SelectedItem.ToString() == "IMS")
        //                {
        //                    if ((AdData4 <= mSpec.HeightLimitCurr) || ((mSpec.HeightLimitTime * 1000) <= (ComF.timeGetTimems() - SpeedCheckStartTime[2]))) Flag = true;
        //                }
        //                else
        //                {
        //                    if (NotImsHeightSpeedCheckFlag == true)
        //                    {
        //                        if ((mSpec.HeightLimitCurr <= AdData4) || ((mSpec.HeightLimitTime * 1000) <= (ComF.timeGetTimems() - SpeedCheckStartTime[2]))) Flag = true;
        //                    }
        //                    else
        //                    {
        //                        if ((5 <= AdData4) || ((mSpec.HeightLimitTime * 1000) <= (ComF.timeGetTimems() - SpeedCheckStartTime[2]))) Flag = true;
        //                    }
        //                }
        //                if (Flag == true)
        //                {
        //                    if (NotImsHeightSpeedCheckFlag == false)
        //                        HeightDnOnOff = false;
        //                    else HeightUpOnOff = false;

        //                    float EndTime = (ComF.timeGetTimems() - SpeedCheckStartTime[2]) / 1000F;

        //                    if (NotImsHeightSpeedCheckFlag == false)
        //                    {
        //                        mData.HeightDnTime.Test = true;
        //                        mData.HeightDnTime.Data = (mSpec.MovingStroke.Height / EndTime) * 1F;
        //                        if ((mSpec.MovingSpeed.HeightDn.Min <= mData.HeightDnTime.Data) && (mData.HeightDnTime.Data <= mSpec.MovingSpeed.HeightDn.Max))
        //                            mData.HeightDnTime.Result = RESULT.PASS;
        //                        else mData.HeightDnTime.Result = RESULT.REJECT;

        //                        fpSpread1.Sheets[0].Cells[10, 6].Text = mData.HeightDnTime.Data.ToString("0.0");
        //                        fpSpread1.Sheets[0].Cells[10, 7].Text = mData.HeightDnTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                        fpSpread1.Sheets[0].Cells[10, 7].ForeColor = mData.HeightDnTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                    }
        //                    else
        //                    {
        //                        mData.HeightUpTime.Test = true;
        //                        mData.HeightUpTime.Data = (mSpec.MovingStroke.Height / EndTime) * 1F;
        //                        if ((mSpec.MovingSpeed.HeightUp.Min <= mData.HeightUpTime.Data) && (mData.HeightUpTime.Data <= mSpec.MovingSpeed.HeightUp.Max))
        //                            mData.HeightUpTime.Result = RESULT.PASS;
        //                        else mData.HeightUpTime.Result = RESULT.REJECT;

        //                        fpSpread1.Sheets[0].Cells[9, 6].Text = mData.HeightUpTime.Data.ToString("0.0");
        //                        fpSpread1.Sheets[0].Cells[9, 7].Text = mData.HeightUpTime.Result == RESULT.PASS ? "O.K" : "N.G";
        //                        fpSpread1.Sheets[0].Cells[9, 7].ForeColor = mData.HeightUpTime.Result == RESULT.PASS ? Color.Green : Color.Red;
        //                    }

        //                    if (NotImsHeightSpeedCheckFlag == false)
        //                    {
        //                        ComF.timedelay(200);
        //                        NotImsHeightSpeedCheckFlag = true;
        //                        HeightUpOnOff = true;
        //                        SpeedCheckStartTime[3] = ComF.timeGetTimems();
        //                    }
        //                    else
        //                    {
        //                        NotImsHeightSpeedCheckEnd = true;
        //                    }
        //                }
        //            }
        //        }
        //    }

        //    if ((NotImsSlideSpeedCheckEnd == true) && (NotImsTiltSpeedCheckEnd == true) && (NotImsHeightSpeedCheckEnd == true) && (NotImsReclineSpeedCheckEnd == true))
        //    {
        //        SpecOutputFlag = false;
        //        Step++;
        //    }

        //    return;
        //}
        private void SoundCheck(short Pos)
        {
            if (CheckItem.Sound == false) return;

            

            if (Pos == SoundLevePos.SLIDE_FWD)
            {
                if (mData.SoundSlide.Test == true)
                {
                    if (mSpec.Sound.StartMax < mData.SoundSlide.StartData)
                        mData.SoundSlide.StartResult = RESULT.REJECT;
                    else mData.SoundSlide.StartResult = RESULT.PASS;

                    if (mSpec.Sound.RunMax < mData.SoundSlide.RunData)
                        mData.SoundSlide.RunResult = RESULT.REJECT;
                    else mData.SoundSlide.RunResult = RESULT.PASS;
                }
            }
            else if (Pos == SoundLevePos.RECLINE_BWD)
            {
                if (mData.SoundRecline.Test == true)
                {
                    if (mSpec.Sound.StartMax < mData.SoundRecline.StartData)
                        mData.SoundRecline.StartResult = RESULT.REJECT;
                    else mData.SoundRecline.StartResult = RESULT.PASS;

                    if (mSpec.Sound.RunMax < mData.SoundRecline.RunData)
                        mData.SoundRecline.RunResult = RESULT.REJECT;
                    else mData.SoundRecline.RunResult = RESULT.PASS;
                }
            }
            else if (Pos == SoundLevePos.TILT_UP)
            {
                if (mData.SoundTilt.Test == true)
                {
                    if (mSpec.Sound.StartMax < mData.SoundTilt.StartData)
                        mData.SoundTilt.StartResult = RESULT.REJECT;
                    else mData.SoundTilt.StartResult = RESULT.PASS;

                    if (mSpec.Sound.RunMax < mData.SoundTilt.RunData)
                        mData.SoundTilt.RunResult = RESULT.REJECT;
                    else mData.SoundTilt.RunResult = RESULT.PASS;
                }
            }
            else if (Pos == SoundLevePos.HEIGHT_UP)
            {
                if (mData.SoundHeight.Test == true)
                {
                    if (mSpec.Sound.StartMax < mData.SoundHeight.StartData)
                        mData.SoundHeight.StartResult = RESULT.REJECT;
                    else mData.SoundHeight.StartResult = RESULT.PASS;

                    if (mSpec.Sound.RunMax < mData.SoundHeight.RunData)
                        mData.SoundHeight.RunResult = RESULT.REJECT;
                    else mData.SoundHeight.RunResult = RESULT.PASS;
                }
            }            
            else if ((Pos == SoundLevePos.LUMBER_FWD) || (Pos == SoundLevePos.LUMBER_UP))
            {
                if ((CheckItem.PSeat12Way == true) || (CheckItem.PSeat10Way == true))
                {
                    if (Pos == SoundLevePos.LUMBER_FWD)
                    {
                        if (mData.SoundLumberFwdBwd.Test == true)
                        {
                            if (mSpec.Sound.StartMax < mData.SoundLumberFwdBwd.StartData)
                                mData.SoundLumberFwdBwd.StartResult = RESULT.REJECT;
                            else mData.SoundLumberFwdBwd.StartResult = RESULT.PASS;

                            if (mSpec.Sound.RunMax < mData.SoundLumberFwdBwd.RunData)
                                mData.SoundLumberFwdBwd.RunResult = RESULT.REJECT;
                            else mData.SoundLumberFwdBwd.RunResult = RESULT.PASS;
                        }
                    }
                    else
                    {
                        if (CheckItem.PSeat12Way == true)
                        {
                            if (mData.SoundLumberUpDn.Test == true)
                            {
                                if (mSpec.Sound.StartMax < mData.SoundLumberUpDn.StartData)
                                    mData.SoundLumberUpDn.StartResult = RESULT.REJECT;
                                else mData.SoundLumberUpDn.StartResult = RESULT.PASS;

                                if (mSpec.Sound.RunMax < mData.SoundLumberUpDn.RunData)
                                    mData.SoundLumberUpDn.RunResult = RESULT.REJECT;
                                else mData.SoundLumberUpDn.RunResult = RESULT.PASS;
                            }
                        }
                    }
                }
            }

            if (Pos == SoundLevePos.SLIDE_FWD)
            {
                fpSpread1.ActiveSheet.Cells[17, 6].Text = mData.SoundSlide.StartData.ToString("0");
                fpSpread1.ActiveSheet.Cells[17, 7].Text = (mData.SoundSlide.StartResult == RESULT.PASS) ? "O.K" : "N.G";
                fpSpread1.ActiveSheet.Cells[17, 7].ForeColor = (mData.SoundSlide.StartResult == RESULT.PASS) ? Color.Lime : Color.Red;

                fpSpread1.ActiveSheet.Cells[18, 6].Text = mData.SoundSlide.RunData.ToString("0");
                fpSpread1.ActiveSheet.Cells[18, 7].Text = (mData.SoundSlide.RunResult == RESULT.PASS) ? "O.K" : "N.G";
                fpSpread1.ActiveSheet.Cells[18, 7].ForeColor = (mData.SoundSlide.RunResult == RESULT.PASS) ? Color.Lime : Color.Red;
            }
            else if (Pos == SoundLevePos.RECLINE_BWD)
            {
                fpSpread1.ActiveSheet.Cells[19, 6].Text = mData.SoundRecline.StartData.ToString("0");
                fpSpread1.ActiveSheet.Cells[19, 7].Text = (mData.SoundRecline.StartResult == RESULT.PASS) ? "O.K" : "N.G";
                fpSpread1.ActiveSheet.Cells[19, 7].ForeColor = (mData.SoundRecline.StartResult == RESULT.PASS) ? Color.Lime : Color.Red;

                fpSpread1.ActiveSheet.Cells[20, 6].Text = mData.SoundRecline.RunData.ToString("0");
                fpSpread1.ActiveSheet.Cells[20, 7].Text = (mData.SoundRecline.RunResult == RESULT.PASS) ? "O.K" : "N.G";
                fpSpread1.ActiveSheet.Cells[20, 7].ForeColor = (mData.SoundRecline.RunResult == RESULT.PASS) ? Color.Lime : Color.Red;
            }
            else if (Pos == SoundLevePos.TILT_UP)
            {
                fpSpread1.ActiveSheet.Cells[21, 6].Text = mData.SoundTilt.StartData.ToString("0");
                fpSpread1.ActiveSheet.Cells[21, 7].Text = (mData.SoundTilt.StartResult == RESULT.PASS) ? "O.K" : "N.G";
                fpSpread1.ActiveSheet.Cells[21, 7].ForeColor = (mData.SoundTilt.StartResult == RESULT.PASS) ? Color.Lime : Color.Red;

                fpSpread1.ActiveSheet.Cells[22, 6].Text = mData.SoundTilt.RunData.ToString("0");
                fpSpread1.ActiveSheet.Cells[22, 7].Text = (mData.SoundTilt.RunResult == RESULT.PASS) ? "O.K" : "N.G";
                fpSpread1.ActiveSheet.Cells[22, 7].ForeColor = (mData.SoundTilt.RunResult == RESULT.PASS) ? Color.Lime : Color.Red;
            }
            else if (Pos == SoundLevePos.HEIGHT_UP)
            {
                fpSpread1.ActiveSheet.Cells[23, 6].Text = mData.SoundHeight.StartData.ToString("0");
                fpSpread1.ActiveSheet.Cells[23, 7].Text = (mData.SoundHeight.StartResult == RESULT.PASS) ? "O.K" : "N.G";
                fpSpread1.ActiveSheet.Cells[23, 7].ForeColor = (mData.SoundHeight.StartResult == RESULT.PASS) ? Color.Lime : Color.Red;

                fpSpread1.ActiveSheet.Cells[24, 6].Text = mData.SoundHeight.RunData.ToString("0");
                fpSpread1.ActiveSheet.Cells[24, 7].Text = (mData.SoundHeight.RunResult == RESULT.PASS) ? "O.K" : "N.G";
                fpSpread1.ActiveSheet.Cells[24, 7].ForeColor = (mData.SoundHeight.RunResult == RESULT.PASS) ? Color.Lime : Color.Red;
            }
            else if ((Pos == SoundLevePos.LUMBER_FWD) || (Pos == SoundLevePos.LUMBER_UP))
            {
                if ((CheckItem.PSeat12Way == true) || (CheckItem.PSeat10Way == true))
                {
                    if (Pos == SoundLevePos.LUMBER_FWD)
                    {
                        fpSpread1.ActiveSheet.Cells[25, 6].Text = mData.SoundLumberFwdBwd.StartData.ToString("0");
                        fpSpread1.ActiveSheet.Cells[25, 7].Text = (mData.SoundLumberFwdBwd.StartResult == RESULT.PASS) ? "O.K" : "N.G";
                        fpSpread1.ActiveSheet.Cells[25, 7].ForeColor = (mData.SoundLumberFwdBwd.StartResult == RESULT.PASS) ? Color.Lime : Color.Red;

                        fpSpread1.ActiveSheet.Cells[26, 6].Text = mData.SoundLumberFwdBwd.RunData.ToString("0");
                        fpSpread1.ActiveSheet.Cells[26, 7].Text = (mData.SoundLumberFwdBwd.RunResult == RESULT.PASS) ? "O.K" : "N.G";
                        fpSpread1.ActiveSheet.Cells[26, 7].ForeColor = (mData.SoundLumberFwdBwd.RunResult == RESULT.PASS) ? Color.Lime : Color.Red;
                    }
                    else
                    {
                        if (CheckItem.PSeat12Way == true)
                        {
                            fpSpread1.ActiveSheet.Cells[27, 6].Text = mData.SoundLumberUpDn.StartData.ToString("0");
                            fpSpread1.ActiveSheet.Cells[27, 7].Text = (mData.SoundLumberUpDn.StartResult == RESULT.PASS) ? "O.K" : "N.G";
                            fpSpread1.ActiveSheet.Cells[27, 7].ForeColor = (mData.SoundLumberUpDn.StartResult == RESULT.PASS) ? Color.Lime : Color.Red;

                            fpSpread1.ActiveSheet.Cells[28, 6].Text = mData.SoundLumberUpDn.RunData.ToString("0");
                            fpSpread1.ActiveSheet.Cells[28, 7].Text = (mData.SoundLumberUpDn.RunResult == RESULT.PASS) ? "O.K" : "N.G";
                            fpSpread1.ActiveSheet.Cells[28, 7].ForeColor = (mData.SoundLumberUpDn.RunResult == RESULT.PASS) ? Color.Lime : Color.Red;
                        }
                    }
                }
                //SpreadPositionSelect(28, 7);
                //SpecOutputFlag = false;
                //Step++;
            }
            return;
        }

        private bool SlideDeliveryPosEndFlag { get; set; }
        private bool SlideDeliveryPosFlag { get; set; }
        private short DeliveryPosStep = 0;
        private bool LumberDeliveryMovePos { get; set; }
        private void DeliveryPosMoveing()
        {
            if (DeliveryPosStep == 0)
            {
                if (SpecOutputFlag == false)
                {
                    SlideDeliveryPosEndFlag = false;
                    if (CheckItem.PSeat12Way == true)
                    {
                        LumberBwd = true;
                        LumberDn = true;
                        LumberDeliveryMovePos = false;
                    }
                    else if (CheckItem.PSeat10Way == true)
                    {
                        LumberBwd = true;
                        LumberDn = false;
                        LumberDeliveryMovePos = false;
                    }
                    else
                    {
                        LumberDeliveryMovePos = true;
                    }

                    ReclineMotorMoveEndFlag = false;
                    SpecOutputFlag = true;

                    if (mSpec.DeliveryPos.Slide == 0)//Fwd
                    {
                        //SlideFwdOnOff = true;
                        SlideMotorMoveEndFlag = true;
                    }
                    else if (mSpec.DeliveryPos.Slide == 1)//Mid
                    {
                        if (IOPort.GetSlideMidSensor == false)
                        {
                            SlideBwdOnOff = true;
                            SlideDeliveryPosFlag = false;
                        }
                        else
                        {
                            SlideFwdOnOff = true;
                            SlideDeliveryPosFlag = true;
                        }
                        SlideMotorMoveEndFlag = false;
                    }
                    else
                    {
                        SlideBwdOnOff = true;
                        SlideMotorMoveEndFlag = false;
                    }

                    //if (comboBox4.SelectedItem.ToString() == "IMS")
                    //{
                    //    if (mSpec.DeliveryPos.Tilt == 0)//Up
                    //    {
                    //        TiltUpOnOff = true;
                    //        TiltMotorMoveEndFlag = false;
                    //    }
                    //    else
                    //    {
                    //        TiltMotorMoveEndFlag = true;
                    //    }

                    //    if (mSpec.DeliveryPos.Height == 0)//Up
                    //    {
                    //        HeightUpOnOff = true;
                    //        HeightMotorMoveEndFlag = false;
                    //    }
                    //    else
                    //    {
                    //        HeightMotorMoveEndFlag = true;
                    //    }
                    //}
                    //else
                    //{
                    //    if (mSpec.DeliveryPos.Tilt == 0)//Up
                    //    {
                    //        TiltMotorMoveEndFlag = true;
                    //    }
                    //    else
                    //    {
                    //        TiltDnOnOff = true;
                    //        TiltMotorMoveEndFlag = false;
                    //    }

                    //    if (mSpec.DeliveryPos.Height == 0)//Up
                    //    {
                    //        //HeightUpOnOff = true;
                    //        HeightMotorMoveEndFlag = true;
                    //    }
                    //    else
                    //    {
                    //        HeightDnOnOff = true;
                    //        HeightMotorMoveEndFlag = false;
                    //    }
                    //}

                    if (mSpec.DeliveryPos.Tilt == 0)//Up
                    {
                        TiltUpOnOff = true;
                        TiltMotorMoveEndFlag = false;
                    }
                    else
                    {
                        TiltDnOnOff = true;
                        TiltMotorMoveEndFlag = false;
                    }

                    if (mSpec.DeliveryPos.Height == 0)//Up
                    {
                        HeightUpOnOff = true;
                        HeightMotorMoveEndFlag = false;
                    }
                    else
                    {
                        HeightDnOnOff = true;
                        HeightMotorMoveEndFlag = false;
                    }
                    StepTimeFirst = ComF.timeGetTimems();
                    StepTimeLast = ComF.timeGetTimems();

                    label5.Text = "출하 위치로 이동 중 입니다.";
                }
                else
                {
                    StepTimeLast = ComF.timeGetTimems();

                    //float AdData;
                    if (SlideMotorMoveEndFlag == false)
                    {
                        if (mSpec.DeliveryPos.Slide == 1)//Mid
                        {
                            //납입위치 센서 감지 모드이면 
                            if (SlideDeliveryPosFlag == false) //Bwd
                            {
                                if (IOPort.GetSlideMidSensor == true)
                                {
                                    //ComF.timedelay(300);
                                    SlideDeliveryPosEndFlag = true;
                                    //센서 검지도면 정지
                                    SlideBwdOnOff = false;
                                    SlideMotorMoveEndFlag = true;
                                }
                            }
                            else
                            {
                                //납입 취리를 이동시키기 위해 센서가 꺼지는 시점까지 전진 시킨다.
                                //Fwd
                                ComF.timedelay(1000);
                                SlideFwdOnOff = false;
                                SlideDeliveryPosFlag = false;
                                ComF.timedelay(500);
                                SlideBwdOnOff = true;
                            }
                        }

                        if (mSpec.DeliveryPos.Slide != 1)
                        {
                            if (comboBox4.SelectedItem.ToString() == "IMS")
                            {
                                if (1000 <= (StepTimeLast - StepTimeFirst))
                                {
                                    if (pMeter.GetPSeat < PSEAT_OFF_CURRENT)
                                    {
                                        if (mSpec.DeliveryPos.Slide == 0)//Fwd
                                            SlideFwdOnOff = false;
                                        else SlideBwdOnOff = false;
                                        SlideMotorMoveEndFlag = true;
                                    }
                                }
                            }
                            else
                            {
                                //if (mSpec.DeliveryPos.Slide == 0)
                                //    AdData = IOPort.ADRead[mSpec.PinMap.SlideFWD.PinNo];
                                //else AdData = IOPort.ADRead[mSpec.PinMap.SlideBWD.PinNo];

                                //if (mSpec.SlideLimitCurr <= AdData)
                                //{
                                //    if (mSpec.DeliveryPos.Slide == 0)//Fwd
                                //        SlideFwdOnOff = false;
                                //    else SlideBwdOnOff = false;
                                //    SlideMotorMoveEndFlag = true;
                                //}
                                //else if (IOPort.GetSlideMidSensor == true)
                                //{
                                //    if (mSpec.DeliveryPos.Slide == 0)//Fwd
                                //        SlideFwdOnOff = false;
                                //    else SlideBwdOnOff = false;
                                //    SlideMotorMoveEndFlag = true;
                                //}
                            }
                        }
                    }

                    if (TiltMotorMoveEndFlag == false)
                    {
                        if (comboBox4.SelectedItem.ToString() == "IMS")
                        {
                            if (1000 <= (StepTimeLast - StepTimeFirst))
                            {
                                if (pMeter.GetPSeat < PSEAT_OFF_CURRENT)
                                {
                                    if (mSpec.DeliveryPos.Tilt == 0)//Up
                                        TiltUpOnOff = false;
                                    else TiltDnOnOff = false;
                                    TiltMotorMoveEndFlag = true;
                                }
                            }
                        }
                        else
                        {
                            //if (mSpec.DeliveryPos.Tilt == 0)
                            //    AdData = IOPort.ADRead[mSpec.PinMap.TiltUp.PinNo];
                            //else AdData = IOPort.ADRead[mSpec.PinMap.TiltDn.PinNo];

                            //if (mSpec.TiltLimitCurr <= AdData)
                            //{
                            //    if (mSpec.DeliveryPos.Tilt == 0)//Up
                            //        TiltUpOnOff = false;
                            //    else TiltDnOnOff = false;
                            //    TiltMotorMoveEndFlag = true;
                            //}
                        }
                    }
                    if (HeightMotorMoveEndFlag == false)
                    {
                        if (comboBox4.SelectedItem.ToString() == "IMS")
                        {
                            if (1000 <= (StepTimeLast - StepTimeFirst))
                            {
                                if (pMeter.GetPSeat < PSEAT_OFF_CURRENT)
                                {
                                    if (mSpec.DeliveryPos.Height == 0)//Up
                                        HeightUpOnOff = false;
                                    else HeightDnOnOff = false;
                                    HeightMotorMoveEndFlag = true;
                                }
                            }
                        }
                        else
                        {
                            //bool Flag = false;
                            //if (mSpec.DeliveryPos.Height == 0)
                            //    AdData = IOPort.ADRead[mSpec.PinMap.HeightUp.PinNo];
                            //else AdData = IOPort.ADRead[mSpec.PinMap.HeightDn.PinNo];

                            //if (mSpec.DeliveryPos.Height == 0)
                            //{
                            //    if (mSpec.HeightLimitCurr <= AdData) Flag = true;
                            //}
                            //else
                            //{
                            //    if (5 <= AdData) Flag = true;
                            //}
                            //if (Flag == true)
                            //{
                            //    if (mSpec.DeliveryPos.Height == 0)//Up
                            //        HeightUpOnOff = false;
                            //    else HeightDnOnOff = false;
                            //    HeightMotorMoveEndFlag = true;
                            //}
                        }
                    }


                    if (SlideMotorMoveEndFlag == false)
                    {
                        if ((mSpec.SlideLimitTime * 1000F) <= (StepTimeLast - StepTimeFirst))
                        {
                            if (mSpec.DeliveryPos.Slide == 0)//Fwd
                                SlideFwdOnOff = false;
                            else SlideBwdOnOff = false;
                            SlideMotorMoveEndFlag = true;
                        }
                    }
                    if (TiltMotorMoveEndFlag == false)
                    {
                        if ((mSpec.TiltLimitTime * 1000F) <= (StepTimeLast - StepTimeFirst))
                        {
                            if (mSpec.DeliveryPos.Tilt == 0)//Up
                                TiltUpOnOff = false;
                            else TiltDnOnOff = false;
                            TiltMotorMoveEndFlag = true;
                        }
                    }
                    if (HeightMotorMoveEndFlag == false)
                    {
                        if ((mSpec.HeightLimitTime * 1000F) <= (StepTimeLast - StepTimeFirst))
                        {
                            if (mSpec.DeliveryPos.Height == 0)//Up
                                HeightUpOnOff = false;
                            else HeightDnOnOff = false;
                            HeightMotorMoveEndFlag = true;
                        }
                    }

                    if (LumberDeliveryMovePos == false)
                    {
                        if ((mSpec.LumberLimitTime * 1000F) <= (StepTimeLast - StepTimeFirst))
                        {
                            LumberBwd = false;
                            LumberDn = false;
                            LumberDeliveryMovePos = true;
                        }
                    }

                    if ((SlideMotorMoveEndFlag == true) && (TiltMotorMoveEndFlag == true) && (HeightMotorMoveEndFlag == true) && (LumberDeliveryMovePos == true))
                    {
                        SpecOutputFlag = false;
                        DeliveryPosStep = 1;
                    }
                }
            }
            else
            {
                if (SpecOutputFlag == false)
                {
                    SpecOutputFlag = true;

                    if (mSpec.DeliveryPos.Recliner == 0)//Fwd
                    {
                        ReclineMotorMoveEndFlag = false;
                        ReclineFwdOnOff = true;
                    }
                    else
                    {
                        ReclineBwdOnOff = true;
                        ReclineMotorMoveEndFlag = false;
                    }

                    StepTimeFirst = ComF.timeGetTimems();
                    StepTimeLast = ComF.timeGetTimems();
                }
                else
                {
                    //float AdData;

                    StepTimeLast = ComF.timeGetTimems();

                    //if (mSpec.DeliveryPos.Recliner == 0)//Fwd
                    //{
                    //    if (IOPort.GetReclineMidSensor == false)
                    //    {
                    //        ReclineFwdOnOff = false;
                    //        ReclineMotorMoveEndFlag = true;
                    //    }
                    //}

                    if (ReclineMotorMoveEndFlag == false)
                    {
                        if ((mSpec.ReclinerLimitTime * 1000F) <= (StepTimeLast - StepTimeFirst))
                        {
                            if (mSpec.DeliveryPos.Recliner == 0)//Fwd
                                ReclineFwdOnOff = false;
                            else ReclineFwdOnOff = false;
                            ReclineMotorMoveEndFlag = true;
                        }

                        if (1000 <= (StepTimeLast - StepTimeFirst))
                        {
                            if (comboBox4.SelectedItem.ToString() == "IMS")
                            {
                                if (pMeter.GetPSeat < PSEAT_OFF_CURRENT)
                                {
                                    if (mSpec.DeliveryPos.Recliner == 0)//Fwd
                                        ReclineFwdOnOff = false;
                                    else ReclineFwdOnOff = false;

                                    ReclineMotorMoveEndFlag = true;
                                }
                            }
                            else
                            {
                                bool Flag = false;
                                float AdData;

                                //if (mSpec.DeliveryPos.Recliner == 0)
                                //    AdData = IOPort.ADRead[mSpec.PinMap.ReclineFWD.PinNo];
                                //else AdData = IOPort.ADRead[mSpec.PinMap.ReclineBWD.PinNo];

                                AdData = pMeter.GetPSeat;

                                if (mSpec.DeliveryPos.Recliner == 0)
                                {
                                    if (mSpec.ReclinerLimitCurr <= AdData) Flag = true;
                                }
                                else
                                {
                                    if (5 <= AdData) Flag = true;
                                }
                                if (Flag == true)
                                {
                                    if (mSpec.DeliveryPos.Recliner == 0)//Fwd
                                        ReclineFwdOnOff = false;
                                    else ReclineFwdOnOff = false;

                                    ReclineMotorMoveEndFlag = true;

                                    //ReclineMotorMoveEndFlag = true;

                                    //if (mSpec.DeliveryPos.Recliner == 0)//Fwd
                                    //{
                                    //    if (IOPort.GetReclineMidSensor == false)
                                    //    {
                                    //        ReclineMotorMoveEndFlag = true;
                                    //    }
                                    //}
                                }
                            }
                        }
                    }
                }
            }

            if ((SlideMotorMoveEndFlag == true) && (TiltMotorMoveEndFlag == true) && (HeightMotorMoveEndFlag == true) && (ReclineMotorMoveEndFlag == true))
            {
                //if (mSpec.DeliveryPos.Recliner == 0)//Fwd
                //{
                //    if (IOPort.GetReclineMidSensor == false)
                //    {                   
                //        panel7.Visible = true;
                //        panel7.Parent = this;
                //        panel7.Location = new Point(this.Left + (this.Width / 2) - (panel7.Width / 2), 150);
                //    }
                //}
                //if (mSpec.DeliveryPos.Slide == 1)//Mid
                //{
                //    if (IOPort.GetSlideMidSensor == false)
                //    {
                //        panel7.Visible = true;
                //        panel7.Parent = this;
                //        panel7.Location = new Point(this.Left + (this.Width / 2) - (panel7.Width / 2), 150);
                //    }
                //}
                SpecOutputFlag = false;
                Step++;
            }
            return;
        }

        private void ResultCheck()
        {
            mData.Result = RESULT.PASS;
            if ((mData.Height.Test == true) && (mData.Height.Result == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.Tilt.Test == true) && (mData.Tilt.Result == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.Slide.Test == true) && (mData.Slide.Result == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.Recline.Test == true) && (mData.Recline.Result == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.LumberFwdBwd.Test == true) && (mData.LumberFwdBwd.Result == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.LumberUpDn.Test == true) && (mData.LumberUpDn.Result == RESULT.REJECT)) mData.Result = RESULT.REJECT;

            if ((mData.SlideFwdTime.Test == true) && (mData.SlideFwdTime.Result == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.SlideBwdTime.Test == true) && (mData.SlideBwdTime.Result == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.ReclineFwdTime.Test == true) && (mData.ReclineFwdTime.Result == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.ReclineBwdTime.Test == true) && (mData.ReclineBwdTime.Result == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if (((mData.TiltUpTime.Test == true) && mData.TiltUpTime.Result == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.TiltDnTime.Test == true) && (mData.TiltDnTime.Result == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.HeightUpTime.Test == true) && (mData.HeightUpTime.Result == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.HeightDnTime.Test == true) && (mData.HeightDnTime.Result == RESULT.REJECT)) mData.Result = RESULT.REJECT;

            if ((mData.Ims.Test == true) && (mData.Ims.Result == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.SoundSlide.Test == true) && (mData.SoundSlide.StartResult == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.SoundSlide.Test == true) && (mData.SoundSlide.RunResult == RESULT.REJECT)) mData.Result = RESULT.REJECT;

            if ((mData.SoundRecline.Test == true) && (mData.SoundRecline.StartResult == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.SoundRecline.Test == true) && (mData.SoundRecline.RunResult == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.SoundTilt.Test == true) && (mData.SoundTilt.StartResult == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.SoundTilt.Test == true) && (mData.SoundTilt.RunResult == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.SoundHeight.Test == true) && (mData.SoundHeight.StartResult == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.SoundHeight.Test == true) && (mData.SoundHeight.RunResult == RESULT.REJECT)) mData.Result = RESULT.REJECT;

            if ((mData.SoundLumberFwdBwd.Test == true) && (mData.SoundLumberFwdBwd.StartResult == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.SoundLumberFwdBwd.Test == true) && (mData.SoundLumberFwdBwd.RunResult == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.SoundLumberUpDn.Test == true) && (mData.SoundLumberUpDn.StartResult == RESULT.REJECT)) mData.Result = RESULT.REJECT;
            if ((mData.SoundLumberUpDn.Test == true) && (mData.SoundLumberUpDn.RunResult == RESULT.REJECT)) mData.Result = RESULT.REJECT;


            if (mData.Result == RESULT.REJECT)
            {
                //label16.Text = "불량";
                //label16.ForeColor = Color.Red;

                IOPort.YellowLampOnOff = false;
                IOPort.RedLampOnOff = true;
                IOPort.GreenLampOnOff = false;
                Infor.NgCount++;

                IOPort.BuzzerOnOff = true;
                BuzzerRunFlag = true;
                BuzerOnOff = true;
                BuzzerLast = ComF.timeGetTimems();
                BuzzerFirst = ComF.timeGetTimems();
                BuzzerOnCount = 0;
            }
            else
            {
                //label16.Text = "양품";
                //label16.ForeColor = Color.Lime;

                IOPort.YellowLampOnOff = false;
                IOPort.RedLampOnOff = false;
                IOPort.GreenLampOnOff = true;
                Infor.OkCount++;

                IOPort.BuzzerOnOff = true;
                ComF.timedelay(500);
                IOPort.BuzzerOnOff = false;
            }
            if (mData.Result == RESULT.REJECT)
                label5.Text = "불량 제품 입니다.";
            else label5.Text = "양품 제품 입니다.";

            return;
        }

        private void SendTestData()
        {
            float StartSound = 0;
            float RunSound = 0;

            StartSound = Math.Max(StartSound, mData.SoundSlide.StartData);
            StartSound = Math.Max(StartSound, mData.SoundRecline.StartData);
            StartSound = Math.Max(StartSound, mData.SoundTilt.StartData);
            StartSound = Math.Max(StartSound, mData.SoundHeight.StartData);
            StartSound = Math.Max(StartSound, mData.SoundLumberFwdBwd.StartData);
            StartSound = Math.Max(StartSound, mData.SoundLumberUpDn.StartData);

            RunSound = Math.Max(RunSound, mData.SoundSlide.RunData);
            RunSound = Math.Max(RunSound, mData.SoundRecline.RunData);
            RunSound = Math.Max(RunSound, mData.SoundTilt.RunData);
            RunSound = Math.Max(RunSound, mData.SoundHeight.RunData);
            RunSound = Math.Max(RunSound, mData.SoundLumberFwdBwd.RunData);
            RunSound = Math.Max(RunSound, mData.SoundLumberUpDn.RunData);


            string s = PopCtrl.CrateData(
                Serial: Infor.TotalCount - Infor.NgCount,
                TestCount: Infor.TotalCount,
                Min_1: mData.Ims.Test == true ? 1 : -9999,
                Max_1: mData.Ims.Test == true ? 1 : -9999,
                Min_2: CheckItem.Slide == true ? (float)mSpec.Current.Slide.Min : -9999,
                Max_2: CheckItem.Slide == true ? (float)mSpec.Current.Slide.Max : -9999,
                Min_3: CheckItem.Tilt == true ? (float)mSpec.Current.Tilt.Min : -9999,
                Max_3: CheckItem.Tilt == true ? (float)mSpec.Current.Tilt.Max : -9999,
                Min_4: CheckItem.Height == true ? (float)mSpec.Current.Height.Min : -9999,
                Max_4: CheckItem.Height == true ? (float)mSpec.Current.Height.Max : -9999,
                Min_5: CheckItem.Recline == true ? (float)mSpec.Current.Recliner.Min : -9999,
                Max_5: CheckItem.Recline == true ? (float)mSpec.Current.Recliner.Max : -9999,
                Min_6: (mData.LumberFwdBwd.Test == true) ? (float)mSpec.Current.LumberFwdBwd.Min : -9999,
                Max_6: (mData.LumberFwdBwd.Test == true) ? (float)mSpec.Current.LumberFwdBwd.Min : -9999,
                Min_7: (mData.LumberUpDn.Test == true) ? (float)mSpec.Current.LumberUpDn.Min : -9999,
                Max_7: (mData.LumberUpDn.Test == true) ? (float)mSpec.Current.LumberUpDn.Max : -9999,
                Min_8: (CheckItem.Sound == true) ? 0 : -9999,
                Max_8: (CheckItem.Sound == true) ? mSpec.Sound.StartMax : -9999,
                Min_9: (CheckItem.Sound == true) ? 0 : -9999,
                Max_9: (CheckItem.Sound == true) ? mSpec.Sound.RunMax : -9999,
                Min_10: -9999,
                Max_10: -9999,
                Min_11: -9999,
                Max_11: -9999,
                Min_12: -9999,
                Max_12: -9999,
                Min_13: -9999,
                Max_13: -9999,
                Min_14: -9999,
                Max_14: -9999,
                Min_15: -9999,
                Max_15: -9999,
                Value1: (mData.Ims.Test == false) ? -9999 : (mData.Ims.Result == RESULT.REJECT ? 9 : 1),
                Value2: (mData.Slide.Test == false) ? -9999 : mData.Slide.Data,
                Value3: (mData.Tilt.Test == false) ? -9999 : mData.Tilt.Data,
                Value4: (mData.Height.Test == false) ? -9999 : mData.Height.Data,
                Value5: (mData.Recline.Test == false) ? -9999 : mData.Recline.Data,
                Value6: (mData.LumberFwdBwd.Test == false) ? -9999 : mData.LumberFwdBwd.Data,
                Value7: (mData.LumberUpDn.Test == false) ? -9999 : mData.LumberUpDn.Data,
                Value8: (mData.SoundSlide.Test == false) ? -9999 : StartSound,
                Value9: (mData.SoundSlide.Test == false) ? -9999 : RunSound,
                Value10: -9999,
                Value11: -9999,
                Value12: -9999,
                Value13: -9999,
                Value14: -9999,
                Value15: -9999,
                Result: mData.Result
                );

            if (IOPort.GetAuto == true)
            {
                //   if (PopCtrl.isServerConnection == true) PopCtrl.ServerSend = s;
                if (PopCtrl.isClientConnection == true) PopCtrl.Write = s + ",1";
            }
            else
            {
                if (PopCtrl.isClientConnection == true) PopCtrl.Write = s + ",2";
            }

            SaveLogData = "SENDING - " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
            SaveLogData = s;
            return;
        }


        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            if (IOPort.GetAuto == false) panel4.Visible = !panel4.Visible;
            if (panel4.Visible == true)
            {
                //panel4.Parent = this;
                panel4.BringToFront();
            }
            else
            {
                IOPort.SetPinConnection = false;
            }
            IOPort.FunctionIOInit();
            return;
        }

        private void imageButton1_Click(object sender, EventArgs e)
        {
            UserImageButton.ImageButton But = sender as UserImageButton.ImageButton;
            if (But.ButtonColor == Color.Black)
            {
                PowerOnOff = true;
                BattOnOff = true;
                But.ButtonColor = Color.Red;
            }
            else
            {
                PowerOnOff = false;
                BattOnOff = false;
                But.ButtonColor = Color.Black;
            }
            return;
        }

        private void ledArrow1_ValueChanged(object sender, Iocomp.Classes.ValueBooleanEventArgs e)
        {
            //Slide Fwd
            SlideFwdOnOff = e.ValueNew;
            return;
        }

        private void ledArrow2_ValueChanged(object sender, Iocomp.Classes.ValueBooleanEventArgs e)
        {
            //Slide Bwd
            SlideBwdOnOff = e.ValueNew;
            return;
        }

        private void ledArrow4_ValueChanged(object sender, Iocomp.Classes.ValueBooleanEventArgs e)
        {
            //Recline Fwd
            ReclineFwdOnOff = e.ValueNew;
            return;
        }

        private void ledArrow3_ValueChanged(object sender, Iocomp.Classes.ValueBooleanEventArgs e)
        {
            //Recline Bwd
            ReclineBwdOnOff = e.ValueNew;
            return;
        }

        private void ledArrow7_ValueChanged(object sender, Iocomp.Classes.ValueBooleanEventArgs e)
        {
            //Tilt Up
            TiltUpOnOff = e.ValueNew;
            return;
        }

        private void ledArrow8_ValueChanged(object sender, Iocomp.Classes.ValueBooleanEventArgs e)
        {
            //Tilt Down
            TiltDnOnOff = e.ValueNew;
            return;
        }

        private void ledArrow5_ValueChanged(object sender, Iocomp.Classes.ValueBooleanEventArgs e)
        {
            //Height Up
            HeightUpOnOff = e.ValueNew;
            return;
        }

        private void ledArrow6_ValueChanged(object sender, Iocomp.Classes.ValueBooleanEventArgs e)
        {
            //Height Dn
            HeightDnOnOff = e.ValueNew;
            return;
        }

        private void ledArrow11_ValueChanged(object sender, Iocomp.Classes.ValueBooleanEventArgs e)
        {
            //Lumber Fwd            
            LumberFwd = e.ValueNew;
            return;
        }

        private void ledArrow12_ValueChanged(object sender, Iocomp.Classes.ValueBooleanEventArgs e)
        {
            //Lumber Bwd
            LumberBwd = e.ValueNew;
            return;
        }

        private void ledArrow9_ValueChanged(object sender, Iocomp.Classes.ValueBooleanEventArgs e)
        {
            //Lumber Up
            LumberUp = e.ValueNew;
            return;
        }

        private void ledArrow10_ValueChanged(object sender, Iocomp.Classes.ValueBooleanEventArgs e)
        {
            //LumberBwd Down
            LumberDn = e.ValueNew;
            return;
        }

        private void ledArrow1_MouseDown(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = true;
            return;
        }

        private void ledArrow1_MouseUp(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = false;
            return;
        }

        private void ledArrow2_MouseDown(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = true;
            return;
        }

        private void ledArrow2_MouseUp(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = false;
            return;
        }

        private void ledArrow4_MouseDown(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = true;
            return;
        }

        private void ledArrow4_MouseUp(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = false;
            return;
        }

        private void ledArrow3_MouseDown(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = true;
            return;
        }

        private void ledArrow3_MouseUp(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = false;
            return;
        }

        private void ledArrow7_MouseDown(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = true;
            return;
        }

        private void ledArrow7_MouseUp(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = false;
            return;
        }

        private void ledArrow8_MouseDown(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = true;
            return;
        }

        private void ledArrow8_MouseUp(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = false;
            return;
        }

        private void ledArrow5_MouseDown(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = true;
            return;
        }

        private void ledArrow5_MouseUp(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = false;
            return;
        }

        private void ledArrow6_MouseDown(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = true;
            return;
        }

        private void ledArrow6_MouseUp(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = false;
            return;
        }

        private void ledArrow11_MouseDown(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = true;
            return;
        }

        private void ledArrow11_MouseUp(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = false;
            return;
        }

        private void ledArrow12_MouseDown(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = true;
            return;
        }

        private void ledArrow12_MouseUp(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = false;
            return;
        }

        private void ledArrow9_MouseDown(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = true;
            return;
        }

        private void ledArrow9_MouseUp(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = false;
            return;
        }

        private void ledArrow10_MouseDown(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = true;
            return;
        }

        private void ledArrow10_MouseUp(object sender, MouseEventArgs e)
        {
            Iocomp.Instrumentation.Professional.LedArrow LBut = sender as Iocomp.Instrumentation.Professional.LedArrow;
            LBut.Value.AsBoolean = false;
            return;
        }

        private void imageButton2_Click(object sender, EventArgs e)
        {
            if (RunningFlag == false)
            {
                if (IOPort.GetAuto == true)
                {
                    if (PopReadOk == true) StartSetting();
                }
                else
                {
                    StartSetting();
                }
            }
            return;
        }

        private void DisplayAD()
        {
            if (panel4.Visible == true)
            {
                if (fpSpread2.ActiveSheet.Cells[2, 1].Text != IOPort.ADRead[0].ToString("0.0")) fpSpread2.ActiveSheet.Cells[2, 1].Text = IOPort.ADRead[0].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[3, 1].Text != IOPort.ADRead[1].ToString("0.0")) fpSpread2.ActiveSheet.Cells[3, 1].Text = IOPort.ADRead[1].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[4, 1].Text != IOPort.ADRead[2].ToString("0.0")) fpSpread2.ActiveSheet.Cells[4, 1].Text = IOPort.ADRead[2].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[5, 1].Text != IOPort.ADRead[3].ToString("0.0")) fpSpread2.ActiveSheet.Cells[5, 1].Text = IOPort.ADRead[3].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[6, 1].Text != IOPort.ADRead[4].ToString("0.0")) fpSpread2.ActiveSheet.Cells[6, 1].Text = IOPort.ADRead[4].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[7, 1].Text != IOPort.ADRead[5].ToString("0.0")) fpSpread2.ActiveSheet.Cells[7, 1].Text = IOPort.ADRead[5].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[8, 1].Text != IOPort.ADRead[6].ToString("0.0")) fpSpread2.ActiveSheet.Cells[8, 1].Text = IOPort.ADRead[6].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[9, 1].Text != IOPort.ADRead[7].ToString("0.0")) fpSpread2.ActiveSheet.Cells[9, 1].Text = IOPort.ADRead[7].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[10, 1].Text != IOPort.ADRead[8].ToString("0.0")) fpSpread2.ActiveSheet.Cells[10, 1].Text = IOPort.ADRead[8].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[11, 1].Text != IOPort.ADRead[9].ToString("0.0")) fpSpread2.ActiveSheet.Cells[11, 1].Text = IOPort.ADRead[9].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[12, 1].Text != IOPort.ADRead[10].ToString("0.0")) fpSpread2.ActiveSheet.Cells[12, 1].Text = IOPort.ADRead[10].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[13, 1].Text != IOPort.ADRead[11].ToString("0.0")) fpSpread2.ActiveSheet.Cells[13, 1].Text = IOPort.ADRead[11].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[14, 1].Text != IOPort.ADRead[12].ToString("0.0")) fpSpread2.ActiveSheet.Cells[14, 1].Text = IOPort.ADRead[12].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[15, 1].Text != IOPort.ADRead[13].ToString("0.0")) fpSpread2.ActiveSheet.Cells[15, 1].Text = IOPort.ADRead[13].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[16, 1].Text != IOPort.ADRead[14].ToString("0.0")) fpSpread2.ActiveSheet.Cells[16, 1].Text = IOPort.ADRead[14].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[17, 1].Text != IOPort.ADRead[15].ToString("0.0")) fpSpread2.ActiveSheet.Cells[17, 1].Text = IOPort.ADRead[15].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[18, 1].Text != IOPort.ADRead[16].ToString("0.0")) fpSpread2.ActiveSheet.Cells[18, 1].Text = IOPort.ADRead[16].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[19, 1].Text != IOPort.ADRead[17].ToString("0.0")) fpSpread2.ActiveSheet.Cells[19, 1].Text = IOPort.ADRead[17].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[20, 1].Text != IOPort.ADRead[18].ToString("0.0")) fpSpread2.ActiveSheet.Cells[20, 1].Text = IOPort.ADRead[18].ToString("0.0");
                if (fpSpread2.ActiveSheet.Cells[21, 1].Text != IOPort.ADRead[19].ToString("0.0")) fpSpread2.ActiveSheet.Cells[21, 1].Text = IOPort.ADRead[19].ToString("0.0");

                byte[,] OutData = IOPort.GetOutData;

                string s;

                s = "";
                for (int i = 0; i < 8; i++) s += string.Format("{0:X2}, ", OutData[1, i]);
                if (label41.Text != s) label41.Text = s;
                s = "";
                for (int i = 0; i < 8; i++) s += string.Format("{0:X2}, ", OutData[2, i]);
                if (label40.Text != s) label40.Text = s;
                s = "";
                for (int i = 0; i < 8; i++) s += string.Format("{0:X2}, ", OutData[3, i]);
                if (label39.Text != s) label39.Text = s;
            }
            return;
        }

        private bool MemSetButton { get; set; }
        private long MemSetButtonToTime = 0;
        private void imageButton3_Click(object sender, EventArgs e)
        {
            MemSetButton = true;
            ledBulb1.On = true;
            MemSetButtonToTime = ComF.timeGetTimems();
            return;
        }

        private void imageButton4_Click(object sender, EventArgs e)
        {
            ledBulb2.On = true;
            if (MemSetButton == true)
                IMSSetButton(0);
            else IMSM1Button();
            ledBulb1.On = false;
            ledBulb2.On = false;
            MemSetButton = false;
            return;
        }

        private void imageButton5_Click(object sender, EventArgs e)
        {
            ledBulb3.On = true;
            if (MemSetButton == true)
                IMSSetButton(1);
            else IMSM2Button();
            ledBulb1.On = false;
            ledBulb3.On = false;
            MemSetButton = false;
            return;
        }

        private int RowCount { get; set; }

        private void CreateFileName()
        {
            string Path = Program.DATA_PATH.ToString() + "\\" + DateTime.Now.ToString("yyyyMM") + ".xls";

            //if ((Infor.DataName != "") && (Infor.DataName != null))
            //{
            //    if (File.Exists(Infor.DataName))
            //    {
            //        if (Infor.DataName != Path)
            //        {
            //            Infor.Date = DateTime.Now.ToString("yyyyMMdd");
            //            Infor.DataName = Path;
            //            Infor.TotalCount = 0;
            //            Infor.OkCount = 0;
            //            Infor.NgCount = 0;
            //            SaveInfor();
            //        }
            //    }
            //}
            //else
            //{
            //    Infor.Date = DateTime.Now.ToString("yyyyMMdd");

            //    Infor.DataName = Path;
            //    Infor.TotalCount = 0;
            //    Infor.OkCount = 0;
            //    Infor.NgCount = 0;
            //    SaveInfor();
            //}

            if (File.Exists(Path) == false)
            {
                CreateDataFile(Path);
            }
            else
            {
                fpSpread3.OpenExcel(Path);
                fpSpread3.ActiveSheet.Protect = false;
                fpSpread3.Visible = false;
                RowCount = 6;
                for (int i = RowCount; i < fpSpread3.ActiveSheet.RowCount; i++)
                {
                    if (fpSpread3.ActiveSheet.Cells[RowCount, 0].Text == "") break;
                    if (fpSpread3.ActiveSheet.Cells[RowCount, 0].Text == null) break;
                    RowCount++;
                }

                int Col = 0;
                //bool Flag = false;

                for (int i = 0; i < fpSpread3.ActiveSheet.ColumnCount; i++)
                {
                    //if (fpSpread3.ActiveSheet.Cells[4, Col].Text == "LUMBER UP/DN") Flag = true;

                    //if (Flag == true)
                    //{
                    //    if (fpSpread3.ActiveSheet.Cells[5, Col].Text == "구동음")
                    //        break;
                    //    else Col++;
                    //}
                    //else
                    //{
                    //    Col++;
                    //}

                    if (fpSpread3.ActiveSheet.Cells[3, Col].Text == "납입위치")
                        break;
                    else Col++;
                }
                fpSpread3.ActiveSheet.ColumnCount = Col + 1;
            }
            return;
        }

        private void SaveData()
        {
            string Path = Program.DATA_PATH.ToString() + "\\" + DateTime.Now.ToString("yyyyMM") + ".xls";

            //if ((Infor.DataName != "") && (Infor.DataName != null))
            //{
            //    if (File.Exists(Infor.DataName))
            //    {
            //        if (Infor.DataName != Path)
            //        {
            //            Infor.Date = DateTime.Now.ToString("yyyyMMdd");
            //            Infor.DataName = Path;
            //            Infor.TotalCount = 1;
            //            if (mData.Result == RESULT.PASS)
            //                Infor.OkCount = 1;
            //            else Infor.NgCount = 1;
            //            SaveInfor();
            //        }
            //        else
            //        {
            //            SaveInfor();
            //        }
            //    }
            //    else
            //    {
            //        SaveInfor();
            //    }
            //}
            //else
            //{
            //    Infor.Date = DateTime.Now.ToString("yyyyMMdd");
            //    Infor.DataName = Path;
            //    Infor.TotalCount = 1;
            //    if (mData.Result == RESULT.PASS)
            //        Infor.OkCount = 1;
            //    else Infor.NgCount = 1;
            //    SaveInfor();
            //}

            //if (Infor.DataName != Path)
            //{
            //    Infor.Date = DateTime.Now.ToString("yyyyMMdd");
            //    Infor.DataName = Path;
            //    Infor.TotalCount = 1;
            //    if (mData.Result == RESULT.PASS)
            //        Infor.OkCount = 1;
            //    else Infor.NgCount = 1;
            //    SaveInfor();
            //}
            if (File.Exists(Infor.DataName) == false) CreateDataFile(Path);

            int Col = 0;

            fpSpread3.SuspendLayout();
            fpSpread3.ActiveSheet.RowCount = RowCount + 1;

            fpSpread3.ActiveSheet.SetRowHeight(RowCount, 21);
            //if (fpSpread3.ActiveSheet.ColumnCount != 10) fpSpread3.ActiveSheet.ColumnCount = 10;

            for (int i = 0; i < fpSpread3.ActiveSheet.ColumnCount; i++)
            {
                fpSpread3.ActiveSheet.Cells[RowCount, i].CellType = new FarPoint.Win.Spread.CellType.TextCellType();
                fpSpread3.ActiveSheet.Cells[RowCount, i].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
                fpSpread3.ActiveSheet.Cells[RowCount, i].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
                fpSpread3.ActiveSheet.Cells[RowCount, i].Border = LineBorderToData;
                fpSpread3.ActiveSheet.Cells[RowCount, i].Text = "";
            }
            //No.
            fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
            fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = ((RowCount - 6) + 1).ToString();
            Col++;

            //Time
            fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
            fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = DateTime.Now.ToString("yyyy-MM-dd") + " " + DateTime.Now.ToLongTimeString();
            Col++;

            //Barcode
            fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
            fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;

            if (IOPort.GetAuto == true)
                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = label17.Text;
            else fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = textBox1.Text;
            Col++;

            //자동,수동
            fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
            fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;

            if (IOPort.GetAuto == true)
                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = "자동";
            else fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = "수동";
            Col++;

            if (mData.Result == RESULT.PASS)
            {
                fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
            }
            else if (mData.Result == RESULT.REJECT)
            {
                fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
            }
            if (mData.Result == RESULT.PASS)
                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = "O.K";
            else if (mData.Result == RESULT.REJECT)
                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = "N.G";
            Col++;

            if (mData.Ims.Test == true)
            {
                //fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                //fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                //fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = fpSpread1.ActiveSheet.Cells[2, 3].Text + " ~ " + fpSpread1.ActiveSheet.Cells[2, 4].Text;
                //Col++;

                //if (mData.Ims.Result != RESULT.REJECT)
                //{
                //    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                //    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                //}
                //else
                //{
                //    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                //    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                //}
                //fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.Ims.Data);

                if (mData.Ims.Result != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = "OK";
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = "NG";
                }
            }
            //else
            //{
            //    Col++;
            //}
            Col++;

            //Slide Curr
            if (mData.Slide.Test == true)
            {
                fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = fpSpread1.ActiveSheet.Cells[3, 3].Text + " ~ " + fpSpread1.ActiveSheet.Cells[3, 4].Text;
                Col++;

                if (mData.Slide.Result != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }
                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.Slide.Data);
            }
            else
            {
                Col++;
            }
            Col++;
            //Recline Curr
            if (mData.Recline.Test == true)
            {
                fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = fpSpread1.ActiveSheet.Cells[4, 3].Text + " ~ " + fpSpread1.ActiveSheet.Cells[4, 4].Text;
                Col++;

                if (mData.Recline.Result != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }
                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.Recline.Data);
            }
            else
            {
                Col++;
            }
            Col++;


            //Tilt Curr
            if (mData.Tilt.Test == true)
            {
                fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = fpSpread1.ActiveSheet.Cells[5, 3].Text + " ~ " + fpSpread1.ActiveSheet.Cells[5, 4].Text;
                Col++;

                if (mData.Tilt.Result != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }
                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.Tilt.Data);
            }
            else
            {
                Col++;
            }
            Col++;

            //Height Curr
            if (mData.Height.Test == true)
            {
                fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = fpSpread1.ActiveSheet.Cells[6, 3].Text + " ~ " + fpSpread1.ActiveSheet.Cells[6, 4].Text;
                Col++;

                if (mData.Height.Result != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }
                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.Height.Data);
            }
            else
            {
                Col++;
            }
            Col++;

            //Lumber Fwd / bwd Curr
            if (mData.LumberFwdBwd.Test == true)
            {
                fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = fpSpread1.ActiveSheet.Cells[7, 3].Text + " ~ " + fpSpread1.ActiveSheet.Cells[7, 4].Text;
                Col++;

                if (mData.LumberFwdBwd.Result != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }
                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.LumberFwdBwd.Data);
            }
            else
            {
                Col++;
            }
            Col++;


            //Lumber Up / Dn Curr
            if (mData.LumberUpDn.Test == true)
            {
                fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = fpSpread1.ActiveSheet.Cells[8, 3].Text + " ~ " + fpSpread1.ActiveSheet.Cells[8, 4].Text;
                Col++;

                if (mData.LumberUpDn.Result != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }
                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.LumberUpDn.Data);
            }
            else
            {
                Col++;
            }
            Col++;


            //Sound Slide
            if (mData.SoundSlide.Test == true)
            {
                if (mData.SoundSlide.StartResult != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }

                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.SoundSlide.StartData);
                Col++;

                if (mData.SoundSlide.RunResult != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }

                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.SoundSlide.RunData);
            }
            else
            {
                Col++;
            }
            Col++;


            //Sound Recline
            if (mData.SoundRecline.Test == true)
            {
                if (mData.SoundRecline.StartResult != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }

                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.SoundRecline.StartData);
                Col++;

                if (mData.SoundRecline.RunResult != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }

                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.SoundRecline.RunData);
            }
            else
            {
                Col++;
            }
            Col++;


            //Sound Tilt
            if (mData.SoundTilt.Test == true)
            {
                if (mData.SoundTilt.StartResult != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }

                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.SoundTilt.StartData);
                Col++;

                if (mData.SoundTilt.RunResult != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }

                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.SoundTilt.RunData);
            }
            else
            {
                Col++;
            }
            Col++;



            //Sound Height
            if (mData.SoundHeight.Test == true)
            {
                if (mData.SoundHeight.StartResult != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }

                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.SoundHeight.StartData);
                Col++;

                if (mData.SoundHeight.RunResult != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }

                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.SoundHeight.RunData);
            }
            else
            {
                Col++;
            }
            Col++;



            //Sound LumberFwdBwd
            if (mData.SoundLumberFwdBwd.Test == true)
            {
                if (mData.SoundLumberFwdBwd.StartResult != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }

                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.SoundLumberFwdBwd.StartData);
                Col++;

                if (mData.SoundLumberFwdBwd.RunResult != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }

                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.SoundLumberFwdBwd.RunData);
            }
            else
            {
                Col++;
            }
            Col++;


            //Sound LumberUpDn
            if (mData.SoundLumberUpDn.Test == true)
            {
                if (mData.SoundLumberUpDn.StartResult != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }

                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.SoundLumberUpDn.StartData);
                Col++;

                if (mData.SoundLumberUpDn.RunResult != RESULT.REJECT)
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
                }
                else
                {
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                    fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
                }

                fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = string.Format("{0:0.0}", mData.SoundLumberUpDn.RunData);
            }
            else
            {
                Col++;             
            }
            Col++;

            //납입위치
            if (InComingResult != RESULT.REJECT)
            {
                fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.White;
                fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.Black;
            }
            else
            {
                fpSpread3.ActiveSheet.Cells[RowCount, Col].BackColor = Color.Red;
                fpSpread3.ActiveSheet.Cells[RowCount, Col].ForeColor = Color.White;
            }

            fpSpread3.ActiveSheet.Cells[RowCount, Col].Text = (InComingResult != RESULT.REJECT) ? "OK" : "NG";
            Col++;


            //Col++;
            //if(fpSpread3.ActiveSheet.ColumnCount < Col) fpSpread3.ActiveSheet.ColumnCount = Col;
            RowCount++;
            fpSpread3.ResumeLayout();
            fpSpread3.SaveExcel(Path);
            return;
        }


        private FarPoint.Win.LineBorder LineBorderToHeader = new FarPoint.Win.LineBorder(Color.Black, 1/*RowHeight*/, true, true, true, true);//line color,line style,left,top,right,buttom                       
        private FarPoint.Win.LineBorder LineBorderToData = new FarPoint.Win.LineBorder(Color.Black, 1/*RowHeight*/, true, false, true, true);//line color,line style,left,top,right,buttom                       
        private void CreateDataFile(string dName)
        {
            fpSpread3.ActiveSheet.Reset();
            fpSpread2.ActiveSheet.Protect = false;
            fpSpread3.SuspendLayout();

            fpSpread3.ActiveSheet.RowCount = 9;
            //용지 방향
            fpSpread3.ActiveSheet.PrintInfo.Orientation = FarPoint.Win.Spread.PrintOrientation.Landscape;
            //프린트 할 때 가로,세로 중앙에 프린트 할 수 있도록 설정
            fpSpread3.ActiveSheet.PrintInfo.Centering = FarPoint.Win.Spread.Centering.Horizontal; //좌/우 중앙                        
            //fpSpread3.ActiveSheet.PrintInfo.PrintCenterOnPageV = false; //Top 쪽으로간다. 만약 true로 설정할 경우 상,하 중간에 프린트가 된다.

            //여백
            fpSpread3.ActiveSheet.PrintInfo.Margin.Bottom = 1;
            fpSpread3.ActiveSheet.PrintInfo.Margin.Left = 1;
            fpSpread3.ActiveSheet.PrintInfo.Margin.Right = 1;
            fpSpread3.ActiveSheet.PrintInfo.Margin.Top = 2;

            //프린트에서 컬러 표시
            fpSpread3.ActiveSheet.PrintInfo.ShowColor = true;
            //프린트에서 셀 라인 표시여부 (true일경우 내가 그린 라인 말고 셀에 사각 표시 라인도 같이 프린트가 된다.
            fpSpread3.ActiveSheet.PrintInfo.ShowGrid = false;

            fpSpread3.ActiveSheet.VerticalGridLine = new FarPoint.Win.Spread.GridLine(FarPoint.Win.Spread.GridLineType.None);
            fpSpread3.ActiveSheet.HorizontalGridLine = new FarPoint.Win.Spread.GridLine(FarPoint.Win.Spread.GridLineType.None);
            //용지 넓이에 페이지 맞춤
            fpSpread3.ActiveSheet.PrintInfo.UseSmartPrint = true;

            //그리드를 표시할 경우 저장할 때나 프린트 할 때 화면에 같이 표시,그리드가 프린트 되기 때문에 지저분해 보인다.
            fpSpread3.ActiveSheet.PrintInfo.ShowColumnFooter = FarPoint.Win.Spread.PrintHeader.Hide;
            fpSpread3.ActiveSheet.PrintInfo.ShowColumnFooterEachPage = false;


            //헤더와 밖같 라인이 같이 프린트 되지 않도록 한다.
            fpSpread3.ActiveSheet.PrintInfo.ShowBorder = false;
            fpSpread3.ActiveSheet.PrintInfo.ShowColumnHeader = FarPoint.Win.Spread.PrintHeader.Hide;
            fpSpread3.ActiveSheet.PrintInfo.ShowRowHeader = FarPoint.Win.Spread.PrintHeader.Hide;
            fpSpread3.ActiveSheet.PrintInfo.ShowShadows = false;
            fpSpread3.ActiveSheet.PrintInfo.ShowTitle = FarPoint.Win.Spread.PrintTitle.Hide;
            fpSpread3.ActiveSheet.PrintInfo.ShowSubtitle = FarPoint.Win.Spread.PrintTitle.Hide;

            //시트 보호를 해지 한다.
            //fpSpread3.ActiveSheet.PrintInfo.PrintType = FarPoint.Win.Spread.PrintType.All;
            //fpSpread3.ActiveSheet.PrintInfo.SmartPrintRules.Add(new ReadOnlyAttribute(false));
            //axfpSpread1.Protect = false;


            //for (int i = 0; i < 24; i++) fpSpread3.ActiveSheet.SetColumnWidth(i, 80);

            //틀 고정
            fpSpread3.ActiveSheet.FrozenColumnCount = 3;
            fpSpread3.ActiveSheet.FrozenRowCount = 6;
            fpSpread3.ActiveSheet.Cells[1, 0].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[1, 0].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right;
            fpSpread3.ActiveSheet.Cells[1, 0].Text = "날짜 :";
            fpSpread3.ActiveSheet.AddSpanCell(1, 1, 1, 22);

            fpSpread3.ActiveSheet.Cells[1, 1].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[1, 1].Text = DateTime.Now.ToLongDateString();
            fpSpread3.ActiveSheet.Cells[1, 1].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[1, 1].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Left;

            int Col;

            Col = 0;
            //No
            fpSpread3.ActiveSheet.SetRowHeight(3, 31);
            fpSpread3.ActiveSheet.SetRowHeight(4, 31);
            fpSpread3.ActiveSheet.SetRowHeight(5, 31);
            fpSpread3.ActiveSheet.AddSpanCell(3, Col, 3, 1);
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[3, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[3, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].Text = "NO.";
            fpSpread3.ActiveSheet.Cells[3, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[3, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[3, Col].Border = LineBorderToHeader;
            Col++;

            //Time
            fpSpread3.ActiveSheet.AddSpanCell(3, Col, 3, 1);
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 300);
            fpSpread3.ActiveSheet.Cells[3, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[3, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].Text = "생산 시간";
            fpSpread3.ActiveSheet.Cells[3, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[3, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[3, Col].Border = LineBorderToHeader;
            Col++;

            //바코드
            fpSpread3.ActiveSheet.AddSpanCell(3, Col, 3, 1);
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 400);
            fpSpread3.ActiveSheet.Cells[3, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[3, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].Text = "바코드";
            fpSpread3.ActiveSheet.Cells[3, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[3, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[3, Col].Border = LineBorderToHeader;
            Col++;
            //자동,수동
            fpSpread3.ActiveSheet.AddSpanCell(3, Col, 3, 1);
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[3, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[3, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].Text = "자동/수동";
            fpSpread3.ActiveSheet.Cells[3, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[3, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[3, Col].Border = LineBorderToHeader;
            Col++;

            //판정
            fpSpread3.ActiveSheet.AddSpanCell(3, Col, 3, 1);
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[3, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[3, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].Text = "판정";
            fpSpread3.ActiveSheet.Cells[3, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[3, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[3, Col].Border = LineBorderToHeader;

            Col++;

            //IMS
            fpSpread3.ActiveSheet.AddSpanCell(3, Col, 3, 1);
            fpSpread3.ActiveSheet.Cells[3, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[3, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].Text = "IMS";
            fpSpread3.ActiveSheet.Cells[3, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[3, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[3, Col].Border = LineBorderToHeader;

            //fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            //fpSpread3.ActiveSheet.AddSpanCell(4, Col, 1, 2);
            //fpSpread3.ActiveSheet.Cells[4, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            //fpSpread3.ActiveSheet.Cells[4, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            //fpSpread3.ActiveSheet.Cells[4, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            //fpSpread3.ActiveSheet.Cells[4, Col].Text = "전류 [A]";
            //fpSpread3.ActiveSheet.Cells[4, Col].BackColor = Color.WhiteSmoke;
            //fpSpread3.ActiveSheet.Cells[4, Col].ForeColor = Color.Black;
            //fpSpread3.ActiveSheet.Cells[4, Col].Border = LineBorderToData;

            //fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            //fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            //fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            //fpSpread3.ActiveSheet.Cells[5, Col].Text = "스팩";
            //fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            //fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            //fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            //Col++;
            //fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            //fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            //fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            //fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            //fpSpread3.ActiveSheet.Cells[5, Col].Text = "데이타";
            //fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            //fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            //fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;


            //전류
            fpSpread3.ActiveSheet.AddSpanCell(3, Col, 1, 8);
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[3, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[3, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].Text = "P/SEAT 전류 [A]";
            fpSpread3.ActiveSheet.Cells[3, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[3, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[3, Col].Border = LineBorderToHeader;

            fpSpread3.ActiveSheet.AddSpanCell(4, Col, 1, 2);
            fpSpread3.ActiveSheet.Cells[4, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[4, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].Text = "SLIDE";
            fpSpread3.ActiveSheet.Cells[4, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[4, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[4, Col].Border = LineBorderToHeader;

            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "스팩";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "데이타";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.AddSpanCell(4, Col, 1, 2);
            fpSpread3.ActiveSheet.Cells[4, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[4, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].Text = "RECLINE";
            fpSpread3.ActiveSheet.Cells[4, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[4, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[4, Col].Border = LineBorderToHeader;

            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "스팩";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "데이타";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.AddSpanCell(4, Col, 1, 2);
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[4, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[4, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].Text = "TILT";
            fpSpread3.ActiveSheet.Cells[4, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[4, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[4, Col].Border = LineBorderToHeader;

            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "스팩";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "데이타";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;

            fpSpread3.ActiveSheet.AddSpanCell(4, Col, 1, 2);
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[4, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[4, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].Text = "HEIGHT";
            fpSpread3.ActiveSheet.Cells[4, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[4, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[4, Col].Border = LineBorderToHeader;

            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "스팩";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "데이타";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;


            //전류
            fpSpread3.ActiveSheet.AddSpanCell(3, Col, 1, 4);
            fpSpread3.ActiveSheet.Cells[3, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[3, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].Text = "LUMBER 전류 [A]";
            fpSpread3.ActiveSheet.Cells[3, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[3, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[3, Col].Border = LineBorderToHeader;

            fpSpread3.ActiveSheet.AddSpanCell(4, Col, 1, 2);
            fpSpread3.ActiveSheet.Cells[4, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[4, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].Text = "FWD/BWD";
            fpSpread3.ActiveSheet.Cells[4, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[4, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[4, Col].Border = LineBorderToHeader;

            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "스팩";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "데이타";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.AddSpanCell(4, Col, 1, 2);
            fpSpread3.ActiveSheet.Cells[4, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[4, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].Text = "UP/DN";
            fpSpread3.ActiveSheet.Cells[4, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[4, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[4, Col].Border = LineBorderToHeader;

            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "스팩";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "데이타";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;

            //소음
            fpSpread3.ActiveSheet.AddSpanCell(3, Col, 1, 12);
            fpSpread3.ActiveSheet.Cells[3, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[3, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].Text = "소음[dB]";
            fpSpread3.ActiveSheet.Cells[3, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[3, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[3, Col].Border = LineBorderToHeader;

            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.AddSpanCell(4, Col, 1, 2);
            fpSpread3.ActiveSheet.Cells[4, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[4, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].Text = "SLIDE";
            fpSpread3.ActiveSheet.Cells[4, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[4, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[4, Col].Border = LineBorderToData;

            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "기동음";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "구동음";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.AddSpanCell(4, Col, 1, 2);
            fpSpread3.ActiveSheet.Cells[4, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[4, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].Text = "RECLINE";
            fpSpread3.ActiveSheet.Cells[4, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[4, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[4, Col].Border = LineBorderToData;

            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "기동음";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "구동음";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.AddSpanCell(4, Col, 1, 2);
            fpSpread3.ActiveSheet.Cells[4, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[4, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].Text = "TILT";
            fpSpread3.ActiveSheet.Cells[4, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[4, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[4, Col].Border = LineBorderToData;

            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "기동음";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "구동음";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.AddSpanCell(4, Col, 1, 2);
            fpSpread3.ActiveSheet.Cells[4, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[4, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].Text = "HEIGHT";
            fpSpread3.ActiveSheet.Cells[4, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[4, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[4, Col].Border = LineBorderToData;

            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "기동음";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "구동음";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;

            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.AddSpanCell(4, Col, 1, 2);
            fpSpread3.ActiveSheet.Cells[4, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[4, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].Text = "LUMBER FWD/BWD";
            fpSpread3.ActiveSheet.Cells[4, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[4, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[4, Col].Border = LineBorderToData;

            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "기동음";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "구동음";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;

            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.AddSpanCell(4, Col, 1, 2);
            fpSpread3.ActiveSheet.Cells[4, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[4, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[4, Col].Text = "LUMBER UP/ DN";
            fpSpread3.ActiveSheet.Cells[4, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[4, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[4, Col].Border = LineBorderToData;

            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "기동음";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[5, Col].Text = "구동음";
            fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            Col++;

            //fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            //fpSpread3.ActiveSheet.AddSpanCell(4, Col, 1, 2);
            //fpSpread3.ActiveSheet.Cells[4, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            //fpSpread3.ActiveSheet.Cells[4, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            //fpSpread3.ActiveSheet.Cells[4, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            //fpSpread3.ActiveSheet.Cells[4, Col].Text = "L/UP DN";
            //fpSpread3.ActiveSheet.Cells[4, Col].BackColor = Color.WhiteSmoke;
            //fpSpread3.ActiveSheet.Cells[4, Col].ForeColor = Color.Black;
            //fpSpread3.ActiveSheet.Cells[4, Col].Border = LineBorderToData;

            //fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            //fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            //fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            //fpSpread3.ActiveSheet.Cells[5, Col].Text = "기동음";
            //fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            //fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            //fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            //Col++;
            //fpSpread3.ActiveSheet.SetColumnWidth(Col, 100);
            //fpSpread3.ActiveSheet.Cells[5, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            //fpSpread3.ActiveSheet.Cells[5, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            //fpSpread3.ActiveSheet.Cells[5, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            //fpSpread3.ActiveSheet.Cells[5, Col].Text = "구동음";
            //fpSpread3.ActiveSheet.Cells[5, Col].BackColor = Color.WhiteSmoke;
            //fpSpread3.ActiveSheet.Cells[5, Col].ForeColor = Color.Black;
            //fpSpread3.ActiveSheet.Cells[5, Col].Border = LineBorderToHeader;
            //Col++;


            //납입위치
            fpSpread3.ActiveSheet.AddSpanCell(3, Col, 3, 1);
            fpSpread3.ActiveSheet.SetColumnWidth(Col, 150);
            fpSpread3.ActiveSheet.Cells[3, Col].CellType = new FarPoint.Win.Spread.CellType.EditBaseCellType();
            fpSpread3.ActiveSheet.Cells[3, Col].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[3, Col].Text = "납입위치";
            fpSpread3.ActiveSheet.Cells[3, Col].BackColor = Color.WhiteSmoke;
            fpSpread3.ActiveSheet.Cells[3, Col].ForeColor = Color.Black;
            fpSpread3.ActiveSheet.Cells[3, Col].Border = LineBorderToHeader;


            fpSpread2.ActiveSheet.ColumnCount = Col;
            //Header
            fpSpread3.ActiveSheet.AddSpanCell(0, 0, 1, Col);
            fpSpread3.ActiveSheet.SetRowHeight(0, 100);
            fpSpread3.ActiveSheet.Cells[0, 0].Font = new Font("맑은 고딕", 26);
            fpSpread3.ActiveSheet.SetText(0, 0, "레포트");
            fpSpread3.ActiveSheet.Cells[0, 0].VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center;
            fpSpread3.ActiveSheet.Cells[0, 0].HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center;
            RowCount = 6;

            fpSpread3.ResumeLayout();
            fpSpread3.SaveExcel(dName);
            return;
        }

        private void SaveInfor()
        {
            string Path = Program.INFOR_PATH.ToString() + "\\" + DateTime.Now.ToString("yyyyMMdd") + ".inf";


            TIniFile Ini = new TIniFile(Path);

            Ini.WriteInteger("COUNT", "TOTAL", Infor.TotalCount);
            Ini.WriteInteger("COUNT", "OK", Infor.TotalCount - Infor.NgCount);
            Ini.WriteInteger("COUNT", "NG", Infor.NgCount);

            Ini.WriteString("NAME", "DATA", Infor.DataName);
            Ini.WriteString("NAME", "DATE", Infor.Date);
            Ini.WriteBool("OPTION", "VALUE", Infor.ReBootingFlag);
            return;
        }

        private void OpenInfor()
        {
            string Path = Program.INFOR_PATH.ToString() + "\\" + DateTime.Now.ToString("yyyyMMdd") + ".inf";
            string dPath = Program.DATA_PATH.ToString() + "\\" + DateTime.Now.ToString("yyyyMM") + ".xls";

            if (File.Exists(Path) == false)
            {
                //CreateFileName();
                Infor.Date = DateTime.Now.ToString("yyyyMMdd");
                Infor.DataName = dPath;
                Infor.TotalCount = 0;
                Infor.OkCount = 0;
                Infor.NgCount = 0;
                SaveInfor();
                return;
            }
            else
            {
                TIniFile Ini = new TIniFile(Path);

                if (Ini.ReadInteger("COUNT", "TOTAL", ref Infor.TotalCount) == false) Infor.TotalCount = 0;
                if (Ini.ReadInteger("COUNT", "OK", ref Infor.OkCount) == false) Infor.OkCount = 0;
                if (Ini.ReadInteger("COUNT", "NG", ref Infor.NgCount) == false) Infor.NgCount = 0;

                if (Ini.ReadString("NAME", "DATA", ref Infor.DataName) == false) Infor.DataName = dPath;
                if (Ini.ReadString("NAME", "DATE", ref Infor.Date) == false) Infor.Date = DateTime.Now.ToString("yyyyMMdd");
                if (Ini.ReadBool("OPTION", "VALUE", ref Infor.ReBootingFlag) == false) Infor.ReBootingFlag = true;
            }
            
            
            if (File.Exists(dPath) == false)
            {
                CreateFileName();
            }
            else
            {
                fpSpread3.OpenExcel(dPath);
                fpSpread2.ActiveSheet.Protect = false;
                fpSpread3.Visible = false;
                RowCount = 6;
                for (int i = RowCount; i < fpSpread3.ActiveSheet.RowCount; i++)
                {
                    if (fpSpread3.ActiveSheet.Cells[RowCount, 0].Text == "") break;
                    if (fpSpread3.ActiveSheet.Cells[RowCount, 0].Text == null) break;
                    RowCount++;
                }

                int Col = 0;
                //bool Flag = false;

                for (int i = 0; i < fpSpread3.ActiveSheet.ColumnCount; i++)
                {
                    //if (fpSpread3.ActiveSheet.Cells[4, Col].Text == "LUMBER UP/DN") Flag = true;

                    //if (Flag == true)
                    //{
                    //    if (fpSpread3.ActiveSheet.Cells[5, Col].Text == "구동음")
                    //        break;
                    //    else Col++;
                    //}
                    //else
                    //{
                    //    Col++;
                    //}

                    if (fpSpread3.ActiveSheet.Cells[3, Col].Text == "납입위치")
                        break;
                    else Col++;
                }

                fpSpread3.ActiveSheet.ColumnCount = Col + 1;
            }
            return;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            panel7.Visible = false;
            return;
        }

        private bool OpenSpec(bool sFlag = false)
        {
            bool Flag = false;

            string LHType = comboBox3.SelectedItem.ToString();
            string PSeat = comboBox4.SelectedItem.ToString();
            string LHDType;
            string ModelType = "";
            
            if (comboBox2.SelectedItem != null)
                LHDType = comboBox2.SelectedItem.ToString();
            else LHDType = "LHD";

            if (comboBox3.SelectedItem != null)
                LHType = comboBox3.SelectedItem.ToString();
            else LHType = "LH";


            if (((LHDType == "LHD") && (LHType == "LH")) || ((LHDType == "RHD") && (LHType == "RH")))
            {
                //운전석

                //PSeat -> IMS OR POWER
                if ((CheckItem.PSeat4Way == true) || (PSeat == "IMS"))
                {
                    if (PSeat == "IMS")
                        ModelType = LHDType + "_" + LHType + "_" + PSeat; //LHD_LH_IMS or LHD_LH_POWER or RHD_RH_IMS or RHD_RH_POWER
                    else ModelType = LHDType + "_" + LHType + "_" + "4WAY";
                }
                else ModelType = LHDType + "_" + LHType + "_" + "2WAY"; //LHD_LH_2WAY or RHD_RH_2WAY


                if ((CheckItem.PSeat4Way == true) || (PSeat == "IMS"))
                {
                    CheckItem.PSeat12Way = true;
                    CheckItem.PSeat10Way = false;
                }
                else
                {
                    CheckItem.PSeat12Way = false;
                    CheckItem.PSeat10Way = true;
                }
                if (panel8.Visible == true) panel8.Visible = false;
            }
            else
            {
                if (PSeat != "IMS")
                {
                    //if (CheckItem.PSeat4Way == true)
                    if (CheckItem.WalkIn == true)
                    {
                        //ModelType = "RH_8WAY_WALKIN";
                        ModelType = ModelType = LHDType + "_" + LHType + "_" + "WALKIN"; //"LHD_RH_WALKIN";
                        //CheckItem.WalkIn = true;

                        if (sFlag == false)
                        {
                            CheckItem.PSeat12Way = false;
                            CheckItem.PSeat10Way = false;
                            CheckItem.PSeat8Way = true;
                        }
                    }
                    else if (CheckItem.PSeat4Way == false)
                    {
                        ModelType = ModelType = LHDType + "_" + LHType + "_" + "8WAY";// "LHD_RH_8WAY";

                        if (sFlag == false)
                        {
                            CheckItem.WalkIn = false;
                            CheckItem.PSeat12Way = false;
                            CheckItem.PSeat10Way = false;
                            CheckItem.PSeat8Way = true;
                        }
                    }                    
                }
                else
                {
                    ModelType = ModelType = LHDType + "_" + LHType + "_" + PSeat; //"LHD_RH_WALKIN";
                }                
            }

            //if (panel8.Visible == true) panel8.Visible = false;
            string s = comboBox1.SelectedItem.ToString();
            bool xFlag = false;
            //if (s.IndexOf(ModelType) < 0)
            if (s.ToUpper() != ModelType)
            {
                //foreach (string Item in comboBox1.Items)
                //{
                //    //if (0 <= Item.ToUpper().IndexOf(ModelType))
                //    if (Item.ToUpper() == ModelType)
                //    {
                //        comboBox1.SelectedItem = Item;
                //        Flag = true;
                //        xFlag = true;
                //        break;
                //    }
                //}


                if (comboBox1.Items.Contains(ModelType) == true) comboBox1.SelectedItem = ModelType;
            }
            else
            {
                xFlag = true;
            }
            if (xFlag == false)
            {
                if (panel8.Visible == false) panel8.Visible = true;
            }
            else
            {
                if (panel8.Visible == true) panel8.Visible = false;
            }

            return Flag;
        }

        private bool isCheckItem
        {
            get
            {
                if ((CheckItem.Height == true) || (CheckItem.Slide == true) || (CheckItem.Recline == true) || (CheckItem.Tilt == true) || (CheckItem.WalkIn == true) || (CheckItem.PSeat12Way == true) || (CheckItem.Sound == true) || (CheckItem.PSeat == PSEAT_TYPE.IMS))
                    return true;
                else return false;
            }
        }
        private string SaveLogData
        {
            set
            {
                string xPath = Program.LOG_PATH.ToString() + "\\" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
                using (FileStream fp = File.Open(xPath, FileMode.Append, FileAccess.Write))
                {
                    StreamWriter writer = new StreamWriter(fp);
                    writer.Write(value + "\n");
                    writer.Close();
                    fp.Close();
                }

                //using (File.Open(xPath, FileMode.Open))
                //{
                //    File.AppendAllText(value + "\n", xPath);
                //}
            }
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", Program.DATA_PATH.ToString());
            return;
        }

        private void 시작프로그램등록ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ComF.StartProgramRegistrySet();
            return;
        }

        private void 시작프로그램삭제ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ComF.StartProgramRegistryDel();
            return;
        }

        private void 프로그램재시작ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExitFlag = true;
            ComF.timedelay(1000);
            this.Dispose();
#if PROGRAM_RUNNING
            if (Sound != null)
            {
                if (Sound.Connection == true) Sound.StartStopMesurement(false);
                ComF.timedelay(1000);
            }
            if (CanCtrl.isOpen(0) == true) CanCtrl.CanClose(0);
                

            if (PwCtrl.IsOpen == true) PwCtrl.Close();
            if (pMeter.isOpen == true) pMeter.Close();

            ComF.WindowRestartToDelay(0);
#endif
            System.Diagnostics.Process[] mProcess = System.Diagnostics.Process.GetProcessesByName(Application.ProductName);
            foreach (System.Diagnostics.Process p in mProcess) p.Kill();
            return;
        }

        private void 플ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExitFlag = true;
            ComF.timedelay(1000);
            this.Dispose();
#if PROGRAM_RUNNING
            if (Sound != null)
            {
                if (Sound.Connection == true) Sound.StartStopMesurement(false);
                ComF.timedelay(1000);
            }
            if (CanCtrl.isOpen(0) == true) CanCtrl.CanClose(0);


            if (PwCtrl.IsOpen == true) PwCtrl.Close();
            if (pMeter.isOpen == true) pMeter.Close();

            ComF.WindowExit();
#endif
            System.Diagnostics.Process[] mProcess = System.Diagnostics.Process.GetProcessesByName(Application.ProductName);
            foreach (System.Diagnostics.Process p in mProcess) p.Kill();
            return;
        }
    }
    public class KALMAN_FILETER
    {
        MyInterface mControl = null;
        private struct myKalmanFilterType
        {
            public float z_Din;
            public float Q;
            public float R;
            public float A;
            public float B_uk;
            public float H;
            public float x_Predict;
            public float Xk;
            public float p_Predict;
            public float Pk;
            public float K_gain;
        }

        private myKalmanFilterType[] KalmanStruct = new myKalmanFilterType[20];

        public KALMAN_FILETER()
        {

        }

        public KALMAN_FILETER(MyInterface mControl)
        {
            this.mControl = mControl;
        }

        //public myKalmanFilterType[] myKalmanBuff = new myKalmanFilterType[5];
        //public myKalmanFilterType myKalmanBuff = new myKalmanFilterType();

        ~KALMAN_FILETER()
        {

        }

        //public const double Speed = 0.01;
        //public const double Speed = 0.15;
        private void init_kalmanFiltering(short Ch, ref myKalmanFilterType[] kf)
        {
            //kf->Q = pow(0.01,2);

            //kf.Q = (float)Math.Pow(0.01, 2);
            //kf.R = (float)Math.Pow(0.5, 2);
            kf[Ch].Q = (float)Math.Pow(mControl.GetConfig.KalmanSpeed, 2);
            kf[Ch].R = (float)Math.Pow(0.5, 2);
            kf[Ch].A = 1;
            kf[Ch].B_uk = 0;
            kf[Ch].H = (float)1.0;
            kf[Ch].Xk = 0;     //25
            kf[Ch].Pk = (float)1.0;
            kf[Ch].K_gain = (float)1.0;

            return;
        }

        private void init_kalmanFiltering(ref myKalmanFilterType[] kf)
        {
            //kf->Q = pow(0.01,2);

            for (int i = 0; i < kf.Length; i++)
            {
                kf[i].Q = (float)Math.Pow(mControl.GetConfig.KalmanSpeed, 2);
                kf[i].R = (float)Math.Pow(0.5, 2);
                kf[i].A = 1;
                kf[i].B_uk = 0;
                kf[i].H = (float)1.0;
                kf[i].Xk = 0;     //25
                kf[i].Pk = (float)1.0;
                kf[i].K_gain = (float)1.0;
            }
            return;
        }

        private float kalmanFilter(ref myKalmanFilterType kf)
        {

            kf.x_Predict = (kf.A * kf.Xk) + kf.B_uk;

            kf.p_Predict = (kf.A * kf.Pk) + kf.Q;

            kf.K_gain = kf.p_Predict / (kf.H * kf.p_Predict + kf.R);

            kf.Xk = kf.x_Predict + kf.K_gain * (kf.z_Din - kf.x_Predict);

            kf.Pk = (1 - kf.K_gain) * kf.p_Predict;

            return kf.Xk;

        }

        private float kalmanFilter_(short Ch, float getdata, ref myKalmanFilterType myKalmanBuff)
        {
            myKalmanBuff.z_Din = getdata;

            return kalmanFilter(ref myKalmanBuff);
        }

        //초기화 함수
        private void init_kalmanFilter(ref myKalmanFilterType[] myKalmanBuff)
        {
            //init_kalmanFiltering(ref myKalmanBuff[0]);
            //init_kalmanFiltering(ref myKalmanBuff[1]);
            //init_kalmanFiltering(ref myKalmanBuff[2]);
            //init_kalmanFiltering(ref myKalmanBuff[3]);
            //init_kalmanFiltering(ref myKalmanBuff[4]);
            init_kalmanFiltering(ref myKalmanBuff);
            return;
        }
        private void init_kalmanFilter(short Ch, ref myKalmanFilterType[] myKalmanBuff)
        {
            //init_kalmanFiltering(ref myKalmanBuff[0]);
            //init_kalmanFiltering(ref myKalmanBuff[1]);
            //init_kalmanFiltering(ref myKalmanBuff[2]);
            //init_kalmanFiltering(ref myKalmanBuff[3]);
            //init_kalmanFiltering(ref myKalmanBuff[4]);
            init_kalmanFiltering(Ch, ref myKalmanBuff);
            return;
        }
        public void InitAll()
        {
            init_kalmanFilter(ref KalmanStruct);
            return;
        }

        public void Init(short Ch)
        {
            init_kalmanFilter(Ch, ref KalmanStruct);
            return;
        }

        public float CheckData(short Ch, float Data)
        {
            float rData = kalmanFilter_(Ch, Data, ref KalmanStruct[Ch]);
            rData = KalmanStruct[Ch].Xk;
            return rData;
            //return Data;
        }
    }
}
