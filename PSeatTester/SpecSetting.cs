using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PSeatTester
{
    public partial class SpecSetting : Form
    {
        private IOMaping IOMapForm = null;
        private __Spec__ Spec = new __Spec__();
        private MyInterface mControl = null;
        private PinMapStruct PinMap;

        public SpecSetting()
        {
            InitializeComponent();
        }
        public SpecSetting(MyInterface mControl, string mName)
        {
            InitializeComponent();
            this.mControl = mControl;
            this.mName = mName;
        }

        private bool ModelBoxChangeFlag = false;
        private void SpecSetting_Load(object sender, EventArgs e)
        {
            mControl.공용함수.ReadFileListNotExt(Program.SPEC_PATH.ToString(), "*.Spc", COMMON_FUCTION.FileSortMode.FILENAME_ODERBY);
            List<string> FList = mControl.공용함수.GetFileList;

            if (0 < FList.Count)
            {
                ModelBoxChangeFlag = true;
                comboBox1.Items.Clear();
                foreach (string s in FList) comboBox1.Items.Add(s);

                if (0 < comboBox1.Items.Count)
                {
                    if (0 < comboBox1.Items.Count)
                    {
                        if ((mName != null) && (mName != "") && (mName != string.Empty))
                        {
                            if (comboBox1.Items.Contains(mName) == true) comboBox1.SelectedItem = mName;
                        }
                    }

                }
                if (mName != null)
                {
                    string sName = Program.SPEC_PATH.ToString() + "\\" + mName + ".Spc";
                    if(File.Exists(sName) == true) mControl.공용함수.OpenSpec(sName,ref Spec);
                }
            }
            DisplaySpec();
            ModelBoxChangeFlag = false;
            return;
        }

        private void imageButton5_Click(object sender, EventArgs e)
        {
            if (IOMapForm == null)
            {
                IOMapForm = new IOMaping(PinMap:PinMap, mControl: mControl)
                {
                    Text = "핀 맵 설정",
                    FormBorderStyle = FormBorderStyle.FixedSingle,
                    ShowIcon = true,
                    MinimizeBox = false,
                    MaximizeBox = false,
                    ControlBox = false,
                    TopMost = true,
                    TopLevel = true,
                    Owner = this,
                };

                IOMapForm.FormClosing += delegate (object sender1, FormClosingEventArgs e1)
                {
                    PinMap = IOMapForm.GetPinMap;
                    e1.Cancel = false;
                    IOMapForm.Dispose();
                    IOMapForm = null;
                };

                IOMapForm.Show();
            }

            return;
        }
        private string mName;

        private void imageButton1_Click(object sender, EventArgs e)
        {
            //저장
            MoveSpec();

            string sName;
            string ModName = null;
            if (comboBox1.SelectedItem == null)
            {
                if (MessageBox.Show("선택된 모델이 없습니다.\n모델을 생성하시겠습니까?", "경고", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    if (InputBox.Show("모델명 입력", "모델명", ref ModName) == DialogResult.OK)
                    {
                        if ((ModName != null) && (ModName != "") && (ModName != string.Empty))
                        {
                            mName = ModName;
                            sName = Program.SPEC_PATH.ToString() + "\\" + ModName + ".Spc";

                        }
                        else
                        {
                            return;
                        }
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    return;
                }
            }
            else
            {
                ModName = comboBox1.SelectedItem.ToString();
                sName = Program.SPEC_PATH.ToString() + "\\" + comboBox1.SelectedItem.ToString() + ".Spc";
            }
            Spec.ModelName = ModName;
            mControl.공용함수.SaveSpec(Spec: Spec, Name: sName);
            return;
        }

        private void imageButton2_Click(object sender, EventArgs e)
        {
            //모델 추가
            string ModName = null;
            if (InputBox.Show("모델명 입력", "모델명", ref ModName) == DialogResult.OK)
            {
                if ((ModName != null) && (ModName != "") && (ModName != string.Empty))
                {
                    string sName;
                    mName = ModName;
                    sName = Program.SPEC_PATH.ToString() + "\\" + ModName + ".Spc";

                    ModelBoxChangeFlag = true;
                    comboBox1.Items.Add(ModName);

                    InitSpec();
                    Spec.ModelName = mName;
                    DisplaySpec();
                    mControl.공용함수.SaveSpec(sName, Spec);
                    ModelBoxChangeFlag = false;
                }
            }
            return;
        }

        private void imageButton3_Click(object sender, EventArgs e)
        {
            //삭제
            if (comboBox1.SelectedItem == null) return;
            string sName = Program.SPEC_PATH.ToString() + "\\" + comboBox1.SelectedItem.ToString() + ".Spc";

            if (File.Exists(sName) == true) File.Delete(sName);

            InitSpec();
            DisplaySpec();
            comboBox1.Items.Remove(comboBox1.SelectedItem);
            return;
        }

        private void imageButton4_Click(object sender, EventArgs e)
        {
            //다른 이름으로 저장
            MoveSpec();

            string ModName = null;
            if (InputBox.Show("모델명 입력", "모델명", ref ModName) == DialogResult.OK)
            {
                if ((ModName != null) && (ModName != "") && (ModName != string.Empty))
                {
                    string sName;
                    mName = ModName;
                    sName = Program.SPEC_PATH.ToString() + "\\" + ModName + ".Spc";

                    Spec.ModelName = ModName;
                    mControl.공용함수.SaveSpec(sName, Spec);
                    ModelBoxChangeFlag = true;
                    comboBox1.Items.Add(ModName);
                    comboBox1.SelectedItem = ModName;
                    ModelBoxChangeFlag = false;
                }
            }
            return;
        }
        private void DisplaySpec()
        {
            PinMap = Spec.PinMap;
            fpSpread1.Sheets[0].Cells[3, 3].Text = Spec.Current.IMS.Min.ToString("0.00");
            fpSpread1.Sheets[0].Cells[3, 4].Text = Spec.Current.IMS.Max.ToString("0.00");
            fpSpread1.Sheets[0].Cells[4, 3].Text = Spec.Current.Slide.Min.ToString("0.00");
            fpSpread1.Sheets[0].Cells[4, 4].Text = Spec.Current.Slide.Max.ToString("0.00");
            fpSpread1.Sheets[0].Cells[5, 3].Text = Spec.Current.Tilt.Min.ToString("0.00");
            fpSpread1.Sheets[0].Cells[5, 4].Text = Spec.Current.Tilt.Max.ToString("0.00");
            fpSpread1.Sheets[0].Cells[6, 3].Text = Spec.Current.Height.Min.ToString("0.00");
            fpSpread1.Sheets[0].Cells[6, 4].Text = Spec.Current.Height.Max.ToString("0.00");
            fpSpread1.Sheets[0].Cells[7, 3].Text = Spec.Current.Recliner.Min.ToString("0.00");
            fpSpread1.Sheets[0].Cells[7, 4].Text = Spec.Current.Recliner.Max.ToString("0.00");
            fpSpread1.Sheets[0].Cells[8, 3].Text = Spec.Current.LumberFwdBwd.Min.ToString("0.00");
            fpSpread1.Sheets[0].Cells[8, 4].Text = Spec.Current.LumberFwdBwd.Max.ToString("0.00");
            fpSpread1.Sheets[0].Cells[9, 3].Text = Spec.Current.LumberUpDn.Min.ToString("0.00");
            fpSpread1.Sheets[0].Cells[9, 4].Text = Spec.Current.LumberUpDn.Max.ToString("0.00");

            fpSpread1.Sheets[0].Cells[10, 3].Text = Spec.MovingSpeed.SlideFwd.Min.ToString("0.00");
            fpSpread1.Sheets[0].Cells[10, 4].Text = Spec.MovingSpeed.SlideFwd.Max.ToString("0.00");
            fpSpread1.Sheets[0].Cells[11, 3].Text = Spec.MovingSpeed.SlideBwd.Min.ToString("0.00");
            fpSpread1.Sheets[0].Cells[11, 4].Text = Spec.MovingSpeed.SlideBwd.Max.ToString("0.00");

            fpSpread1.Sheets[0].Cells[12, 3].Text = Spec.MovingSpeed.ReclinerFwd.Min.ToString("0.00");
            fpSpread1.Sheets[0].Cells[12, 4].Text = Spec.MovingSpeed.ReclinerFwd.Max.ToString("0.00");
            fpSpread1.Sheets[0].Cells[13, 3].Text = Spec.MovingSpeed.ReclinerBwd.Min.ToString("0.00");
            fpSpread1.Sheets[0].Cells[13, 4].Text = Spec.MovingSpeed.ReclinerBwd.Max.ToString("0.00");

            fpSpread1.Sheets[0].Cells[14, 3].Text = Spec.MovingSpeed.TiltUp.Min.ToString("0.00");
            fpSpread1.Sheets[0].Cells[14, 4].Text = Spec.MovingSpeed.TiltUp.Max.ToString("0.00");
            fpSpread1.Sheets[0].Cells[15, 3].Text = Spec.MovingSpeed.TiltDn.Min.ToString("0.00");
            fpSpread1.Sheets[0].Cells[15, 4].Text = Spec.MovingSpeed.TiltDn.Max.ToString("0.00");

            fpSpread1.Sheets[0].Cells[16, 3].Text = Spec.MovingSpeed.HeightUp.Min.ToString("0.00");
            fpSpread1.Sheets[0].Cells[16, 4].Text = Spec.MovingSpeed.HeightUp.Max.ToString("0.00");
            fpSpread1.Sheets[0].Cells[17, 3].Text = Spec.MovingSpeed.HeightDn.Min.ToString("0.00");
            fpSpread1.Sheets[0].Cells[17, 4].Text = Spec.MovingSpeed.HeightDn.Max.ToString("0.00");

            fpSpread1.Sheets[0].Cells[18, 3].Text = Spec.Sound.StartMax.ToString("0.0");
            fpSpread1.Sheets[0].Cells[18, 4].Text = Spec.Sound.RunMax.ToString("0.00");
            fpSpread1.Sheets[0].Cells[19, 4].Value = Spec.Sound.RMSMode == true ? "True" : "False";

            fpSpread1.Sheets[0].Cells[20, 4].Text = Spec.Sound.기동음구동음구분사이시간.ToString("0.0");
            fpSpread1.Sheets[0].Cells[21, 4].Text = Spec.Sound.Offset.ToString("0.0");

            fpSpread1.Sheets[0].Cells[23, 4].Text = Spec.SlideLimitTime.ToString("0.0");
            fpSpread1.Sheets[0].Cells[24, 4].Text = Spec.TiltLimitTime.ToString("0.0");
            fpSpread1.Sheets[0].Cells[25, 4].Text = Spec.HeightLimitTime.ToString("0.0");
            fpSpread1.Sheets[0].Cells[26, 4].Text = Spec.ReclinerLimitTime.ToString("0.0");
            fpSpread1.Sheets[0].Cells[27, 4].Text = Spec.LumberLimitTime.ToString("0.0");
            //fpSpread1.Sheets[0].Cells[8, 8].Value = Spec.Can == true ? "True" : "False";

            fpSpread1.Sheets[0].Cells[11, 8].Text = Spec.MovingStroke.Slide.ToString("0.0");
            fpSpread1.Sheets[0].Cells[12, 8].Text = Spec.MovingStroke.Tilt.ToString("0.0");
            fpSpread1.Sheets[0].Cells[13, 8].Text = Spec.MovingStroke.Height.ToString("0.0");
            fpSpread1.Sheets[0].Cells[14, 8].Text = Spec.MovingStroke.Recliner.ToString("0.0");

            fpSpread1.Sheets[0].Cells[23, 2].Value = Spec.DeliveryPos.Slide;
            fpSpread1.Sheets[0].Cells[24, 2].Value = Spec.DeliveryPos.Tilt;
            fpSpread1.Sheets[0].Cells[25, 2].Value = Spec.DeliveryPos.Height;
            fpSpread1.Sheets[0].Cells[26, 2].Value = Spec.DeliveryPos.Recliner;

            fpSpread1.Sheets[0].Cells[2, 8].Text = Spec.SlideTestTime.ToString("0.0");
            fpSpread1.Sheets[0].Cells[3, 8].Text = Spec.HeightTestTime.ToString("0.0");
            fpSpread1.Sheets[0].Cells[4, 8].Text = Spec.TiltTestTime.ToString("0.0");
            fpSpread1.Sheets[0].Cells[5, 8].Text = Spec.ReclinerTestTime.ToString("0.0");           
            fpSpread1.Sheets[0].Cells[6, 8].Text = Spec.LumberFwdBwdTestTime.ToString("0.0");
            fpSpread1.Sheets[0].Cells[7, 8].Text = Spec.LumberUpDnTestTime.ToString("0.0");

            fpSpread1.Sheets[0].Cells[16, 8].Text = Spec.Sound.StartTime.ToString("0.0");
            fpSpread1.Sheets[0].Cells[17, 8].Value = Spec.SlideSoundCheckTime.ToString("0.0");
            fpSpread1.Sheets[0].Cells[18, 8].Value = Spec.TiltSoundCheckTime.ToString("0.0");
            fpSpread1.Sheets[0].Cells[19, 8].Value = Spec.HeightSoundCheckTime.ToString("0.0");
            fpSpread1.Sheets[0].Cells[20, 8].Value = Spec.ReclinerSoundCheckTime.ToString("0.0");
            fpSpread1.Sheets[0].Cells[21, 8].Value = Spec.LumberSoundCheckTime.ToString("0.0");
            //fpSpread1.Sheets[0].Cells[22, 8].Value = Spec.Sound.SoundCheckRange.ToString("0.0");

            fpSpread1.Sheets[0].Cells[25, 8].Value = Spec.TestVolt.ToString("0.0");
            fpSpread1.Sheets[0].Cells[27, 8].Value = Spec.SlideLimitCurr.ToString("0.0");
            fpSpread1.Sheets[0].Cells[28, 8].Value = Spec.TiltLimitCurr.ToString("0.0");
            fpSpread1.Sheets[0].Cells[29, 8].Value = Spec.HeightLimitCurr.ToString("0.0");
            fpSpread1.Sheets[0].Cells[30, 8].Value = Spec.ReclinerLimitCurr.ToString("0.0");
            
            return;
        }
        private void MoveSpec()
        {
            InitSpec();
            Spec.PinMap = PinMap;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[3, 3].Text, out Spec.Current.IMS.Min) == false) Spec.Current.IMS.Min = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[3, 4].Text, out Spec.Current.IMS.Max) == false) Spec.Current.IMS.Max = 0;

            if (double.TryParse(fpSpread1.Sheets[0].Cells[4, 3].Text, out Spec.Current.Slide.Min) == false) Spec.Current.Slide.Min = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[4, 4].Text, out Spec.Current.Slide.Max) == false) Spec.Current.Slide.Max = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[5, 3].Text, out Spec.Current.Tilt.Min) == false) Spec.Current.Tilt.Min = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[5, 4].Text, out Spec.Current.Tilt.Max) == false) Spec.Current.Tilt.Max = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[6, 3].Text, out Spec.Current.Height.Min) == false) Spec.Current.Height.Min = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[6, 4].Text, out Spec.Current.Height.Max) == false) Spec.Current.Height.Max = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[7, 3].Text, out Spec.Current.Recliner.Min) == false) Spec.Current.Recliner.Min = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[7, 4].Text, out Spec.Current.Recliner.Max) == false) Spec.Current.Recliner.Max = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[8, 3].Text, out Spec.Current.LumberFwdBwd.Min) == false) Spec.Current.LumberFwdBwd.Min = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[8, 4].Text, out Spec.Current.LumberFwdBwd.Max) == false) Spec.Current.LumberFwdBwd.Max = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[9, 3].Text, out Spec.Current.LumberUpDn.Min) == false) Spec.Current.LumberUpDn.Min = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[9, 4].Text, out Spec.Current.LumberUpDn.Max) == false) Spec.Current.LumberUpDn.Max = 0;
            

            if (double.TryParse(fpSpread1.Sheets[0].Cells[10, 3].Text, out Spec.MovingSpeed.SlideFwd.Min) == false) Spec.MovingSpeed.SlideFwd.Min = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[10, 4].Text, out Spec.MovingSpeed.SlideFwd.Max) == false) Spec.MovingSpeed.SlideFwd.Max = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[11, 3].Text, out Spec.MovingSpeed.SlideBwd.Min) == false) Spec.MovingSpeed.SlideBwd.Min = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[11, 4].Text, out Spec.MovingSpeed.SlideBwd.Max) == false) Spec.MovingSpeed.SlideBwd.Max = 0;

            if (double.TryParse(fpSpread1.Sheets[0].Cells[12, 3].Text, out Spec.MovingSpeed.ReclinerFwd.Min) == false) Spec.MovingSpeed.ReclinerFwd.Min = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[12, 4].Text, out Spec.MovingSpeed.ReclinerFwd.Max) == false) Spec.MovingSpeed.ReclinerFwd.Max = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[13, 3].Text, out Spec.MovingSpeed.ReclinerBwd.Min) == false) Spec.MovingSpeed.ReclinerBwd.Min = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[13, 4].Text, out Spec.MovingSpeed.ReclinerBwd.Max) == false) Spec.MovingSpeed.ReclinerBwd.Max = 0;

            if (double.TryParse(fpSpread1.Sheets[0].Cells[14, 3].Text, out Spec.MovingSpeed.TiltUp.Min) == false) Spec.MovingSpeed.TiltUp.Min = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[14, 4].Text, out Spec.MovingSpeed.TiltUp.Max) == false) Spec.MovingSpeed.TiltUp.Max = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[15, 3].Text, out Spec.MovingSpeed.TiltDn.Min) == false) Spec.MovingSpeed.TiltDn.Min = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[15, 4].Text, out Spec.MovingSpeed.TiltDn.Max) == false) Spec.MovingSpeed.TiltDn.Max = 0;


            if (double.TryParse(fpSpread1.Sheets[0].Cells[16, 3].Text, out Spec.MovingSpeed.HeightUp.Min) == false) Spec.MovingSpeed.HeightUp.Min = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[16, 4].Text, out Spec.MovingSpeed.HeightUp.Max) == false) Spec.MovingSpeed.HeightUp.Max = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[17, 3].Text, out Spec.MovingSpeed.HeightDn.Min) == false) Spec.MovingSpeed.HeightDn.Min = 0;
            if (double.TryParse(fpSpread1.Sheets[0].Cells[17, 4].Text, out Spec.MovingSpeed.HeightDn.Max) == false) Spec.MovingSpeed.HeightDn.Max = 0;


            if (float.TryParse(fpSpread1.Sheets[0].Cells[18, 3].Text, out Spec.Sound.StartMax) == false) Spec.Sound.StartMax = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[18, 4].Text, out Spec.Sound.RunMax) == false) Spec.Sound.RunMax = 0;
            Spec.Sound.RMSMode = fpSpread1.Sheets[0].Cells[19, 4].Text == "True" ? true : false;            

            if (float.TryParse(fpSpread1.Sheets[0].Cells[20, 4].Text, out Spec.Sound.기동음구동음구분사이시간) == false) Spec.Sound.기동음구동음구분사이시간 = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[21, 4].Text, out Spec.Sound.Offset) == false) Spec.Sound.Offset = 0;

            if (float.TryParse(fpSpread1.Sheets[0].Cells[23, 4].Text, out Spec.SlideLimitTime) == false) Spec.SlideLimitTime = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[24, 4].Text, out Spec.TiltLimitTime) == false) Spec.TiltLimitTime = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[25, 4].Text, out Spec.HeightLimitTime) == false) Spec.HeightLimitTime = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[26, 4].Text, out Spec.ReclinerLimitTime) == false) Spec.ReclinerLimitTime = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[27, 4].Text, out Spec.LumberLimitTime) == false) Spec.LumberLimitTime = 0;
            

            //Spec.Can = fpSpread1.Sheets[0].Cells[8, 8].Text == "True" ? true : false;

            if (float.TryParse(fpSpread1.Sheets[0].Cells[11, 8].Text, out Spec.MovingStroke.Slide) == false) Spec.MovingStroke.Slide = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[12, 8].Text, out Spec.MovingStroke.Tilt) == false) Spec.MovingStroke.Tilt = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[13, 8].Text, out Spec.MovingStroke.Height) == false) Spec.MovingStroke.Height = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[14, 8].Text, out Spec.MovingStroke.Recliner) == false) Spec.MovingStroke.Recliner = 0;

            if (float.TryParse(fpSpread1.Sheets[0].Cells[2, 8].Text, out Spec.SlideTestTime) == false) Spec.SlideTestTime = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[3, 8].Text, out Spec.TiltTestTime) == false) Spec.TiltTestTime = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[4, 8].Text, out Spec.HeightTestTime) == false) Spec.HeightTestTime = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[5, 8].Text, out Spec.ReclinerTestTime) == false) Spec.ReclinerTestTime = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[6, 8].Text, out Spec.LumberFwdBwdTestTime) == false) Spec.LumberFwdBwdTestTime = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[7, 8].Text, out Spec.LumberUpDnTestTime) == false) Spec.LumberUpDnTestTime = 0;
            

            var s1 = fpSpread1.Sheets[0].Cells[23, 2].Text;
            var s2 = fpSpread1.Sheets[0].Cells[24, 2].Text;
            var s3 = fpSpread1.Sheets[0].Cells[25, 2].Text;
            var s4 = fpSpread1.Sheets[0].Cells[26, 2].Text;

            if((string)s1 == "FWD")
                Spec.DeliveryPos.Slide = 0;
            else if ((string)s1 == "BWD")
                Spec.DeliveryPos.Slide = 2;
            else Spec.DeliveryPos.Slide = 1;

            Spec.DeliveryPos.Tilt = (string)s2 == "UP" ? (short)0 : (short)1;
            Spec.DeliveryPos.Height = (string)s3 == "UP" ? (short)0 : (short)1;
            Spec.DeliveryPos.Recliner = (string)s4 == "FWD" ? (short)0 : (short)1;

            if (float.TryParse(fpSpread1.Sheets[0].Cells[16, 8].Text, out Spec.Sound.StartTime) == false) Spec.Sound.StartTime = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[17, 8].Text, out Spec.SlideSoundCheckTime) == false) Spec.SlideSoundCheckTime = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[18, 8].Text, out Spec.TiltSoundCheckTime) == false) Spec.TiltSoundCheckTime = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[19, 8].Text, out Spec.HeightSoundCheckTime) == false) Spec.HeightSoundCheckTime = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[20, 8].Text, out Spec.ReclinerSoundCheckTime) == false) Spec.ReclinerSoundCheckTime = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[21, 8].Text, out Spec.LumberSoundCheckTime) == false) Spec.LumberSoundCheckTime = 0;            

            //if (float.TryParse(fpSpread1.Sheets[0].Cells[22, 8].Text, out Spec.Sound.SoundCheckRange) == false) Spec.Sound.SoundCheckRange = 0;

            if (float.TryParse(fpSpread1.Sheets[0].Cells[25, 8].Text, out Spec.TestVolt) == false) Spec.TestVolt = 13.6F;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[27, 8].Text, out Spec.SlideLimitCurr) == false) Spec.SlideLimitCurr = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[28, 8].Text, out Spec.TiltLimitCurr) == false) Spec.TiltLimitCurr = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[29, 8].Text, out Spec.HeightLimitCurr) == false) Spec.HeightLimitCurr = 0;
            if (float.TryParse(fpSpread1.Sheets[0].Cells[30, 8].Text, out Spec.ReclinerLimitCurr) == false) Spec.ReclinerLimitCurr = 0;
            
            return;
        }
        public void InitSpec()
        {
            Spec.ModelName = string.Empty;
            
            //Spec.PinMap.IMSSet.Mode = 0;

            //Spec.PinMap.IMSM1.PinNo = 0;
            //Spec.PinMap.IMSM1.Mode = 0;

            //Spec.PinMap.IMSM2.PinNo = 0;
            //Spec.PinMap.IMSM2.Mode = 0;

            Spec.PinMap.SlideFWD.PinNo = 0;
            Spec.PinMap.SlideFWD.Mode = 0;

            Spec.PinMap.SlideBWD.PinNo = 0;
            Spec.PinMap.SlideBWD.Mode = 0;

            Spec.PinMap.ReclineFWD.PinNo = 0;
            Spec.PinMap.ReclineFWD.Mode = 0;

            Spec.PinMap.ReclineBWD.PinNo = 0;
            Spec.PinMap.ReclineBWD.Mode = 0;

            Spec.PinMap.TiltUp.PinNo = 0;
            Spec.PinMap.TiltUp.Mode = 0;

            Spec.PinMap.TiltDn.PinNo = 0;
            Spec.PinMap.TiltDn.Mode = 0;

            Spec.PinMap.HeightUp.PinNo = 0;
            Spec.PinMap.HeightUp.Mode = 0;

            Spec.PinMap.HeightDn.PinNo = 0;
            Spec.PinMap.HeightDn.Mode = 0;

            Spec.PinMap.Power.Batt1.Batt = 0;
            Spec.PinMap.Power.Batt1.Gnd = 0;

            Spec.PinMap.Power.Batt2.Batt = 0;
            Spec.PinMap.Power.Batt2.Gnd = 0;

            Spec.PinMap.LumberFwdBwd.Batt = 0;
            Spec.PinMap.LumberFwdBwd.Gnd = 0;

            Spec.PinMap.LumberUpDn.Batt = 0;
            Spec.PinMap.LumberUpDn.Gnd = 0;


            Spec.Current.IMS.Min = 0;
            Spec.Current.IMS.Max = 0;
            Spec.Current.Slide.Min = 0;
            Spec.Current.Slide.Max = 0;
            Spec.Current.Tilt.Min = 0;
            Spec.Current.Tilt.Max = 0;
            Spec.Current.Height.Min = 0;
            Spec.Current.Height.Max = 0;
            Spec.Current.Recliner.Min = 0;
            Spec.Current.Recliner.Max = 0;
            Spec.Current.LumberFwdBwd.Min = 0;
            Spec.Current.LumberFwdBwd.Max = 0;
            Spec.Current.LumberUpDn.Min = 0;
            Spec.Current.LumberUpDn.Max = 0;

            Spec.MovingSpeed.SlideFwd.Min = 0;
            Spec.MovingSpeed.SlideFwd.Max = 0;
            Spec.MovingSpeed.SlideBwd.Min = 0;
            Spec.MovingSpeed.SlideBwd.Max = 0;
            Spec.MovingSpeed.TiltUp.Min = 0;
            Spec.MovingSpeed.TiltUp.Max = 0;
            Spec.MovingSpeed.TiltDn.Min = 0;
            Spec.MovingSpeed.TiltDn.Max = 0;
            Spec.MovingSpeed.HeightUp.Min = 0;
            Spec.MovingSpeed.HeightUp.Max = 0;
            Spec.MovingSpeed.HeightDn.Min = 0;
            Spec.MovingSpeed.HeightDn.Max = 0;
            Spec.MovingSpeed.ReclinerFwd.Min = 0;
            Spec.MovingSpeed.ReclinerFwd.Max = 0;
            Spec.MovingSpeed.ReclinerBwd.Min = 0;
            Spec.MovingSpeed.ReclinerBwd.Max = 0;

            Spec.MovingStroke.Slide = 0F;
            Spec.MovingStroke.Tilt = 0F;
            Spec.MovingStroke.Height = 0F;
            Spec.MovingStroke.Recliner = 0F;
            Spec.Sound.RMSMode = false;
            Spec.Sound.기동음구동음구분사이시간 = 0;
            Spec.Sound.StartMax = 0;
            Spec.Sound.RunMax = 0;
            Spec.Sound.StartTime = 0;
            Spec.Sound.Offset = 0;
            Spec.SlideLimitTime = 5;
            Spec.TiltLimitTime = 5;
            Spec.HeightLimitTime = 5;
            Spec.ReclinerLimitTime = 5;
            
            //Spec.Can = false;

            Spec.SlideTestTime = 0;
            Spec.TiltTestTime = 0;
            Spec.HeightTestTime = 0;
            Spec.ReclinerTestTime = 0;
            Spec.LumberFwdBwdTestTime = 0;
            Spec.LumberUpDnTestTime = 0;

            Spec.DeliveryPos.Slide = 0;
            Spec.DeliveryPos.Tilt = 0;
            Spec.DeliveryPos.Height = 0;
            Spec.DeliveryPos.Recliner = 0;

            //Spec.SoundCheckTimeRange = 0;
            Spec.SlideSoundCheckTime = 0;
            Spec.TiltSoundCheckTime = 0;
            Spec.HeightSoundCheckTime = 0;
            Spec.ReclinerSoundCheckTime = 0;

            Spec.SlideLimitCurr = 0;
            Spec.TiltLimitCurr = 0;
            Spec.HeightLimitCurr = 0;
            Spec.ReclinerLimitCurr = 0;

            Spec.TestVolt = 13.6F;
            return;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ModelBoxChangeFlag == true) return;
            string s = comboBox1.SelectedItem.ToString();

            string sName = Program.SPEC_PATH.ToString() + "\\" + s + ".Spc";

            InitSpec();
            if (File.Exists(sName) == true) mControl.공용함수.OpenSpec(sName, ref Spec);
            mName = Spec.ModelName;
            DisplaySpec();
            return;
        }
    }
}
