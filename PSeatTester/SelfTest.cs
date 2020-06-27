using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PSeatTester
{
    public partial class SelfTest : Form
    {
        private MyInterface mControl = null;
        public SelfTest()
        {
            InitializeComponent();
        }
        public SelfTest(MyInterface mControl)
        {
            InitializeComponent();
            this.mControl = mControl;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox Chk = sender as CheckBox;
            string sTag = Chk.Tag.ToString();

            int Tag = 0;
            if (int.TryParse(sTag, out Tag) == false) Tag = -1;

            if (Tag != -1)
            {
                if (Chk.Checked == true)
                {
                    switch (Tag)
                    {
                        case 20: checkBox21.Checked = false; break;
                        case 19: checkBox22.Checked = false; break;
                        case 18: checkBox23.Checked = false; break;
                        case 17: checkBox24.Checked = false; break;
                        case 16: checkBox25.Checked = false; break;
                        case 15: checkBox26.Checked = false; break;
                        case 14: checkBox27.Checked = false; break;
                        case 13: checkBox28.Checked = false; break;
                        case 12: checkBox29.Checked = false; break;
                        case 11: checkBox30.Checked = false; break;
                        case 10: checkBox31.Checked = false; break;
                        case 9: checkBox32.Checked = false; break;
                        case 8: checkBox33.Checked = false; break;
                        case 7: checkBox34.Checked = false; break;
                        case 6: checkBox35.Checked = false; break;
                        case 5: checkBox36.Checked = false; break;
                        case 4: checkBox37.Checked = false; break;
                        case 3: checkBox38.Checked = false; break;
                        case 2: checkBox39.Checked = false; break;
                        case 1: checkBox40.Checked = false; break;
                    }
                }

                switch (Tag)
                {
                    case 1: ledBulb30.On = Chk.Checked; break;
                    case 2: ledBulb31.On = Chk.Checked; break;
                    case 3: ledBulb32.On = Chk.Checked; break;
                    case 4: ledBulb33.On = Chk.Checked; break;
                    case 5: ledBulb34.On = Chk.Checked; break;
                    case 6: ledBulb35.On = Chk.Checked; break;
                    case 7: ledBulb36.On = Chk.Checked; break;
                    case 8: ledBulb37.On = Chk.Checked; break;
                    case 9: ledBulb38.On = Chk.Checked; break;
                    case 10: ledBulb39.On = Chk.Checked; break;
                    case 11: ledBulb40.On = Chk.Checked; break;
                    case 12: ledBulb41.On = Chk.Checked; break;
                    case 13: ledBulb42.On = Chk.Checked; break;
                    case 14: ledBulb43.On = Chk.Checked; break;
                    case 15: ledBulb44.On = Chk.Checked; break;
                    case 16: ledBulb45.On = Chk.Checked; break;
                    case 17: ledBulb46.On = Chk.Checked; break;
                    case 18: ledBulb47.On = Chk.Checked; break;
                    case 19: ledBulb48.On = Chk.Checked; break;
                    case 20: ledBulb49.On = Chk.Checked; break;
                }

                mControl.GetIO.PinSelectToBattOnOff((short)(Tag - 1), Chk.Checked);
            }
            return;
        }

        private void checkBox40_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox Chk = sender as CheckBox;
            string sTag = Chk.Tag.ToString();

            int Tag;
            if (int.TryParse(sTag, out Tag) == false) Tag = -1;
            if (Tag != -1)
            {
                if (Chk.Checked == true)
                {
                    switch (Tag)
                    {
                        case 20: checkBox1.Checked = false; break;
                        case 19: checkBox2.Checked = false; break;
                        case 18: checkBox3.Checked = false; break;
                        case 17: checkBox4.Checked = false; break;
                        case 16: checkBox5.Checked = false; break;
                        case 15: checkBox6.Checked = false; break;
                        case 14: checkBox7.Checked = false; break;
                        case 13: checkBox8.Checked = false; break;
                        case 12: checkBox9.Checked = false; break;
                        case 11: checkBox10.Checked = false; break;
                        case 10: checkBox11.Checked = false; break;
                        case 9: checkBox12.Checked = false; break;
                        case 8: checkBox13.Checked = false; break;
                        case 7: checkBox14.Checked = false; break;
                        case 6: checkBox15.Checked = false; break;
                        case 5: checkBox16.Checked = false; break;
                        case 4: checkBox17.Checked = false; break;
                        case 3: checkBox18.Checked = false; break;
                        case 2: checkBox19.Checked = false; break;
                        case 1: checkBox20.Checked = false; break;
                    }
                }

                switch (Tag)
                {
                    case 1: ledBulb50.On = Chk.Checked; break;
                    case 2: ledBulb51.On = Chk.Checked; break;
                    case 3: ledBulb52.On = Chk.Checked; break;
                    case 4: ledBulb53.On = Chk.Checked; break;
                    case 5: ledBulb54.On = Chk.Checked; break;
                    case 6: ledBulb55.On = Chk.Checked; break;
                    case 7: ledBulb56.On = Chk.Checked; break;
                    case 8: ledBulb57.On = Chk.Checked; break;
                    case 9: ledBulb58.On = Chk.Checked; break;
                    case 10: ledBulb59.On = Chk.Checked; break;
                    case 11: ledBulb60.On = Chk.Checked; break;
                    case 12: ledBulb61.On = Chk.Checked; break;
                    case 13: ledBulb62.On = Chk.Checked; break;
                    case 14: ledBulb63.On = Chk.Checked; break;
                    case 15: ledBulb64.On = Chk.Checked; break;
                    case 16: ledBulb65.On = Chk.Checked; break;
                    case 17: ledBulb66.On = Chk.Checked; break;
                    case 18: ledBulb67.On = Chk.Checked; break;
                    case 19: ledBulb68.On = Chk.Checked; break;
                    case 20: ledBulb69.On = Chk.Checked; break;
                }
                mControl.GetIO.PinSelectToGndOnOff((short)(Tag - 1), Chk.Checked);
            }
            return;
        }

        private void switchLever1_ValueChanged(object sender, Iocomp.Classes.ValueBooleanEventArgs e)
        {
            ledBulb70.On = e.ValueNew;

            if (e.ValueNew == true)
                mControl.GetPower.POWER_PWON();
            else mControl.GetPower.POWER_PWOFF();

            return;
        }

        private bool TextChange { get; set; }
        private bool PowerChnageFlag { get; set; }
        private bool PowerTextChange { get; set; }
        private long PowerChangeTimeToFirst { get; set; }

        private void knob1_ValueChanged(object sender, Iocomp.Classes.ValueDoubleEventArgs e)
        {
            if (TextChange == true) return;
            string Value = e.ValueNew.ToString("0.00") + " [V]";

            if (textBox1.Text != Value) textBox1.Text = Value;
            if (PowerTextChange == true) return;

            PowerChangeTimeToFirst = mControl.공용함수.timeGetTimems();
            PowerChnageFlag = true;
            return;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                timer1.Enabled = false;

                if (PowerChnageFlag == true)
                {
                    if (100 <= (mControl.공용함수.timeGetTimems() - PowerChangeTimeToFirst))
                    {
                        PowerChnageFlag = false;
                        mControl.GetPower.POWER_PWSetting((float)knob1.Value.AsDouble);
                    }
                }

                DisplayIOIn();
                DisplayAD();
            }
            catch { }
            finally { timer1.Enabled = !mControl.isExit; }
        }

        private void DisplayIOIn()
        {
            if (ledBulb2.On != mControl.GetIO.GetStartSw) ledBulb2.On = mControl.GetIO.GetStartSw;
            if (ledBulb3.On != mControl.GetIO.GetResetSw) ledBulb3.On = mControl.GetIO.GetResetSw;
            if (ledBulb4.On != mControl.GetIO.GetRHSelect) ledBulb4.On = mControl.GetIO.GetRHSelect;
            if (ledBulb5.On != mControl.GetIO.GetIMS) ledBulb5.On = mControl.GetIO.GetIMS;
            if (ledBulb6.On != mControl.GetIO.GetSlideMidSensor) ledBulb6.On = mControl.GetIO.GetSlideMidSensor;
            if (ledBulb1.On != mControl.GetIO.GetReclineMidSensor) ledBulb1.On = mControl.GetIO.GetReclineMidSensor;
            if (ledBulb12.On != mControl.GetIO.GetAuto) ledBulb12.On = mControl.GetIO.GetAuto;
            if (ledBulb16.On != mControl.GetIO.GetProductIn) ledBulb16.On = mControl.GetIO.GetProductIn;
            if (ledBulb17.On != mControl.GetIO.GetJigUp) ledBulb17.On = mControl.GetIO.GetJigUp;

            if (ledBulb8.On != mControl.GetIO.GetPinConnectionSw) ledBulb8.On = mControl.GetIO.GetPinConnectionSw;
            if (ledBulb7.On != mControl.GetIO.GetPinConnectionFwd) ledBulb7.On = mControl.GetIO.GetPinConnectionFwd;
            if (ledBulb9.On != mControl.GetIO.GetPinConnectionBwd) ledBulb9.On = mControl.GetIO.GetPinConnectionBwd;
            if (ledBulb10.On != mControl.GetIO.GetRHDSelect) ledBulb13.On = mControl.GetIO.GetRHDSelect;
            if (ledBulb11.On != mControl.GetIO.GetLumber2Way4WaySelect) ledBulb11.On = mControl.GetIO.GetLumber2Way4WaySelect;
            if (ledBulb75.On != mControl.GetIO.GetSeatPower) ledBulb75.On = mControl.GetIO.GetSeatPower;

            if (ledBulb74.On != mControl.GetIO.GetIMSSet) ledBulb74.On = mControl.GetIO.GetIMSSet;
            if (ledBulb73.On != mControl.GetIO.GetM1Sw) ledBulb73.On = mControl.GetIO.GetM1Sw;
            if (ledBulb72.On != mControl.GetIO.GetM2Sw) ledBulb72.On = mControl.GetIO.GetM2Sw;
            if (ledBulb71.On != mControl.GetIO.GetWalkIn) ledBulb71.On = mControl.GetIO.GetWalkIn;

            if (ledBulb18.On != mControl.GetIO.GetSlideFwd) ledBulb18.On = mControl.GetIO.GetSlideFwd;
            if (ledBulb19.On != mControl.GetIO.GetSlideBwd) ledBulb19.On = mControl.GetIO.GetSlideBwd;
            if (ledBulb20.On != mControl.GetIO.GetReclinerFwd) ledBulb20.On = mControl.GetIO.GetReclinerFwd;
            if (ledBulb21.On != mControl.GetIO.GetReclinerBwd) ledBulb21.On = mControl.GetIO.GetReclinerBwd;
            if (ledBulb22.On != mControl.GetIO.GetTiltUp) ledBulb22.On = mControl.GetIO.GetTiltUp;
            if (ledBulb23.On != mControl.GetIO.GetTiltDn) ledBulb23.On = mControl.GetIO.GetTiltDn;
            if (ledBulb24.On != mControl.GetIO.GetHeightUp) ledBulb24.On = mControl.GetIO.GetHeightUp;
            if (ledBulb25.On != mControl.GetIO.GetHeightDn) ledBulb25.On = mControl.GetIO.GetHeightDn;
            if (ledBulb26.On != mControl.GetIO.GetLumberFwd) ledBulb26.On = mControl.GetIO.GetLumberFwd;
            if (ledBulb27.On != mControl.GetIO.GetLumberBwd) ledBulb27.On = mControl.GetIO.GetLumberBwd;
            if (ledBulb28.On != mControl.GetIO.GetLumberUp) ledBulb28.On = mControl.GetIO.GetLumberUp;
            if (ledBulb29.On != mControl.GetIO.GetLumberDn) ledBulb29.On = mControl.GetIO.GetLumberDn;

            return;
        }
        private void DisplayAD()
        {
            if (fpSpread1.ActiveSheet.Cells[2, 1].Text != mControl.GetIO.ADRead[0].ToString("0.00")) fpSpread1.ActiveSheet.Cells[2,1].Text = mControl.GetIO.ADRead[0].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[3, 1].Text != mControl.GetIO.ADRead[1].ToString("0.00")) fpSpread1.ActiveSheet.Cells[3, 1].Text = mControl.GetIO.ADRead[1].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[4, 1].Text != mControl.GetIO.ADRead[2].ToString("0.00")) fpSpread1.ActiveSheet.Cells[4, 1].Text = mControl.GetIO.ADRead[2].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[5, 1].Text != mControl.GetIO.ADRead[3].ToString("0.00")) fpSpread1.ActiveSheet.Cells[5, 1].Text = mControl.GetIO.ADRead[3].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[6, 1].Text != mControl.GetIO.ADRead[4].ToString("0.00")) fpSpread1.ActiveSheet.Cells[6, 1].Text = mControl.GetIO.ADRead[4].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[7, 1].Text != mControl.GetIO.ADRead[5].ToString("0.00")) fpSpread1.ActiveSheet.Cells[7, 1].Text = mControl.GetIO.ADRead[5].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[8, 1].Text != mControl.GetIO.ADRead[6].ToString("0.00")) fpSpread1.ActiveSheet.Cells[8, 1].Text = mControl.GetIO.ADRead[6].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[9, 1].Text != mControl.GetIO.ADRead[7].ToString("0.00")) fpSpread1.ActiveSheet.Cells[9, 1].Text = mControl.GetIO.ADRead[7].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[10, 1].Text != mControl.GetIO.ADRead[8].ToString("0.00")) fpSpread1.ActiveSheet.Cells[10, 1].Text = mControl.GetIO.ADRead[8].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[11, 1].Text != mControl.GetIO.ADRead[9].ToString("0.00")) fpSpread1.ActiveSheet.Cells[11, 1].Text = mControl.GetIO.ADRead[9].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[12, 1].Text != mControl.GetIO.ADRead[10].ToString("0.00")) fpSpread1.ActiveSheet.Cells[12, 1].Text = mControl.GetIO.ADRead[10].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[13, 1].Text != mControl.GetIO.ADRead[11].ToString("0.00")) fpSpread1.ActiveSheet.Cells[13, 1].Text = mControl.GetIO.ADRead[11].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[14, 1].Text != mControl.GetIO.ADRead[12].ToString("0.00")) fpSpread1.ActiveSheet.Cells[14, 1].Text = mControl.GetIO.ADRead[12].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[15, 1].Text != mControl.GetIO.ADRead[13].ToString("0.00")) fpSpread1.ActiveSheet.Cells[15, 1].Text = mControl.GetIO.ADRead[13].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[16, 1].Text != mControl.GetIO.ADRead[14].ToString("0.00")) fpSpread1.ActiveSheet.Cells[16, 1].Text = mControl.GetIO.ADRead[14].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[17, 1].Text != mControl.GetIO.ADRead[15].ToString("0.00")) fpSpread1.ActiveSheet.Cells[17, 1].Text = mControl.GetIO.ADRead[15].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[18, 1].Text != mControl.GetIO.ADRead[16].ToString("0.00")) fpSpread1.ActiveSheet.Cells[18, 1].Text = mControl.GetIO.ADRead[16].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[19, 1].Text != mControl.GetIO.ADRead[17].ToString("0.00")) fpSpread1.ActiveSheet.Cells[19, 1].Text = mControl.GetIO.ADRead[17].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[20, 1].Text != mControl.GetIO.ADRead[18].ToString("0.00")) fpSpread1.ActiveSheet.Cells[20, 1].Text = mControl.GetIO.ADRead[18].ToString("0.00");
            if (fpSpread1.ActiveSheet.Cells[21, 1].Text != mControl.GetIO.ADRead[19].ToString("0.00")) fpSpread1.ActiveSheet.Cells[21, 1].Text = mControl.GetIO.ADRead[19].ToString("0.00");

            plot1.Channels[0].AddXY(plot1.Channels[0].Count, mControl.GetIO.ADRead[SelectIndex]);

            if (sevenSegmentAnalog2.Value.AsDouble != mControl.GetPMeter.GetBatt) sevenSegmentAnalog2.Value.AsDouble = mControl.GetPMeter.GetBatt;
            if (sevenSegmentAnalog5.Value.AsDouble != mControl.GetPMeter.GetPSeat) sevenSegmentAnalog5.Value.AsDouble = mControl.GetPMeter.GetPSeat;
            if (sevenSegmentAnalog1.Value.AsDouble != mControl.GetSound.GetSound) sevenSegmentAnalog1.Value.AsDouble = mControl.GetSound.GetSound;
            return;
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {            
            if(e.KeyChar == (char)ConsoleKey.Enter)
            {
                TextChange = true;
                string s = textBox1.Text;

                if(0 <= s.IndexOf("["))
                {
                    s = s.Substring(0, s.IndexOf("[") - 1);
                }
                float Volt;

                if (float.TryParse(s, out Volt) == false) Volt = -1;
                if(Volt != -1) mControl.GetPower.POWER_PWSetting(Volt);
                knob1.Value.AsDouble = Volt;
                TextChange = false;
            }
            return;
        }

        private int SelectIndex = 0;
        private void SelfTest_Load(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            SelectIndex = comboBox1.SelectedIndex;
            comboBox1.SelectedIndex = 0;
            return;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            plot1.Channels[0].Clear();
            SelectIndex = comboBox1.SelectedIndex;
            return;
        }

        private void imageButton1_Click(object sender, EventArgs e)
        {
            //제품감지
            UserImageButton.ImageButton But = sender as UserImageButton.ImageButton;

            if (But.ButtonColor == Color.DeepPink)
                But.ButtonColor = Color.Black;
            else But.ButtonColor = Color.DeepPink;

            mControl.GetIO.ProductInOut = (But.ButtonColor == Color.DeepPink) ? true : false;

            return;
        }

        private void imageButton2_Click(object sender, EventArgs e)
        {
            //TEST OK
            UserImageButton.ImageButton But = sender as UserImageButton.ImageButton;

            if (But.ButtonColor == Color.DeepPink)
                But.ButtonColor = Color.Black;
            else But.ButtonColor = Color.DeepPink;

            mControl.GetIO.TestOKOnOff = (But.ButtonColor == Color.DeepPink) ? true : false;
        }

        private void imageButton3_Click(object sender, EventArgs e)
        {
            //TEST ING
            UserImageButton.ImageButton But = sender as UserImageButton.ImageButton;

            if (But.ButtonColor == Color.DeepPink)
                But.ButtonColor = Color.Black;
            else But.ButtonColor = Color.DeepPink;

            mControl.GetIO.TestINGOnOff = (But.ButtonColor == Color.DeepPink) ? true : false;
        }

        private void imageButton6_Click(object sender, EventArgs e)
        {
            UserImageButton.ImageButton But = sender as UserImageButton.ImageButton;

            if (But.ButtonColor == Color.Yellow)
                But.ButtonColor = Color.Black;
            else But.ButtonColor = Color.Yellow;

            mControl.GetIO.YellowLampOnOff = (But.ButtonColor == Color.Yellow) ? true : false;
            return;
        }

        private void imageButton5_Click(object sender, EventArgs e)
        {
            UserImageButton.ImageButton But = sender as UserImageButton.ImageButton;

            if (But.ButtonColor == Color.Lime)
                But.ButtonColor = Color.Black;
            else But.ButtonColor = Color.Lime;

            mControl.GetIO.GreenLampOnOff = (But.ButtonColor == Color.Lime) ? true : false;
            return;
        }

        private void imageButton4_Click(object sender, EventArgs e)
        {
            UserImageButton.ImageButton But = sender as UserImageButton.ImageButton;

            if (But.ButtonColor == Color.Red)
                But.ButtonColor = Color.Black;
            else But.ButtonColor = Color.Red;

            mControl.GetIO.RedLampOnOff = (But.ButtonColor == Color.Red) ? true : false;
            return;
        }

        private void imageButton7_Click(object sender, EventArgs e)
        {
            UserImageButton.ImageButton But = sender as UserImageButton.ImageButton;

            if (But.ButtonColor == Color.DeepPink)
                But.ButtonColor = Color.Black;
            else But.ButtonColor = Color.DeepPink;

            mControl.GetIO.BuzzerOnOff = (But.ButtonColor == Color.DeepPink) ? true : false;
            return;
        }

        private void imageButton8_Click(object sender, EventArgs e)
        {
            UserImageButton.ImageButton But = sender as UserImageButton.ImageButton;

            if (But.ButtonColor == Color.DeepPink)
                But.ButtonColor = Color.Black;
            else But.ButtonColor = Color.DeepPink;

            mControl.GetIO.SetPinConnection = (But.ButtonColor == Color.DeepPink) ? true : false;
            return; 
        }

        private void imageButton9_Click(object sender, EventArgs e)
        {
            UserImageButton.ImageButton But = sender as UserImageButton.ImageButton;

            if (But.ButtonColor == Color.Red)
                But.ButtonColor = Color.Black;
            else But.ButtonColor = Color.Red;

            if (But.ButtonColor == Color.Red)
            {
                knob1.Value.AsDouble = 13.6F;
                switchLever1.Value.AsBoolean = true;
            }
            else
            {
                knob1.Value.AsDouble = 0F;
                switchLever1.Value.AsBoolean = false;
            }
            return;
        }

       

        private void imageButton10_Click(object sender, EventArgs e)
        {
            UserImageButton.ImageButton But = sender as UserImageButton.ImageButton;

            if (But.ButtonColor == Color.Red)
                But.ButtonColor = Color.Black;
            else But.ButtonColor = Color.Red;

            if (But.ButtonColor == Color.Red)
            {
                mControl.GetIO.SetDoorOpen  = true;
            }
            else
            {
                mControl.GetIO.SetDoorOpen= false;
            }
            return;
        }
    }
}
