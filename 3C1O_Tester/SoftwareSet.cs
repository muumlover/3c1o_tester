using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace _3C1O_Tester
{
    public partial class SoftwareSet : Form
    {
        string[] HA = new string[10];
        public SoftwareSet()
        {
            InitializeComponent();
            HA[0] = "只要你相信，奇迹一定会实现";
            HA[1] = "用爱心来做事，用感恩的心做人";
            HA[2] = "励志是给人快乐，激励是给人痛苦";
            HA[3] = "成功者绝不给自己软弱的借口";
            HA[4] = "你只有一定要，才一定会得到";
            HA[5] = "成功者绝不放弃";
            HA[6] = "成功永远属于马上行动的人";
            HA[7] = "命运是可以改变的";
            HA[8] = "决心是成功的开始";
            HA[9] = "如果你相信自己，你可以做任何事";
            label28.Text = HA[new Random().Next(1, 10)];
        }
        private void SoftwareSet_Load(object sender, EventArgs e)
        {
            textBox1.Text = Config.Default.CurrentErrorPercentageSemi.ToString();
            textBox2.Text = Config.Default.CurrentErrorValueSemi.ToString();
            textBox6.Text = Config.Default.CurrentErrorPercentage.ToString();
            textBox7.Text = Config.Default.CurrentErrorValue.ToString();
            textBox9.Text = Config.Default.FullCalibrationThreshold.ToString();
            //textBox10.Text = Config.Default.A_Dev2Count.ToString();
            textBox8.Text = Config.Default.TimesForErrorCorrection.ToString();
            checkBox1.Checked = Config.Default.ReadAfterSet;
            textBox5.Text = Config.Default.CountForStandardCurrentRetry.ToString();
            textBox11.Text = Config.Default.ErrorPercentageForStandardCurrent.ToString();
            textBox3.Text = Config.Default.TimeForStandardCurrentStability.ToString();
            textBox14.Text = Config.Default.A_SpeedCurRead.ToString();
            textBox4.Text = Config.Default.TimeForWaitFIRead.ToString();
            textBox12.Text = Config.Default.TimeForWaitFICalibration.ToString();
            textBox13.Text = Config.Default.TimeForLEDBlinkInterval.ToString();
            textBox15.Text = Config.Default.TimeForReadAfterWrite.ToString();
            textBox16.Text = Config.Default.TimeForFive.ToString();
            if (Config.Default.Mode == "3C1O")
                radioButton4.Checked = true;
            else if (Config.Default.Mode == "3C1O-T")
                radioButton5.Checked = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Config.Default.CurrentErrorPercentageSemi = double.Parse(textBox1.Text);
            Config.Default.CurrentErrorValueSemi = double.Parse(textBox2.Text);
            Config.Default.FullCalibrationThreshold = int.Parse(textBox9.Text);
            //Config.Default.A_Dev2Count = int.Parse(textBox10.Text);
            Config.Default.TimesForErrorCorrection = int.Parse(textBox8.Text);
            Config.Default.ReadAfterSet = checkBox1.Checked;
            Config.Default.CurrentErrorPercentage = int.Parse(textBox6.Text);
            Config.Default.CountForStandardCurrentRetry = int.Parse(textBox5.Text);
            Config.Default.ErrorPercentageForStandardCurrent = int.Parse(textBox11.Text);
            Config.Default.TimeForStandardCurrentStability = int.Parse(textBox3.Text);
            Config.Default.A_SpeedCurRead = int.Parse(textBox14.Text);
            Config.Default.TimeForWaitFIRead = int.Parse(textBox4.Text);
            Config.Default.CurrentErrorValue = int.Parse(textBox7.Text);
            Config.Default.TimeForWaitFICalibration = int.Parse(textBox12.Text);
            Config.Default.TimeForLEDBlinkInterval = int.Parse(textBox13.Text);
            Config.Default.TimeForReadAfterWrite = int.Parse(textBox15.Text);
            Config.Default.TimeForFive = int.Parse(textBox16.Text);
            Config.Default.Save();
        }

        private void radioButton4_Click(object sender, EventArgs e)
        {
            radioButton4.Checked = true;
            radioButton5.Checked = false;
            Config.Default.Mode = "3C1O";
            Config.Default.Save();
        }

        private void radioButton5_Click(object sender, EventArgs e)
        {
            radioButton5.Checked = true;
            radioButton4.Checked = false;
            Config.Default.Mode = "3C1O-T";
            Config.Default.Save();
        }

        private void label13_Click(object sender, EventArgs e)
        {

        }
    }
}
