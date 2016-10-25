using Microsoft.Win32;
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
    public partial class SerialSet : Form
    {
        public SerialSet()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Config.Default.Serial1 = comboBox1.Text;
            Config.Default.Serial2 = comboBox2.Text;
            Config.Default.Serial3 = comboBox3.Text;
            Config.Default.Serial4 = comboBox4.Text;
            Config.Default.Serial5 = comboBox5.Text;
            Config.Default.SerialA = comboBoxA.Text;
            Config.Default.SerialB = comboBoxB.Text;
            Config.Default.Save();
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void SerialSet_Load(object sender, EventArgs e)
        {
            comboBox1.Text = Config.Default.Serial1;
            comboBox2.Text = Config.Default.Serial2;
            comboBox3.Text = Config.Default.Serial3;
            comboBox4.Text = Config.Default.Serial4;
            comboBox5.Text = Config.Default.Serial5;
            comboBoxA.Text = Config.Default.SerialA;
            comboBoxB.Text = Config.Default.SerialB;
            GetComList();
        }
        public void GetComList()
        {
            RegistryKey keyCom = Registry.LocalMachine.OpenSubKey(@"Hardware\DeviceMap\SerialComm");
            if (keyCom != null)
            {
                string[] sSubKeys = keyCom.GetValueNames();
                this.comboBox1.Items.Clear();
                this.comboBox2.Items.Clear();
                this.comboBox3.Items.Clear();
                this.comboBox4.Items.Clear();
                this.comboBox5.Items.Clear();
                this.comboBoxA.Items.Clear();
                this.comboBoxB.Items.Clear();
                foreach (string sName in sSubKeys)
                {
                    string sValue = (string)keyCom.GetValue(sName);
                    this.comboBox1.Items.Add(sValue);
                    this.comboBox2.Items.Add(sValue);
                    this.comboBox3.Items.Add(sValue);
                    this.comboBox4.Items.Add(sValue);
                    this.comboBox5.Items.Add(sValue);
                    this.comboBoxA.Items.Add(sValue);
                    this.comboBoxB.Items.Add(sValue);
                }
            }
        }
    }
}
