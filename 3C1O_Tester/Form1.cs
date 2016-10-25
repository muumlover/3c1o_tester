using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;


namespace _3C1O_Tester
{
    public partial class Form1 : Form
    {
        M_Log log = new M_Log();
        public Form1(DialogResult b)
        {
            InitializeComponent();
            this.Height = 680;
            this.Width = 960;

            this.IsMdiContainer = true;
            FileVersionInfo myFileVersion = FileVersionInfo.GetVersionInfo(System.Windows.Forms.Application.ExecutablePath);
            this.Text += " 软件版本：" + myFileVersion.FileVersion;

            if (b == DialogResult.Yes)
            {
                designToolStripMenuItem.Visible = false;
                floor1TestToolStripMenuItem.Visible = false;
                floor1StartToolStripMenuItem.Visible = false;
                floor1StopToolStripMenuItem.Visible = false;
                floor2TestToolStripMenuItem.Visible = false;
                floor2StartToolStripMenuItem.Visible = false;
                floor2StopToolStripMenuItem.Visible = false;
                floor3TestToolStripMenuItem.Visible = false;
                floor3StartToolStripMenuItem.Visible = false;
                floor3StopToolStripMenuItem.Visible = false;
                floor4TestToolStripMenuItem.Visible = false;
                floor4StartToolStripMenuItem.Visible = false;
                floor4StopToolStripMenuItem.Visible = false;
                floor5TestToolStripMenuItem.Visible = false;
                floor5StartToolStripMenuItem.Visible = false;
                floor5StopToolStripMenuItem.Visible = false;

                richTextBox1.Visible = false;
                richTextBox2.Visible = false;
                richTextBox3.Visible = false;
                richTextBox4.Visible = false;
                richTextBox5.Visible = false;
                richTextBox6.Visible = false;

            }
        }

        #region 欢迎界面
        Thread hello;
        public void Hello_Show()
        {
            try
            {
                Application.Run(new Loding());
            }
            catch { }
            //MessageBox.Show("");
            //Hello A = new Hello();
            //A.Show();
        }
        #endregion

        #region 窗体基本功能
        private void Form1_Load(object sender, EventArgs e)
        {
            hello = new Thread(new ThreadStart(Hello_Show));
            hello.IsBackground = true;
            hello.Start();

            toolStripButtonStop.Enabled = false;
            toolStripButtonStart.Enabled = true;

            connectTestToolStripButton.Enabled = true;
            ledTestToolStripButton.Enabled = true;
            toolStripButton1.Enabled = true;
            toolStripButton2.Enabled = true;
            toolStripButton3.Enabled = true;
            toolStripButton4.Enabled = true;
            toolStripButton5.Enabled = true;
            toolStripSplitButton1.Enabled = true;

            if (Config.Default.SoftWareMode == "自动模式")
            {
                toolStripComboBox6.Text = "自动模式";
                toolStrip4.Visible = false;
                toolStrip5.Visible = false;
                toolStrip6.Visible = false;
                toolStrip7.Visible = false;
                toolStrip8.Visible = false;
                //floor1ToolStripMenuItem.Enabled = false;
                //floor2ToolStripMenuItem.Enabled = false;
                //floor3ToolStripMenuItem.Enabled = false;
                //floor4ToolStripMenuItem.Enabled = false;
                //floor5ToolStripMenuItem.Enabled = false;
            }
            else
            {
                toolStripComboBox6.Text = "手动模式";
                toolStrip1.Visible = false;
                toolStrip2.Visible = false;
                toolStrip3.Visible = false;
                //modeAutoToolStripMenuItem.Enabled = false;
            }
            if (Config.Default.HardMod == "V0")
            {
                toolStrip2.Visible = false;
            }

            modeAutoToolStripMenuItem.Checked = toolStrip1.Visible;
            modeFloorToolStripMenuItem.Checked = toolStrip2.Visible;
            modeModeToolStripMenuItem.Checked = toolStrip3.Visible;

            floor1ToolStripMenuItem.Checked = toolStrip4.Visible;
            floor2ToolStripMenuItem.Checked = toolStrip5.Visible;
            floor3ToolStripMenuItem.Checked = toolStrip6.Visible;
            floor4ToolStripMenuItem.Checked = toolStrip7.Visible;
            floor5ToolStripMenuItem.Checked = toolStrip8.Visible;


            int height = SystemInformation.WorkingArea.Height;
            int width = SystemInformation.WorkingArea.Width;

            int formheight = this.Size.Height;
            int formwidth = this.Size.Width;

            int newformx = width / 2 - formwidth / 2;
            int newformy = height / 2 - formheight / 2;

            SetDesktopLocation(newformx, newformy);

            splitContainer1.SplitterDistance = 2 + dataGridView1.Rows.GetRowsHeight(DataGridViewElementStates.Visible) + dataGridView1.ColumnHeadersHeight;

            //WindowState = FormWindowState.Maximized;

            //this.splitContainer1.SplitterDistance = 412;
            dataGridView1.Invalidate();

            dataGridView1.ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.AllowUserToResizeColumns = false;
            dataGridView1.AllowUserToOrderColumns = false;

            dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dataGridView1.MultiSelect = false;

            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            foreach (DataGridViewColumn item in dataGridView1.Columns)
            {
                item.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                if (item.HeaderText == "详情")
                {
                    item.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    item.Width = 25;
                }
                else if (item.HeaderText == "编号")
                {
                    item.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    item.Width = 30;
                }
                else
                {
                    item.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    item.MinimumWidth = 30;
                    //item.Width = 30;
                }
            }
            for (int i = 0; i < 28; i++)
            {
                int newRowIndex = dataGridView1.Rows.Add();
                DataGridViewRow newRow = dataGridView1.Rows[newRowIndex];
                newRow.Cells["colSFI_ID1"].Value = "1-" + (i + 1).ToString("00");
                newRow.Cells["colSFI_ID2"].Value = "2-" + (i + 1).ToString("00");
                newRow.Cells["colSFI_ID3"].Value = "3-" + (i + 1).ToString("00");
                newRow.Cells["colSFI_ID4"].Value = "4-" + (i + 1).ToString("00");
                newRow.Cells["colSFI_ID5"].Value = "5-" + (i + 1).ToString("00");
                for (int j = 0; j < 5; j++)
                {
                    newRow.Cells[6 * j + 1].Value = "0A";
                    newRow.Cells[6 * j + 2].Value = "0%";
                    newRow.Cells[6 * j + 3].Value = "0℃";
                    newRow.Cells[6 * j + 4].Value = "0℃";
                }
            }


            toolStripButton1.Tag = "DisConnect";
            toolStripButton2.Tag = "DisConnect";
            toolStripButton3.Tag = "DisConnect";
            toolStripButton4.Tag = "DisConnect";
            toolStripButton5.Tag = "DisConnect";


            this.cur1SetToolStripMenuItem.Text = "  1段(" + Config.Default.Ia01 + "A)";
            this.cur2SetToolStripMenuItem.Text = "  2段(" + Config.Default.Ia02 + "A)";
            this.cur3SetToolStripMenuItem.Text = "  3段(" + Config.Default.Ia03 + "A)";
            this.cur4SetToolStripMenuItem.Text = "  4段(" + Config.Default.Ia04 + "A)";
            this.cur5SetToolStripMenuItem.Text = "  5段(" + Config.Default.Ia05 + "A)";
            this.cur6SetToolStripMenuItem.Text = "  6段(" + Config.Default.Ia06 + "A)";
            this.cur7SetToolStripMenuItem.Text = "  7段(" + Config.Default.Ia07 + "A)";
            this.cur8SetToolStripMenuItem.Text = "  8段(" + Config.Default.Ia08 + "A)";
            this.cur9SetToolStripMenuItem.Text = "  9段(" + Config.Default.Ia09 + "A)";
            this.cur10SetToolStripMenuItem.Text = " 10段(" + Config.Default.Ia10 + "A)";
            this.cur11SetToolStripMenuItem.Text = " 11段(" + Config.Default.Ia11 + "A)";
            this.cur12SetToolStripMenuItem.Text = " 12段(" + Config.Default.Ia12 + "A)";
            this.cur13SetToolStripMenuItem.Text = " 13段(" + Config.Default.Ia13 + "A)";
            this.cur14SetToolStripMenuItem.Text = " 14段(" + Config.Default.Ia14 + "A)";
            this.cur15SetToolStripMenuItem.Text = " 15段(" + Config.Default.Ia15 + "A)";
            this.cur16SetToolStripMenuItem.Text = " 16段(" + Config.Default.Ia16 + "A)";
            this.cur17SetToolStripMenuItem.Text = " 17段(" + Config.Default.Ia17 + "A)";
            this.cur18SetToolStripMenuItem.Text = " 18段(" + Config.Default.Ia18 + "A)";


            serialPort1.ReadTimeout = Config.Default.TimeForReadAfterWrite;
            serialPort2.ReadTimeout = Config.Default.TimeForReadAfterWrite;
            serialPort3.ReadTimeout = Config.Default.TimeForReadAfterWrite;
            serialPort4.ReadTimeout = Config.Default.TimeForReadAfterWrite;
            serialPort5.ReadTimeout = Config.Default.TimeForReadAfterWrite;
            serialPortA.ReadTimeout = Config.Default.TimeForReadAfterWrite;
            serialPortB.ReadTimeout = Config.Default.TimeForReadAfterWrite;

            serialPort1.ReadBufferSize = Config.Default.ReadLenght;
            serialPort2.ReadBufferSize = Config.Default.ReadLenght;
            serialPort3.ReadBufferSize = Config.Default.ReadLenght;
            serialPort4.ReadBufferSize = Config.Default.ReadLenght;
            serialPort5.ReadBufferSize = Config.Default.ReadLenght;
            serialPortA.ReadBufferSize = Config.Default.ReadLenght;
            serialPortB.ReadBufferSize = Config.Default.ReadLenght;

            toolStripSplitButton1.Text = Config.Default.Mode;

            splitContainer1.SplitterDistance = 2 + dataGridView1.Rows.GetRowsHeight(DataGridViewElementStates.Visible) + dataGridView1.ColumnHeadersHeight;

            dataGridView1.Update();

            hello.Abort();
            toolStripProgressBar1.Value = 60;
            //this.TopMost = false;
        }

        FormWindowState OldState = FormWindowState.Maximized;
        private void Form1_Resize(object sender, EventArgs e)
        {
            richTextBox6.Text = this.Width.ToString() + " " + this.Height.ToString();
            hello = new Thread(new ThreadStart(Hello_Show));
            hello.IsBackground = true;
            hello.Start();
            if (OldState == FormWindowState.Maximized && this.WindowState != FormWindowState.Minimized)
            {
                splitContainer1.SplitterDistance = 2 + dataGridView1.Rows.GetRowsHeight(DataGridViewElementStates.Visible) + dataGridView1.ColumnHeadersHeight;
                //MessageBox.Show(this.Width.ToString());
                //dataGridView1.Visible = false;
                toolStripProgressBar1.Value = 0;
                dataGridView1.Invalidate();
                foreach (DataGridViewColumn item in dataGridView1.Columns)
                {
                    item.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    toolStripProgressBar1.Value++;
                }
                foreach (DataGridViewColumn item in dataGridView1.Columns)
                {
                    toolStripProgressBar1.Value++;
                    if (item.HeaderText == "详情")
                    {
                        item.Width = 25;
                    }
                    else if (item.HeaderText == "编号")
                    {
                        item.Width = 30;
                    }
                    else
                    {
                        item.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    }
                }
                dataGridView1.Update();
            }

            toolStripProgressBar1.Value = 60;
            //dataGridView1.Visible = true;
            hello.Abort();
            OldState = this.WindowState;
        }

        private void Form1_ResizeEnd(object sender, EventArgs e)
        {
            hello = new Thread(new ThreadStart(Hello_Show));
            hello.IsBackground = true;
            hello.Start();
            if (OldState == FormWindowState.Minimized)
            {
                OldState = this.WindowState;
                hello.Abort();
                return;
            }
            OldState = this.WindowState;
            if (OldState == FormWindowState.Maximized)
            {
                splitContainer1.SplitterDistance = 2 + dataGridView1.Rows.GetRowsHeight(DataGridViewElementStates.Visible) + dataGridView1.ColumnHeadersHeight;
                hello.Abort();
                return;
            }
            if (OldState == FormWindowState.Minimized)
            {
                hello.Abort();
                return;
            }
            splitContainer1.SplitterDistance = 2 + dataGridView1.Rows.GetRowsHeight(DataGridViewElementStates.Visible) + dataGridView1.ColumnHeadersHeight;
            //MessageBox.Show(this.Width.ToString());
            //dataGridView1.Visible = false;
            toolStripProgressBar1.Value = 0;
            dataGridView1.Invalidate();
            foreach (DataGridViewColumn item in dataGridView1.Columns)
            {
                item.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                toolStripProgressBar1.Value++;
            }
            foreach (DataGridViewColumn item in dataGridView1.Columns)
            {
                toolStripProgressBar1.Value++;
                if (item.HeaderText == "详情")
                {
                    item.Width = 25;
                }
                else if (item.HeaderText == "编号")
                {
                    item.Width = 30;
                }
                else
                {
                    item.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                }
            }
            dataGridView1.Update();
            toolStripProgressBar1.Value = 60;
            //dataGridView1.Visible = true;
            hello.Abort();
        }
        #endregion

        #region 自动按ESC
        private void autoESC()
        {
            Thread threadPressESC = new Thread(new ThreadStart(pressESC));
            threadPressESC.IsBackground = true;
            threadPressESC.Start();
        }
        private void pressESC()
        {
            Thread.Sleep(Config.Default.AutoEsc);
            try
            {
                SendKeys.SendWait("{ESC}");
            }
            catch { }
        }
        #endregion

        #region 串口打开
        private void serialClose()
        {
            if (serialPort1.IsOpen) serialPort1.Close();
            if (serialPort2.IsOpen) serialPort2.Close();
            if (serialPort3.IsOpen) serialPort3.Close();
            if (serialPort4.IsOpen) serialPort4.Close();
            if (serialPort5.IsOpen) serialPort5.Close();
        }
        private bool serialOpen1()
        {
            try
            {
                if (serialPort1.IsOpen) serialPort1.Close();
                serialPort1.PortName = Config.Default.Serial1;
                serialPort1.Open();
                toolStripButton1.Image = Properties.Resources.success;
                toolStripButton1.Tag = "Selected";
                for (int j = 0; j < 6; j++)
                {
                    dataGridView1.Columns[j].DefaultCellStyle.ForeColor = Color.Black;
                }
                return true;
            }
            catch (Exception)
            {
                toolStripButton1.Image = Properties.Resources.delete;
                toolStripButton1.Tag = "CantConnect";
                for (int j = 0; j < 6; j++)
                {
                    dataGridView1.Columns[j].DefaultCellStyle.ForeColor = Color.Gray;
                }
                return false;
            }
        }
        private bool serialOpen2()
        {
            try
            {
                if (serialPort2.IsOpen) serialPort2.Close();
                serialPort2.PortName = Config.Default.Serial2;
                serialPort2.Open();
                toolStripButton2.Image = Properties.Resources.success;
                toolStripButton2.Tag = "Selected";
                for (int j = 0; j < 6; j++)
                {
                    dataGridView1.Columns[j + 6].DefaultCellStyle.ForeColor = Color.Black;
                }
                return true;
            }
            catch (Exception)
            {
                toolStripButton2.Image = Properties.Resources.delete;
                toolStripButton2.Tag = "CantConnect";
                for (int j = 0; j < 6; j++)
                {
                    dataGridView1.Columns[j + 6].DefaultCellStyle.ForeColor = Color.Gray;
                }
                return false;
            }
        }
        private bool serialOpen3()
        {
            try
            {
                if (serialPort3.IsOpen) serialPort3.Close();
                serialPort3.PortName = Config.Default.Serial3;
                serialPort3.Open();
                toolStripButton3.Image = Properties.Resources.success;
                toolStripButton3.Tag = "Selected";
                for (int j = 0; j < 6; j++)
                {
                    dataGridView1.Columns[j + 6 + 6].DefaultCellStyle.ForeColor = Color.Black;
                }
                return true;
            }
            catch (Exception)
            {
                toolStripButton3.Image = Properties.Resources.delete;
                toolStripButton3.Tag = "CantConnect";
                for (int j = 0; j < 6; j++)
                {
                    dataGridView1.Columns[j + 6 + 6].DefaultCellStyle.ForeColor = Color.Gray;
                }
                return false;
            }
        }
        private bool serialOpen4()
        {
            try
            {
                if (serialPort4.IsOpen) serialPort4.Close();
                serialPort4.PortName = Config.Default.Serial4;
                serialPort4.Open();
                toolStripButton4.Image = Properties.Resources.success;
                toolStripButton4.Tag = "Selected";
                for (int j = 0; j < 6; j++)
                {
                    dataGridView1.Columns[j + 6 + 6 + 6].DefaultCellStyle.ForeColor = Color.Black;
                }
                return true;
            }
            catch (Exception)
            {
                toolStripButton4.Image = Properties.Resources.delete;
                toolStripButton4.Tag = "CantConnect";
                for (int j = 0; j < 6; j++)
                {
                    dataGridView1.Columns[j + 6 + 6 + 6].DefaultCellStyle.ForeColor = Color.Gray;
                }
                return false;
            }
        }
        private bool serialOpen5()
        {
            try
            {
                if (serialPort5.IsOpen) serialPort5.Close();
                serialPort5.PortName = Config.Default.Serial5;
                serialPort5.Open();
                toolStripButton5.Image = Properties.Resources.success;
                toolStripButton5.Tag = "Selected";
                for (int j = 0; j < 6; j++)
                {
                    dataGridView1.Columns[j + 6 + 6 + 6 + 6].DefaultCellStyle.ForeColor = Color.Black;
                }
                return true;
            }
            catch (Exception)
            {
                toolStripButton5.Image = Properties.Resources.delete;
                toolStripButton5.Tag = "CantConnect";
                for (int j = 0; j < 6; j++)
                {
                    dataGridView1.Columns[j + 6 + 6 + 6 + 6].DefaultCellStyle.ForeColor = Color.Gray;
                }
                return false;
            }
        }
        private bool serialOpenAB()
        {
            try
            {
                if (serialPortA.IsOpen) serialPortA.Close();
                serialPortA.PortName = Config.Default.SerialA;
                serialPortA.Open();
                if (serialPortB.IsOpen) serialPortB.Close();
                serialPortB.PortName = Config.Default.SerialB;
                serialPortB.Open();
                toolStripButton21.Image = Properties.Resources.success;
                return true;
            }
            catch (Exception)
            {
                toolStripButton21.Image = Properties.Resources.delete;
                return false;
            }
        }
        #endregion

        #region 报文发送
        /// <summary>
        /// 报文显示委托
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private delegate void DataSend(byte[] data);
        /// <summary>
        /// 报文显示A
        /// </summary>
        /// <param name="str"></param>
        private void sendA(byte[] buffer)
        {
            if (serialPortA.IsOpen == false) return;
            serialPortA.Write(buffer, 0, buffer.Length);
            string str = "Send:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
            showRichTextBox(richTextBox1, str);
        }
        /// <summary>
        /// 报文显示B
        /// </summary>
        /// <param name="str"></param>
        private void sendB(byte[] buffer)
        {
            if (serialPortB.IsOpen == false) return;
            serialPortB.Write(buffer, 0, buffer.Length);
            string str = "Send:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
            showRichTextBox(richTextBox1, str);
        }
        /// <summary>
        /// 报文显示1
        /// </summary>
        /// <param name="str"></param>
        private void send1(byte[] buffer)
        {
            if (serialPort1.IsOpen == false) return;
            serialPort1.Write(buffer, 0, buffer.Length);
            string str = "Send:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
            showRichTextBox(richTextBox1, str);
        }
        /// <summary>
        /// 报文显示2
        /// </summary>
        /// <param name="str"></param>
        private void send2(byte[] buffer)
        {
            if (serialPort2.IsOpen == false) return;
            serialPort2.Write(buffer, 0, buffer.Length);
            string str = "Send:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
            showRichTextBox(richTextBox2, str);
        }
        /// <summary>
        /// 报文显示3
        /// </summary>
        /// <param name="str"></param>
        private void send3(byte[] buffer)
        {
            if (serialPort3.IsOpen == false) return;
            serialPort3.Write(buffer, 0, buffer.Length);
            string str = "Send:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
            showRichTextBox(richTextBox3, str);
        }
        /// <summary>
        /// 报文显示4
        /// </summary>
        /// <param name="str"></param>
        private void send4(byte[] buffer)
        {
            if (serialPort4.IsOpen == false) return;
            serialPort4.Write(buffer, 0, buffer.Length);
            string str = "Send:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
            showRichTextBox(richTextBox4, str);
        }
        /// <summary>
        /// 报文显示5
        /// </summary>
        /// <param name="str"></param>
        private void send5(byte[] buffer)
        {
            if (serialPort5.IsOpen == false) return;
            serialPort5.Write(buffer, 0, buffer.Length);
            string str = "Send:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
            showRichTextBox(richTextBox5, str);
        }
        private void sends(int floor, byte[] buffer)
        {
            switch (floor)
            {
                case 1:
                    send1(buffer);
                    break;
                case 2:
                    send2(buffer);
                    break;
                case 3:
                    send3(buffer);
                    break;
                case 4:
                    send4(buffer);
                    break;
                case 5:
                    send5(buffer);
                    break;
            }
        }
        #endregion

        #region 自定义命令发送
        internal class Traget
        {
            public static byte ToOne = 0x55;
            public static byte ToAll = 0x66;
            public static byte ForPend = 0x77;
        }
        /// <summary>
        /// 带数据命令
        /// </summary>
        /// <param name="traget">
        /// 目标地址:0-17
        /// </param>
        /// <param name="commond">
        /// 命令功能
        /// </param>
        /// <param name="data">
        /// 附加数据
        /// </param>
        private void MessageSend(byte traget, byte commond, byte[] data, int floor)
        {
            //AA num/ 0xFF    55 / 66   11  0C  55  55  55  AA  75  E4  01  01  01  01  01  01  02  01  01(Test) / 00(Work)   Xor 10 - 19
            byte[] buffer;
            if (data != null)
                buffer = new byte[data.Length + 20];
            else
                buffer = new byte[20];
            int i = 0;
            //0
            buffer[i++] = 0xAA;
            //1
            //2
            if (traget == 0xFF)
            {
                buffer[i++] = traget;
                buffer[i++] = Traget.ToAll;
            }
            else
            {
                buffer[i++] = ++traget;
                buffer[i++] = Traget.ToOne;
            }
            //3
            if (data == null)
                buffer[i++] = 0x10;
            else
                buffer[i++] = (byte)(data.Length + 0x10);
            if (commond == 0x04)
                buffer[i - 1]++;
            //4
            switch (commond)
            {
                case 0x01:
                    buffer[i++] = 0x0C;
                    break;
                case 0x02:
                case 0x03:
                    buffer[i++] = 0x05;
                    break;
                case 0x04:
                    buffer[i++] = 0x07;
                    break;
                case 0x05:
                case 0x06:
                    buffer[i++] = 0x0B;
                    break;
                case 0x07:
                    break;
                case 0x08:
                    buffer[i++] = 0x49;
                    break;
                case 0x09:
                    break;
                default:
                    buffer[i++] = 0x00;
                    break;
            }
            {//固定块5-16
                buffer[i++] = 0x55;
                buffer[i++] = 0x55;
                buffer[i++] = 0x55;
                buffer[i++] = 0xAA;
                buffer[i++] = 0x75;
                buffer[i++] = 0xE4;
                buffer[i++] = 0x01;
                buffer[i++] = 0x01;
                buffer[i++] = 0x01;
                buffer[i++] = 0x01;
                buffer[i++] = 0x01;
                buffer[i++] = 0x01;
            }
            //17
            if (data == null)
                buffer[i++] = 0x01;
            else
                buffer[i++] = (byte)(data.Length + 1);
            //18
            buffer[i++] = commond;

            if (data != null)
            {
                foreach (byte item in data)
                {
                    buffer[i++] = item;
                }
            }
            buffer[i] = 0;
            for (int j = 10; j < i; j++)
                buffer[i] ^= buffer[j];
            sends(floor, buffer);
        }
        /// <summary>
        ///    
        /// </summary>
        /// <param name="traget">
        /// 发送目标：0-17
        /// Traget.ToOne    对单个子地址发送
        /// Traget.ToAll    对所有子地址发送
        /// Traget.ForPend  请求返回数据
        /// </param>
        /// <param name="commond">
        /// 命令功能
        /// </param>
        private void MessageSend(byte traget, byte[] commond, byte[] data, int floor)
        {
            byte[] buffer;
            if (data == null && commond == null)
                buffer = new byte[3];
            else
                return;
            int i = 0;
            //0
            buffer[i++] = 0xAA;
            //1
            buffer[i++] = ++traget;
            //2
            buffer[i++] = Traget.ForPend;

            sends(floor, buffer);
        }

        private bool MessageSends(byte traget, byte commond, byte[] data)
        {
            int i = 0;
            if ((string)toolStripButton1.Tag == "Selected")
            {
                i++;
                MessageSend(traget, commond, data, 1);
            }
            if ((string)toolStripButton2.Tag == "Selected")
            {
                i++;
                MessageSend(traget, commond, data, 2);
            }
            if ((string)toolStripButton3.Tag == "Selected")
            {
                i++;
                MessageSend(traget, commond, data, 3);
            }
            if ((string)toolStripButton4.Tag == "Selected")
            {
                i++;
                MessageSend(traget, commond, data, 4);
            }
            if ((string)toolStripButton5.Tag == "Selected")
            {
                i++;
                MessageSend(traget, commond, data, 5);
            }
            if (i == 0)
            {
                if (connectFlag == true)
                {
                    connectFlag = false;
                    connectTestToolStripButton.Enabled = true;
                }
                autoESC();
                MessageBox.Show("未选中任何测试层");
                tshow("执行终止:未选中层");
                return false;
            }
            return true;
        }
        #endregion

        #region 报文接收
        private delegate byte[] DataRead(byte item);
        private byte[] read1(byte item)
        {
            serialPort1.ReadExisting();
            MessageSend(item, null, null, 1);
            Thread.Sleep(Config.Default.CurrentErrorValue);
            byte[] buffer = new byte[serialPort1.ReadBufferSize];
            try
            {
                serialPort1.Read(buffer, 0, buffer.Length);
                string str = "Recive:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
                showRichTextBox(richTextBox1, str);
                return buffer;
            }
            catch (TimeoutException)
            {
                return null;
            }
            catch { return null; }
        }
        private byte[] read2(byte item)
        {
            serialPort2.ReadExisting();
            MessageSend(item, null, null, 2);
            Thread.Sleep(Config.Default.CurrentErrorValue);
            byte[] buffer = new byte[serialPort2.ReadBufferSize];
            try
            {
                serialPort2.Read(buffer, 0, buffer.Length);
                string str = "Recive:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
                showRichTextBox(richTextBox2, str);
                return buffer;
            }
            catch (TimeoutException)
            {
                return null;
            }
            catch { return null; }
        }
        private byte[] read3(byte item)
        {
            serialPort3.ReadExisting();
            MessageSend(item, null, null, 3);
            Thread.Sleep(Config.Default.CurrentErrorValue);
            byte[] buffer = new byte[serialPort3.ReadBufferSize];
            try
            {
                serialPort3.Read(buffer, 0, buffer.Length);
                string str = "Recive:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
                showRichTextBox(richTextBox3, str);
                return buffer;
            }
            catch (TimeoutException)
            {
                return null;
            }
            catch { return null; }
        }
        private byte[] read4(byte item)
        {
            serialPort4.ReadExisting();
            MessageSend(item, null, null, 4);
            Thread.Sleep(Config.Default.CurrentErrorValue);
            byte[] buffer = new byte[serialPort4.ReadBufferSize];
            try
            {
                serialPort4.Read(buffer, 0, buffer.Length);
                string str = "Recive:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
                showRichTextBox(richTextBox4, str);
                return buffer;
            }
            catch (TimeoutException)
            {
                return null;
            }
            catch { return null; }
        }
        private byte[] read5(byte item)
        {
            serialPort5.ReadExisting();
            MessageSend(item, null, null, 5);
            Thread.Sleep(Config.Default.CurrentErrorValue);
            byte[] buffer = new byte[serialPort5.ReadBufferSize];
            try
            {
                serialPort5.Read(buffer, 0, buffer.Length);
                string str = "Recive:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
                showRichTextBox(richTextBox5, str);
                return buffer;
            }
            catch (TimeoutException)
            {
                return null;
            }
            catch { return null; }
        }
        #endregion

        #region 自定义读取命令
        private double[] ReadOne(int floor, byte i)
        {
            double[] data = new double[5];
            byte[] buffer = null;
            string tmp;
            switch (floor)
            {
                case 1:
                    buffer = read1(i);
                    break;
                case 2:
                    buffer = read2(i);
                    break;
                case 3:
                    buffer = read3(i);
                    break;
                case 4:
                    buffer = read4(i);
                    break;
                case 5:
                    buffer = read5(i);
                    break;
            }

            if (buffer != null)                                 //接收数据正常
            {
                data[0] = 1;                                 //更改数据存在标志
                                                             //if (v)                                          //如果需要返回数据
                                                             //{
                tmp = BitConverter.ToString(buffer);
                tmp = tmp.Replace("-", "");
                int s = tmp.IndexOf("FF55");
                if (s != -1)                                //数据合法
                {
                    byte[] iitt = new byte[4];
                    iitt = Conv.strToToHexByte(tmp.Substring(s + 4, 8));

                    if (iitt[1] > 0x80)
                    {                                       //小通道
                        data[1] = (((iitt[1] - 0x80) << 8) + iitt[0]) / 10;
                    }
                    else
                    {                                       //大通道
                        data[1] = (iitt[1] << 8) + iitt[0];
                    }


                    if (Config.Default.Mode == "3C1O-T")  //指示器返回温度数据
                    {
                        //外温
                        if (iitt[2] == 0x80)                //温度传感器故障
                        {
                            data[2] = -99;
                        }
                        else if (iitt[2] > 0x80)            //温度值小于零度
                        {
                            if (iitt[2] >= 0xC4)            //温度值小于-60度
                                data[2] = iitt[2] - 256;
                            else
                                data[2] = iitt[2];
                        }
                        else if (iitt[2] < 0x80)            //温度值大于零度
                        {
                            data[2] = iitt[2];
                        }

                        //内温
                        if (iitt[3] == 0x80)                //温度传感器故障
                        {
                            data[3] = -99;
                        }
                        else if (iitt[3] > 0x80)            //温度值小于零度
                        {
                            if (iitt[3] >= 0xC4)            //温度值小于-60度
                                data[3] = iitt[3] - 256;
                            else
                                data[3] = iitt[3];
                        }
                        else if (iitt[3] < 0x80)            //温度值大于零度
                        {
                            data[3] = iitt[3];
                        }
                        data[4] = CurTrue(data[1], label1.Text);
                    }
                    else                                    //指示器不返回温度数据
                    {
                        data[2] = 0;
                        data[3] = 0;
                    }
                }
                else                                        //数据不合法
                {
                    data[0] = 0;
                    data[1] = 0;
                    data[2] = 0;
                    data[3] = 0;
                }
                //}
                //else                                            //不需要返回数据
                //{
                //    data[1] = 0;
                //    data[2] = 0;
                //    data[3] = 0;
                //}
            }
            else                                                //接收数据超时
            {
                data[0] = 0;
                data[1] = 0;
                data[2] = 0;
                data[3] = 0;
            }
            return data;
        }



        //读取报文解析
        private double[,] Read(bool v, int floor)
        {
            double[,] data = new double[28, 5];
            byte[] buffer = null;
            string tmp;
            byte i = 0;
            for (i = 0; i < 28; i++)
            {
                data[i, 0] = 0;                                 //更改数据存在标志
                switch (floor)
                {
                    case 1:
                        buffer = read1(i);
                        break;
                    case 2:
                        buffer = read2(i);
                        break;
                    case 3:
                        buffer = read3(i);
                        break;
                    case 4:
                        buffer = read4(i);
                        break;
                    case 5:
                        buffer = read5(i);
                        break;
                }

                if (buffer != null)                                 //接收数据正常
                {
                    //if (v)                                          //如果需要返回数据
                    //{
                    tmp = BitConverter.ToString(buffer);
                    tmp = tmp.Replace("-", "");
                    int s = tmp.IndexOf("FF55");
                    if (s != -1)                                //数据合法
                    {
                        data[i, 0] = 1;                                 //更改数据存在标志
                        byte[] iitt = new byte[4];
                        iitt = Conv.strToToHexByte(tmp.Substring(s + 4, 8));

                        if (iitt[1] >= 0x80)
                        {                                       //小通道
                            data[i, 1] = (((iitt[1] - 0x80) << 8) + iitt[0]) / 10.0;
                        }
                        else
                        {                                       //大通道
                            data[i, 1] = (iitt[1] << 8) + iitt[0];
                        }
                        if (iitt[0] == 0x00 && iitt[1] == 0x00)
                            data[i, 0] = 2;
                        if (iitt[0] == 0x00 && iitt[1] == 0x00 && iitt[2] == 0x00 && iitt[3] == 0x00)
                            data[i, 0] = 0;

                        if (Config.Default.Mode == "3C1O-T")  //指示器返回温度数据
                        {
                            //外温
                            if (iitt[2] == 0x80)                //温度传感器故障
                            {
                                data[i, 2] = -99;
                            }
                            else if (iitt[2] > 0x80)            //温度值小于零度
                            {
                                if (iitt[2] >= 0xC4)            //温度值小于-60度
                                    data[i, 2] = iitt[2] - 256;
                                else
                                    data[i, 2] = iitt[2];
                            }
                            else if (iitt[2] < 0x80)            //温度值大于零度
                            {
                                data[i, 2] = iitt[2];
                            }

                            //内温
                            if (iitt[3] == 0x80)                //温度传感器故障
                            {
                                data[i, 3] = -99;
                            }
                            else if (iitt[3] > 0x80)            //温度值小于零度
                            {
                                if (iitt[3] >= 0xC4)            //温度值小于-60度
                                    data[i, 3] = iitt[3] - 256;
                                else
                                    data[i, 3] = iitt[3];
                            }
                            else if (iitt[3] < 0x80)            //温度值大于零度
                            {
                                data[i, 3] = iitt[3];
                            }
                        }
                        else                                    //指示器不返回温度数据
                        {
                            data[i, 2] = 0;
                            data[i, 3] = 0;
                        }
                        switch (floor)
                        {
                            case 1:
                                data[i, 4] = CurTrue(data[i, 1], label1.Text);
                                break;
                            case 2:
                                data[i, 4] = CurTrue(data[i, 1], label2.Text);
                                break;
                            case 3:
                                data[i, 4] = CurTrue(data[i, 1], label3.Text);
                                break;
                            case 4:
                                data[i, 4] = CurTrue(data[i, 1], label4.Text);
                                break;
                            case 5:
                                data[i, 4] = CurTrue(data[i, 1], label5.Text);
                                break;
                        }
                    }
                    else                                        //数据不合法
                    {
                        data[i, 0] = 2;
                        data[i, 1] = 0;
                        data[i, 2] = 0;
                        data[i, 3] = 0;
                    }
                    //}
                    //else                                            //不需要返回数据
                    //{
                    //    data[i, 1] = 0;
                    //    data[i, 2] = 0;
                    //    data[i, 3] = 0;
                    //}
                }
                else                                                //接收数据超时
                {
                    data[i, 0] = 0;
                    data[i, 1] = 0;
                    data[i, 2] = 0;
                    data[i, 3] = 0;
                }
            }
            //FF  55  VL VH  ET IT
            DealReadData(v, data, floor);
            return data;
        }
        private double CurTrue(double cur, string rel)
        {
            double relCur;
            //注释原因，指示器电流为0不是异常状态
            //if (cur == 0) return 0;
            if (!double.TryParse(rel, out relCur)) return 100;
            if (double.Parse(rel) == -99) return 100;
            double Dev = (Math.Abs((cur - relCur)) / relCur) * 100;
            return Dev;
        }
        //显示解析数据double.Parse(Config.Default.Dev)
        private void DealReadData(bool s, double[,] data, int floor)
        {
            //double[,] data = Read(s, read1);
            floor--;
            double dev = Config.Default.CurrentErrorPercentageSemi;

            for (int i = 0; i < 28; i++)
            {
                if (data[i, 0] == 0)
                {
                    for (int j = 0; j < 5; j++)
                    {
                        dataGridView1.Rows[i].Cells[floor * 6 + j].Style.BackColor = Color.White;
                        dataGridView1.Rows[i].Cells[floor * 6 + j].Style.ForeColor = Color.White;//与测试架通讯失败
                    }
                }
                else if (data[i, 0] == 2)
                {
                    for (int j = 0; j < 5; j++)
                    {
                        dataGridView1.Rows[i].Cells[floor * 6 + j].Style.BackColor = Color.White;
                        dataGridView1.Rows[i].Cells[floor * 6 + j].Style.ForeColor = Color.Black;
                    }
                    dataGridView1.Rows[i].Cells[floor * 6 + 0].Style.BackColor = Color.Yellow;//与指示器通讯失败
                    //dataGridView1.Rows[i].Cells[floor * 6 + 0].Style.ForeColor = Color.Black;
                }
                else
                {
                    for (int j = 0; j < 5; j++)
                    {
                        dataGridView1.Rows[i].Cells[floor * 6 + j].Style.BackColor = Color.White;
                        dataGridView1.Rows[i].Cells[floor * 6 + j].Style.ForeColor = Color.Black;
                    }
                }
                if (s)
                {
                    dataGridView1.Rows[i].Cells[floor * 6 + 1].Value = data[i, 1].ToString("0.0") + "A";

                    dataGridView1.Rows[i].Cells[floor * 6 + 2].Value = data[i, 4].ToString("0.0") + "%";
                    if (data[i, 4] > dev)
                        dataGridView1.Rows[i].Cells[floor * 6 + 2].Style.BackColor = Color.Red;
                    else
                        dataGridView1.Rows[i].Cells[floor * 6 + 2].Style.BackColor = Color.White;
                    if (Config.Default.Mode == "3C1O-T")
                    {
                        dataGridView1.Rows[i].Cells[floor * 6 + 3].Value = data[i, 2] + "℃";
                        if (data[i, 2] == -99)
                            dataGridView1.Rows[i].Cells[floor * 6 + 3].Style.BackColor = Color.Red;
                        else
                            dataGridView1.Rows[i].Cells[floor * 6 + 3].Style.BackColor = Color.White;

                        dataGridView1.Rows[i].Cells[floor * 6 + 4].Value = data[i, 3] + "℃";
                        if (data[i, 3] == -99)
                            dataGridView1.Rows[i].Cells[floor * 6 + 4].Style.BackColor = Color.Red;
                        else
                            dataGridView1.Rows[i].Cells[floor * 6 + 4].Style.BackColor = Color.White;
                    }
                }
            }
        }
        private void Read1(object tmp)
        {
            bool s = (bool)tmp;
            Read(s, 1);
            tshow("执行完毕");
            if (connectFlag == true)
                Invoke((MethodInvoker)delegate ()
                {
                    connectFlag = false;
                    connectTestToolStripButton.Enabled = true;
                });
        }
        private void Read2(object tmp)
        {
            bool s = (bool)tmp;
            Read(s, 2);
            tshow("执行完毕");
            if (connectFlag == true)
                Invoke((MethodInvoker)delegate ()
                {
                    connectFlag = false;
                    connectTestToolStripButton.Enabled = true;
                }); 
        }
        private void Read3(object tmp)
        {
            bool s = (bool)tmp;
            Read(s, 3);
            tshow("执行完毕");
            if (connectFlag == true)
                Invoke((MethodInvoker)delegate ()
                {
                    connectFlag = false;
                    connectTestToolStripButton.Enabled = true;
                });
        }
        private void Read4(object tmp)
        {
            bool s = (bool)tmp;
            Read(s, 4);
            tshow("执行完毕");
            if (connectFlag == true)
                Invoke((MethodInvoker)delegate ()
                {
                    connectFlag = false;
                    connectTestToolStripButton.Enabled = true;
                });
        }
        private void Read5(object tmp)
        {
            bool s = (bool)tmp;
            Read(s, 5);
            tshow("执行完毕");
            if (connectFlag == true)
                Invoke((MethodInvoker)delegate ()
                {
                    connectFlag = false;
                    connectTestToolStripButton.Enabled = true;
                });
        }
        //多线程启动读取
        private bool Reads(bool v)
        {
            int i = 0;
            if ((string)toolStripButton1.Tag == "Selected")
            {
                i++;
                threadFloor1 = new Thread(new ParameterizedThreadStart(Read1));
                threadFloor1.IsBackground = true;
                threadFloor1.Start(v);
            }
            if ((string)toolStripButton2.Tag == "Selected")
            {
                i++;
                threadFloor2 = new Thread(new ParameterizedThreadStart(Read2));
                threadFloor2.IsBackground = true;
                threadFloor2.Start(v);
            }
            if ((string)toolStripButton3.Tag == "Selected")
            {
                i++;
                threadFloor3 = new Thread(new ParameterizedThreadStart(Read3));
                threadFloor3.IsBackground = true;
                threadFloor3.Start(v);
            }
            if ((string)toolStripButton4.Tag == "Selected")
            {
                i++;
                threadFloor4 = new Thread(new ParameterizedThreadStart(Read4));
                threadFloor4.IsBackground = true;
                threadFloor4.Start(v);
            }
            if ((string)toolStripButton5.Tag == "Selected")
            {
                i++;
                threadFloor5 = new Thread(new ParameterizedThreadStart(Read5));
                threadFloor5.IsBackground = true;
                threadFloor5.Start(v);
            }
            if (i == 0)
            {
                autoESC();
                MessageBox.Show("未选中任何测试层");
                tshow("执行终止:未选中层");
                return false;
            }
            return true;
        }
        #endregion

        #region 报文显示
        private void showRichTextBox(RichTextBox richTextBox, string str)
        {
            richTextBox.Invoke((MethodInvoker)delegate ()
            {
                richTextBox.AppendText(str);

                //int length = richTextBox.Lines[richTextBox.Lines.Length - 2].Length;
                //richTextBox.Select(richTextBox.TextLength - length - 1, length);
                int start = richTextBox.TextLength - str.Length;
                richTextBox.Select(start + 1, str.Length - 1);
                if (str.IndexOf("Send") != -1)
                    richTextBox.SelectionColor = Color.Green;
                else if (str.IndexOf("Recive") != -1)
                    richTextBox.SelectionColor = Color.Blue;
                else
                    richTextBox.SelectionColor = Color.Red;
                richTextBox.SelectionLength = 0;

                richTextBox.ScrollToCaret();
            });
            int nameNumber = 0;
            switch (richTextBox.Name)
            {
                case "richTextBox1":
                    nameNumber = 1;
                    break;
                case "richTextBox2":
                    nameNumber = 2;
                    break;
                case "richTextBox3":
                    nameNumber = 3;
                    break;
                case "richTextBox4":
                    nameNumber = 4;
                    break;
                case "richTextBox5":
                    nameNumber = 5;
                    break;
            }
            //string name = "";
            //if (logD[nameNumber] == null)
            //{
            //    name = nameNumber + "." + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();
            //    name = name.Replace("/", "-");
            //    name = name.Replace(":", "_");
            //}
            //else name = logD[nameNumber];
            //write_txt(name, str);

            log.Out_Log("第" + nameNumber + "层报文", str);
        }
        #endregion

        #region 报文自动清除
        private void richTextBox_TextChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(richTextBox1.Lines.Length.ToString());
            RichTextBox One = (RichTextBox)sender;
            if (One.Text.Length > 10000)
                One.Text = One.Text.Remove(0, 1000);
        }
        #endregion

        #region 层开启、关闭选择
        private void toolStripButton_ButtonClick(object sender, EventArgs e)
        {
            ToolStripSplitButton item = (ToolStripSplitButton)sender;
            if (item.Tag.ToString() == "Selected")
            {
                item.Image = Properties.Resources.Prohibit;
                item.Tag = "DisSelected";
                CurClose(int.Parse(item.Text.Substring(3, 1)));
                tshow(item.Text + "被关闭");
            }
            else if (item.Tag.ToString() == "DisSelected")
            {
                item.Image = Properties.Resources.success;
                item.Tag = "Selected";
                tshow(item.Text + "被打开");
            }
            else if (item.Tag.ToString() == "DisConnect" || item.Tag.ToString() == "CantConnect")
            {
                tshow(item.Text + "被打开");
                switch (int.Parse(item.Text.Substring(3, 1)))
                {
                    case 1:
                        if (!serialOpen1())
                        {
                            autoESC();
                            MessageBox.Show("串口" + Config.Default.Serial1.Substring(3) + "打开失败");
                            tshow(item.Text + "打开失败");
                        }
                        break;
                    case 2:
                        if (!serialOpen2())
                        {
                            autoESC();
                            MessageBox.Show("串口" + Config.Default.Serial2.Substring(3) + "打开失败");
                            tshow(item.Text + "打开失败");
                        }
                        break;
                    case 3:
                        if (!serialOpen3())
                        {
                            autoESC();
                            MessageBox.Show("串口" + Config.Default.Serial3.Substring(3) + "打开失败");
                            tshow(item.Text + "打开失败");
                        }
                        break;
                    case 4:
                        if (!serialOpen4())
                        {
                            autoESC();
                            MessageBox.Show("串口" + Config.Default.Serial4.Substring(3) + "打开失败");
                            tshow(item.Text + "打开失败");
                        }
                        break;
                    case 5:
                        if (!serialOpen5())
                        {
                            autoESC();
                            MessageBox.Show("串口" + Config.Default.Serial5.Substring(3) + "打开失败");
                            tshow(item.Text + "打开失败");
                        }
                        break;
                }
            }
        }
        #endregion

        #region 列表选择修复
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            dataGridView1.ClearSelection();
        }
        #endregion

        #region 多线程合一
        Thread threadFloor1, threadFloor2, threadFloor3, threadFloor4, threadFloor5;
        Thread thread1, thread2, thread3, thread4, thread5;
        private void floor1Worker()
        {
            autoRun(1);
        }
        private void floor2Worker()
        {
            autoRun(2);
        }
        private void floor3Worker()
        {
            autoRun(3);
        }
        private void floor4Worker()
        {
            autoRun(4);
        }
        private void floor5Worker()
        {
            autoRun(5);
        }
        private bool floorsWorker()
        {
            int i = 0;
            if ((string)toolStripButton1.Tag == "Selected")
            {
                i++;
                thread1 = new Thread(new ThreadStart(floor1Worker));
                thread1.IsBackground = true;
                thread1.Start();
            }
            if ((string)toolStripButton2.Tag == "Selected")
            {
                i++;
                thread2 = new Thread(new ThreadStart(floor2Worker));
                thread2.IsBackground = true;
                thread2.Start();
            }
            if ((string)toolStripButton3.Tag == "Selected")
            {
                i++;
                thread3 = new Thread(new ThreadStart(floor3Worker));
                thread3.IsBackground = true;
                thread3.Start();
            }
            if ((string)toolStripButton4.Tag == "Selected")
            {
                i++;
                thread4 = new Thread(new ThreadStart(floor4Worker));
                thread4.IsBackground = true;
                thread4.Start();
            }
            if ((string)toolStripButton5.Tag == "Selected")
            {
                i++;
                thread5 = new Thread(new ThreadStart(floor5Worker));
                thread5.IsBackground = true;
                thread5.Start();
            }
            if (i == 0)
            {
                autoESC();
                MessageBox.Show("未选中任何测试层");
                tshow("执行终止:未选中层");
                return false;
            }
            return true;
        }
        #endregion

        #region 自动测试主函数
        bool[] finish = new bool[6];
        int[,] FI = new int[6, 29];
        /*
        指示器标志位说明：
        0：正常
        1：通讯失败
        3：误差大
        */
        //bool[] haveTest = new bool[6];
        string[] logD = new string[5];
        private void autoRun(int floor)
        {
            logD[floor] = floor + "." + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "_" + DateTime.Now.Second;
            logD[floor] = logD[floor].Replace("/", "-");
            logD[floor] = logD[floor].Replace(":", "_");
            FI = new int[6, 29];
            int i;
            int ConnectRetryCount = 0;
            int ConnectRetryCount_bak = 0;
            int CurRetryCount = 0;
            int ConnectNeed = 0;
            int ConnectError = 0;
            finish[floor] = false;
            double[,] data = null;
            //MessageSend(0x01, 0x02, null, new DataSend(send));
            for (i = 1; i <= 18; i++)   //18段电流选择
            {
                ishow("", "\r\n");
                ishow("***************************************");
                ishow("第" + floor + "层:进入第" + i + "段模式");
                ishow("***************************************");
                ishow("", "\r\n");
                Invoke((MethodInvoker)delegate ()
                {
                    switch (floor)
                    {
                        case 1:
                            toolStripProgressBar2.Value = i;
                            break;
                        case 2:
                            toolStripProgressBar3.Value = i;
                            break;
                        case 3:
                            toolStripProgressBar4.Value = i;
                            break;
                        case 4:
                            toolStripProgressBar5.Value = i;
                            break;
                        case 5:
                            toolStripProgressBar6.Value = i;
                            break;
                    }
                }); //更新界面  进度条 与信息显示框

                #region 标准电流
                //设置标准电流
                CurOpen(i, floor);                              //设置电流 段数为 i
                ishow("第" + floor + "层:等待测试架电流稳定");
                Thread.Sleep(Config.Default.TimeForStandardCurrentStability);        //等待电流稳定
                if (i == 5)
                {
                    ishow("第" + floor + "层:第五段长延时处理");
                    Thread.Sleep(Config.Default.TimeForFive);        //等待电流稳定
                    ishow("第" + floor + "层:第五段长延时处理完毕");
                }
                //读取标准电流    
                CurRead(floor);                                 //读取标准电流
                ishow("第" + floor + "层:标准电流为"+ double.Parse(label1.Text).ToString("0.0") + "A");
                ConnectRetryCount = 0;                          //清空标准电流读取标志
                while (!CmpCur(floor))                          //判断标准电流是否符合误差范围
                {
                    ishow("第" + floor + "层:标准电流第" + (ConnectRetryCount + 1) + "次失败");
                    Thread.Sleep(Config.Default.TimeForStandardCurrentStability);    //不符合则 等待 
                    //读取标准电流
                    CurRead(floor);                             //再次读取标准电流
                    if (ConnectRetryCount > Config.Default.CountForStandardCurrentRetry)   //达到设定读取次数则 提示错误
                    {
                        CurClose(floor);
                        tshow("自动测试中止");
                        ishow("非正常错误：第" + i + "段标准电流读取错误,请联系系统维护人员！！");
                        MessageBox.Show("非正常错误：第" + i + "段标准电流读取错误\r\n请联系系统维护人员！！",
                            "警告！第" + floor + "层提示",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        goto FISTOP;                            //终止此层测试进程
                    }
                    ConnectRetryCount++;                        //读取错误记数+1
                }
                CurRead(floor);                             //再次读取标准电流
                ishow("第" + floor + "层:标准电流为" + double.Parse(label1.Text).ToString("0.0") + "A");
                ishow("第" + floor + "层:标准电流稳定，进入指示器连接模式");
                #endregion

                #region 指示器连接
                ConnectRetryCount = 0;                          //清空电流读取次数
                SetCur(0xFF, floor);                            //修正指示器参数
                ishow("第" + floor + "层:校准指示器参数");
FIRead:
                Thread.Sleep(Config.Default.TimeForWaitFICalibration);       //进入一个长等待
                ishow("第" + floor + "层:读取标准电流值");
                CurRead(floor);                                 //读取标准电流
                ishow("第" + floor + "层:标准电流为" + double.Parse(label1.Text).ToString("0.0") + "A");
                MessageSend(0xFF, 0x02, null, floor);           //发送读取指示器电流命令
                Thread.Sleep(Config.Default.TimeForWaitFIRead);           //读取延时
                ishow("第" + floor + "层:读取指示器参数");
                data = Read(false, floor);                      //读取指示器电流

                ishow("第" + floor + "层:开始统计连接失败指示器个数");
                //统计连接失败数量
                ConnectNeed=0;
                ConnectError=0;                                 //清空通讯失败指示器计数
                for (byte item = 0; item < 28; item++)          //统计连接失败的指示器
                {
                    if (i == 1)//第一段时
                    {
                        ConnectNeed++;
                        if (data[item, 0] != 1)
                            ConnectError++;
                        else
                            FI[floor, item] = 2;
                    }
                    else
                    {
                        if(FI[floor, item] == 0 && data[item, 0] == 1)
                            FI[floor, item] = 1;
                        if (FI[floor, item] == 2)
                        {
                            ConnectNeed++;
                            if (data[item, 0] != 1)
                                ConnectError++;
                        }
                    }
                }

                //判断连接失败数量是否太多（大于设定值）
                ishow("第" + floor + "层:" + ConnectError + "个指示器连接失败");
                //2016-01-15 由下面if移出：找不到指示器就没法测试了，不只是第一次而已。
                if (ConnectError == ConnectNeed)                      //指示器全部连接失败
                {
                    //2016-02-22 先关闭电流
                    CurClose(floor);
                    tshow("自动测试中止");
                    ishow("第" + floor + "层没有找到合格的指示器");
                    MessageBox.Show(
                        "第" + floor + "层没有找到合格的指示器",
                        "第" + floor + "层提示",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Stop);
                    goto FISTOP;                            //终止此层测试进程
                }
                #endregion
                
                #region

                //if (!FIAutoOut && cantconnect > Config.Default.A_CantConnet)
                //{
                //    //2016年02月22日 不判断数量，节省人为操作
                //    //DialogResult res = MessageBox.Show(
                //    //    "第" + floor + "层只能识别出 " + (28 - cantconnect) + " 个指示器，是否继续测试？",
                //    //    "第" + floor + "层提示",
                //    //    MessageBoxButtons.YesNo,
                //    //    MessageBoxIcon.Question);           //提示
                //    //if (res == DialogResult.Yes)            //得到信息：挂的指示器少
                //    //{
                //    //2016-02-02 添加选择是时直接屏蔽
                //    for (int item = 0; item < 28; item++)
                //    {
                //        if (data[item, 0] == 0)
                //        {
                //            //2016年2月22日 FI默认为0 即未找到指示器
                //            //FI[floor, item] = 1;                //屏蔽指示器，原因：无通讯
                //            //                                    //输出：指示器被屏蔽
                //            Invoke((MethodInvoker)delegate ()
                //            {
                //                richTextBox6.AppendText("指示器 " + floor + "-" + (item + 1) + ":连接失败被自动屏蔽\r\n");

                //                richTextBox6.Select(richTextBox6.TextLength, 0);
                //                richTextBox6.ScrollToCaret();
                //            });
                //        }
                //        else
                //        {
                //            //2016年2月22日 重定义FI标志 0：未找到指示器1：找到指示器2：通讯失败3：误差大
                //            FI[floor, item] = 1;
                //        }
                //    }
                //    FIAutoOut = true;                   //自动排除连接失败指示器
                //    //2016年02月22日 不判断数量，节省人为操作
                //    //}
                //    //else if (res == DialogResult.No)        //得到信息：挂的指示器不少
                //    //{
                //    //    goto FISTOP;                            //终止此层测试进程
                //    //}
                //    //2016-01-15 注释原因：本次失败仅作为统计（并不是所有指示器都连接失败）
                //    //ConnectRetryCount++;                      //连接失败计数+1
                //    goto FIRead;
                //}
                //else //连接失败指示器数目较少 或 进入自动排除模式
                //{
                //FIAutoOut = true;                           //自动得到信息：挂的指示器不少
                #endregion


                #region 指示器连接
                //以此判断28个指示器
                ishow("第" + floor + "层:屏蔽连接失败指示器");
                ConnectRetryCount_bak = ConnectRetryCount;
                for (byte item = 0; item < 28; item++)
                {
                    ConnectRetryCount = ConnectRetryCount_bak;
                    //if (FI[floor, item] == 0)               //指示器未被屏蔽，开始判断
                    if (i == 1 || FI[floor, item] > 1)               //指示器未被屏蔽，开始判断
                    {
                        if (data[item, 0] != 1)             //通讯失败，进入重试
                        {
                            if (data[item, 0] == 0)
                                ishow("指示器 " + floor + "-" + (item + 1) + ":测试架连接失败");
                            else if (data[item, 0] == 2)
                                ishow("指示器 " + floor + "-" + (item + 1) + ":指示器连接失败");
                            double[] datatmp = { 0, 0, 0 };
                            do
                            {
                                if (ConnectRetryCount < Config.Default.CountForStandardCurrentRetry)
                                {
                                    //读取指令发送
                                    MessageSend(item, 0x02, null, floor);               //发送读取电流命令
                                    Thread.Sleep(Config.Default.TimeForWaitFIRead);     //等待接收数据
                                    datatmp = ReadOne(floor, item);                     //读取电流数据
                                    if (datatmp[0] == 1)                                //如果接收到数据
                                    {
                                        if (data[item, 0] == 0)
                                            ishow("测试架 " + floor + "-" + (item + 1) + ":重新连接成功");
                                        else if (data[item, 0] == 2)
                                            ishow("指示器 " + floor + "-" + (item + 1) + ":重新连接成功");
                                        data[item, 0] = datatmp[0];
                                        data[item, 1] = datatmp[1];
                                        data[item, 2] = datatmp[2];
                                        data[item, 3] = datatmp[3];
                                        data[item, 4] = datatmp[4];
                                    }
                                    //失败次数 ＋
                                    ConnectRetryCount++;
                                }
                                else
                                {
                                    //屏蔽指示器

                                    //2015-01-15 注释：通讯失败就应该被屏蔽！不需要理由
                                    if (data[item, 0] == 2)                 //
                                        FI[floor, item] = 1;                //屏蔽指示器，原因：无通讯
                                                                            //输出：指示器被屏蔽
                                    Invoke((MethodInvoker)delegate ()
                                    {
                                        if (data[item, 0] == 0)
                                            ishow("测试架 " + floor + "-" + (item + 1) + ":连接失败被自动屏蔽");
                                        else if (data[item, 0] == 2)
                                            ishow("指示器 " + floor + "-" + (item + 1) + ":连接失败被自动屏蔽");
                                    });
                                    ConnectRetryCount = 0;
                                    break;
                                }
                            }
                            while (datatmp[0] != 1);        //判断指示器是否通讯正常
                        }
                        //数据读取成功
                    }
                    else if (FI[floor, item] == 0)
                    {
                        //if (data[item, 0] != 1)
                        //测试过程中检测到指示器数据，表示存在指示器，但是通讯失败
                        if (data[item, 0] == 1)
                            FI[floor, item] = 1;
                    }
                    // }//28个指示器判断
                }
                #endregion

                #region
                //2016-01-15 注释：不区分指示器不在位状态（未挂与失败）
                //for (byte item = 0; item < 28; item++)
                //{
                //    if (FI[floor, item] == 0)
                //    {
                //        if (data[item, 0] == 0)
                //        {
                //            FI[floor, item] = 2;
                //            dataGridView1.Rows[item].Cells[floor * 6 - 6].Style.BackColor = Color.Yellow;
                //            //输出：指示器连接失败
                //            Invoke((MethodInvoker)delegate ()
                //            {
                //                richTextBox6.AppendText("指示器 " + floor + "-" + (item + 1) + ":连接失败\r\n");

                //                richTextBox6.Select(richTextBox6.TextLength, 0);
                //                richTextBox6.ScrollToCaret();
                //            });
                //        }
                //    }
                //}//屏蔽通讯失败指示器
                #endregion

                ishow("第" + floor + "层:指示器误差判断","");
                #region 数据处理
                //数据处理--误差判断
                //double dev = Config.Default.Dev;                //读取系统设定合格范围
                int errorCount = 0;
                //以此判断28个指示器
                for (byte item = 0; item < 28; item++)          //遍历指示器，统计合格数量与最佳状态数量
                {
                    if (FI[floor, item] > 1)                   //
                    {
                        if (CmpDev(data[item, 1], double.Parse(label1.Text)))
                            errorCount++;
                    }
                    if(item%7==0)
                        ishow("", "\r\n");
                    ishow(floor + "-" + (item + 1).ToString("00") + ":" + data[item, 1].ToString("0.0") + "A " + data[item, 4].ToString("0.0") + "%", "\t");
                }//统计指示器故障数量
                ishow("", "\r\n");
                ishow("第" + floor + "层:误差大" + errorCount + "个");
                //if (Count == 0) ;
                if ( errorCount > (Config.Default.FullCalibrationThreshold / 28 * ConnectNeed))         //优秀的小于9个
                {                                                               //统一处理
                    ishow("第" + floor + "层:指示器集中校准");
                    if (CurRetryCount < Config.Default.TimesForErrorCorrection)
                    {
                        CurRead(floor);                                         //读取标准电流
                        ishow("第" + floor + "层:标准电流为" + double.Parse(label1.Text) + "A");
                        SetCur(0xFF, floor);                                    //设置指示器
                        ishow("第" + floor + "层:校准所有指示器");
                        Thread.Sleep(Config.Default.TimeForWaitFIRead);
                        CurRetryCount++;                                        //校正次数+1
                        goto FIRead;                                            //重新读取指示器
                    }
                    else
                    {
                        CurRetryCount = 0;                                      //清空校正次数
                    }
                }
                else if (errorCount > 0)                                       //不合格超过0个
                {                                                               //单点处理
                    ishow("第" + floor + "层:指示器单独校准");
                    if (CurRetryCount < Config.Default.TimesForErrorCorrection)
                    {
                        CurRead(floor);                             //读取标准电流
                        ishow("第" + floor + "层:标准电流为" + double.Parse(label1.Text) + "A");
                        for (byte item = 0; item < 28; item++)                  //遍历    
                        {
                            //if (FI[floor, item] == 0)                           //状态正常正常
                            if (FI[floor, item] > 1)                           //状态正常正常
                            {
                                if (CmpDev(data[item, 1], double.Parse(label1.Text)))
                                {
                                    ishow("第" + floor + "层:校准" + floor + "-" + item + " 电流" + data[item, 1] + "A");
                                    //CurRead(floor);                             //读取标准电流
                                    //ishow("第" + floor + "层:校准" + floor + "-" + item + " 电流" + data[item, 1] + "A");
                                    SetCur(item, floor);                        //校正指示器
                                    //Thread.Sleep(Config.Default.A_Speed);
                                }
                            }
                        }
                        CurRetryCount++;                                        //校正次数+1
                        goto FIRead;                                            //重新读取指示器
                    }
                    else
                    {
                        CurRetryCount = 0;                                      //清空校正次数
                    }
                }
                else                                                            //全都合格！！！
                {
                    ishow("第" + floor + "层:完美！没有误差大的指示器");
                    CurRetryCount = 0;                                          //清空校正次数
                    goto NEXT;                                                  //下一步操作
                }
                //挑出故障的指示器
                for (byte item = 0; item < 28; item++)
                {
                    //if (FI[floor, item] == 0)
                    if (FI[floor, item] > 1)
                    {
                        if (CmpDev(data[item, 1], double.Parse(label1.Text)))
                        {
                            FI[floor, item] = 3;                                //标记为误差大
                            Invoke((MethodInvoker)delegate ()
                            {
                                ishow("指示器 " + floor + " - " + (item + 1) + ": 第" + i + "段电流值为" + data[item, 1].ToString("0.0") + "误差为" + data[item, 4].ToString("0.0") + " %> " + Config.Default.CurrentErrorPercentageSemi + " %");
                            });
                        }
                    }
                }
                NEXT:                                                           //下一段
                continue;
                #endregion
            }                                                                   //测试结束
            ishow("第" + floor + "层:测试结束，统计指示器");
            for (byte item = 0; item < 28; item++)
            {
                dataGridView1.Rows[item].Cells[floor * 6 - 6].Style.BackColor = Color.White;
                //if (FI[floor, item] == 0)
                //{
                //    dataGridView1.Rows[item].Cells[floor * 6 - 6].Style.BackColor = Color.Red;
                //}
                if (FI[floor, item] == 0)
                {
                    dataGridView1.Rows[item].Cells[floor * 6 - 4].Style.BackColor = Color.Gray;
                    dataGridView1.Rows[item].Cells[floor * 6 - 5].Style.BackColor = Color.Gray;
                    dataGridView1.Rows[item].Cells[floor * 6 - 6].Style.BackColor = Color.Gray;
                }
                else if (FI[floor, item] == 1)
                {
                    dataGridView1.Rows[item].Cells[floor * 6 - 4].Style.BackColor = Color.Yellow;
                    dataGridView1.Rows[item].Cells[floor * 6 - 5].Style.BackColor = Color.Yellow;
                    dataGridView1.Rows[item].Cells[floor * 6 - 6].Style.BackColor = Color.Yellow;
                }
                else if (FI[floor, item] == 2)
                {
                    dataGridView1.Rows[item].Cells[floor * 6 - 4].Style.BackColor = Color.LightGreen;
                    dataGridView1.Rows[item].Cells[floor * 6 - 5].Style.BackColor = Color.LightGreen;
                    dataGridView1.Rows[item].Cells[floor * 6 - 6].Style.BackColor = Color.LightGreen;
                }
                else if (FI[floor, item] == 3)
                {
                    dataGridView1.Rows[item].Cells[floor * 6 - 4].Style.BackColor = Color.Red;
                    dataGridView1.Rows[item].Cells[floor * 6 - 5].Style.BackColor = Color.Red;
                    dataGridView1.Rows[item].Cells[floor * 6 - 6].Style.BackColor = Color.Red;
                }

                if (Config.Default.Mode == "3C1O-T" && i > 1)  //指示器返回温度数据
                {
                    dataGridView1.Rows[item].Cells[floor * 6 - 3].Value = data[item, 2] + "℃";
                    if (data[item, 2] == -99)
                        dataGridView1.Rows[item].Cells[floor * 6 - 3].Style.BackColor = Color.Red;
                    else
                        dataGridView1.Rows[item].Cells[floor * 6 - 3].Style.BackColor = Color.White;

                    dataGridView1.Rows[item].Cells[floor * 6 - 2].Value = data[item, 3] + "℃";
                    if (data[item, 3] == -99)
                        dataGridView1.Rows[item].Cells[floor * 6 - 2].Style.BackColor = Color.Red;
                    else
                        dataGridView1.Rows[item].Cells[floor * 6 - 2].Style.BackColor = Color.White;
                }
                //统计指示器故障数量
            }

            ishow("第" + floor + "层:统计完毕，自动结束测试");
            #region 测试结束
            CurClose(floor);
            Thread.Sleep(Config.Default.TimeForStandardCurrentStability);        //等待电流稳定   
            CurRead(floor);                                 //读取标准电流
            finish[floor] = true;
            if (
                !(finish[1] ^ ((string)toolStripButton1.Tag == "Selected"))
                & !(finish[2] ^ ((string)toolStripButton2.Tag == "Selected"))
                & !(finish[3] ^ ((string)toolStripButton3.Tag == "Selected"))
                & !(finish[4] ^ ((string)toolStripButton4.Tag == "Selected"))
                & !(finish[5] ^ ((string)toolStripButton5.Tag == "Selected"))
                )
            {
                //string name = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "_" + DateTime.Now.Second;
                //name = name.Replace("/", "-");
                //name = name.Replace(":", "_");
                //write_txt(name, richTextBox6.Text);
                log.Out_Log("第" + floor + "层测试信息", richTextBox6.Text);

                richTextBox6.Text = "";
                FinishFlg = true;
                Invoke((MethodInvoker)delegate ()
                {
                    toolStripButtonStop.PerformClick();
                });
            }
            return;
            FISTOP:
            tshow("自动测试已经中止");
            switch (floor)
            {
                case 1:
                    thread1.Abort();
                    break;
                case 2:
                    thread2.Abort();
                    break;
                case 3:
                    thread3.Abort();
                    break;
                case 4:
                    thread4.Abort();
                    break;
                case 5:
                    thread5.Abort();
                    break;
            }
            #endregion

        }

        private static void write_txt(string fileName, string content)
        {
            string FILE_NAME = fileName + ".txt";//每天按照日期建立一个不同的文件名
            StreamWriter sr;
            if (File.Exists(FILE_NAME)) //如果文件存在,则创建File.AppendText对象
            {
                sr = File.AppendText(FILE_NAME);
            }
            else  //如果文件不存在,则创建File.CreateText对象
            {
                sr = File.CreateText(FILE_NAME);
            }

            sr.WriteLine(content);//将传入的字符串加上时间写入文本文件一行
            sr.Close();
        }
        
        private bool CmpDev(double cur, double curs)
        {
            double curt = 0, devt = 0;
            double curts = 0, devts = 0;
            if (cur > curs)
                curt = cur - curs;
            else
                curt = curs - cur;
            devt = (curt / curs) * 100;
            if (Config.Default.TestMode == "成品模式")
            {
                curts = Config.Default.CurrentErrorValue;
                devts = Config.Default.CurrentErrorPercentage;
            }
            else if (Config.Default.TestMode == "半成品模式")
            {
                curts = Config.Default.CurrentErrorValueSemi;
                devts = Config.Default.CurrentErrorPercentageSemi;
            }
            if (curt > curts && devt >= devts)
                return true;
            else
                return false;
        }
        private bool CmpCur(int floor)
        {
            Label lab1 = null, lab2 = null;
            switch (floor)
            {
                case 1:
                    lab1 = label1;
                    lab2 = label7;
                    break;
                case 2:
                    lab1 = label2;
                    lab2 = label9;
                    break;
                case 3:
                    lab1 = label3;
                    lab2 = label11;
                    break;
                case 4:
                    lab1 = label4;
                    lab2 = label13;
                    break;
                case 5:
                    lab1 = label5;
                    lab2 = label15;
                    break;
            }

            string s2 = lab2.Text.Remove(lab2.Text.Length - 1);
            string s1 = lab1.Text;
            if (LabelShow(lab1, lab2))
            {
                double d2 = double.Parse(s2);
                double d1 = double.Parse(s1);
                if ((Math.Abs(d1 - d2) / d2) * 100 < Config.Default.ErrorPercentageForStandardCurrent)
                    return true;
                else
                    return false;
            }
            else return false;
        }
        #endregion

        #region 十八段参数
        byte[,] cur = {
            { 0x16, 0x10 },
            { 0x16, 0x11 },
            { 0x16, 0x12 },
            { 0x16, 0x13 },
            { 0x16, 0x14 },
            { 0x16, 0x15 },
            { 0x17, 0x10 },
            { 0x17, 0x11 },
            { 0x17, 0x12 },
            { 0x17, 0x13 },
            { 0x17, 0x14 },
            { 0x17, 0x15 },
            { 0x18, 0x10 },
            { 0x18, 0x11 },
            { 0x18, 0x12 },
            { 0x18, 0x13 },
            { 0x18, 0x14 },
            { 0x18, 0x15 }
        };
        #endregion

        #region 标准电流设置
        //CRC校验
        private int GetModBusCRC(byte[] pByte, int nNumberOfBytes)
        {
            ushort nShiftedBit;
            int pChecksum = 0xFFFF;

            for (int nByte = 0; nByte < nNumberOfBytes; nByte++)
            {
                pChecksum ^= pByte[nByte];
                for (int nBit = 0; nBit < 8; nBit++)
                {
                    if ((pChecksum & 0x01) == 0x01)
                        nShiftedBit = 1;
                    else
                        nShiftedBit = 0;
                    pChecksum >>= 1;
                    if (nShiftedBit == 1)
                        pChecksum ^= 0xA001;
                }
            }
            return pChecksum;
        }

        //源方法 设置标准电流
        private void CurOpen(int item, int floor)
        {
            CurClose(floor);
            int crc;
            byte[] buffer = new byte[8];

            buffer[0] = 0x01;
            buffer[1] = 0x05;
            buffer[2] = 0x00;
            buffer[5] = 0x00;

            //buffer[3] = (byte)(0xF0 + 0x0F);
            //buffer[4] = 0x00;
            //crc = GetModBusCRC(buffer, 6);
            //buffer[6] = (byte)(crc & 0xFF);
            //buffer[7] = (byte)(crc >> 8 & 0xFF);
            //Thread.Sleep(Config.Default.A_SpeedCurRead);
            //sends(floor, buffer);
            //Thread.Sleep(Config.Default.A_SpeedCurRead);
            //sends(floor, buffer);

            buffer[3] = cur[item - 1, 0];
            buffer[4] = 0xFF;
            crc = GetModBusCRC(buffer, 6);
            buffer[6] = (byte)(crc & 0xFF);
            buffer[7] = (byte)(crc >> 8 & 0xFF);
            if (Config.Default.HardMod == "V0")
                sendA(buffer);
            else if (Config.Default.HardMod == "V1")
                sends(floor, buffer);
            //Thread.Sleep(Config.Default.A_SpeedCurRead);
            //sends(floor, buffer);

            buffer[3] = cur[item - 1, 1];
            buffer[4] = 0xFF;
            crc = GetModBusCRC(buffer, 6);
            buffer[6] = (byte)(crc & 0xFF);
            buffer[7] = (byte)(crc >> 8 & 0xFF);
            Thread.Sleep(Config.Default.A_SpeedCurRead);
            if (Config.Default.HardMod == "V0")
                sendA(buffer);
            else if (Config.Default.HardMod == "V1")
                sends(floor, buffer);
            //Thread.Sleep(Config.Default.A_SpeedCurRead);
            //sends(floor, buffer);

            Invoke((MethodInvoker)delegate ()
            {
                switch (floor)
                {
                    case 1:
                        label6.Text = "第" + (item) + "段";
                        label7.Text = ShowCur(item) + "A";
                        break;
                    case 2:
                        label8.Text = "第" + (item) + "段";
                        label9.Text = ShowCur(item) + "A";
                        break;
                    case 3:
                        label10.Text = "第" + (item) + "段";
                        label11.Text = ShowCur(item) + "A";
                        break;
                    case 4:
                        label12.Text = "第" + (item) + "段";
                        label13.Text = ShowCur(item) + "A";
                        break;
                    case 5:
                        label14.Text = "第" + (item) + "段";
                        label15.Text = ShowCur(item) + "A";
                        break;
                }
            });
        }
        //设置标准电流
        private bool CurOpens(int item)
        {
            int i = 0;
            if ((string)toolStripButton1.Tag == "Selected")
            {
                threadFloor1 = new Thread(new ThreadStart(delegate
                {
                    CurOpen(item, 1);
                }));
                threadFloor1.IsBackground = true;
                threadFloor1.Start();
                i++;
            }
            if ((string)toolStripButton2.Tag == "Selected")
            {
                threadFloor2 = new Thread(new ThreadStart(delegate
                {
                    CurOpen(item, 2);
                }));
                threadFloor2.IsBackground = true;
                threadFloor2.Start();
                i++;
            }
            if ((string)toolStripButton3.Tag == "Selected")
            {
                threadFloor3 = new Thread(new ThreadStart(delegate
                {
                    CurOpen(item, 3);
                }));
                threadFloor3.IsBackground = true;
                threadFloor3.Start();
                i++;
            }
            if ((string)toolStripButton4.Tag == "Selected")
            {
                threadFloor4 = new Thread(new ThreadStart(delegate
                {
                    CurOpen(item, 4);
                }));
                threadFloor4.IsBackground = true;
                threadFloor4.Start();
                i++;
            }
            if ((string)toolStripButton5.Tag == "Selected")
            {
                threadFloor5 = new Thread(new ThreadStart(delegate
                {
                    CurOpen(item, 5);
                }));
                threadFloor5.IsBackground = true;
                threadFloor5.Start();
                i++;
            }
            if (i == 0)
            {
                autoESC();
                MessageBox.Show("未选中任何测试层");
                tshow("执行终止:未选中层");
                return false;
            }
            return true;
        }
        private string ShowCur(int i)
        {
            switch (i)
            {
                case 1:
                    return Config.Default.Ia01;
                case 2:
                    return Config.Default.Ia02;
                case 3:
                    return Config.Default.Ia03;
                case 4:
                    return Config.Default.Ia04;
                case 5:
                    return Config.Default.Ia05;
                case 6:
                    return Config.Default.Ia06;
                case 7:
                    return Config.Default.Ia07;
                case 8:
                    return Config.Default.Ia08;
                case 9:
                    return Config.Default.Ia09;
                case 10:
                    return Config.Default.Ia10;
                case 11:
                    return Config.Default.Ia11;
                case 12:
                    return Config.Default.Ia12;
                case 13:
                    return Config.Default.Ia13;
                case 14:
                    return Config.Default.Ia14;
                case 15:
                    return Config.Default.Ia15;
                case 16:
                    return Config.Default.Ia16;
                case 17:
                    return Config.Default.Ia17;
                case 18:
                    return Config.Default.Ia18;
            }
            return "";
        }


        //源方法 关闭电流
        private void CurClose(int floor)
        {
            int crc;
            byte[] buffer = new byte[8];

            buffer[0] = 0x01;
            buffer[1] = 0x05;
            buffer[2] = 0x00;

            buffer[5] = 0x00;

            Thread.Sleep(Config.Default.A_SpeedCurRead);
            if (Config.Default.HardMod == "V0")
            {
                for (int i = 0; i < 18; i++)
                {
                    buffer[3] = (byte)(0x10 + i);
                    buffer[4] = 0x00;
                    crc = GetModBusCRC(buffer, 6);
                    buffer[6] = (byte)(crc & 0xFF);
                    buffer[7] = (byte)(crc >> 8 & 0xFF);
                    sendA(buffer);
                    Thread.Sleep(Config.Default.A_SpeedCurRead);
                }
            }
            else if (Config.Default.HardMod == "V1")
            {
                buffer[3] = (byte)(0xF0 + 0x0F);
                buffer[4] = 0x00;
                crc = GetModBusCRC(buffer, 6);
                buffer[6] = (byte)(crc & 0xFF);
                buffer[7] = (byte)(crc >> 8 & 0xFF);
                sends(floor, buffer);
                Thread.Sleep(Config.Default.A_SpeedCurRead);
            }
            Invoke((MethodInvoker)delegate ()
            {
                switch (floor)
                {
                    case 1:
                        label6.Text = "未开启";
                        label7.Text = "0A";
                        break;
                    case 2:
                        label8.Text = "未开启";
                        label9.Text = "0A";
                        break;
                    case 3:
                        label10.Text = "未开启";
                        label11.Text = "0A";
                        break;
                    case 4:
                        label12.Text = "未开启";
                        label13.Text = "0A";
                        break;
                    case 5:
                        label14.Text = "未开启";
                        label15.Text = "0A";
                        break;
                }
            });
        }
        //关闭电流
        private bool CurCloses()
        {
            int i = 0;
            if ((string)toolStripButton1.Tag == "Selected")
            {
                CurClose(1); i++;
            }
            if ((string)toolStripButton2.Tag == "Selected")
            {
                CurClose(2); i++;
            }
            if ((string)toolStripButton3.Tag == "Selected")
            {
                CurClose(3); i++;
            }
            if ((string)toolStripButton4.Tag == "Selected")
            {
                CurClose(4); i++;
            }
            if ((string)toolStripButton5.Tag == "Selected")
            {
                CurClose(5); i++;
            }
            if (i == 0)
            {
                autoESC();
                MessageBox.Show("未选中任何测试层");
                tshow("执行终止:未选中层");
                return false;
            }
            return true;
        }
        //关闭所有电流
        private void CurCloseAll()
        {
            if (Config.Default.HardMod == "V0")
            {
                CurClose(0);
                label6.Text = "未开启";
                label7.Text = "0A";
                label8.Text = "未开启";
                label9.Text = "0A";
                label10.Text = "未开启";
                label11.Text = "0A";
                label12.Text = "未开启";
                label13.Text = "0A";
                label14.Text = "未开启";
                label15.Text = "0A";
            }
            else if (Config.Default.HardMod == "V1")
            {
                CurClose(1);
                label6.Text = "未开启";
                label7.Text = "0A";
                CurClose(2);
                label8.Text = "未开启";
                label9.Text = "0A";
                CurClose(3);
                label10.Text = "未开启";
                label11.Text = "0A";
                CurClose(4);
                label12.Text = "未开启";
                label13.Text = "0A";
                CurClose(5);
                label14.Text = "未开启";
                label15.Text = "0A";
            }
        }
        #endregion

        #region 标准电流读取
        private delegate byte[] CurReadr();
        //标准电流读取报文接收
        private byte[] creadB()
        {
            //Send：01 03 00 00 00 01 84 0A
            try
            {
                serialPortB.ReadExisting();
            }
            catch (InvalidOperationException)
            {
                MessageBox.Show("串口连接错误");
                Invoke((MethodInvoker)delegate ()
                {
                    toolStripMenuItemSerialPort.PerformClick();
                });
                //ToolStripMenuItemSerialPort.
            }
            byte[] buffer = { 0x01, 0x03, 0x00, 0x00, 0x00, 0x01, 0x84, 0x0A };
            sendB(buffer);

            buffer = new byte[serialPortB.ReadBufferSize];
            try
            {
                Thread.Sleep(Config.Default.TimeForReadAfterWrite);
                serialPortB.Read(buffer, 0, buffer.Length);
                string str = "Recive:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
                showRichTextBox(richTextBox1, str);
                return buffer;
            }
            catch (TimeoutException)
            {
                MessageBox.Show("MM,这里超时了");
                return null;
            }
            catch { return null; }
        }
        private byte[] cread1()
        {
            //Send：01 03 00 00 00 01 84 0A
            try
            {
                serialPort1.ReadExisting();
            }
            catch (InvalidOperationException)
            {
                MessageBox.Show("串口连接错误");
                Invoke((MethodInvoker)delegate ()
                {
                    toolStripMenuItemSerialPort.PerformClick();
                });
                //ToolStripMenuItemSerialPort.
            }
            byte[] buffer = { 0x01, 0x03, 0x00, 0x00, 0x00, 0x01, 0x84, 0x0A };
            send1(buffer);

            buffer = new byte[serialPort1.ReadBufferSize];
            try
            {
                Thread.Sleep(Config.Default.TimeForReadAfterWrite);
                serialPort1.Read(buffer, 0, buffer.Length);
                string str = "Recive:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
                showRichTextBox(richTextBox1, str);
                return buffer;
            }
            catch (TimeoutException)
            {
                return null;
            }
            catch { return null; }
        }
        private byte[] cread2()
        {
            //Send：01 03 00 00 00 01 84 0A
            try
            {
                serialPort2.ReadExisting();
            }
            catch (InvalidOperationException)
            {
                MessageBox.Show("串口连接错误");
                Invoke((MethodInvoker)delegate ()
                {
                    toolStripMenuItemSerialPort.PerformClick();
                });
                //ToolStripMenuItemSerialPort.
            }
            byte[] buffer = { 0x01, 0x03, 0x00, 0x00, 0x00, 0x01, 0x84, 0x0A };
            send2(buffer);

            buffer = new byte[serialPort2.ReadBufferSize];
            try
            {
                Thread.Sleep(Config.Default.TimeForReadAfterWrite);
                serialPort2.Read(buffer, 0, buffer.Length);
                string str = "Recive:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
                showRichTextBox(richTextBox2, str);
                return buffer;
            }
            catch (TimeoutException)
            {
                return null;
            }
            catch { return null; }
        }
        private byte[] cread3()
        {
            //Send：01 03 00 00 00 01 84 0A
            try
            {
                serialPort3.ReadExisting();
            }
            catch (InvalidOperationException)
            {
                MessageBox.Show("串口连接错误");
                Invoke((MethodInvoker)delegate ()
                {
                    toolStripMenuItemSerialPort.PerformClick();
                });
                //ToolStripMenuItemSerialPort.
            }
            byte[] buffer = { 0x01, 0x03, 0x00, 0x00, 0x00, 0x01, 0x84, 0x0A };
            send3(buffer);

            buffer = new byte[serialPort3.ReadBufferSize];
            try
            {
                Thread.Sleep(Config.Default.TimeForReadAfterWrite);
                serialPort3.Read(buffer, 0, buffer.Length);
                string str = "Recive:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
                showRichTextBox(richTextBox3, str);
                return buffer;
            }
            catch (TimeoutException)
            {
                return null;
            }
            catch { return null; }
        }
        private byte[] cread4()
        {
            //Send：01 03 00 00 00 01 84 0A
            try
            {
                serialPort4.ReadExisting();
            }
            catch (InvalidOperationException)
            {
                MessageBox.Show("串口连接错误");
                Invoke((MethodInvoker)delegate ()
                {
                    toolStripMenuItemSerialPort.PerformClick();
                });
                //ToolStripMenuItemSerialPort.
            }
            byte[] buffer = { 0x01, 0x03, 0x00, 0x00, 0x00, 0x01, 0x84, 0x0A };
            send4(buffer);

            buffer = new byte[serialPort4.ReadBufferSize];
            try
            {
                Thread.Sleep(Config.Default.TimeForReadAfterWrite);
                serialPort4.Read(buffer, 0, buffer.Length);
                string str = "Recive:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
                showRichTextBox(richTextBox4, str);
                return buffer;
            }
            catch (TimeoutException)
            {
                return null;
            }
            catch { return null; }
        }
        private byte[] cread5()
        {
            //Send：01 03 00 00 00 01 84 0A
            try
            {
                serialPort5.ReadExisting();
            }
            catch (InvalidOperationException)
            {
                MessageBox.Show("串口连接错误");
                Invoke((MethodInvoker)delegate ()
                {
                    toolStripMenuItemSerialPort.PerformClick();
                });
                //ToolStripMenuItemSerialPort.
            }
            byte[] buffer = { 0x01, 0x03, 0x00, 0x00, 0x00, 0x01, 0x84, 0x0A };
            send5(buffer);

            buffer = new byte[serialPort5.ReadBufferSize];
            try
            {
                Thread.Sleep(Config.Default.TimeForReadAfterWrite);
                serialPort5.Read(buffer, 0, buffer.Length);
                string str = "Recive:" + BitConverter.ToString(buffer) + "\t" + DateTime.Now.ToString() + "\r\n";
                showRichTextBox(richTextBox5, str);
                return buffer;
            }
            catch (TimeoutException)
            {
                return null;
            }
            catch { return null; }
        }
        //标准电流读取报文解析
        private void CurRead(int floor)
        {
            //Receive：01 03 02 IH IL
            //电流计算：I = (IH & IL) / 10
            double data = -99;
            byte[] buffer = null;
            if (Config.Default.HardMod == "V0")
                buffer = creadB();
            else if (Config.Default.HardMod == "V1")
            {
                switch (floor)
                {
                    case 1:
                        buffer = cread1();
                        break;
                    case 2:
                        buffer = cread2();
                        break;
                    case 3:
                        buffer = cread3();
                        break;
                    case 4:
                        buffer = cread4();
                        break;
                    case 5:
                        buffer = cread5();
                        break;
                }
            }
            string tmp;
            if (buffer != null)                                      //未接收到数据
            {
                tmp = BitConverter.ToString(buffer);
                tmp = tmp.Replace("-", "");
                int s = tmp.IndexOf("010302");
                if (s != -1)                                         //数据不合法
                {
                    byte[] iitt = new byte[2];
                    iitt = Conv.strToToHexByte(tmp.Substring(s + 6, 4));
                    data = (iitt[0] << 8) + iitt[1];
                    data /= 10;
                }
            }
            Invoke((MethodInvoker)delegate ()
            {
                switch (floor)
                {
                    case 1:
                        label1.Text = data.ToString();
                        label1.BackColor = this.BackColor;
                        break;
                    case 2:
                        label2.Text = data.ToString();
                        label2.BackColor = this.BackColor;
                        break;
                    case 3:
                        label3.Text = data.ToString();
                        label3.BackColor = this.BackColor;
                        break;
                    case 4:
                        label4.Text = data.ToString();
                        label4.BackColor = this.BackColor;
                        break;
                    case 5:
                        label5.Text = data.ToString();
                        label5.BackColor = this.BackColor;
                        break;
                }
            });
        }
        //标准电流读取数据显示
        private void CurRead1()
        {
            CurRead(1);
        }
        private void CurRead2()
        {
            CurRead(2);
        }
        private void CurRead3()
        {
            CurRead(3);
        }
        private void CurRead4()
        {
            CurRead(4);
        }
        private void CurRead5()
        {
            CurRead(5);
        }
        //标准电流读取启动
        private bool CurReads()
        {
            int i = 0;
            if ((string)toolStripButton1.Tag == "Selected")
            {
                threadFloor1 = new Thread(new ThreadStart(CurRead1));
                threadFloor1.IsBackground = true;
                threadFloor1.Start();
                i++;
            }
            if ((string)toolStripButton2.Tag == "Selected")
            {
                threadFloor2 = new Thread(new ThreadStart(CurRead2));
                threadFloor2.IsBackground = true;
                threadFloor2.Start();
                i++;
            }
            if ((string)toolStripButton3.Tag == "Selected")
            {
                threadFloor3 = new Thread(new ThreadStart(CurRead3));
                threadFloor3.IsBackground = true;
                threadFloor3.Start();
                i++;
            }
            if ((string)toolStripButton4.Tag == "Selected")
            {
                threadFloor4 = new Thread(new ThreadStart(CurRead4));
                threadFloor4.IsBackground = true;
                threadFloor4.Start();
                i++;
            }
            if ((string)toolStripButton5.Tag == "Selected")
            {
                threadFloor5 = new Thread(new ThreadStart(CurRead5));
                threadFloor5.IsBackground = true;
                threadFloor5.Start();
                i++;
            }
            if (i == 0)
            {
                autoESC();
                MessageBox.Show("未选中任何测试层");
                tshow("执行终止:未选中层");
                return false;
            }
            return true;
        }
        #endregion

        #region 指示器电流设置

        private bool SetCur(byte traget, int floor)
        {
            Label lab1 = null, lab2 = null;
            switch (floor)
            {
                case 1:
                    lab1 = label1;
                    lab2 = label6;
                    break;
                case 2:
                    lab1 = label2;
                    lab2 = label8;
                    break;
                case 3:
                    lab1 = label3;
                    lab2 = label10;
                    break;
                case 4:
                    lab1 = label4;
                    lab2 = label12;
                    break;
                case 5:
                    lab1 = label5;
                    lab2 = label14;
                    break;
            }

            if (LabelShow(lab1, lab2))
            {
                string data = lab1.Text;
                string data2 = lab2.Text;
                double tmp = double.Parse(data);
                int tmp2 = (int)(tmp * 10);
                byte tmp3 = byte.Parse(data2.Substring(1, data2.Length - 2));
                byte[] tmp4 = { (byte)tmp2, (byte)(tmp2 >> 8), tmp3 };
                MessageSend(traget, 0x04, tmp4, floor);
                return true;
            }
            else
                return false;
        }

        private bool LabelShow(Label lab1, Label lab2)
        {
            if (lab2.Text == "未开启")
            {
                tshow("标准电流未设置");
                for (int i = 0; i < 10; i++)
                {
                    Invoke((MethodInvoker)delegate ()
                    {
                        lab2.BackColor = Color.Yellow;
                    });
                    Thread.Sleep(100);
                    Invoke((MethodInvoker)delegate ()
                    {
                        lab2.BackColor = this.BackColor;
                    });
                    Thread.Sleep(100);
                }
            }
            else if (lab1.Text == "Close" || lab1.Text == "None")
            {
                tshow("标准电流未读取");
                for (int i = 0; i < 10; i++)
                {
                    Invoke((MethodInvoker)delegate ()
                    {
                        lab1.BackColor = Color.Yellow;
                    });
                    Thread.Sleep(100);
                    Invoke((MethodInvoker)delegate ()
                    {
                        lab1.BackColor = this.BackColor;
                    });
                    Thread.Sleep(100);
                }
            }
            else if (lab1.Text == "-99")
            {
                tshow("标准电流读取失败");
                Invoke((MethodInvoker)delegate ()
                {
                    lab1.BackColor = Color.Yellow;
                });
            }
            else
            {
                return true;
            }
            return false;
        }

        private void SetCur1()
        {
            SetCur(0xFF, 1);
        }
        private void SetCur2()
        {
            SetCur(0xFF, 2);
        }
        private void SetCur3()
        {
            SetCur(0xFF, 3);
        }
        private void SetCur4()
        {
            SetCur(0xFF, 4);
        }
        private void SetCur5()
        {
            SetCur(0xFF, 5);
        }
        private bool SetCurs()
        {
            int i = 0;
            if ((string)toolStripButton1.Tag == "Selected")
            {
                threadFloor1 = new Thread(new ThreadStart(SetCur1));
                threadFloor1.IsBackground = true;
                threadFloor1.Start();
                i++;
            }
            if ((string)toolStripButton2.Tag == "Selected")
            {
                threadFloor2 = new Thread(new ThreadStart(SetCur2));
                threadFloor2.IsBackground = true;
                threadFloor2.Start();
                i++;
            }
            if ((string)toolStripButton3.Tag == "Selected")
            {
                threadFloor3 = new Thread(new ThreadStart(SetCur3));
                threadFloor3.IsBackground = true;
                threadFloor3.Start();
                i++;
            }
            if ((string)toolStripButton4.Tag == "Selected")
            {
                threadFloor4 = new Thread(new ThreadStart(SetCur4));
                threadFloor4.IsBackground = true;
                threadFloor4.Start();
                i++;
            }
            if ((string)toolStripButton5.Tag == "Selected")
            {
                threadFloor5 = new Thread(new ThreadStart(SetCur5));
                threadFloor5.IsBackground = true;
                threadFloor5.Start();
                i++;
            }
            if (i == 0)
            {
                autoESC();
                MessageBox.Show("未选中任何测试层");
                tshow("执行终止:未选中层");
                return false;
            }
            return true;
        }

        #endregion

        #region Design
        private void ToolStripMenuItemSend_Click(object sender, EventArgs e)
        {
            byte[] by1 = { 0x00 };
            byte[] by4 = { 0xAA, 0x55, 0x10 };
            MessageSend(0x01, 0x01, by1, 1);
            MessageSend(0x01, 0x02, null, 2);
            MessageSend(0x01, 0x03, null, 3);
            MessageSend(0x01, 0x04, by4, 4);
            MessageSend(0x01, 0x05, null, 5);
            MessageSend(0x01, 0x06, null, 1);
            MessageSend(0x01, 0x08, null, 2);
            MessageSend(0xFF, 0x01, by1, 3);
            MessageSend(0xFF, 0x02, null, 4);
            MessageSend(0xFF, 0x03, null, 5);
            MessageSend(0xFF, 0x04, by4, 1);
            MessageSend(0xFF, 0x05, null, 2);
            MessageSend(0xFF, 0x06, null, 3);
            MessageSend(0xFF, 0x08, null, 4);
            MessageSend(0x01, null, null, 5);
        }
        private void treadTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            threadFloor1 = new Thread(new ThreadStart(floor1Worker));
            threadFloor1.IsBackground = true;
            threadFloor1.Start();
            threadFloor2 = new Thread(new ThreadStart(floor2Worker));
            threadFloor2.IsBackground = true;
            threadFloor2.Start();
            threadFloor3 = new Thread(new ThreadStart(floor3Worker));
            threadFloor3.IsBackground = true;
            threadFloor3.Start();
            threadFloor4 = new Thread(new ThreadStart(floor4Worker));
            threadFloor4.IsBackground = true;
            threadFloor4.Start();
            threadFloor5 = new Thread(new ThreadStart(floor5Worker));
            threadFloor5.IsBackground = true;
            threadFloor5.Start();
        }
        private void serialOpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            serialOpen1();
            serialOpen2();
            serialOpen3();
            serialOpen4();
            serialOpen5();
        }
        private void toolStripMenuItemDatatest_Click(object sender, EventArgs e)
        {

            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                foreach (DataGridViewColumn item2 in dataGridView1.Columns)
                {
                    if (item2.HeaderText == "详情")
                    {
                    }
                    else if (item2.HeaderText == "编号")
                    {
                    }
                    else
                    {
                        item.Cells[item2.Name].Value = 100;
                    }
                }
            }

        }
        private void dataResetToolStripMenuItem_Click(object sender, EventArgs e)
        {

            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                for (int j = 0; j < 5; j++)
                {
                    item.Cells[6 * j + 1].Value = "0A";
                    item.Cells[6 * j + 2].Value = "0%";
                    item.Cells[6 * j + 3].Value = "0℃";
                    item.Cells[6 * j + 4].Value = "0℃";

                    item.Cells[6 * j + 0].Style.BackColor = Color.White;
                    item.Cells[6 * j + 1].Style.BackColor = Color.White;
                    item.Cells[6 * j + 2].Style.BackColor = Color.White;
                    item.Cells[6 * j + 3].Style.BackColor = Color.White;
                    item.Cells[6 * j + 4].Style.BackColor = Color.White;
                    item.Cells[6 * j + 5].Style.BackColor = Color.White;
                }
            }
        }
        #endregion

        #region 菜单
        private void ToolStripMenuItemSerialPort_Click(object sender, EventArgs e)
        {

            toolStripButtonStop.PerformClick();
            serialClose();

            SerialSet windows = new SerialSet();
            windows.ShowInTaskbar = false;
            windows.ShowDialog();

            serialOpen1();
            serialOpen2();
            serialOpen3();
            serialOpen4();
            serialOpen5();
        }

        private void toolStripMenuItemSoftWareSet_Click(object sender, EventArgs e)
        {
            SoftwareSet windows = new SoftwareSet();
            windows.ShowInTaskbar = false;
            windows.ShowDialog();
            toolStripSplitButton1.Text = Config.Default.Mode;
            testModeChange();
        }

        private void ToolStripMenuItemShutdown_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion

        #region 控制
        private void cutCurToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CurCloses();
        }
        private void curSetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem item = (ToolStripMenuItem)sender;
            int i = int.Parse((string)item.Tag);
            CurOpens(i);
            //注释原因：手动设置电流后，应手动读取电流
            //if (Config.Default.ReadAfterSet)
            //{
            //    CurReads();
            //}
        }
        private void curReadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CurReads();
        }
        private void cutAllCurToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CurCloseAll();
        }
        #endregion

        #region 测试
        private void lEDTestToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MessageSends(0xFF, 0x05, null);
        }

        private void resetToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MessageSends(0xFF, 0x06, null);
        }
        #endregion

        #region 功能
        private void readDataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageSends(0xFF, 0x02, null))
                Reads(true);
        }

        private void readRealToolStripMenuItem_Click(object sender, EventArgs e)
        {
            autoESC();
            MessageBox.Show("还没编写");
        }

        private void setCurrentToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetCurs();
        }
        #endregion

        #region 其他
        private void workModeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageSends(0xFF, 0x01, new byte[] { 0x00 });
            workModeToolStripMenuItem.Checked = true;
            testModeToolStripMenuItem.Checked = false;
        }
        private void testModeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageSends(0xFF, 0x01, new byte[] { 0x01 }))
            {
                testModeToolStripMenuItem.Checked = true;
                workModeToolStripMenuItem.Checked = false;
            }
        }
        private void systemReadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageSends(0x01, 0x02, null);
        }

        private void readFloor1Item1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ReadOne(1, 2);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            CurCloseAll();
            try
            {
                thread1.Abort();
            }
            catch { }
            try
            {
                thread2.Abort();
            }
            catch { }
            try
            {
                thread3.Abort();
            }
            catch { }
            try
            {
                thread4.Abort();
            }
            catch { }
            try
            {
                thread5.Abort();
            }
            catch { }

            serialPort1.Close();
            serialPort2.Close();
            serialPort3.Close();
            serialPort4.Close();
            serialPort5.Close();
        }

        private void setTo3AToolStripMenuItem_Click(object sender, EventArgs e)
        {
            label1.Text = "3";
            label2.Text = "3";
            label3.Text = "3";
            label4.Text = "3";
            label5.Text = "3";
        }

        #endregion

        #region 关于
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            connectFlag = false;
            connectTestToolStripButton.Enabled = true;
            autoESC();
            Random a = new Random();
            switch (a.Next(0, 10))
            {
                case 1:
                    MessageBox.Show("要成功，自身硬", "关于");
                    break;
                case 2:
                    MessageBox.Show("北京科锐配电自动化股份有限公司", "关于");
                    break;
                case 3:
                    MessageBox.Show("自动化事业部", "关于");
                    break;
                case 4:
                    MessageBox.Show("故障定位系统产品部", "关于");
                    break;
                case 5:
                    MessageBox.Show("自动化-指示器-技术组", "关于");
                    break;
                case 6:
                    MessageBox.Show("自动化-指示器-技术组", "关于");
                    break;
                case 7:
                    MessageBox.Show("3C1O和3C1O-T的测试程序", "关于");
                    break;
                case 8:
                    MessageBox.Show("要成功，自身硬", "关于");
                    break;
                case 9:
                    MessageBox.Show("自动化-指示器-技术组", "关于");
                    break;
            }

        }



        #endregion

        #region 快捷测试
        Thread threadLEDTest;

        private void toolAutoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (modeAutoToolStripMenuItem.Checked == true)
            {
                modeAutoToolStripMenuItem.Checked = false;
                toolStrip1.Visible = false;
            }
            else
            {
                modeAutoToolStripMenuItem.Checked = true;
                toolStrip1.Visible = true;
            }
        }

        private void toolFloorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (modeFloorToolStripMenuItem.Checked == true)
            {
                modeFloorToolStripMenuItem.Checked = false;
                toolStrip2.Visible = false;
            }
            else
            {
                modeFloorToolStripMenuItem.Checked = true;
                toolStrip2.Visible = true;
            }
        }

        private void toolModeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (modeModeToolStripMenuItem.Checked == true)
            {
                modeModeToolStripMenuItem.Checked = false;
                toolStrip3.Visible = false;
            }
            else
            {
                modeModeToolStripMenuItem.Checked = true;
                toolStrip3.Visible = true;
            }
        }

        private void toolStripComboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

            Config.Default.SoftWareMode = toolStripComboBox6.Text;
            Config.Default.Save();

            if (Config.Default.SoftWareMode == "自动模式")
            {
                toolStripComboBox6.Text = "自动模式";

                //floor1ToolStripMenuItem.Enabled = false;
                //floor2ToolStripMenuItem.Enabled = false;
                //floor3ToolStripMenuItem.Enabled = false;
                //floor4ToolStripMenuItem.Enabled = false;
                //floor5ToolStripMenuItem.Enabled = false;

                floor5ToolStripMenuItem.Checked = false;
                floor4ToolStripMenuItem.Checked = false;
                floor3ToolStripMenuItem.Checked = false;
                floor2ToolStripMenuItem.Checked = false;
                floor1ToolStripMenuItem.Checked = false;

                toolStrip8.Visible = false;
                toolStrip7.Visible = false;
                toolStrip6.Visible = false;
                toolStrip5.Visible = false;
                toolStrip4.Visible = false;

                //modeAutoToolStripMenuItem.Enabled = true;
                //modeFloorToolStripMenuItem.Enabled = true;
                //modeModeToolStripMenuItem.Enabled = true;

                modeAutoToolStripMenuItem.Checked = true;
                modeFloorToolStripMenuItem.Checked = true;
                modeModeToolStripMenuItem.Checked = true;

                toolStrip1.Visible = true;
                toolStrip2.Visible = true;
                toolStrip3.Visible = true;


            }
            else
            {
                toolStripComboBox6.Text = "手动模式";

                //modeAutoToolStripMenuItem.Enabled = false;
                //modeFloorToolStripMenuItem.Enabled = false;
                //modeModeToolStripMenuItem.Enabled = false;

                modeModeToolStripMenuItem.Checked = false;
                modeFloorToolStripMenuItem.Checked = false;
                modeAutoToolStripMenuItem.Checked = false;

                toolStrip3.Visible = false;
                toolStrip2.Visible = false;
                toolStrip1.Visible = false;

                //floor1ToolStripMenuItem.Enabled = true;
                //floor2ToolStripMenuItem.Enabled = true;
                //floor3ToolStripMenuItem.Enabled = true;
                //floor4ToolStripMenuItem.Enabled = true;
                //floor5ToolStripMenuItem.Enabled = true;

                floor1ToolStripMenuItem.Checked = true;
                floor2ToolStripMenuItem.Checked = true;
                floor3ToolStripMenuItem.Checked = true;
                floor4ToolStripMenuItem.Checked = true;
                floor5ToolStripMenuItem.Checked = true;

                toolStrip4.Visible = true;
                toolStrip5.Visible = true;
                toolStrip6.Visible = true;
                toolStrip7.Visible = true;
                toolStrip8.Visible = true;

            }

        }

        private void floor1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (floor1ToolStripMenuItem.Checked == true)
            {
                floor1ToolStripMenuItem.Checked = false;
                toolStrip4.Visible = false;
            }
            else
            {
                floor1ToolStripMenuItem.Checked = true;
                toolStrip4.Visible = true;
            }
        }

        private void floor2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (floor2ToolStripMenuItem.Checked == true)
            {
                floor2ToolStripMenuItem.Checked = false;
                toolStrip5.Visible = false;
            }
            else
            {
                floor2ToolStripMenuItem.Checked = true;
                toolStrip5.Visible = true;
            }
        }

        private void floor3ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (floor3ToolStripMenuItem.Checked == true)
            {
                floor3ToolStripMenuItem.Checked = false;
                toolStrip6.Visible = false;
            }
            else
            {
                floor3ToolStripMenuItem.Checked = true;
                toolStrip6.Visible = true;
            }
        }

        private void floor4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (floor4ToolStripMenuItem.Checked == true)
            {
                floor4ToolStripMenuItem.Checked = false;
                toolStrip7.Visible = false;
            }
            else
            {
                floor4ToolStripMenuItem.Checked = true;
                toolStrip7.Visible = true;
            }
        }

        private void floor5ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (floor5ToolStripMenuItem.Checked == true)
            {
                floor5ToolStripMenuItem.Checked = false;
                toolStrip8.Visible = false;
            }
            else
            {
                floor5ToolStripMenuItem.Checked = true;
                toolStrip8.Visible = true;
            }
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {

            MessageBox.Show(((ToolStripButton)sender).Tag.ToString());
            /*
            if (MessageSends(0xFF, 0x02, null))
                Reads(true);*/
        }

        private void toolStripButtonSet_Click(object sender, EventArgs e)
        {
            int floor = int.Parse(((ToolStripButton)sender).Tag.ToString());
            try {
                ((ToolStripButton)sender).ForeColor = Color.Black;
                Read(true, floor);
            }
            catch { ((ToolStripButton)sender).ForeColor = Color.Red; }
        }

        private void toolStripButtonRead_Click(object sender, EventArgs e)
        {
            int floor = int.Parse(((ToolStripButton)sender).Tag.ToString());
            try {
                ((ToolStripButton)sender).ForeColor = Color.Black;
                SetCur(0xFF, floor);
            }
            catch { ((ToolStripButton)sender).ForeColor = Color.Red; }
        }

        private void toolStripLabelOpen_Click(object sender, EventArgs e)
        {
            int floor = int.Parse(((ToolStripLabel)sender).Tag.ToString());

            Control[] a = Controls.Find("toolStrip" + (floor + 3), true);
            ToolStrip theComboBox = a[0] as ToolStrip;
            //((ToolStripLabel)Controls.Find("toolStripLabel" + (floor + 1), true)[0]).ForeColor = Color.Green;
            theComboBox.Items[0].ForeColor = Color.Green;
            //if (serialOpen1()) ((ToolStripLabel)a[0]).ForeColor = Color.Green;
            //else ((ToolStripLabel)a[0]).ForeColor = Color.Red;


            /*switch (floor)
            {
                case 1:
                    if (serialOpen1()) toolStripLabel2.ForeColor = Color.Green;
                    else toolStripLabel2.ForeColor = Color.Red;
                    break;
                case 2:
                    if (serialOpen2()) toolStripLabel3.ForeColor = Color.Green;
                    else toolStripLabel3.ForeColor = Color.Red;
                    break;
                case 3:
                    if (serialOpen3()) toolStripLabel4.ForeColor = Color.Green;
                    else toolStripLabel4.ForeColor = Color.Red;
                    break;
                case 4:
                    if (serialOpen4()) toolStripLabel5.ForeColor = Color.Green;
                    else toolStripLabel5.ForeColor = Color.Red;
                    break;
                case 5:
                    if (serialOpen5()) toolStripLabel6.ForeColor = Color.Green;
                    else toolStripLabel6.ForeColor = Color.Red;
                    break;
            }*/
        }

        private void toolStripButtonNext_Click(object sender, EventArgs e)
        {
            int floor = int.Parse(((ToolStripButton)sender).Tag.ToString());
            switch (floor)
            {
                case 1:
                    if (toolStripSplitButton2.SelectedIndex >= 17)
                        toolStripSplitButton2.SelectedIndex = 0;
                    else
                        toolStripSplitButton2.SelectedIndex++;
                    break;
                case 2:
                    if (toolStripComboBox2.SelectedIndex >= 17)
                        toolStripComboBox2.SelectedIndex = 0;
                    else
                        toolStripComboBox2.SelectedIndex++;
                    break;
                case 3:
                    if (toolStripComboBox3.SelectedIndex >= 17)
                        toolStripComboBox3.SelectedIndex = 0;
                    else
                        toolStripComboBox3.SelectedIndex++;
                    break;
                case 4:
                    if (toolStripComboBox4.SelectedIndex >= 17)
                        toolStripComboBox4.SelectedIndex = 0;
                    else
                        toolStripComboBox4.SelectedIndex++;
                    break;
                case 5:
                    if (toolStripComboBox5.SelectedIndex >= 17)
                        toolStripComboBox5.SelectedIndex = 0;
                    else
                        toolStripComboBox5.SelectedIndex++;
                    break;
            }

        }

        private void toolStripButton20_Click(object sender, EventArgs e)
        {

        }

        private void toolStripSplitButton2_Click(object sender, EventArgs e)
        {

        }

        private void toolStripComboBox2_Click(object sender, EventArgs e)
        {

        }

        private void lastTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(Config.Default.lastTest, Config.Default.lastTime);
        }

        private void ledTest()
        {
            if (MessageSends(0xFF, 0x05, null))
            {
                MessageBox.Show("已开始LED测试，再次点击停止测试");
                //Invoke((MethodInvoker)delegate ()
                //{
                //});
            }
            else
            {
                Invoke((MethodInvoker)delegate ()
                {
                    ledTestToolStripButton.Checked = false;
                });
                return;
            }
            while (true)
            {
                if (!MessageSends(0xFF, 0x05, null)) break;
                Thread.Sleep(Config.Default.TimeForLEDBlinkInterval);
            }

        }

        private void Form1_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void Form1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F2 && connectTestToolStripButton.Enabled == true) connectTestToolStripButton_Click(sender, e);
            else if (e.KeyCode == Keys.F3 && ledTestToolStripButton.Enabled == true) ledTestToolStripButton_Click(sender, e);
            else if (e.KeyCode == Keys.F4 && toolStripButtonStart.Enabled == true) toolStripButtonStart_Click(sender, e);
            else if (e.KeyCode == Keys.F5 && toolStripButtonStop.Enabled == true) toolStripButtonStop_Click(sender, e);
        }

        private void ledTestToolStripButton_Click(object sender, EventArgs e)
        {
            ledTestToolStripButton.Checked = !ledTestToolStripButton.Checked;
            if (ledTestToolStripButton.Checked == true)
            {
                threadLEDTest = new Thread(new ThreadStart(ledTest));
                threadLEDTest.IsBackground = true;
                threadLEDTest.Start();
                autoESC();
                tshow("正在进行LED测试……");
            }
            else
            {
                threadLEDTest.Abort();
                tshow("LED测试手动结束");
            }
        }

        private void toolStripButton21_Click(object sender, EventArgs e)
        {
            if (!serialPortA.IsOpen || !serialPortB.IsOpen)
            {
                if (!serialOpenAB())
                {
                    autoESC();
                    MessageBox.Show("串口" + Config.Default.SerialA.Substring(3) + "与" + Config.Default.SerialB.Substring(3) + "打开失败");
                    tshow("标准电流端口打开失败");
                }
            }
            //123
        }

        bool connectFlag = false;
        private void connectTestToolStripButton_Click(object sender, EventArgs e)
        {
            if (connectFlag == false)
            {
                connectFlag = true;
                connectTestToolStripButton.Enabled = false;
                tshow("正在执行通讯准备……");
                if (!serialPort1.IsOpen)
                    serialOpen1();
                if (!serialPort2.IsOpen)
                    serialOpen2();
                if (!serialPort3.IsOpen)
                    serialOpen3();
                if (!serialPort4.IsOpen)
                    serialOpen4();
                if (!serialPort5.IsOpen)
                    serialOpen5();
                if (!serialPortA.IsOpen || !serialPortB.IsOpen)
                    serialOpenAB();

                if (threadLEDTest != null)
                    if (threadLEDTest.IsAlive)
                    {
                        threadLEDTest.Abort();
                        ledTestToolStripButton.Checked = false;
                    }
                if (MessageSends(0xFF, 0x02, null))
                    Reads(false);
            }
        }

        private void 报文ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (richTextBox6.Visible)
                richTextBox6.Visible = false;
            else
                richTextBox6.Visible = true;
        }

        private void toolStripButtonStart_Click(object sender, EventArgs e)
        {
            toolStripStatusLabel9.Visible = true;
            tshow("自动测试开始");
            richTextBox6.Text = "";
            if (threadLEDTest != null)
                if (threadLEDTest.IsAlive)
                {
                    threadLEDTest.Abort();
                    ledTestToolStripButton.Checked = false;
                }
            if (floorsWorker())
            {
                toolStripButtonStop.Enabled = true;
                toolStripButtonStart.Enabled = false;

                connectTestToolStripButton.Enabled = false;
                ledTestToolStripButton.Enabled = false;
                toolStripButton1.Enabled = false;
                toolStripButton2.Enabled = false;
                toolStripButton3.Enabled = false;
                toolStripButton4.Enabled = false;
                toolStripButton5.Enabled = false;
                toolStripSplitButton1.Enabled = false;


                menuToolStripMenuItem.Enabled = false;
                viewToolStripMenuItem.Enabled = false;
                ctrlToolStripMenuItem.Enabled = false;
                testToolStripMenuItem.Enabled = false;
                commendToolStripMenuItem.Enabled = false;
                otherToolStripMenuItem.Enabled = false;
            }
        }

        private void toolStripButton22_Click(object sender, EventArgs e)
        {
            if (floorsWorker0())
            {
                toolStripButtonStop.Enabled = true;
                toolStripButtonStart.Enabled = false;
                toolStripButton22.Enabled = false;

                connectTestToolStripButton.Enabled = false;
                ledTestToolStripButton.Enabled = false;
                toolStripButton1.Enabled = false;
                toolStripButton2.Enabled = false;
                toolStripButton3.Enabled = false;
                toolStripButton4.Enabled = false;
                toolStripButton5.Enabled = false;
                toolStripSplitButton1.Enabled = false;


                menuToolStripMenuItem.Enabled = false;
                viewToolStripMenuItem.Enabled = false;
                ctrlToolStripMenuItem.Enabled = false;
                testToolStripMenuItem.Enabled = false;
                commendToolStripMenuItem.Enabled = false;
                otherToolStripMenuItem.Enabled = false;
            }
        }
        private void floor1Worker0()
        {
            autoRun(-1);
        }
        private void floor2Worker0()
        {
            autoRun(-2);
        }
        private void floor3Worker0()
        {
            autoRun(-3);
        }
        private void floor4Worker0()
        {
            autoRun(-4);
        }
        private void floor5Worker0()
        {
            autoRun(-5);
        }
        private bool floorsWorker0()
        {
            int i = 0;
            if ((string)toolStripButton1.Tag == "Selected")
            {
                i++;
                thread1 = new Thread(new ThreadStart(floor1Worker0));
                thread1.IsBackground = true;
                thread1.Start();
            }
            if ((string)toolStripButton2.Tag == "Selected")
            {
                i++;
                thread2 = new Thread(new ThreadStart(floor2Worker0));
                thread2.IsBackground = true;
                thread2.Start();
            }
            if ((string)toolStripButton3.Tag == "Selected")
            {
                i++;
                thread3 = new Thread(new ThreadStart(floor3Worker0));
                thread3.IsBackground = true;
                thread3.Start();
            }
            if ((string)toolStripButton4.Tag == "Selected")
            {
                i++;
                thread4 = new Thread(new ThreadStart(floor4Worker0));
                thread4.IsBackground = true;
                thread4.Start();
            }
            if ((string)toolStripButton5.Tag == "Selected")
            {
                i++;
                thread5 = new Thread(new ThreadStart(floor5Worker0));
                thread5.IsBackground = true;
                thread5.Start();
            }
            if (i == 0)
            {
                autoESC();
                MessageBox.Show("未选中任何测试层");
                tshow("执行终止:未选中层");
                return false;
            }
            return true;
        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //string name = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "_" + DateTime.Now.Second;
            //name=name.Replace("/", "-");
            //name=name.Replace(":", "_");
            //write_txt(name, richTextBox6.Text);

            log.Out_Log("测试信息", richTextBox6.Text);
        }

        bool FinishFlg = false;
        private void toolStripButtonStop_Click(object sender, EventArgs e)
        {
            if (!FinishFlg)
                if (MessageBox.Show("测试进程正在运行，确定要提前退出本次测试吗？", "提示", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
            FinishFlg = false;
            toolStripStatusLabel9.Visible = false;
            CurCloseAll();
            toolStripProgressBar2.Value = 0;
            toolStripProgressBar3.Value = 0;
            toolStripProgressBar4.Value = 0;
            toolStripProgressBar5.Value = 0;
            toolStripProgressBar6.Value = 0;
            try
            {
                thread1.Abort();
            }
            catch { }
            try
            {
                thread2.Abort();
            }
            catch { }
            try
            {
                thread3.Abort();
            }
            catch { }
            try
            {
                thread4.Abort();
            }
            catch { }
            try
            {
                thread5.Abort();
            }
            catch { }
            //2016-01-22 防止误操作
            //autoESC();
            //MessageBox.Show("测试已结束");


            string mes = "";
            string wcd = "";
            string wtx = "";
            string good = "";
            int wcds = 0, wtxs = 0, goods = 0;
            int wcdsa = 0, wtxsa = 0, gooda = 0;
            for (int i = 0; i < 5; i++)
            {
                logD[i] = null;
                wcds = 0; wtxs = 0; goods = 0;
                mes += "第" + (i + 1) + "层:";
                wcd = "误差大：";
                wtx = "无通讯：";
                for (int j = 0; j < 28; j++)
                {
                    if (FI[i + 1, j + 1] == 3) //误差大
                    {
                        wcd += (i + 1) + "-" + (j + 2) + ";";
                        wcds++;
                    }
                    if (FI[i + 1, j + 1] == 1) //无通讯
                    {
                        wtx += (i + 1) + "-" + (j + 2) + ";";
                        wtxs++;
                    }
                    if (FI[i + 1, j + 1] == 2) //haode
                    {
                        good += (i + 1) + "-" + (j + 2) + ";";
                        goods++;
                    }
                }
                mes += "\t" + "测试通过" + goods + "个";
                mes += "\r\n\t" + wcd;
                if (wtxs != 0)
                    mes += "\r\n\t" + wtx;
                mes += "\r\n";
                wcdsa += wcds;
                wtxsa += wtxs;
                gooda += goods;
            }
            mes += "\r\n总计：测试通过" + gooda + "个\t误差大" + wcdsa + "个";
            mes += "\t无通讯" + wtxsa + "个";
            mes += "\r\n";
            Config.Default.lastTest = mes;
            Config.Default.lastTime = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "的测试结果";
            Config.Default.Save();

            //string name = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "_" + DateTime.Now.Second;
            //name = name.Replace("/", "-");
            //name = name.Replace(":", "_");
            //write_txt(name, mes);

            log.Out_Log("测试结果", "\r\n" + mes);

            MessageBox.Show(mes, "测试结果");


            toolStripButtonStop.Enabled = false;
            toolStripButtonStart.Enabled = true;

            connectTestToolStripButton.Enabled = true;
            ledTestToolStripButton.Enabled = true;
            toolStripButton1.Enabled = true;
            toolStripButton2.Enabled = true;
            toolStripButton3.Enabled = true;
            toolStripButton4.Enabled = true;
            toolStripButton5.Enabled = true;
            toolStripSplitButton1.Enabled = true;

            menuToolStripMenuItem.Enabled = true;
            viewToolStripMenuItem.Enabled = true;
            ctrlToolStripMenuItem.Enabled = true;
            testToolStripMenuItem.Enabled = true;
            commendToolStripMenuItem.Enabled = true;
            otherToolStripMenuItem.Enabled = true;

            label6.Text = "未开启";
            label8.Text = "未开启";
            label10.Text = "未开启";
            label12.Text = "未开启";
            label14.Text = "未开启";

            label7.Text = "000A";
            label9.Text = "000A";
            label11.Text = "000A";
            label13.Text = "000A";
            label15.Text = "000A";

            label1.Text = "Close";
            label2.Text = "Close";
            label3.Text = "Close";
            label4.Text = "Close";
            label5.Text = "Close";
            tshow("自动测试终止");

        }
        #endregion

        private void toolStripSplitButton1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Config.Default.Mode = toolStripSplitButton1.Text;
            Config.Default.Save();
            testModeChange();
            tshow("测试模式切换为“" + toolStripSplitButton1.Text + "”测试模式");
        }

        private void testModeChange()
        {
            if (Config.Default.Mode == "3C1O-T")
            {
                toolStripProgressBar1.Value = 0;
                colSFI_Tin1.Visible = true;
                colSFI_Tou1.Visible = true;
                colSFI_Tin2.Visible = true;
                colSFI_Tou2.Visible = true;
                colSFI_Tin3.Visible = true;
                colSFI_Tou3.Visible = true;
                colSFI_Tin4.Visible = true;
                colSFI_Tou4.Visible = true;
                colSFI_Tin5.Visible = true;
                colSFI_Tou5.Visible = true;
                dataGridView1.Invalidate();
                toolStripProgressBar1.Value = 60;
                //dataGridView1.Update();
            }
            else
            {
                toolStripProgressBar1.Value = 0;
                colSFI_Tin1.Visible = false;
                colSFI_Tou1.Visible = false;
                colSFI_Tin2.Visible = false;
                colSFI_Tou2.Visible = false;
                colSFI_Tin3.Visible = false;
                colSFI_Tou3.Visible = false;
                colSFI_Tin4.Visible = false;
                colSFI_Tou4.Visible = false;
                colSFI_Tin5.Visible = false;
                colSFI_Tou5.Visible = false;
                dataGridView1.Invalidate();
                dataGridView1.Update();
                toolStripProgressBar1.Value = 60;
            }

        }

        private void tshow(string s)
        {
            Invoke((MethodInvoker)delegate ()
            {
                toolStripStatusLabel7.Text = s;
            });
        }
        private void ishow(string s)
        {
            Invoke((MethodInvoker)delegate ()
            {
                toolStripStatusLabel9.Text = s;
            });
            s += "\t" + DateTime.Now.ToShortDateString() 
                + " " + DateTime.Now.ToShortTimeString() 
                + ":" + DateTime.Now.Second.ToString() 
                + ":" + DateTime.Now.Millisecond.ToString();
            Invoke((MethodInvoker)delegate ()
            {
                richTextBox6.AppendText(s + "\r\n");
                richTextBox6.Select(richTextBox6.TextLength, 0);
                richTextBox6.ScrollToCaret();
            });
        }
        private void ishow(string s, string a)
        {
            Invoke((MethodInvoker)delegate ()
            {
                toolStripStatusLabel9.Text = s;
                richTextBox6.AppendText(s);
                richTextBox6.AppendText(a);
                richTextBox6.Select(richTextBox6.TextLength, 0);
                richTextBox6.ScrollToCaret();
            });
        }
        #region 16进制转换
        class Conv
        {
            /// <summary> 
            /// 字符串转16进制字节数组 
            /// </summary> 
            /// <param name="hexString"></param> 
            /// <returns></returns> 
            public static byte[] strToToHexByte(string hexString)
            {
                hexString = hexString.Replace(" ", "");
                if ((hexString.Length % 2) != 0)
                    hexString += " ";
                byte[] returnBytes = new byte[hexString.Length / 2];
                for (int i = 0; i < returnBytes.Length; i++)
                    returnBytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
                return returnBytes;
            }
            /// <summary> 
            /// 字节数组转16进制字符串 
            /// </summary> 
            /// <param name="bytes"></param> 
            /// <returns></returns> 
            public static string byteToHexStr(byte[] bytes)
            {
                string returnStr = "";
                if (bytes != null)
                {
                    for (int i = 0; i < bytes.Length; i++)
                    {
                        returnStr += bytes[i].ToString("X2");
                    }
                }
                return returnStr;
            }
            /// <summary> 
            /// 从汉字转换到16进制 
            /// </summary> 
            /// <param name="s"></param> 
            /// <param name="charset">编码,如"utf-8","gb2312"</param> 
            /// <param name="fenge">是否每字符用逗号分隔</param> 
            /// <returns></returns> 
            public static string ToHex(string s, string charset, bool fenge)
            {
                if ((s.Length % 2) != 0)
                {
                    s += " ";//空格 
                             //throw new ArgumentException("s is not valid chinese string!"); 
                }
                System.Text.Encoding chs = System.Text.Encoding.GetEncoding(charset);
                byte[] bytes = chs.GetBytes(s);
                string str = "";
                for (int i = 0; i < bytes.Length; i++)
                {
                    str += string.Format("{0:X}", bytes[i]);
                    if (fenge && (i != bytes.Length - 1))
                    {
                        str += string.Format("{0}", ",");
                    }
                }
                return str.ToLower();
            }
            ///<summary> 
            /// 从16进制转换成汉字 
            /// </summary> 
            /// <param name="hex"></param> 
            /// <param name="charset">编码,如"utf-8","gb2312"</param> 
            /// <returns></returns> 
            public static string UnHex(string hex, string charset)
            {
                if (hex == null)
                    throw new ArgumentNullException("hex");
                hex = hex.Replace(",", "");
                hex = hex.Replace("\n", "");
                hex = hex.Replace("\\", "");
                hex = hex.Replace(" ", "");
                if (hex.Length % 2 != 0)
                {
                    hex += "20";//空格 
                }
                // 需要将 hex 转换成 byte 数组。 
                byte[] bytes = new byte[hex.Length / 2];
                for (int i = 0; i < bytes.Length; i++)
                {
                    try
                    {
                        // 每两个字符是一个 byte。 
                        bytes[i] = byte.Parse(hex.Substring(i * 2, 2),
                        System.Globalization.NumberStyles.HexNumber);
                    }
                    catch
                    {
                        // Rethrow an exception with custom message. 
                        throw new ArgumentException("hex is not a valid hex number!", "hex");
                    }
                }
                System.Text.Encoding chs = System.Text.Encoding.GetEncoding(charset);
                return chs.GetString(bytes);
            }
        }
        #endregion
        

    }
}
