using System;
using System.Windows.Forms;

//using CCWin;
using System.Runtime.InteropServices;

namespace _3C1O_Tester
{
    public partial class Hello : Form
    {
        string[] HA = new string[10];
        public Hello()
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
            label2.Text = HA[new Random().Next(1, 10)];
        }


        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();
        private void Form1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void skinButton1_Click(object sender, EventArgs e)
        {
            if (skinTextBox1.Text == Properties.Resources.Design)
            {
                if (skinTextBox2.Text == Properties.Resources.Pass)
                {
                    this.DialogResult = DialogResult.OK;
                }
            }
            else if (skinTextBox1.Text == Properties.Resources.User)
                if (skinTextBox2.Text == Properties.Resources.Pass)
                {
                    this.DialogResult = DialogResult.Yes;
                }
                else MessageBox.Show("密码错误");
            else MessageBox.Show("用户名错误");
        }

        private void skinButton2_Click(object sender, EventArgs e)
        {
            Application.ExitThread();
        }
    }
}
