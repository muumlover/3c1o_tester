using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ADOX;
using System.IO;
using System.Data.OleDb;

namespace DataBaseTest
{
    public partial class Form1 : Form
    {
        //private OleDbConnection conn;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ADOX.Catalog catalog = new Catalog();
            string filePath = new DirectoryInfo(".").FullName + @"\data.accdb";
            string connStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}", filePath);
            catalog.Create(connStr);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string filePath = new DirectoryInfo(".").FullName + @"\data.accdb";
            string connStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}", filePath);
            OleDbConnection conn = new OleDbConnection(connStr);
            conn.Open();

            OleDbCommand cmdStr = new OleDbCommand();
            cmdStr.Connection = conn;
            cmdStr.CommandText = "create table usrInfo (usrID autoincrement primary key, usrName char(16), grpName char(16), usrSignature char(255))";
            //cmdStr.CommandText = "create table usrInfo (usrID int not null primary key, usrName char(16), grpName char(16), usrSignature char(255))";
            cmdStr.ExecuteNonQuery();
            conn.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string filePath = new DirectoryInfo(".").FullName + @"\data.accdb";
            string connStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}", filePath);
            OleDbConnection conn = new OleDbConnection(connStr);
            conn.Open();

            OleDbCommand cmdStr = new OleDbCommand();
            cmdStr.Connection = conn;
    //        string sql = "insert into " + table + "(" +
    //name[0] + "," + name[1] + "," + name[2] + "," + name[3] + "," + name[4] + ")values(" +
    //value[0] + "," + value[1] + "," + value[2] + "," + value[3] + "," + value[4] + ")";
            cmdStr.CommandText = "insert into usrInfo (usrName,grpName,usrSignature)values('admin','group','happy')";
            cmdStr.ExecuteNonQuery();
            conn.Close();
        }
    }
}
