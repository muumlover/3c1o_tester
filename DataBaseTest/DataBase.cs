using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//操作SQL数据库必须引入此包
using System.Data.SqlClient;
//使用DataSet类必须引入此包
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using ADOX;
using System.IO;

namespace _3C1O_Tester
{
    class DataBase
    {
        private string filePath;
        private string connStr;
        private OleDbConnection conn;
        private OleDbCommand cmdStr;

        private void Create()
        {
            ADOX.Catalog catalog = new Catalog();
            filePath = new DirectoryInfo(".").FullName + @"\data.accdb";
            connStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}", filePath);
            catalog.Create(connStr);

            conn = new OleDbConnection(connStr);
            conn.Open();

            cmdStr = new OleDbCommand();
            cmdStr.Connection = conn;
            cmdStr.CommandText = "create table usrInfo (usrID autoincrement primary key, usrName char(16), grpName char(16), usrSignature char(255))";
            //cmdStr.CommandText = "create table usrInfo (usrID int not null primary key, usrName char(16), grpName char(16), usrSignature char(255))";
            cmdStr.ExecuteNonQuery();

            conn.Close();
        }
        private void Prep()
        {
            filePath = new DirectoryInfo(".").FullName + @"\data.accdb";
            connStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}", filePath);
            conn = new OleDbConnection(connStr);
            cmdStr = new OleDbCommand();
        }

        private void insert()
        {
            if (conn == null) Prep();
            conn.Open();
            cmdStr.Connection = conn;
            cmdStr.CommandText = "insert into usrInfo (usrName,grpName,usrSignature)values('admin','group','happy')";
            cmdStr.ExecuteNonQuery();
            conn.Close();
        }




    }
}
