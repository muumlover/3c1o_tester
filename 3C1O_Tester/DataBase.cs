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

namespace _3C1O_Tester
{
    class DataBase
    {

        DataGridView dataGridViewBp = new DataGridView();
        static string exePath = System.Environment.CurrentDirectory;//本程序所在路径
                //创建连接对象
        OleDbConnection conn = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;data source=" + exePath + @"\文件名.mdb");

        private void 获取数据表()
        {
            conn.Open();
            //获取数据表
            string sql = "select * from 表名 order by 字段1";
            //查询
            //string sql = "select * from 表名 where 字段2="...;

            OleDbDataAdapter da = new OleDbDataAdapter(sql, conn); //创建适配对象
            DataTable dt = new DataTable(); //新建表对象
            da.Fill(dt); //用适配对象填充表对象
            dataGridViewBp.DataSource = dt; //将表对象作为DataGridView的数据源

            conn.Close();
        }

        //private void 查询数据表()
        private void select(string 表名,string 字段,string 值)
        {
            conn.Open();
            //获取数据表
            //string sql = "select * from 表名 order by 字段1";
            //查询
            string sql = "select * from " + 表名 + " where " + 字段 + "=" + 值;

            OleDbDataAdapter da = new OleDbDataAdapter(sql, conn); //创建适配对象
            DataTable dt = new DataTable(); //新建表对象
            da.Fill(dt); //用适配对象填充表对象
            dataGridViewBp.DataSource = dt; //将表对象作为DataGridView的数据源

            conn.Close();
        }

        //private void 增()
        private void insert(string table, string[] name, string[] value)
        {
            conn.Open();
                //增
                string sql = "insert into " + table + "(" +
                    name[0] + "," + name[1] + "," + name[2] + "," + name[3] + "," + name[4] + ")values(" +
                    value[0] + "," + value[1] + "," + value[2] + "," + value[3] + "," + value[4] + ")";
            //删 
            //string sql = "delete from 表名 where 字段1="...; 
            //改 
            //string sql = "update student set 学号=" ...; 

            OleDbCommand comm = new OleDbCommand(sql, conn);
            comm.ExecuteNonQuery();

            conn.Close();
        }
        private void 删()
        {
            conn.Open();
            //增
            //string sql = "insert into 表名(字段1,字段2,字段3,字段4)values(...)";
            //删 
            string sql = "delete from 表名 where 字段1=";//...; 
            //改 
            //string sql = "update student set 学号=" ...; 

            OleDbCommand comm = new OleDbCommand(sql, conn);
            comm.ExecuteNonQuery();

            conn.Close();
        }
        private void 改()
        {
            conn.Open();
            //增
            //string sql = "insert into 表名(字段1,字段2,字段3,字段4)values(...)";
            //删 
            //string sql = "delete from 表名 where 字段1="...; 
            //改 
            string sql = "update student set 学号=";// ...; 

            OleDbCommand comm = new OleDbCommand(sql, conn);
            comm.ExecuteNonQuery();

            conn.Close();
        }

        private void saveData2()
        {
            dataGridViewBp.EndEdit();

            string sql = "select * from 表名";

            OleDbDataAdapter da = new OleDbDataAdapter(sql, conn);

            OleDbCommandBuilder bld = new OleDbCommandBuilder(da);
            da.UpdateCommand = bld.GetUpdateCommand();

            //把DataGridView赋值给dataTbale。(DataTable)的意思是类型转换，前提是后面紧跟着的东西要能转换成dataTable类型
            DataTable dt = (DataTable)dataGridViewBp.DataSource;

            da.Update(dt);
            dt.AcceptChanges();

            conn.Close();
        }


    }
}
