using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Welcome
{
    public partial class AdminForm : Form
    {
        public AdminForm()
        {
            InitializeComponent();
            panel3.Dock = DockStyle.Fill;
        }

        public MySqlConnection DatabaseCon()
        {
            string constr = "server=162.243.150.192;Uid=admin;password=admin123456;Database=NewPartner;Charset=utf8";
            try
            {
                MySqlConnection mycon = new MySqlConnection(constr);
                mycon.Open();
                mycon.Close();
                toolStripStatusLabel1.Text = "数据库连接成功";
                return mycon;
            }
            catch (Exception)
            {
                toolStripStatusLabel1.Text = "数据库连接失败";
                return null;
            }
        }

        MySqlConnection myConn = new MySqlConnection();
        private void AdminForm_Load(object sender, EventArgs e)
        {
            Thread ConDB = new Thread(() =>
            {
                myConn = DatabaseCon();
            });
            ConDB.Start();
        }

        private void button8_Click(object sender, EventArgs e)//刷新列表
        {
            myConn.Open();
            DataTable NewPerson = new DataTable();
            MySqlDataAdapter disAdapter = new MySqlDataAdapter("select * from newperson", myConn);
            disAdapter.Fill(NewPerson);
            dataGridView1.DataSource = NewPerson;
            myConn.Close();
        }

        private void button3_Click(object sender, EventArgs e)//发布招新
        {
            string createStatement =
                @"CREATE TABLE `newperson`(
                   `Id` VARCHAR(9) NOT NULL,
                   `Name` VARCHAR(20) NOT NULL,
                   `Sex` SET('男','女'),
                   `Tel` VARCHAR(11) NOT NULL,
                   `Group` VARCHAR(10) NOT NULL,
                   `Professor` VARCHAR(30) NOT NULL,
                   PRIMARY KEY ( `Id` )
                )ENGINE=InnoDB DEFAULT CHARSET=utf8;";
            string dropStatement = "DROP TABLE newperson";

            //try
            //{
                myConn.Open();
                // 建立新表  
                using (MySqlCommand cmdDrop = new MySqlCommand(dropStatement, myConn))
                {
                    cmdDrop.ExecuteNonQuery();
                    toolStripStatusLabel1.Text = "删除成功";
                }
                using (MySqlCommand cmdCreate = new MySqlCommand(createStatement, myConn))
                {
                    cmdCreate.ExecuteNonQuery();
                    toolStripStatusLabel1.Text = "数据表建立成功";
                }
                myConn.Close();
            //}
            //catch
            //{
            //    toolStripStatusLabel1.Text = "数据表建立失败";
            //}
           
        }
    }

    public class Person
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Tel { get; set; }
        public string Sex { get; set; }
        public string Group { get; set; }
        public string Professor { get; set; }
    }
}
