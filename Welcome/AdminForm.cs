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
        }

        public class Person
        {
            public int Id { get; set; }
            public string Name { get; set; }
        }

        private void DatabaseCon()
        {
            string constr = "server=162.243.150.192;Uid=admin;password=admin123456;Database=NewPartner";
            try
            {
                MySqlConnection mycon = new MySqlConnection(constr);
                mycon.Open();
                MessageBox.Show("数据库连接成功！");
                MySqlDataAdapter disAdapter = new MySqlDataAdapter("select * from newperson", mycon);
                DataTable NewPerson = new DataTable();
                disAdapter.Fill(NewPerson);
                dataGridView1.DataSource = NewPerson;
                mycon.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("数据库连接失败！");
                Application.Exit();
            }
        }

        private void AdminForm_Load(object sender, EventArgs e)
        {
            Thread ConDB = new Thread(() =>
            {
                DatabaseCon();
            });
            ConDB.Start();
        }
    }
}
