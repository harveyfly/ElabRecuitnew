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
            toolStripStatusLabel2.Alignment = ToolStripItemAlignment.Right;
            timer1.Interval = 1000;
            timer1.Enabled = true;
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
                foreach(Control ctr in panel1.Controls)
                {
                    if(ctr is Button)
                    {
                        ctr.Enabled = true;
                    }
                }
                foreach (Control ctr in panel2.Controls)
                {
                    if (ctr is Button)
                    {
                        ctr.Enabled = true;
                    }
                }
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
            try
            {
                Control.CheckForIllegalCrossThreadCalls = false;
                Thread ConDB = new Thread(() =>
                {
                    myConn = DatabaseCon();
                    flushList("select * from newperson");
                });
                ConDB.IsBackground = true;
                ConDB.Start();
                foreach (Control ctr in panel1.Controls)
                {
                    if (ctr is Button)
                    {
                        ctr.Enabled = false;
                    }
                }
                foreach (Control ctr in panel2.Controls)
                {
                    if (ctr is Button)
                    {
                        ctr.Enabled = false;
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
                Application.Exit();
            }
            
        }

        private void flushList(string mySqlCmd)//刷新列表
        {
            myConn.Open();
            DataTable NewPerson = new DataTable();
            MySqlDataAdapter disAdapter = new MySqlDataAdapter(mySqlCmd, myConn);
            disAdapter.Fill(NewPerson);
            dataGridView1.DataSource = NewPerson;
            myConn.Close();
        }

        private void flushData()//更新数据
        {
            flushList("select * from newperson");
            List<string> ProList = new List<string>();
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                string ProName = dataGridView1.Rows[i].Cells[5].Value.ToString();
                if (ProList.Contains(ProName))
                {
                    continue;
                }
                else
                {
                    ProList.Add(ProName);
                }
            }
            cbxPro.DataSource = ProList;
            RowDisplay();
        }
        private void button8_Click(object sender, EventArgs e)//刷新列表
        {
            flushData();
        }

        private void button3_Click(object sender, EventArgs e)//发布招新
        {
            InputBox InBox = new InputBox("请输入管理员密码：", "管理员验证");
            InBox.pasChar = '●';
            DialogResult InBoxResult = InBox.ShowDialog();
            if(InBoxResult == DialogResult.OK && InBox.Value == "admin123456")
            {
                string createStatement =
                    @"CREATE TABLE `newperson`(
                       `Id` VARCHAR(9) NOT NULL,
                       `Name` VARCHAR(20) NOT NULL,
                       `Sex` SET('男','女'),
                       `Tel` VARCHAR(11) NOT NULL,
                       `ElabGroup` VARCHAR(10) NOT NULL,
                       `Professor` VARCHAR(30) NOT NULL,
                       PRIMARY KEY ( `Id` )
                    )ENGINE=InnoDB DEFAULT CHARSET=utf8;";
                string dropStatement = "DROP TABLE newperson";

                try
                {
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
                }
                catch
                {
                    toolStripStatusLabel1.Text = "数据表建立失败";
                }
            }
            else if(InBox.Value != "admin123456" && InBoxResult == DialogResult.OK)
            {
                MessageBox.Show("密码错误");
            }
        }

        private void RowDisplay()//更新显示
        {
            tbxId.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            tbxName.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            cbxSex.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            cbxPro.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            cbxGro.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            RowDisplay();
        }

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            RowDisplay();
        }

        private void button5_Click(object sender, EventArgs e)//查询
        {
            if(tbxId.Text != "")
            {
                string sqlCmd = @"select * from newperson where Id = '" + tbxId.Text + "'";
                flushList(sqlCmd);
            }
            else if(tbxName.Text != "")
            {
                string sqlCmd = @"select * from newperson where Name = '" + tbxName.Text + "'";
                flushList(sqlCmd);
            }
            else if (cbxSex.Text != "")
            {
                string sqlCmd = @"select * from newperson where Sex = '" + cbxSex.Text + "'";
                flushList(sqlCmd);
            }
            else if (cbxGro.Text != "")
            {
                string sqlCmd = @"select * from newperson where ElabGroup = '" + cbxGro.Text + "'";
                flushList(sqlCmd);
            }
            else if (cbxPro.Text != "")
            {
                string sqlCmd = @"select * from newperson where Professor = '" + cbxPro.Text + "'";
                flushList(sqlCmd);
            }
            else
            {
                MessageBox.Show("查询数据不能为空");
            }
        }

        private void panel2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            tbxId.Text = "";
            tbxName.Text = "";
            cbxSex.Text = "";
            cbxPro.Text = "";
            cbxGro.Text = "";
        }

        bool PasInput = false;
        private void deletePerson()
        {
            if(tbxId.Text != "" && tbxName.Text != "")
            {
                DialogResult re = MessageBox.Show("确定删除:" + tbxName.Text + "?", "提示", MessageBoxButtons.OKCancel);
                if(re == DialogResult.OK)
                {
                    myConn.Open();
                    string sqlCmd = "delete from newperson where Id = '" + tbxId.Text + "'";
                    using (MySqlCommand myCmd = new MySqlCommand(sqlCmd, myConn))
                    {
                        myCmd.ExecuteNonQuery();
                        MessageBox.Show("删除成功");
                    }
                    myConn.Close();
                }
            }
            else
            {
                MessageBox.Show("请选择想要删除的学生信息");
            }
            
        }
        private void button6_Click(object sender, EventArgs e)//删除
        {
            if(!PasInput)
            {
                InputBox InBox = new InputBox("请输入管理员密码：", "管理员验证");
                InBox.pasChar = '●';
                DialogResult InBoxResult = InBox.ShowDialog();
                if (InBoxResult == DialogResult.OK && InBox.Value == "admin123456")
                {
                    deletePerson();
                    PasInput = true;
                }
                else if (InBox.Value != "admin123456" && InBoxResult == DialogResult.OK)
                {
                    MessageBox.Show("密码错误");
                }
            }
            else
            {
                deletePerson();
            }
            flushData();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            toolStripStatusLabel2.Text = DateTime.Now.ToString();
        }

        private void button7_Click(object sender, EventArgs e)//导出联系方式
        {
            flushData();
            string contactInfo = null;
            for(int i = 0; i < dataGridView1.RowCount -1;i ++)
            {
                contactInfo += dataGridView1.Rows[i].Cells["Tel"].Value.ToString() + ",";
            }
            Clipboard.SetText(contactInfo);
            MessageBox.Show("联系方式已经导入到剪切板");
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
