using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Windows.Forms;
using MsWord = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Drawing;

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
            panel7.Visible = false;
        }

        private delegate void flushData_Delegate(string SqlCmd, MySqlConnection myConn);//跨线程访问代理
        private delegate void btn_Delegate();
        MySqlConnection yesConn = new MySqlConnection();//mysql连接
        MySqlConnection elabConn = new MySqlConnection();
        private bool PasInput = false;
        private string passwd = "elabadmin"; //设置密码

        private void buttonStatusChange()//button状态转变
        {
            if (panel1.InvokeRequired || panel2.InvokeRequired)
            {
                btn_Delegate df = new btn_Delegate(buttonStatusChange);
                this.Invoke(df);
            }
            else
            {
                foreach (Control ctr in panel1.Controls)
                {
                    if (ctr is Button)
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
            }
            
        }

        private MySqlConnection DatabaseCon() //连接数据库
        {
            string constr = "server=yespace.xyz;Uid=admin;password=admin123456;Database=NewPartner;Charset=utf8";
            try
            {
                MySqlConnection mycon = new MySqlConnection(constr);
                mycon.Open();
                mycon.Close();
                toolStripStatusLabel1.Text = "数据库连接成功";
                Thread btsChange = new Thread(() =>
                {
                    buttonStatusChange();
                });
                btsChange.IsBackground = true;
                btsChange.Start();
                return mycon;
            }
            catch (Exception)
            {
                toolStripStatusLabel1.Text = "数据库连接失败";
                return null;
            }
        }

        private void AdminForm_Load(object sender, EventArgs e)
        {
            try
            {
                Thread ConDB = new Thread(() =>
                {
                    yesConn = DatabaseCon();
                    flushData("SELECT * FROM newperson;",yesConn);
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
                MessageBox.Show(ex.ToString(), "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Application.Exit();
            }
            
        }

        private void flushList(string mySqlCmd, MySqlConnection myConn)//刷新列表
        {
            if(dataGridView1.InvokeRequired)
            {
                flushData_Delegate df = new flushData_Delegate(flushData);
                this.Invoke(df, mySqlCmd, myConn);
            }
            else
            {
                try
                {
                    myConn.Open();
                    DataTable NewPerson = new DataTable();
                    MySqlDataAdapter disAdapter = new MySqlDataAdapter(mySqlCmd, myConn);
                    disAdapter.Fill(NewPerson);
                    dataGridView1.DataSource = NewPerson;
                    myConn.Close();
                    toolStripStatusLabel1.Text = "列表刷新成功";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("列表刷新失败" + ex.ToString(), "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void flushData(string mySqlCmd, MySqlConnection myConn)//更新下拉菜单数据
        {
            flushList(mySqlCmd, myConn);
            List<string> ProList = new List<string>();
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                string ProName = dataGridView1.Rows[i].Cells["专业"].Value.ToString();
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
            //RowDisplay();
        }

        private void button8_Click(object sender, EventArgs e)//刷新列表
        {
            if(panel6.Visible && !panel7.Visible)
            {
                string SqlCmd = "SELECT * FROM `newperson`;";
                flushData(SqlCmd, yesConn);
            }
            else if(panel7.Visible && !panel6.Visible)
            {
                string SqlCmd = "SELECT * FROM `TimeStatistics`;";
                flushData(SqlCmd, elabConn);
            }
            else
            {
                toolStripStatusLabel1.Text = "列表刷新失败";
            }
        }

        private void button3_Click(object sender, EventArgs e)//发布招新
        {
            InputBox InBox = new InputBox("请输入管理员密码：", "管理员验证");
            InBox.pasChar = '●';
            DialogResult InBoxResult = InBox.ShowDialog();
            if(InBoxResult == DialogResult.OK && InBox.Value == passwd)
            {
                string createStatement =
                    @"CREATE TABLE `newperson`(
                    `学号` VARCHAR(9) NOT NULL,
                    `姓名` VARCHAR(20) NOT NULL,
                    `性别` SET('男','女'),
                    `Tel` VARCHAR(11) NOT NULL,
                    `组别` VARCHAR(10) NOT NULL,
                    `专业` VARCHAR(30) NOT NULL,
                    `籍贯` VARCHAR(10) NOT NULL,
                    `班级` VARCHAR(10) NOT NULL,
                    `职务` VARCHAR(10) NOT NULL,
                    `社团` VARCHAR(15) NOT NULL,
                    `爱好` VARCHAR(20) NOT NULL,
                    `讲座时间` VARCHAR(10) NOT NULL,
                    `Email` VARCHAR(20) NOT NULL,
                    `经历` VARCHAR(200),
                    `时间安排` VARCHAR(100),
                    `理解` VARCHAR(200),
                    `自我评价` VARCHAR(300),
                     PRIMARY KEY ( `学号` )
                    )ENGINE=InnoDB DEFAULT CHARSET=utf8;";
                string dropStatement = "DROP TABLE newperson;";

                try
                {
                    yesConn.Open();
                    // 建立新表
                    using (MySqlCommand cmdDrop = new MySqlCommand(dropStatement, yesConn))
                    {
                        cmdDrop.ExecuteNonQuery();
                        toolStripStatusLabel1.Text = "数据表建立成功";
                    }
                    yesConn.Close();
                }
                catch
                {
                    toolStripStatusLabel1.Text = "数据表不存在，正在新建...";
                    using (MySqlCommand cmdDrop = new MySqlCommand(createStatement, yesConn))
                    {
                        cmdDrop.ExecuteNonQuery();
                        toolStripStatusLabel1.Text = "数据表建立成功";
                    }
                }
            }
            else if(InBox.Value != passwd && InBoxResult == DialogResult.OK)
            {
                MessageBox.Show("密码错误", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void RowDisplay() //更新显示
        {
            try
            {
                if (dataGridView1.Rows[0].Cells[0].Value != null)
                {
                    tbxId.Text = dataGridView1.CurrentRow.Cells["学号"].Value.ToString();
                    tbxName.Text = dataGridView1.CurrentRow.Cells["姓名"].Value.ToString();
                    cbxSex.Text = dataGridView1.CurrentRow.Cells["性别"].Value.ToString();
                    cbxPro.Text = dataGridView1.CurrentRow.Cells["专业"].Value.ToString();
                    cbxGro.Text = dataGridView1.CurrentRow.Cells["组别"].Value.ToString();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("显示更新失败" + ex.ToString(), "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
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
                string sqlCmd = @"SELECT * FROM `newperson` WHERE `学号` = '" + tbxId.Text + "'";
                flushList(sqlCmd, yesConn);
            }
            else if(tbxName.Text != "")
            {
                string sqlCmd = @"SELECT * FROM `newperson` WHERE `姓名` = '" + tbxName.Text + "'";
                flushList(sqlCmd ,yesConn);
            }
            else if (cbxSex.Text != "")
            {
                string sqlCmd = @"SELECT * FROM `newperson` WHERE `性别` = '" + cbxSex.Text + "'";
                flushList(sqlCmd,yesConn);
            }
            else if (cbxGro.Text != "")
            {
                string sqlCmd = @"SELECT * FROM `newperson` WHERE `组别` = '" + cbxGro.Text + "'";
                flushList(sqlCmd,yesConn);
            }
            else if (cbxPro.Text != "")
            {
                string sqlCmd = @"SELECT * FROM `newperson` WHERE `专业` = '" + cbxPro.Text + "'";
                flushList(sqlCmd, yesConn);
            }
            else
            {
                MessageBox.Show("查询数据不能为空", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

        private void deletePerson()
        {
            if(tbxId.Text != "" && tbxName.Text != "")
            {
                DialogResult re = MessageBox.Show("确定删除:" + tbxName.Text + "?", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if(re == DialogResult.OK)
                {
                    yesConn.Open();
                    string sqlCmd = "DELETE FROM `newperson` WHERE `学号` = '" + tbxId.Text + "'";
                    using (MySqlCommand myCmd = new MySqlCommand(sqlCmd, yesConn))
                    {
                        myCmd.ExecuteNonQuery();
                        MessageBox.Show("删除成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    yesConn.Close();
                }
            }
            else
            {
                MessageBox.Show("请选择想要删除的学生信息", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            
        }

        private void button6_Click(object sender, EventArgs e)//删除
        {
            if(!PasInput)
            {
                InputBox InBox = new InputBox("请输入管理员密码：", "管理员验证");
                InBox.pasChar = '●';
                DialogResult InBoxResult = InBox.ShowDialog();
                if (InBoxResult == DialogResult.OK && InBox.Value == passwd)
                {
                    deletePerson();
                    PasInput = true;
                }
                else if (InBox.Value != passwd && InBoxResult == DialogResult.OK)
                {
                    MessageBox.Show("密码错误", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                deletePerson();
            }
            flushData("SELECT * FROM newperson;", yesConn);
        }

        private void timer1_Tick(object sender, EventArgs e) //时钟
        {
            toolStripStatusLabel2.Text = DateTime.Now.ToString();
        }

        private void button7_Click(object sender, EventArgs e)//导出联系方式
        {
            try
            {
                flushData("SELECT * FROM newperson;", yesConn);
                string contactInfo = null;
                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    contactInfo += dataGridView1.Rows[i].Cells["Tel"].Value.ToString() + ",";
                }
                Clipboard.SetText(contactInfo);
                MessageBox.Show("联系方式已经导入到剪切板", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                MessageBox.Show("数据为空，导出失败!","提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button4_Click(object sender, EventArgs e)//导出报名表
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                try
                {
                    FolderBrowserDialog fbd = new FolderBrowserDialog();
                    fbd.Description = "请选择需要导出到的文件夹";
                    if (fbd.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = fbd.SelectedPath.ToString();
                        progressBar.Refresh();
                        progressBar.Value = 1;
                        progressBar.Visible = true;
                        progressBar.Minimum = 1;
                        progressBar.Maximum = dataGridView1.SelectedRows.Count;
                        progressBar.Step = 1;
                        for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                        {
                            progressBar.PerformStep();
                            if (dataGridView1.Rows[i].Selected)
                            {
                                object oMissing = System.Reflection.Missing.Value;
                                MsWord._Application oWord = new MsWord.Application();
                                oWord.Visible = false;
                                object oTemplate = Directory.GetCurrentDirectory() + "\\Template.dotx";
                                MsWord._Document oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing, ref oMissing, ref oMissing);
                                object[] oBookMark = new object[17];
                                oBookMark[0] = "Name";
                                oBookMark[1] = "TEL";
                                oBookMark[2] = "Group";
                                oBookMark[3] = "Sex";
                                oBookMark[4] = "Native";
                                oBookMark[5] = "Id";
                                oBookMark[6] = "Email";
                                oBookMark[7] = "Class";
                                oBookMark[8] = "Duty";
                                oBookMark[9] = "Professor";
                                oBookMark[10] = "League";
                                oBookMark[11] = "Hobbies";
                                oBookMark[12] = "FreeTime";
                                oBookMark[13] = "Experience";
                                oBookMark[14] = "WeekTime";
                                oBookMark[15] = "Expect";
                                oBookMark[16] = "Evaluation";
                                oDoc.Bookmarks.get_Item(ref oBookMark[0]).Range.Text = dataGridView1.Rows[i].Cells["姓名"].Value.ToString();
                                oDoc.Bookmarks.get_Item(ref oBookMark[1]).Range.Text = dataGridView1.Rows[i].Cells["TEL"].Value.ToString();
                                oDoc.Bookmarks.get_Item(ref oBookMark[2]).Range.Text = dataGridView1.Rows[i].Cells["组别"].Value.ToString();
                                oDoc.Bookmarks.get_Item(ref oBookMark[3]).Range.Text = dataGridView1.Rows[i].Cells["性别"].Value.ToString();
                                oDoc.Bookmarks.get_Item(ref oBookMark[4]).Range.Text = dataGridView1.Rows[i].Cells["籍贯"].Value.ToString();
                                oDoc.Bookmarks.get_Item(ref oBookMark[5]).Range.Text = dataGridView1.Rows[i].Cells["学号"].Value.ToString();
                                oDoc.Bookmarks.get_Item(ref oBookMark[6]).Range.Text = dataGridView1.Rows[i].Cells["Email"].Value.ToString();
                                oDoc.Bookmarks.get_Item(ref oBookMark[7]).Range.Text = dataGridView1.Rows[i].Cells["班级"].Value.ToString();
                                oDoc.Bookmarks.get_Item(ref oBookMark[8]).Range.Text = dataGridView1.Rows[i].Cells["职务"].Value.ToString();
                                oDoc.Bookmarks.get_Item(ref oBookMark[9]).Range.Text = dataGridView1.Rows[i].Cells["专业"].Value.ToString();
                                oDoc.Bookmarks.get_Item(ref oBookMark[10]).Range.Text = dataGridView1.Rows[i].Cells["社团"].Value.ToString();
                                oDoc.Bookmarks.get_Item(ref oBookMark[11]).Range.Text = dataGridView1.Rows[i].Cells["爱好"].Value.ToString();
                                oDoc.Bookmarks.get_Item(ref oBookMark[12]).Range.Text = dataGridView1.Rows[i].Cells["讲座时间"].Value.ToString();
                                oDoc.Bookmarks.get_Item(ref oBookMark[13]).Range.Text = dataGridView1.Rows[i].Cells["经历"].Value.ToString();
                                oDoc.Bookmarks.get_Item(ref oBookMark[14]).Range.Text = dataGridView1.Rows[i].Cells["时间安排"].Value.ToString();
                                oDoc.Bookmarks.get_Item(ref oBookMark[15]).Range.Text = dataGridView1.Rows[i].Cells["理解"].Value.ToString();
                                oDoc.Bookmarks.get_Item(ref oBookMark[16]).Range.Text = dataGridView1.Rows[i].Cells["自我评价"].Value.ToString();
                                object fileName = filePath + "\\" + dataGridView1.Rows[i].Cells["组别"].Value.ToString() + "_" + dataGridView1.Rows[i].Cells["姓名"].Value.ToString() + "_报名表.docx";
                                oDoc.SaveAs(ref fileName, ref oMissing, ref oMissing, ref oMissing,
                                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                    ref oMissing, ref oMissing);
                                oDoc.Close(ref oMissing, ref oMissing, ref oMissing);
                                oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
                            }
                        }
                        progressBar.Visible = false;
                        MessageBox.Show("报名表已经保存到\n" + filePath, "提示 " ,MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show("导出失败!\n\n" + ex.ToString(), "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("请至少选择一行信息", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void button1_Click(object sender, EventArgs e)//保存报名单
        {
            SaveFileDialog sfdlg = new SaveFileDialog();
            sfdlg.FileName = DateTime.Now.ToString("yyyy") + "年电气创新实践基地招新报名表";
            sfdlg.Filter = "Excel文档|*.xlsx ";
            sfdlg.RestoreDirectory = true;
            if(sfdlg.ShowDialog() == DialogResult.OK)
            {
                if(sfdlg.FileName != null)
                {
                    if (dataGridView1.RowCount <= 0)
                    {
                        MessageBox.Show("没有数据可供保存 ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    Excel.Application objExcel = null;
                    Excel.Workbook objWorkbook = null;
                    Excel.Worksheet objsheet = null;
                    try
                    {
                        objExcel = new Excel.Application();
                        objWorkbook = objExcel.Workbooks.Add(Missing.Value);
                        objsheet = (Excel.Worksheet)objWorkbook.ActiveSheet;

                        int excelColumns = 1;
                        for(int i = 0;i < 6;i ++)
                        {
                            if(dataGridView1.Columns[i].Visible)
                            {
                                objExcel.Cells[1, excelColumns] = dataGridView1.Columns[i].HeaderText.Trim();
                                excelColumns++;
                            }
                        }
                        for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                        {
                            excelColumns = 1;
                            for (int j = 0; j < 6; j++)
                            {
                                if (dataGridView1.Columns[j].Visible == true)
                                {
                                    try
                                    {
                                        objExcel.Cells[i + 2, excelColumns] = dataGridView1.Rows[i].Cells[j].Value.ToString().Trim();
                                        excelColumns++;
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show(ex.Message, "警告 ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }
                                }
                            }
                        }
                        objWorkbook.SaveAs(sfdlg.FileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                            Missing.Value, Excel.XlSaveAsAccessMode.xlShared, Missing.Value, Missing.Value, Missing.Value,
                                            Missing.Value, Missing.Value);
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show(ex.Message, "警告 ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    finally
                    {
                        //关闭Excel应用      
                        if (objWorkbook != null) objWorkbook.Close(Missing.Value, Missing.Value, Missing.Value);
                        if (objExcel.Workbooks != null) objExcel.Workbooks.Close();
                        if (objExcel != null) objExcel.Quit();

                        objsheet = null;
                        objWorkbook = null;
                        objExcel = null;
                    }
                    MessageBox.Show(sfdlg.FileName + "\n\n导出完毕! ", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("文件名不能为空！", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private MySqlConnection ElabDBConn() //连接签到数据库
        {
            string constr = "server=192.168.31.165;Uid=elabadmin;password=elab2018;Database=QianDao;Charset=utf8";
            try
            {
                MySqlConnection mycon = new MySqlConnection(constr);
                mycon.Open();
                mycon.Close();
                toolStripStatusLabel1.Text = "签到数据库连接成功";
                return mycon;
            }
            catch (Exception)
            {
                MessageBox.Show("签到数据库连接失败，请确保在科中内网内使用", "提示 ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

        }

        private void btnQianDao_Click(object sender, EventArgs e) //签到管理
        {
            toolStripStatusLabel1.Text = "签到数据库连接中...";
            panel6.Visible = false;
            panel7.Visible = true;
            elabConn = ElabDBConn();
            string SqlCmd = "SELECT * FROM TimeStatistics;";
            flushData(SqlCmd, elabConn);
        }

        private void button2_Click(object sender, EventArgs e) //考核管理
        {
            toolStripStatusLabel1.Text = "远程数据库连接中...";
            flushData("SELECT * FROM newperson;", yesConn);
            panel6.Visible = true;
            panel7.Visible = false;
        }

        private void button12_Click(object sender, EventArgs e) //查询签到时间
        {
            if (tbxId.Text != "")
            {
                string sqlCmd = @"SELECT `学号`,`姓名`,`性别`,`组别`,`专业`,`周次`,`时长` FROM `TimeWeek` WHERE `学号` = '" + tbxId.Text + "' AND `周次`='" + cbxWeek.Text + "';";
                flushList(sqlCmd, elabConn);
            }
            else if (tbxName.Text != "")
            {
                string sqlCmd = @"SELECT `学号`,`姓名`,`性别`,`组别`,`专业`,`周次`,`时长` FROM `TimeWeek` WHERE `姓名` = '" + tbxName.Text + "' AND `周次`='" + cbxWeek.Text + "';";
                flushList(sqlCmd, elabConn);
            }
            else if (cbxSex.Text != "")
            {
                string sqlCmd = @"SELECT `学号`,`姓名`,`性别`,`组别`,`专业`,`周次`,`时长` FROM `TimeWeek` WHERE `性别` = '" + cbxSex.Text + "' AND `周次`='" + cbxWeek.Text + "';";
                flushList(sqlCmd, elabConn);
            }
            else if (cbxGro.Text != "")
            {
                string sqlCmd = @"SELECT `学号`,`姓名`,`性别`,`组别`,`专业`,`周次`,`时长` FROM `TimeWeek` WHERE `组别` = '" + cbxGro.Text + "' AND `周次`='" + cbxWeek.Text + "';";
                flushList(sqlCmd, elabConn);
            }
            else if (cbxPro.Text != "")
            {
                string sqlCmd = @"SELECT `学号`,`姓名`,`性别`,`组别`,`专业`,`周次`,`时长` FROM `TimeWeek` WHERE `专业` = '" + cbxPro.Text + "' AND `周次`='" + cbxWeek.Text + "';";
                flushList(sqlCmd, elabConn);
            }
            else
            {
                string sqlCmd = @"SELECT `学号`,`姓名`,`性别`,`组别`,`专业`,`周次`,`时长` FROM `TimeWeek` WHERE `周次`='" + cbxWeek.Text + "';";
                flushList(sqlCmd, elabConn);
            }
        }

        private void button11_Click(object sender, EventArgs e) //解除MAC绑定
        {
            if(tbxId.Text != "")
            {
                string SqlCmd = "UPDATE TimeStatistics SET Mac = '',IP = '',OnlineStatus=0 WHERE `学号`='" + tbxId.Text + "';";
                MySqlCommand myCom = new MySqlCommand(SqlCmd, elabConn);
                elabConn.Open();
                try
                {
                    myCom.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("操作失败，请检查学号后重试！","提示",MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                MessageBox.Show(tbxId.Text + " 绑定解除成功，请重新签到", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                elabConn.Close();
                flushList("SELECT * FROM `TimeStatistics` WHERE `学号` = '" + tbxId.Text + "';", elabConn);
            }
            else
            {
                MessageBox.Show("输入学号不能为空", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button9_Click(object sender, EventArgs e) //导入成员列表
        {
            InputBox InBox = new InputBox("请输入管理员密码：", "管理员验证");
            InBox.pasChar = '●';
            DialogResult InBoxResult = InBox.ShowDialog();
            if (InBoxResult == DialogResult.OK && InBox.Value == passwd)
            {
                flushList("SELECT `学号`,`姓名`,`性别`,`组别`,`专业` FROM newperson", yesConn);
                if(dataGridView1.RowCount > 0)
                {
                    elabConn.Open();
                    for(int i = 0;i < dataGridView1.RowCount - 1; i ++)
                    {
                        string StuID = dataGridView1.Rows[i].Cells["学号"].Value.ToString();
                        string Name = dataGridView1.Rows[i].Cells["姓名"].Value.ToString();
                        string Sex = dataGridView1.Rows[i].Cells["性别"].Value.ToString();
                        string Group = dataGridView1.Rows[i].Cells["组别"].Value.ToString();
                        string Professor = dataGridView1.Rows[i].Cells["专业"].Value.ToString();
                        string SqlInsert = @"REPLACE INTO TimeStatistics(`学号`,`姓名`,`性别`,`组别`,`专业`) VALUES('"+StuID+"','"+Name+"','"+Sex+"','"+Group+"','"+Professor+"' );";
                        MySqlCommand myCom = new MySqlCommand(SqlInsert, elabConn);
                        try
                        {
                            myCom.ExecuteNonQuery();
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show("数据导入失败！\n\n" + ex.ToString(), "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    elabConn.Close();
                    MessageBox.Show("导入成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    flushList("SELECT * FROM TimeStatistics;", elabConn);
                }
            }
            else if (InBox.Value != passwd && InBoxResult == DialogResult.OK)
            {
                MessageBox.Show("密码错误", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }
    }
}
