using System;
using System.Windows.Forms;

namespace Welcome
{
    public partial class InputBox : Form
    {
        public string Value { get; set; }
        public char pasChar { get; set; }

        public InputBox(string label, string title)
        {
            InitializeComponent();
            this.AcceptButton = this.button1;
            this.CancelButton = this.button2;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            label1.Text = label;
            this.Text = title;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            Value = textBox1.Text;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void InputBox_Load(object sender, EventArgs e)
        {
            textBox1.Focus();
            textBox1.Text = Value;
            textBox1.PasswordChar = pasChar;
        }
    }
}
