using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            dateTimePicker1.MaxDate = DateTime.Now;

            comboBox1.Items.Add("男");
            comboBox1.Items.Add("女");

            comboBox1.SelectedIndex = 0;
            checkBox1.Checked = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string name = textBox1.Text;
            string phone = textBox2.Text;
            string address = textBox3.Text;
            string gender = comboBox1.Text;
            string bir = dateTimePicker1.Value.ToString();
            string mar = checkBox1.Checked.ToString();

            List<string> list = new List<string>() {
                name, 
                phone, 
                address, 
                gender, 
                bir,
                mar
            };

            //2.加序號
            int id = dataGridView1.Rows.Count + 1;

            list.Insert(0,id.ToString());

            dataGridView1.Rows.Add(list.ToArray());


            //3.清掉格子

            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            comboBox1.SelectedIndex = 0;
            checkBox1.Checked = false;
            dateTimePicker1.Value= DateTime.Now.Date;
        }
    }
}
