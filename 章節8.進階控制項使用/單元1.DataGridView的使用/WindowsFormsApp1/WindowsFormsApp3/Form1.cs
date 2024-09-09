using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace WindowsFormsApp3
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
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;

            checkBox1.Checked = false;


           

        }
        string p = "D:\\CodeProject\\線下教學\\student\\C# Basic Tutorial Course File\\章節8.進階控制項使用\\temp\\EXCEL\\member3.xlsx";
        private void button1_Click(object sender, EventArgs e)
        {

            string name = textBox1.Text.Trim();

            if (name == "")
            {
                return;
            }

            string phone = textBox2.Text.Trim();

            if (phone == "")
            {
                return;
            }

            string address = textBox3.Text.Trim();

            if (address == "")
            {
                return;
            }


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

            list.Insert(0, id.ToString());

            dataGridView1.Rows.Add(list.ToArray());


            //3.清掉格子

            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            comboBox1.SelectedIndex = 0;
            checkBox1.Checked = false;
            dateTimePicker1.Value = DateTime.Now.Date;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //MessageBox.Show("213");

            //有可能案到其他格子也會刪除
            int index = e.RowIndex;

            //MessageBox.Show(e.RowIndex.ToString());

            //dataGridView1.Rows.RemoveAt(index);


            //2.只限定最後一個格子 要按到內容上


            //MessageBox.Show(e.RowIndex.ToString()+":"+ e.ColumnIndex.ToString());



            //if (e.ColumnIndex == 7)
            //{
            //    dataGridView1.Rows.RemoveAt(index);
            //}

            //3.加防呆
            if (e.ColumnIndex == 7)
            {
                DialogResult x = MessageBox.Show("是否刪除", "delete~~", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (x == DialogResult.Yes)
                {
                    dataGridView1.Rows.RemoveAt(index);
                }

            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                DataGridViewHandler.DataGridViewExcel c = new DataGridViewHandler.DataGridViewExcel();

                c.Export(dataGridView1, p, true);
                MessageBox.Show("匯出完畢");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                //dataGridView1.Rows.Clear();

                EasyDbHandler.EasyExcel x = new EasyDbHandler.EasyExcel();

                //不抓excel的標頭 不然會出錯
                List<List<string>> cc = x.Output(p, 1, 2);


                //--顯示在gd

                
                foreach (List<string> cc2 in cc)
                {
                    dataGridView1.Rows.Add(cc2.ToArray());
                }

                button3.Enabled = false;
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
