using ClosedXML.Excel;
using DataGridViewHandler;
using EasyDbHandler;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        List<List<string>> cc = new List<List<string>>();

        string p = "D:\\CodeProject\\線下教學\\student\\C# Basic Tutorial Course File\\章節8.進階控制項使用\\temp\\EXCEL\\member.xlsx";

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                EasyExcel x = new EasyDbHandler.EasyExcel();

                //cc = x.Output(p, 1, 1);
                cc = x.Output(p, 1, 1);
            }
            catch (Exception)
            {

                throw;
            }
        }

        //填一整列的概念
        private void button7_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();

            foreach (List<string> row in cc)
            {
                dataGridView1.Rows.Add(row.ToArray());  // 將 List<string> 轉為陣列再加入
            }
        }


        //填格子的概念
        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();

            //雙迴圈要先預設好列數
            dataGridView1.Rows.Add(cc.Count);

            for (int i = 0; i < cc.Count; i++)
            {
                for (int j = 0; j < cc[i].Count; j++)
                {
                    dataGridView1.Rows[i].Cells[j].Value = cc[i][j];  // 逐個填入每一個 cell
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }



        private void button4_Click(object sender, EventArgs e)
        {
            try
            {

                List<List<string>> data = new List<List<string>>();

                //List<string> t = new List<string>();

                //foreach (DataGridViewColumn item in dataGridView1.Columns)
                //{
                //    t.Add(item.HeaderText);
                //}

                //data.Add(t);

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    // 跳過空行
                    if (row.IsNewRow == false)
                    {
                        List<string> rowData = new List<string>();

                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            string x = "";

                            if (cell != null && cell.Value != null)
                            {
                                x = cell.Value.ToString();
                            }

                            rowData.Add(x);
                        }

                        data.Add(rowData);
                    }
                }

                //-------------------存入excel



                // 創建一個新的 Excel 工作簿
                using (var workbook = new XLWorkbook())
                {
                    // 添加工作表
                    var worksheet = workbook.Worksheets.Add(1);

                    int row = 1;

                    foreach (List<string> memberInfo in data)
                    {
                        int column = 1;

                        foreach (string cellvalue in memberInfo)
                        {
                            worksheet.Cell(row, column).Value = cellvalue;

                            column++;
                        }

                        row++;
                    }

                    // 儲存工作簿
                    workbook.SaveAs(p);
                }

                MessageBox.Show("匯出完畢");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                //客戶跟你說明天要 就用這一個 然後要多收錢

                DataGridViewExcel c = new DataGridViewHandler.DataGridViewExcel();

                c.Export(dataGridView1, p, true);
                MessageBox.Show("匯出完畢");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}

