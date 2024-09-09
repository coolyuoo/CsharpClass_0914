using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string url = $"https://favqs.com/api/qotd";

            RestClient client = new RestClient(url);

            // 創建請求
            RestRequest request = new RestRequest();

            // 發送請求並取得回應
            RestResponse response = client.Execute(request);

            if (response.IsSuccessful)
            {
                this.Text = response.StatusCode.ToString();
                textBox1.Text = response.Content;
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string x = textBox1.Text;

            JToken c = JsonConvert.DeserializeObject<JToken>(x);

            JToken dd = c["quote"]["body"];

            textBox2.Text = dd.ToString();
        }
    }
}
