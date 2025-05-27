using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace WindowsFormsApp1
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();

            _connectStrAPS = ConfigurationManager.AppSettings["Localhost_APS"];

        }

        string _connectStrAPS = "Data Source = 127.0.0.1; Initial Catalog = 01_APS_DF; User ID = MES2014; Password = PMCMES;";

        private void Login_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" && textBox2.Text != "")
            {
                try
                {
                    //解析一下json
                    Accunt descJsonStu = JsonConvert.DeserializeObject<Accunt>(textBox2.Text);
                    textBox1.Text = descJsonStu.account;
                    textBox2.Text = descJsonStu.password;
                }
                catch
                {

                }
            }
            if (textBox1.Text.Length != 0 && textBox2.Text.Length != 0)
            {
                bool IsPass = passwordcheck(textBox1.Text, textBox2.Text);
                if (IsPass)
                {
                    Form1 form1 = new Form1(getusername(textBox1.Text));
                    this.Hide();
                    form1.ShowDialog();

                    this.Close();
                }
                else
                {
                    MessageBox.Show("帳號、密碼有錯誤，請重新登入。");
                    textBox1.Text = "";
                    textBox2.Text = "";
                }
            }
            else
            {
                MessageBox.Show("請輸入帳號密碼。");
            }
        }

        private bool passwordcheck(string text1, string text2)
        {
            bool result = false;

            string SqlStr = $@"SELECT *
                              FROM [soco].[dbo].[User]
                              where user_account=@account and user_password=@password
                              " + ConfigurationManager.AppSettings["LIC"];

            using (var conn = new SqlConnection(_connectStrAPS))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    comm.Parameters.Add(("@account"), SqlDbType.VarChar).Value = text1;
                    comm.Parameters.Add(("@password"), SqlDbType.NVarChar).Value = PasswordUtility.SHA512Encryptor(text2);
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            result = true;
                        }
                        else
                        {
                            result = false;
                        }

                    }
                }
            }

            return result;
        }

        private UserInfo getusername(string text1)
        {
            UserInfo result = new UserInfo() ;

            string SqlStr = $@"SELECT [user_name] as Operater ,[user_id]  FROM [soco].[dbo].[User] where user_account=@account";

            using (var conn = new SqlConnection(_connectStrAPS))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    comm.Parameters.Add(("@account"), SqlDbType.VarChar).Value = text1;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            SqlData.Read();
                            result.Number= SqlData["user_id"].ToString();
                            result.Name= SqlData["Operater"].ToString();
                        }
                        else
                        {
                            result.Number = "N/A";
                            result.Name = "N/A";
                        }
                    }
                }
            }

            return result;
        }

        private void textBox2_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1_Click(null, null);
            }
        }
    }

    class Accunt
    {
        public string account { get; set; }
        public string password { get; set; }
    }
}
