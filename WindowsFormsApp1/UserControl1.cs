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
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }

        //string _connectStrAPS = "Data Source = 127.0.0.1; Initial Catalog = 01_APS_DF; User ID = MES2014; Password = PMCMES;";

        private string _AppVersion = "1.0";
        public string AppVersion
        {
            get { return _AppVersion; }
            set { _AppVersion = value; }
        }

        private string _connectStrAPS = "Data Source = 127.0.0.1; Initial Catalog = 01_APS_DF; User ID = MES2014; Password = PMCMES;";
        public string connectStrAPS
        {
            get { return _connectStrAPS; }
            set { _connectStrAPS = value; }
        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                var ans = getAssigmentlist(textBox1.Text.ToUpper());
                comboBox2.Items.Clear();
                foreach (var item in ans)
                {
                    comboBox2.Items.Add(item.Range + "=" + item.OPID + "=" + item.OPLTXA1);
                }
                comboBox2.DroppedDown = true;
            }
        }

        private List<Models.Assignment> getAssigmentlist(string orderid)
        {
            var result = new List<Models.Assignment>();
            string SqlStr = $@"SELECT *
                              FROM Assignment where OrderID='{orderid}'";
            using (SqlConnection conn = new SqlConnection(_connectStrAPS))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                result.Add(new Models.Assignment
                                {
                                    SeriesID = SqlData["SeriesID"].ToString(),
                                    OrderID = SqlData["OrderID"].ToString(),
                                    ERPOrderID = SqlData["ERPOrderID"].ToString(),
                                    OPID = SqlData["OPID"].ToString(),
                                    OPLTXA1 = SqlData["OPLTXA1"].ToString(),
                                    Range = int.Parse(SqlData["Range"].ToString()),
                                    MAKTX= SqlData["MAKTX"].ToString(),
                                });
                            }
                        }
                    }
                }
            }

            return result.OrderBy(x => x.Range).ToList();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> temp = getOPmachines(textBox1.Text.Trim().ToUpper(), comboBox2.SelectedItem.ToString().Split('=')[1]);
            comboBox1.Items.Clear();
            foreach (var item in temp)
            {
                comboBox1.Items.Add(item);
            }
            comboBox1.DroppedDown = true;
        }

        private List<string> getOPmachines(string orderid, string opid)
        {
            var result = new List<string>();
            string SqlStr = $@"SELECT * FROM [01_MRP_DF].[dbo].[Part] as a 
                            left join [01_MRP_DF].dbo.RoutingDetail as b on a.RoutingID=b.RoutingId
                            left join  [01_MRP_DF].dbo.ProcessDetial as c on b.ProcessId=c.ProcessID
                            left join [01_APS_DF].dbo.Device as d on c.MachineID=d.ID

                            where a.Number=(select top(1)MAKTX from [01_APS_DF].dbo.Assignment where OrderID='{orderid}') and b.ProcessId={opid}";
            using (SqlConnection conn = new SqlConnection(_connectStrAPS))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                result.Add(SqlData["remark"].ToString());
                            }
                        }
                    }
                }
            }

            return result;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = string.Empty;
            comboBox1.Text = string.Empty;
            comboBox2.Text = string.Empty;
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
        }
    }
}
