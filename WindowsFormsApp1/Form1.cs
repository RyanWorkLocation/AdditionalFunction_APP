using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1(UserInfo User)
        {
            InitializeComponent();
            _connectStrAPS = ConfigurationManager.AppSettings["Localhost_APS"].ToString();
            APSDB = ConfigurationManager.AppSettings["APSDB"].ToString();
            MRPDB = ConfigurationManager.AppSettings["MRPDB"].ToString();
            MeasureDB = ConfigurationManager.AppSettings["MeasureDB"].ToString();
            ImagePath = ConfigurationManager.AppSettings["ImagePath"].ToString();
            QISFilePath = ConfigurationManager.AppSettings["QISFilePath"].ToString();
            defaultSavePath = ConfigurationManager.AppSettings["defaultSavePath"].ToString();
            _sysUser = User.Number;//ConfigurationManager.AppSettings["USER"].ToString();
            _sysUserName = User.Name;
            this.Text += (" - " + User.Name);

            // 初始化时设置提示文字
            textBox11.Text = placeholderText;
            textBox11.ForeColor = Color.Gray; // 提示文字用灰色显示

            // 绑定事件
            textBox11.Enter += textBox11_Enter; // 进入TextBox时触发
            textBox11.Leave += textBox11_Leave; // 离开TextBox时触发

            //// 設置DataGridView列
            //dataGridViewMachines.ColumnCount = 2;
            //dataGridViewMachines.Columns[0].Name = "機台編號";
            //dataGridViewMachines.Columns[1].Name = "機台名稱";
        }

        #region 參數變數設定
        string _connectStrAPS = "";
        string _sysUser = "";
        string _sysUserName = "";
        string APSDB = "";
        string MRPDB = "";
        string MeasureDB = "";
        string ImagePath = "";
        string QISFilePath = "";
        string defaultSavePath = "";
        string placeholderText = "*請輸入工單編號*"; // 设置提示文字
        #endregion

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                var ans = getAssigmentlist(textBox1.Text.ToUpper());
                comboBox1.Items.Clear();
                foreach (var item in ans)
                {
                    comboBox1.Items.Add(item.Range + "=" + item.OPID + "=" + item.OPLTXA1);
                }
                comboBox1.DroppedDown = true;
            }
        }

        private List<Models.Assignment> getAssigmentlist(string orderid)
        {
            var result = new List<Models.Assignment>();
            string SqlStr = $@"SELECT *
                              FROM {APSDB}.[dbo].Assignment where OrderID='{orderid}'";
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
                                });
                            }
                        }
                    }
                }
            }

            return result.OrderBy(x => x.Range).ToList();
        }

        private List<Models.Assignment> getSelectedlist(string orderid)
        {
            var result = new List<Models.Assignment>();
            string SqlStr = $@"SELECT *
                              FROM {APSDB}.[dbo].Assignment where OrderID='{orderid}' and WorkGroup is not null";
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
                                    WorkGroup = SqlData["WorkGroup"].ToString()
                                });
                            }
                        }
                    }
                }
            }

            return result.OrderBy(x => x.Range).ToList();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            int ans = updatewipdata(textBox1.Text, comboBox1.SelectedIndex);
            // 記錄點擊信息到log檔
            //string logFilePath = @"C:\PMC筆電備份\chase_log.txt";
            //using (StreamWriter sw = new StreamWriter(logFilePath, true)) // `true` 使其以追加模式打開文件
            //{
            //    sw.WriteLine($"User:{_sysUserName}執行，追至OrderID: {textBox1.Text}, OPID: {comboBox1.SelectedIndex}, DateTime: {DateTime.Now}");
            //}
            string SqlStr = $@"INSERT INTO {APSDB}.[dbo].[ChaseUpLog]
                                VALUES(@OrderID, @OPID, @UserID, @UserName, @ExeTime)";

            using (var conn = new SqlConnection(_connectStrAPS))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = textBox1.Text;
                    comm.Parameters.Add(("@OPID"), SqlDbType.NVarChar).Value = comboBox1.Items[comboBox1.SelectedIndex].ToString().Split('=')[1];
                    comm.Parameters.Add(("@UserID"), SqlDbType.NVarChar).Value = _sysUser;
                    comm.Parameters.Add(("@UserName"), SqlDbType.NVarChar).Value = _sysUserName;
                    comm.Parameters.Add(("@ExeTime"), SqlDbType.DateTime).Value = DateTime.Now;

                    // 使用 ExecuteNonQuery 並判斷是否成功執行
                    int rowsAffected = comm.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("記錄成功插入到 ChaseUpLog 表。");
                    }
                    else
                    {
                        MessageBox.Show("記錄插入失敗。");
                    }
                }
            }
            if (ans != 0)
            {
                MessageBox.Show("已經影響" + ans.ToString() + "筆資料");

            }
            else
            {
                MessageBox.Show("更新資料失敗");
            }
        }

        private int updatewipdata(string text, int selectedItemindex)
        {
            int result = 0;
            for (int i = 0; i <= selectedItemindex; i++)
            {

                string SqlStr = $@"
                                IF EXISTS (SELECT * FROM {APSDB}.[dbo].WIP WHERE OrderID = @OrderID and OPID=@OPID and( StartTime is null or EndTime is null))
                                BEGIN
                                    IF EXISTS (SELECT * FROM {APSDB}.[dbo].WIP WHERE OrderID = @OrderID and OPID=@OPID and StartTime is null)
                                        BEGIN
	                                        update {APSDB}.[dbo].WIP set WIPEvent=3, StartTime=GETDATE(),EndTime=GETDATE(),QtyGood=OrderQTY,QtyTol=OrderQTY 
                                            where OrderID=@OrderID and OPID=@OPID

                                            insert into {APSDB}.[dbo].WIPLog Values (@OrderID, @OPID, 0, 0, (select TOP(1) WorkGroup from {APSDB}.[dbo].Assignment where OrderID=@OrderID and OPID=@OPID), 0, 1, GETDATE(),{_sysUser})
                                            insert into {APSDB}.[dbo].WIPLog Values (@OrderID, @OPID, (select TOP(1) OrderQTY from {APSDB}.[dbo].Assignment where OrderID=@OrderID and OPID=@OPID), 0, (select TOP(1) WorkGroup from {APSDB}.[dbo].Assignment where OrderID=@OrderID and OPID=@OPID), 0, 3, GETDATE(),{_sysUser})

                                            update {APSDB}.[dbo].Assignment set Operator={_sysUser} where OrderID=@OrderID and OPID=@OPID
                                            UPDATE {APSDB}.[dbo].WipRegisterLog SET WorkOrderID = NULL,OPID = NULL,OperatorID = NULL,LastUpdateTime = GETDATE() WHERE WorkOrderID = @OrderID AND OPID = @OPID;
	                                    END
                                    ELSE IF EXISTS (SELECT * FROM {APSDB}.[dbo].WIP WHERE OrderID = @OrderID and OPID=@OPID and StartTime is not null and EndTime is null)
                                        BEGIN
	                                        update {APSDB}.[dbo].WIP set WIPEvent=3,EndTime=GETDATE(),QtyGood=OrderQTY,QtyTol=OrderQTY 
                                            where OrderID=@OrderID and OPID=@OPID
                                            
                                            -- 若已有開工紀錄，就新增完工紀錄即可
                                            -- insert into {APSDB}.[dbo].WIPLog Values (@OrderID, @OPID, 0, 0, (select TOP(1) WorkGroup from {APSDB}.[dbo].Assignment where OrderID=@OrderID and OPID=@OPID), 0, 1, GETDATE(),{_sysUser})
                                            insert into {APSDB}.[dbo].WIPLog Values (@OrderID, @OPID, (select TOP(1) OrderQTY from {APSDB}.[dbo].Assignment where OrderID=@OrderID and OPID=@OPID), 0, (select TOP(1) WorkGroup from {APSDB}.[dbo].Assignment where OrderID=@OrderID and OPID=@OPID), 0, 3, GETDATE(),{_sysUser})

                                            update {APSDB}.[dbo].Assignment set Operator={_sysUser} where OrderID=@OrderID and OPID=@OPID
                                            UPDATE {APSDB}.[dbo].WipRegisterLog SET WorkOrderID = NULL,OPID = NULL,OperatorID = NULL,LastUpdateTime = GETDATE() WHERE WorkOrderID = @OrderID AND OPID = @OPID;
                                        END
                                END";

                using (var conn = new SqlConnection(_connectStrAPS))
                {
                    using (var comm = new SqlCommand(SqlStr, conn))
                    {
                        if (conn.State != ConnectionState.Open)
                            conn.Open();
                        comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = text;
                        comm.Parameters.Add(("@OPID"), SqlDbType.NVarChar).Value = comboBox1.Items[i].ToString().Split('=')[1];
                        result += comm.ExecuteNonQuery();
                    }
                }
            }
            return result;
        }


        checkeAppStarted checkeAppStarted = new checkeAppStarted();
        private void Form1_Load(object sender, EventArgs e)
        {
            #region 判斷程式是否已開啟
            {
                checkeAppStarted.IskAppStart();
            }
            #endregion

            textBox1.TabIndex = 0;

        }


        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = string.Empty;
            comboBox1.Items.Clear();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int flag = 0;
            //刪除訂單
            if(textBox2.Text != "" && textBox3.Text == "")
            {
                flag = 1;
                if (deleteOrder(textBox2.Text, textBox3.Text,flag) != 0)
                {
                    MessageBox.Show("刪除成功");
                }
                else
                {
                    MessageBox.Show("刪除失敗");
                }
            }
            //刪除工單
            else if(textBox2.Text != "" && textBox3.Text != "")
            {
                flag = 2;
                if (deleteOrder(textBox2.Text, textBox3.Text,flag) != 0)
                {
                    MessageBox.Show("刪除成功");
                }
                else
                {
                    MessageBox.Show("刪除失敗");
                }
            }
            else
            {
                MessageBox.Show("請輸入正確資訊");
            }
        }

        private int deleteOrder(string order, string workorder,int flag)
        {
            int result = 0;
            string SqlStr = string.Empty;
            switch (flag)
            {
                //刪訂單
                case 1:
                    SqlStr = $@"
                                        delete {APSDB}.[dbo].[OrderOverview] where OrderID='{order}'
                                        delete {APSDB}.[dbo].WorkOrderOverview where OrderID='{order}'
                                        delete {APSDB}.[dbo].OperationOverview where OrderID='{order}'
                                        delete {APSDB}.[dbo].[Assignment] where ERPOrderID ='{order}'
                                        delete from {APSDB}.[dbo].[WIP]
                                        where OrderID in (
                                        SELECT distinct WorkOrderID
                                          FROM {APSDB}.[dbo].[WorkOrderOverview]
                                          where OrderID='{order}')
                                        ";
                    break;
                //刪工單
                case 2:
                    SqlStr = $@"
                                        delete {APSDB}.[dbo].[OrderOverview] where OrderID='{order}' and WorkOrderID='{workorder}'
                                        delete {APSDB}.[dbo].WorkOrderOverview where OrderID='{order}'and WorkOrderID='{workorder}'
                                        delete {APSDB}.[dbo].OperationOverview where OrderID='{order}'and WorkOrderID='{workorder}'
                                        delete {APSDB}.[dbo].[Assignment] where ERPOrderID ='{order}' and OrderID='{workorder}'
                                        delete {APSDB}.[dbo].[WIP] where OrderID='{workorder}'
                                        ";
                    break;
            }
            
            using (var conn = new SqlConnection(_connectStrAPS))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    result += comm.ExecuteNonQuery();
                }
            }
            return result;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var ans = getSelectedlist(textBox1.Text.ToUpper());
            comboBox1.Items.Clear();
            foreach (var item in ans)
            {
                comboBox1.Items.Add(item.Range + "=" + item.OPID + "=" + item.OPLTXA1);
            }
            comboBox1.DroppedDown = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // 清空所有行
            dataGridViewMachines.Rows.Clear();

            // 清空所有列
            dataGridViewMachines.Columns.Clear();


            // 獲取使用者帳號
            string userId = textBox4.Text;
            if(string.IsNullOrEmpty(textBox4.Text))
            {
                MessageBox.Show("請輸入欲查詢人員帳號!", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                if (ExistUser(userId))
                {
                    // 查詢使用者綁定的機台列表
                    var machineLists = getMachineList(userId);

                    if (machineLists.Count > 0)
                    {
                        // 創建機台編號列
                        DataGridViewTextBoxColumn machineIDColumn = new DataGridViewTextBoxColumn();
                        machineIDColumn.Name = "機台編號";
                        machineIDColumn.HeaderText = "機台編號";
                        machineIDColumn.Width = 40; // 設置機台編號列的寬度

                        // 創建機台名稱列
                        DataGridViewTextBoxColumn machineNameColumn = new DataGridViewTextBoxColumn();
                        machineNameColumn.Name = "機台名稱";
                        machineNameColumn.HeaderText = "機台名稱";
                        machineNameColumn.Width = 100; // 設置機台名稱列的寬度

                        // 創建新的DataGridViewCheckBoxColumn
                        DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();
                        checkBoxColumn.Name = "綁定";
                        checkBoxColumn.HeaderText = "綁定";
                        checkBoxColumn.Width = 60; // 設置選擇列的寬度

                        // 將列添加到DataGridView的列集合中
                        dataGridViewMachines.Columns.Add(machineIDColumn);
                        dataGridViewMachines.Columns.Add(machineNameColumn);
                        dataGridViewMachines.Columns.Add(checkBoxColumn);


                        // 將機台列表添加到DataGridView中
                        foreach (var machine in machineLists)
                        {
                            // 創建一行數據，根據機台編號、名稱和Authorize欄位值
                            var row = new DataGridViewRow();

                            row.CreateCells(dataGridViewMachines, machine.MachineID, machine.MachineName, machine.Authorize == 1);

                            // 將行添加到DataGridView中
                            dataGridViewMachines.Rows.Add(row);
                        }

                        // 將所有列的標題對齊方式設置為居中
                        foreach (DataGridViewColumn column in dataGridViewMachines.Columns)
                        {
                            column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        }
                    }
                    else
                    {
                        MessageBox.Show("無相關機台權限資料!", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    MessageBox.Show("無此帳號，請再次確認!", "查無帳號提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            

            



        }

        private List<Models.MachineList> getMachineList(string userid)
        {
            List<Models.MachineList> result = new List<Models.MachineList>();


            var SqlStr = $@"SELECT
                                d.ID,
                                d.remark,
                                CASE
                                    WHEN gd.DeviceId IS NOT NULL THEN 1
                                    ELSE 0
                                END AS Authorize
                            FROM
                                [01_APS_DF].[dbo].[Device] AS d
                            LEFT JOIN (
                                SELECT DISTINCT
                                    b.DeviceId
                                FROM
                                    [soco].[dbo].[User] AS a
                                INNER JOIN
                                    [soco].[dbo].[GroupDevice] AS b ON a.usergroup_id = b.GroupSeq
                                INNER JOIN
                                    [01_APS_DF].[dbo].[Device] AS c ON b.DeviceId = c.ID
                                WHERE
                                    a.user_account = '{userid}'
                            ) AS gd ON d.ID = gd.DeviceId
                            ORDER BY
                                d.ID;";




            using (var conn = new SqlConnection(_connectStrAPS))
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
                                result.Add(new Models.MachineList
                                {
                                    MachineID = !String.IsNullOrEmpty(SqlData["ID"].ToString()) ? SqlData["ID"].ToString() : "None",
                                    MachineName = !String.IsNullOrEmpty(SqlData["remark"].ToString()) ? SqlData["remark"].ToString() : "None",
                                    Authorize = Convert.ToInt16(SqlData["Authorize"])
                                });
                            }
                        }
                    }
                }
            }


            return result;
        }

        private void textBox4_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button5_Click(null, null);
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 檢查當前選中的 TabPage 是否是 tabPage3
            if (tabControl1.SelectedTab == tabPage3)
            {
                // 將焦點設置到 textBox4 上
                textBox4.Focus();
            }
            else if(tabControl1.SelectedTab == tabPage4)
            {
                textBox5.Focus();
            }
            else if(tabControl1.SelectedTab == tabPage1)
            {
                textBox1.Focus();
            }
            else if (tabControl1.SelectedTab == tabPage5)
            {
                textBox11.Focus();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //所有群組綁定資料
            var AllGroupData = FindEachGroup();

            //撈取datagridview資料
            // 創建一個列表來保存 DataGridView 的結果
            List<Models.Machine> SaveData = new List<Models.Machine>();

            // 遍歷 DataGridView 的每一行
            foreach (DataGridViewRow row in dataGridViewMachines.Rows)
            {
                // 跳過新行（通常為空行）
                if (row.IsNewRow)
                {
                    continue;
                }

                if(row.Cells["綁定"].Value != null && Convert.ToInt32(row.Cells["綁定"].Value) == 1)
                {
                    // 創建一個 MachineList 的實例
                    Models.Machine machine = new Models.Machine();

                    // 從 DataGridViewRow 中提取數據
                    machine.MachineID = row.Cells["機台編號"].Value.ToString();
                    machine.MachineName = row.Cells["機台名稱"].Value.ToString();

                    // 將 MachineList 對象添加到列表中
                    SaveData.Add(machine);
                }
                
            }
            bool isEqual=false;
            Models.MachineComparer comparer = new Models.MachineComparer();
            foreach (var eachgroup in AllGroupData)
            {
                //確認是否有相符群組
                // 獲取群組中的機台列表
                List<Models.Machine> groupMachines = eachgroup.Value;
                isEqual = SaveData.SequenceEqual(groupMachines, comparer);
                Console.WriteLine(isEqual); // Output: True
                //若有相同groupid，則修改此此使用者機台群組
                if (isEqual==true)
                {
                    var exeresult = UpdateGroupID(textBox4.Text.Trim(), eachgroup.Key);

                    MessageBox.Show($"{exeresult}","綁定結果提示",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    break;

                }
            }
            if(isEqual==false)
            {
                // 找到 result 中鍵（key）的最大值
                int maxKey = AllGroupData.Keys.Max();
                //創建一個新的群組編號
                var exeresult = InsertNewDevice(SaveData, maxKey+1, textBox4.Text.Trim());
                //更新使用者的群組編號
                UpdateGroupID(textBox4.Text.Trim(), maxKey + 1);
                MessageBox.Show($"{exeresult}", "綁定結果提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            // 在 button2.Click 事件處理程序中直接調用 button1.Click 事件處理程序
            button5_Click(sender, e);




        }

        private Dictionary<int, List<Models.Machine>> FindEachGroup()
        {
            Dictionary<int, List<Models.Machine>> result = new Dictionary<int, List<Models.Machine>>();

            //先加入groupid
            var SqlStr = $@"SELECT distinct 
                                  [GroupSeq]
                              FROM [soco].[dbo].[GroupDevice] as a                              
                              order by GroupSeq;";




            using (var conn = new SqlConnection(_connectStrAPS))
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
                                if (!result.ContainsKey(Convert.ToInt16(SqlData["GroupSeq"])))
                                {
                                    List<Models.Machine> temp = new List<Models.Machine>();
                                    result[Convert.ToInt16(SqlData["GroupSeq"])] = temp;
                                }
                            }
                        }
                    }
                }
            }

            SqlStr = string.Empty;
            SqlStr = $@"SELECT 
                                  a.[GroupSeq]
                                  ,a.[DeviceId]
	                              ,b.remark
                              FROM [soco].[dbo].[GroupDevice] as a
                              left join [01_APS_DF].[dbo].[Device]  as b
                              on a.DeviceId=b.ID
                              where b.remark is not null
                              order by a.GroupSeq,a.DeviceId;";




            using (var conn = new SqlConnection(_connectStrAPS))
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

                                result[Convert.ToInt16(SqlData["GroupSeq"])].Add(new Models.Machine
                                {
                                    MachineID = SqlData["DeviceId"].ToString(),
                                    MachineName = SqlData["remark"].ToString()
                                });
                            }
                        }
                    }
                }
            }


            return result;
        }
        private string UpdateGroupID(string user_account, int groupid)
        {
            var result = "更新成功";


            var SqlStr = $@"update [soco].[dbo].[User]
                              set usergroup_id={groupid}
                              where user_account='{user_account}';";




            using (var conn = new SqlConnection(_connectStrAPS))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    try
                    {
                        // 使用 ExecuteNonQuery 執行查詢
                        // ExecuteNonQuery 返回修改的行數
                        int rowsAffected = comm.ExecuteNonQuery();

                        // 根據修改行數進行相應處理
                        if (rowsAffected > 0)
                        {
                            result = "更新成功";
                        }

                    }
                    catch (Exception ex)
                    {
                        result = $"執行查詢時發生異常: {ex.Message}";
                    }
                }
            }


            return result;
        }

        private string InsertNewDevice(List<Models.Machine> newdevicegroup, int groupid, string user_account)
        {
            var result = "更新成功";

            // 構建一個 INSERT 語句
            string sql = $"INSERT INTO [soco].[dbo].[Units]([usergroup_id],[usergroup_name]) VALUES({groupid},'新增群組')" +
                " INSERT INTO [soco].[dbo].[GroupDevice] VALUES ";
            using (var conn = new SqlConnection(_connectStrAPS))
            {
                

                

                // 添加多個值列表
                List<string> values = new List<string>();
                foreach (var item in newdevicegroup)
                {
                    // 添加一組值
                    values.Add($"({groupid}, {item.MachineID})");
                }

                // 將所有值列表連接起來
                sql += string.Join(", ", values);
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                // 使用 SqlCommand
                using (SqlCommand comm = new SqlCommand(sql, conn))
                {
                    try
                    {
                        // 使用 ExecuteNonQuery 執行查詢
                        // ExecuteNonQuery 返回修改的行數
                        int rowsAffected = comm.ExecuteNonQuery();

                        // 根據修改行數進行相應處理
                        if (rowsAffected > 0)
                        {
                            result = "更新成功";
                        }

                    }
                    catch (Exception ex)
                    {
                        result = $"執行查詢時發生異常: {ex.Message}";
                    }
                }
            }


            return result;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox5.Text))
            {
                MessageBox.Show("請輸入帳號!");
            }
            else
            {
                //確認帳號是否存在，避免重複新增
                bool result = ExistUser(textBox5.Text);
                if (result == true)
                {
                    MessageBox.Show("使用者已存在");
                }
                else
                {
                    int ret = AddUser(textBox5.Text, textBox6.Text, textBox7.Text);
                    if (ret != 0)
                        MessageBox.Show("已新增使用者【" + textBox5.Text + "】!");
                    else
                        MessageBox.Show("資料未填寫完全!");
                }
                textBox5.Text = string.Empty;
                textBox6.Text = string.Empty;
                textBox7.Text = string.Empty;
            }
        }
        /// <summary>
        /// 確認是否已存在帳號
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public bool ExistUser(string id)
        {
            List<string> accountlist = new List<string>();

            string sqlStr = $@"SELECT DISTINCT user_account FROM [soco].[dbo].[User]";
            using (SqlConnection conn = new SqlConnection(_connectStrAPS))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                {
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                accountlist.Add(SqlData["user_account"].ToString());
                            }
                        }
                    }
                }
            }
            if (accountlist.Exists(x => x == id))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public int AddUser(string id, string password, string name)
        {
            int result;



            string sqlStr = $@"INSERT INTO [soco].[dbo].[User]
                            (user_account,user_password,user_name,usergroup_id)
                              VALUES(@account,@password,@name,0)";
            using (SqlConnection conn = new SqlConnection(_connectStrAPS))
            {

                using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.AddWithValue("@account", id);
                    comm.Parameters.AddWithValue("@password", PasswordUtility.SHA512Encryptor(password));
                    comm.Parameters.Add(("@name"), SqlDbType.NVarChar).Value = name;

                    result = comm.ExecuteNonQuery();
                }
            }
            return result;
        }

        public int UpdateUser(string id, string password,string name)
        {
            int result;

            string sqlStr = $@"UPDATE [soco].[dbo].[User]
                                SET user_password=@password, user_name = @name
                                WHERE user_account=@account";
            using (SqlConnection conn = new SqlConnection(_connectStrAPS))
            {

                using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.AddWithValue("@account", id);
                    comm.Parameters.AddWithValue("@password", PasswordUtility.SHA512Encryptor(password));
                    comm.Parameters.AddWithValue("@name", name);
                    result = comm.ExecuteNonQuery();
                }
            }
            return result;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox8.Text))
            {
                MessageBox.Show("請輸入帳號!");
            }
            else
            {
                int ret = UpdateUser(textBox8.Text, textBox9.Text, textBox10.Text);
                if (ret != 0)
                    MessageBox.Show("已修改使用者【" + textBox8.Text + "】!");
                else
                    MessageBox.Show("無此帳號!");
                textBox8.Text = string.Empty;
                textBox9.Text = string.Empty;
            }
        }

        private void textBox11_Enter(object sender, EventArgs e)
        {
            // 当焦点进入时，如果是提示文字就清空
            if (textBox11.Text == placeholderText)
            {
                textBox11.Text = "";
                textBox11.ForeColor = Color.Black; // 输入文字用黑色
            }
        }

        private void textBox11_Leave(object sender, EventArgs e)
        {
            // 当焦点离开时，如果为空就显示提示文字
            if (string.IsNullOrWhiteSpace(textBox11.Text))
            {
                textBox11.Text = placeholderText;
                textBox11.ForeColor = Color.Gray; // 提示文字用灰色
            }
        }

        /// <summary>
        /// 從資料庫撈出品檢資料
        /// </summary>
        private void LoadQISDataToDataGridView()
        {
            string query = $@"
                SELECT
                    OrderVIEW.OrderID,
                    OrderVIEW.CustomerInfo,
                    [WorkOrderID],
                    QCP.[OPID],
                    PRO.ProcessName,
                    QCP.[MAKTX],
                    P.Name AS ProductName,
                    QCP.[QCPoint],
                    [QCPointValue],
                    QCR.QCPointName,
                    QCR.QCCL,
                    CASE 
                        WHEN CONVERT(FLOAT, QCR.QCUSL) = 0 THEN 0
                        ELSE ROUND((CONVERT(FLOAT, QCR.QCUSL) - CONVERT(FLOAT, QCR.QCCL)),2)
                    END as Lower_error,
                    CASE 
                        WHEN CONVERT(FLOAT, QCR.QCLSL) = 0 THEN 0
                        ELSE ROUND((CONVERT(FLOAT, QCR.QCLSL) - CONVERT(FLOAT, QCR.QCCL)),2)
                    END as Upper_error,
                    [QCunit],
                    CONVERT(DATE, [Lastupdatetime]) AS UpdateDate,
                    [QCMan],
                    [QCMode]
                FROM {APSDB}.[dbo].[QCPointValue] AS QCP
                LEFT JOIN {MRPDB}.[dbo].[QCrule] AS QCR
                    ON QCP.OPID = QCR.id AND QCP.QCPoint = QCR.QCPoint
                LEFT JOIN {MRPDB}.[dbo].[Part] AS P
                    ON QCP.MAKTX = P.Number
                LEFT JOIN {MRPDB}.[dbo].[Process] AS PRO
                    ON QCP.OPID = PRO.ID
                LEFT JOIN {APSDB}.[dbo].[Assignment] AS ASSIGN
                    ON QCP.WorkOrderID = ASSIGN.OrderID AND QCP.OPID = ASSIGN.OPID
                INNER JOIN {APSDB}.[dbo].[OrderOverview] AS OrderVIEW
                    ON ASSIGN.ERPOrderID = OrderVIEW.OrderID
                WHERE QCPointValue IS NOT NULL
                    AND PRO.ProcessName IS NOT NULL
                    AND WorkOrderID = '{textBox11.Text}';";

            using (var conn = new SqlConnection(_connectStrAPS))
            {
                using (var comm = new SqlCommand(query, conn))
                {
                    conn.Open();
                    using (SqlDataReader reader = comm.ExecuteReader())
                    {
                        if (!reader.HasRows)
                        {
                            MessageBox.Show("查無檢驗資料!","提示", MessageBoxButtons.OK,MessageBoxIcon.Warning);
                            return;
                        }

                        // 清空並設定DataGridView列
                        dataGridView1.Columns.Clear();
                        dataGridView1.Columns.Add("OrderID", "訂單編號");
                        dataGridView1.Columns.Add("CustomerInfo", "客戶資訊");
                        dataGridView1.Columns.Add("ProcessName", "製程名稱");
                        dataGridView1.Columns.Add("ProductName", "產品名稱");
                        dataGridView1.Columns.Add("QCPointName", "檢驗點名稱");
                        dataGridView1.Columns.Add("QCCL", "標準值");
                        dataGridView1.Columns.Add("Lower_error", "下限誤差");
                        dataGridView1.Columns.Add("Upper_error", "上限誤差");
                        dataGridView1.Columns.Add("QCPointValue", "檢驗值");

                        // 設定唯讀欄位
                        // 設定唯讀欄位並標示為淺灰色
                        SetColumnReadOnlyAndGray("OrderID");
                        SetColumnReadOnlyAndGray("CustomerInfo");
                        SetColumnReadOnlyAndGray("ProcessName");
                        SetColumnReadOnlyAndGray("ProductName");
                        SetColumnReadOnlyAndGray("QCPointName");
                        SetColumnReadOnlyAndGray("QCCL");
                        SetColumnReadOnlyAndGray("Lower_error");
                        SetColumnReadOnlyAndGray("Upper_error");

                        dataGridView1.Columns["CustomerInfo"].Visible = false;
                        dataGridView1.Columns["ProductName"].Visible = false;

                        dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font(dataGridView1.Font, FontStyle.Bold);
                        dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // 置中對齊
                        dataGridView1.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False; // 禁止換行
                        //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells; // 讓欄位自適應大小

                        // 填充數據
                        while (reader.Read())
                        {
                            dataGridView1.Rows.Add(
                                reader["OrderID"],
                                reader["CustomerInfo"],
                                reader["ProcessName"],
                                reader["ProductName"],
                                reader["QCPointName"],
                                reader["QCCL"],
                                reader["Lower_error"],
                                reader["Upper_error"],
                                reader["QCPointValue"]
                            );
                        }
                    }
                }
            }
        }

        // 輔助方法：將指定欄位設為唯讀並標示為淺灰色
        private void SetColumnReadOnlyAndGray(string columnName)
        {
            dataGridView1.Columns[columnName].ReadOnly = true;
            dataGridView1.Columns[columnName].DefaultCellStyle.BackColor = ColorTranslator.FromHtml("#D9D9D9");;
        }

        private void textBox11_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                LoadQISDataToDataGridView();
                txtSavePath.Text = defaultSavePath;
                e.Handled = true; // 防止Enter鍵觸發其他行為（如換行）
            }
        }

        private void folder_preview_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "請選擇儲存路徑";
                folderBrowserDialog.ShowNewFolderButton = true;

                // 設定初始位置
                string initialPath = @"C:\PMC筆電備份\廠商資料\3. 鼎烽\小工具\品質管制標準書測試"; // 替換為你想要的路徑
                if (Directory.Exists(initialPath)) // 檢查路徑是否存在
                {
                    folderBrowserDialog.SelectedPath = initialPath;
                }
                else
                {
                    // 如果初始路徑不存在，預設從桌面開始
                    folderBrowserDialog.RootFolder = Environment.SpecialFolder.Desktop;
                }

                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    txtSavePath.Text = folderBrowserDialog.SelectedPath;
                }
            }
        }

        private void download_Click(object sender, EventArgs e)
        {
            if( dataGridView1.Rows.Count == 0 )
            {
                MessageBox.Show("無檢驗資料可儲存!", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                #region 儲存檔案
                // 設定 Excel 讀取授權
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                // STEP 1. 尋照圖紙檔案
                string ImgFile = string.Empty;
                string query = $@"
                             SELECT ImgPath
                                FROM {APSDB}.[dbo].[Assignment]
                             WHERE OrderID='{textBox11.Text}'
                            ";

                using (var conn = new SqlConnection(_connectStrAPS))
                {
                    using (var comm = new SqlCommand(query, conn))
                    {
                        conn.Open();
                        using (SqlDataReader reader = comm.ExecuteReader())
                        {
                            if (!reader.HasRows)
                            {
                                Console.WriteLine("查無資料");
                                return;
                            }

                            while (reader.Read())  // 需要調用 Read() 方法
                            {
                                ImgFile = reader["ImgPath"].ToString();  // 讀取第一列數據

                            }


                        }
                    }
                }

                // STEP 2. 把datagridview填入excel
                //string excelPath = @"C:\PMC筆電備份\廠商資料\3. 鼎烽\小工具\品質管制標準書測試\鼎烽-品質管製表.xlsx";
                string savePath = string.IsNullOrWhiteSpace(txtSavePath.Text.Trim())
                    ? Path.Combine(defaultSavePath, $"品質管制標準書_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx")
                    : Path.Combine(txtSavePath.Text, $"品質管制標準書_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

                FileInfo existingFile = new FileInfo(QISFilePath);
                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // 插入圖片
                    string baseUrl = ImagePath;
                    if (File.Exists(baseUrl + ImgFile))
                    {
                        ExcelPicture picture = worksheet.Drawings.AddPicture("MyPic", baseUrl + ImgFile);
                        int widthPx = (int)(28.33f * 37.79527559f);
                        int heightPx = (int)(15.90f * 37.79527559f);
                        picture.SetPosition(4, 10, 1, 10);
                        picture.SetSize(widthPx, heightPx);
                    }

                    // 從DataGridView填充Excel
                    int rowIndex = 26;
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.IsNewRow) continue; // 跳過新行

                        if (rowIndex == 26)
                        {
                            worksheet.Cells["B4"].Value = "客戶: " + row.Cells["CustomerInfo"].Value.ToString().Split('/')[1];
                            worksheet.Cells["E4"].Value = "品名: " + row.Cells["ProductName"].Value;
                            worksheet.Cells["N3"].Value = DateTime.Now.ToString("yyyy/MM/dd");
                            worksheet.Cells["F45"].Value = row.Cells["OrderID"].Value;

                            if (row.Cells["ProcessName"].Value.ToString().Contains("出貨前"))
                                worksheet.Cells["L3"].Value = "☑出貨檢";
                            else
                                worksheet.Cells["O2"].Value = "☑抽檢";
                        }

                        var upperError = Convert.ToDouble(row.Cells["Upper_error"].Value);
                        var lowerError = Convert.ToDouble(row.Cells["Lower_error"].Value);
                        var inspectValue = Convert.ToDouble(row.Cells["QCPointValue"].Value);

                        if (rowIndex <= 40)
                        {
                            worksheet.Cells[rowIndex, 3].Value = $"{row.Cells["QCPointName"].Value} {row.Cells["QCCL"].Value}\n{lowerError} ~ {upperError}";
                            worksheet.Cells[rowIndex, 3].Style.WrapText = true;
                            worksheet.Cells[rowIndex, 6].Value = row.Cells["QCPointValue"].Value;
                            worksheet.Cells[rowIndex, 7].Value = (inspectValue <= upperError && inspectValue >= lowerError) ? "V" : "";
                            worksheet.Cells[rowIndex, 8].Value = (inspectValue > upperError || inspectValue < lowerError) ? "V" : "";
                        }
                        else
                        {
                            int rightSideRow = rowIndex - 15;
                            worksheet.Cells[rightSideRow, 10].Value = $"{row.Cells["QCPointName"].Value} {row.Cells["QCCL"].Value}\n{lowerError}~{upperError}";
                            worksheet.Cells[rightSideRow, 10].Style.WrapText = true;
                            worksheet.Cells[rightSideRow, 13].Value = row.Cells["QCPointValue"].Value;
                            worksheet.Cells[rightSideRow, 14].Value = (inspectValue <= upperError && inspectValue >= lowerError) ? "V" : "";
                            worksheet.Cells[rightSideRow, 15].Value = (inspectValue > upperError || inspectValue < lowerError) ? "V" : "";
                        }
                        rowIndex++;
                    }

                    package.SaveAs(new FileInfo(savePath));
                    MessageBox.Show($"文件已成功儲存至: {savePath}", "儲存成功提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }


                #endregion
            }

        }
    }
}
