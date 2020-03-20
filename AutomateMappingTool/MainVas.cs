using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using AutomateMappingTool;

namespace MappingDiscount
{
    public partial class MainVas : Form
    {
        private HomeVas homeVas;
        private OracleConnection ConnectionProd;
        private OracleConnection ConnectionTemp;
        private ChangeFormat changeFormat;
        private DgvSettings settingView;
        private ReserveID reserv;
        private Validation validate;
        private string user;
        private string urNo;
        private string filename;
        private string outputPath;
        private string log = "";
        private bool isCorrect;
        private List<int> indexListbox;

        Excel.Application xlApp;

        //For maximize form
        int maximize = 1;

        public MainVas(HomeVas form, OracleConnection con, string filename,
            string folder, string user, string ur)
        {
            InitializeComponent();

            homeVas = form;
            ConnectionProd = con;
            this.user = user;
            this.filename = filename;
            outputPath = folder;
            urNo = ur;
        }

        private void ViewVasData_SizeChanged(object sender, EventArgs e)
        {
            int w = this.Size.Width;
            int h = this.Size.Height;

            //btnOK.Location = new Point(w - 130, h - 100);

            if (listBox1.Items.Count > 0)
            {
                btnOK.Location = new Point((w - panelHome.Width) + 60, h - 124);
            }
            else
            {
                btnOK.Location = new Point((w - panelHome.Width) + 60, h - 105);
            }

            label1.Location = new Point(w - 23, 0);
            btnMinimize.Location = new Point(w - 88, 3);
            btnMaximize.Location = new Point(w - 53, 3);

            dataGridView1.Size = new Size(w - panelHome.Width, h - 280);

        }

        private void ViewVasData_Load(object sender, EventArgs e)
        {
            homeVas.Hide();
            listBox1.Items.Clear();
            listBox1.Hide();

            Cursor.Current = Cursors.WaitCursor;

            settingView = new DgvSettings();
            changeFormat = new ChangeFormat();
            reserv = new ReserveID();

            this.dataGridView1.AllowUserToAddRows = false;

            List<string> lstHeader = new List<string>();
            lstHeader.Add("VasCode");
            lstHeader.Add("Channel");
            lstHeader.Add("MktCode");
            lstHeader.Add("OrderType");
            lstHeader.Add("Product");
            lstHeader.Add("Speed");
            lstHeader.Add("Province");
            lstHeader.Add("Effective");
            lstHeader.Add("Expire");

            ToVasProduct();

            settingView.SetDgv(dataGridView1, filename, "VAS New Sale(SMART UI)$B3:J", lstHeader);

            validate = new Validation(dataGridView1, ConnectionProd, ConnectionTemp, outputPath);

            listBox1.Items.Clear();
            ListBox ls = validate.Verify("VAS");

            foreach (string v in ls.Items)
            {
                listBox1.Items.Add(v);
            }

            if (listBox1.Items.Count > 0)
            {
                indexListbox = validate.IndexDgv;
                listBox1.Show();
            }
            else
            {
                listBox1.Hide();
            }

            toolStripStatusLabel1.Text = "Rows : " + dataGridView1.RowCount.ToString();

            Cursor.Current = Cursors.Default;
        }

        private void ToVasProduct()
        {
            //Create Excel
            xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = null;

            //Create an OracleCommand object using the connection object
            OracleCommand command = ConnectionProd.CreateCommand();
            OracleTransaction transaction = null;

            try
            {
                //Read data from file input
                xlWorkbook = xlApp.Workbooks.Open(filename);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets["New VAS code (VCare&CCBS)"];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                for (int index = 4; index <= xlRange.Rows.Count; index++)
                {
                    string rangeCode = "B" + index;
                    string vas_code = xlWorksheet.Range[rangeCode].Value2;
                    string vas_name;
                    string vas_type;
                    string vas_channel;
                    string vas_rule;
                    string vas_price;
                    string vas_parent;
                    string vas_start;
                    string vas_user = user.Trim();
                    string vas_end = "";

                    if (vas_code != null)
                    {
                        vas_name = xlWorksheet.Range["C" + index].Value2.Trim();
                        vas_type = xlWorksheet.Range["D" + index].Value2.Trim();
                        vas_rule = xlWorksheet.Range["E" + index].Value2.Trim();
                        vas_price = xlWorksheet.Range["F" + index].Value2.ToString().Trim();
                        vas_channel = xlWorksheet.Range["G" + index].Value2.Trim();
                        vas_parent = xlWorksheet.Range["H" + index].Value2.Trim();
                        vas_start = xlWorksheet.Range["I" + index].Value2.ToString().Trim();

                        if (vas_start != "-" || vas_start != "")
                        {
                            vas_start = changeFormat.FormatDate(vas_start);
                        }
                        else
                        {
                            vas_start = "";
                        }

                        string query = "SELECT VAS_CODE FROM VAS_PRODUCT WHERE VAS_CODE = '" + vas_code + "' AND VAS_TYPE = '" +
                            vas_type + "' AND VAS_CHANNEL = '" + vas_channel + "'";

                        OracleCommand cmd = new OracleCommand(query, ConnectionProd);

                        OracleDataReader reader = cmd.ExecuteReader();

                        if (reader.HasRows)
                        {
                            //write log
                            log += "VasCode already exists! : " + vas_code + "\r\n";
                        }
                        else
                        {
                            using (transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted))
                            {
                                try
                                {
                                    string cmdTxt = "INSERT INTO VAS_PRODUCT VALUES('" + vas_code + "','" + vas_name + "','" + vas_type + "','Active','" +
                                   vas_rule + "','" + vas_price + "', null, '" + vas_channel + "', null, '" + vas_parent + "',TO_DATE('" + vas_start + "','dd/mm/yyyy')," +
                                   " TO_DATE('" + vas_end + "','dd/mm/yyyy'), sysdate,'" + vas_user + "',null,null)";

                                    //command = new OracleCommand(cmdTxt, ConnectionProd, transaction);
                                    command.CommandText = cmdTxt;
                                    command.ExecuteNonQuery();

                                    transaction.Commit();
                                }
                                catch (Exception)
                                {
                                    transaction.Rollback();
                                }
                            }
                        }

                        reader.Close();
                    }
                }

            }
            catch (NullReferenceException)
            { }
            finally
            {
                xlApp.Quit();
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
            this.Close();

            homeVas.Show();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            ConnectionTemp = new OracleConnection();
            string connString = "Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.19.193.20)(PORT = 1560))" +
                   "(CONNECT_DATA = (SID = TEST03)));User Id= TRUREF71; Password= TRUREF71;";

            ConnectionTemp.ConnectionString = connString;

            //ConnectionTemp = ConnectionProd;

            string inactiveCode = CheckVasCode();

            try
            {
                ConnectionTemp.Open();

                bool hasUserInProcess = reserv.CheckStatus(ConnectionTemp, "VAS", user, urNo);

                if (hasUserInProcess == false)
                {
                    if (inactiveCode != "")
                    {
                        DialogResult result = MessageBox.Show("VasCode is inactive in table VAS_PRODUCT, Please review and confirm about these code."
                            + "\r\n" + "Detail : " + inactiveCode, "Invalid Code", MessageBoxButtons.OK);
                    }
                    else
                    {
                        int vasID = reserv.GetMaxID(ConnectionTemp, ConnectionProd, "VAS", user, urNo);

                        Execute(vasID);

                        if (log != "")
                        {
                            string strFilePath = outputPath + "\\log_" + urNo + ".txt";

                            using (StreamWriter writer = new StreamWriter(strFilePath, true))
                            {
                                writer.Write(log);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show("Failed..." + "\r\n" + ex.Message);

                Application.Exit();
            }

            Cursor.Current = Cursors.Default;

        }

        private string CheckVasCode()
        {
            string inactiveCode = "";
            OracleDataReader reader = null;

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                string code = dataGridView1.Rows[i].Cells[0].Value.ToString().Trim();

                string query = "SELECT * FROM VAS_PRODUCT WHERE VAS_CODE = '" + code + "' AND VAS_STATUS = 'Active'";

                OracleCommand cmd = new OracleCommand(query, ConnectionProd);
                reader = cmd.ExecuteReader();
                reader.Read();
                if (reader.HasRows == false)
                {
                    inactiveCode += code + " ";
                }
            }

            reader.Close();

            return inactiveCode;
        }

        private void Execute(int minId)
        {
            //create table view data from table mapping
            DataTable dtTableView = new DataTable();
            dtTableView.Clear();
            dtTableView.Columns.Add("DC_ID");
            dtTableView.Columns.Add("DC_TYPE");
            dtTableView.Columns.Add("DC_VALUE");
            dtTableView.Columns.Add("DC_GROUPID");
            dtTableView.Columns.Add("DC_START_DT");
            dtTableView.Columns.Add("DC_END_DT");
            dtTableView.Columns.Add("DC_ACTIVE_FLAG");

            string vas_code;
            string channel;
            string mkt;
            string speed;
            string product;
            string orderType;
            string effective;
            string expire;
            string province;
            string strMinID = "VAS" + string.Format("{0:0000000}", minId);
            string existID = "";

            try
            {
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    bool hasID = false;

                    vas_code = dataGridView1.Rows[i].Cells[0].Value.ToString().Trim();
                    channel = dataGridView1.Rows[i].Cells[1].Value.ToString();
                    mkt = dataGridView1.Rows[i].Cells[2].Value.ToString();
                    string upMkt = mkt.ToUpper();
                    if (upMkt.Equals("ALL"))
                    {
                        mkt = "ALL";
                    }
                    else
                    {
                        string[] lstMkt = mkt.Split('-');
                        mkt = lstMkt[0].Trim();
                    }

                    orderType = dataGridView1.Rows[i].Cells[3].Value.ToString();
                    product = dataGridView1.Rows[i].Cells[4].Value.ToString();
                    string getSpeed = dataGridView1.Rows[i].Cells[5].Value.ToString();
                    province = dataGridView1.Rows[i].Cells[6].Value.ToString();
                    effective = dataGridView1.Rows[i].Cells[7].Value.ToString();
                    expire = dataGridView1.Rows[i].Cells[8].Value.ToString();

                    if (expire.Equals("-"))
                    {
                        expire = String.Empty;
                    }

                    string[] lstChannel = null;
                    string[] lstorder = null;

                    //Cut each Order_Type
                    if (orderType.Contains(","))
                    {
                        lstorder = orderType.Split(',');
                    }
                    else
                    {
                        lstorder = new string[] { orderType };

                        string upper = lstorder[0].ToUpper().Trim();
                        if (upper.Equals("ALL"))
                        {
                            lstorder[0] = "ALL";
                        }
                    }

                    //Cut each Channel
                    if (channel.Contains(","))
                    {
                        lstChannel = channel.Split(',');
                    }
                    else
                    {
                        lstChannel = new string[] { channel };

                        string upper = lstChannel[0].ToUpper().Trim();
                        if (upper.Equals("ALL") || upper.Equals("DEFAULT"))
                        {
                            lstChannel[0] = "ALL";
                        }
                    }

                    //Speed
                    speed = changeFormat.FormatSpeed(getSpeed);
                    //Change format Date
                    effective = changeFormat.FormatDate(effective);

                    if (String.IsNullOrEmpty(expire))
                    {
                        expire = "";
                    }
                    else
                    {
                        expire = changeFormat.FormatDate(expire);
                    }

                    for (int j = 0; j < lstChannel.Length; j++)
                    {
                        for (int k = 0; k < lstorder.Length; k++)
                        {
                            string ch = lstChannel[j].Trim();
                            string ord = lstorder[k].Trim();

                            string strQuery = "SELECT * FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_ID IN (SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING " +
                               "WHERE DC_VALUE = '" + mkt + "' AND DC_GROUPID = '" + vas_code + "' AND DC_ID IN (SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING " +
                               "WHERE DC_VALUE = '" + ch + "') AND DC_ID IN (SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_VALUE = '" + speed +
                               "') AND DC_ID IN (SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_VALUE = '" + ord + "'))";

                            OracleCommand cmd = new OracleCommand(strQuery, ConnectionProd);
                            OracleDataReader reader = cmd.ExecuteReader();
                            reader.Read();
                            if (reader.HasRows)
                            {
                                hasID = true;
                                string idEx = reader["DC_ID"].ToString();
                                existID += "ID : " + idEx + ", code : " + vas_code + ", MKT : " + mkt + ", Order : " + ord +
                                    ", Channel : " + ch + ", Speed : " + speed + "\r\n";
                            }
                            else
                            {
                                hasID = false;
                                if (lstChannel[0].Equals("ALL"))
                                {
                                    UPDATECHANNEL(mkt, vas_code, speed, effective);
                                }
                            }

                            for (int m = 1; m <= 6; m++)
                            {
                                object[] obj = new object[7];
                                if (hasID == false)
                                {
                                    obj[0] = "VAS" + string.Format("{0:0000000}", minId);
                                }
                                else
                                {
                                    obj[0] = "";
                                }

                                obj[3] = vas_code;
                                obj[4] = effective;

                                if (String.IsNullOrEmpty(expire))
                                {
                                    obj[5] = "";
                                }
                                else
                                {
                                    obj[5] = expire;
                                }

                                obj[6] = "Y";

                                switch (m)
                                {
                                    case 1:
                                        obj[1] = "MARKETING_CODE";
                                        obj[2] = mkt;
                                        break;
                                    case 2:
                                        obj[1] = "ORDER_TYPE";
                                        obj[2] = ord;
                                        break;
                                    case 3:
                                        obj[1] = "PRODUCT";
                                        obj[2] = product;
                                        break;
                                    case 4:
                                        obj[1] = "PROVINCE";
                                        obj[2] = province;
                                        break;
                                    case 5:
                                        obj[1] = "SALE_CHANNEL";
                                        obj[2] = ch;
                                        break;
                                    case 6:
                                        obj[1] = "SPEED";
                                        obj[2] = speed;
                                        break;
                                }

                                dtTableView.Rows.Add(obj);
                            }

                            if (hasID == false)
                            {
                                minId += 1;
                            }
                        }

                    }
                }

                minId = minId - 1;
                string max_id = "VAS" + string.Format("{0:0000000}", minId);

                if (existID != "")
                {
                    log += "\r\n" + "Already exists data in database" + "\r\n" + existID + "\r\n";
                }

                DataTable[] lstTable = new DataTable[2];
                lstTable[0] = dtTableView;

                ViewResult viewResult = new ViewResult(this, lstTable, ConnectionProd, ConnectionTemp, urNo, strMinID,
                    max_id, user, "", outputPath);

                viewResult.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void UPDATECHANNEL(string mkt, string vasCode, string speed, string effective)
        {
            //Create an OracleCommand object using the connection object
            OracleCommand command = ConnectionProd.CreateCommand();
            OracleTransaction transaction;

            List<string> dcID = new List<string>();
            string id = "";

            string cmdTxt = "SELECT * FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_ID IN " +
                "(SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_VALUE = '" + mkt + "' AND DC_GROUPID = '" + vasCode +
                "' AND DC_ID IN(SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_VALUE = '" + speed + "'))";

            OracleCommand cmd = new OracleCommand(cmdTxt, ConnectionProd);
            OracleDataReader reader = cmd.ExecuteReader();

            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    dcID.Add(reader["DC_ID"].ToString());
                }

                reader.Close();
                reader.Dispose();

                // distinct DC_ID
                List<string> distinct = dcID.Distinct().ToList();
                dcID = new List<string>();
                dcID = distinct;

                for (int i = 0; i < dcID.Count; i++)
                {
                    id += "'" + dcID[i] + "',";
                }

                id = id.Substring(0, id.Length - 1);

                //update
                transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted);
                // Assign transaction object for a pending local transaction
                command.Transaction = transaction;

                try
                {
                    command.CommandText = "UPDATE DISCOUNT_CRITERIA_MAPPING SET DC_END_DT = " +
                    "TO_DATE('" + effective + "','dd/MM/yyyy') WHERE DC_ID IN(" + id + ")";

                    command.CommandType = CommandType.Text;

                    command.ExecuteNonQuery();
                    transaction.Commit();

                }
                catch (Exception e)
                {
                    transaction.Rollback();
                    Console.WriteLine("Can't update channel to 'All'" + "\r\n" + e);
                    throw new Exception(e.Message);
                }
            }
        }

        private void ViewVasData_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ConnectionTemp != null)
            {
                try
                {
                    if (ConnectionTemp.State != ConnectionState.Open)
                    {
                        ConnectionTemp.Open();
                    }

                    string query = "SELECT * FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = 'VAS' AND COMPLETE_FLAG = 'N'";

                    OracleCommand cmd = new OracleCommand(query, ConnectionTemp);
                    OracleDataReader reader = cmd.ExecuteReader();
                    reader.Read();
                    if (reader.HasRows)
                    {
                        reader.Close();
                        reader.Dispose();

                        string qryDel = "DELETE FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = 'VAS' AND COMPLETE_FLAG = 'N'";
                        OracleCommand command = new OracleCommand(qryDel, ConnectionTemp);
                        command.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    string log = ex.Message;
                    MessageBox.Show("Can't delete row in field COMPLETE_FLAG = 'N' ... Please manual delete!" + "\r\n" + ex);
                }
                finally
                {
                    ConnectionTemp.Close();
                    ConnectionTemp.Dispose();

                    GC.Collect();
                    Environment.Exit(0);
                }
            }
            else
            {
                GC.Collect();
                Environment.Exit(0);
            }

        }

        private void btnMaximize_Click(object sender, EventArgs e)
        {
            if (this.WindowState != FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Maximized;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;
            }
        }

        private void btnMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void pictureBoxValidate_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            dataGridView1.EndEdit();
            dataGridView1.Update();

            validate = new Validation(dataGridView1, ConnectionProd, ConnectionTemp, outputPath);
            listBox1.Items.Clear();
            ListBox lst = validate.Verify("VAS");

            foreach (string v in lst.Items)
            {
                listBox1.Items.Add(v);
            }

            if (listBox1.Items.Count > 0)
            {
                indexListbox = validate.IndexDgv;
                listBox1.Show();
            }
            else
            {
                listBox1.Hide();
            }

            dataGridView1.Refresh();

            Cursor.Current = Cursors.Default;
        }

        private void listBox1_Click(object sender, EventArgs e)
        {
            dataGridView1.ClearSelection();
            if (listBox1.SelectedItem != null)
            {
                int selected = listBox1.SelectedIndex;
                dataGridView1.Rows[indexListbox[selected]].Selected = true;
            }
        }
    }
}
