using AutomateMappingTool;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MappingDiscount
{
    public partial class MainHispeed : Form
    {
        #region "Private Field"
        /// <summary>
        /// DBConnection TAPRD
        /// </summary>
        private OracleConnection ConnectionProd;
        /// <summary>
        /// DBConnection TRUREF71
        /// </summary>
        private OracleConnection ConnectionTemp;
        /// <summary>
        /// fileName requirement 
        /// </summary>
        private string FILENAME;
        /// <summary>
        /// fileName vcare description (for P_name)
        /// </summary>
        private string FILENAME_DESC;
        /// <summary>
        /// UserName
        /// </summary>
        private string USER;
        /// <summary>
        /// UR_No
        /// </summary>
        private string UR_NO;
        /// <summary>
        /// CoverHispeedForm
        /// </summary>
        private HomeHispeed homeHispeed;
        /// <summary>
        /// List of All Channel in DB
        /// </summary>
        private List<string> lstChannel;
        /// <summary>
        /// Variable keep script for write text
        /// </summary>
        private string Write_SQL = "";
        /// <summary>
        /// boolean for check write sql or not
        /// </summary>
        private bool isExport;
        /// <summary>
        /// Flag for check process complete when you want to close program
        /// </summary>
        private bool flagClose = true;
        /// <summary>
        /// Collect system log
        /// </summary>
        private string systemLog;

        private string outputPath;
        /// <summary>
        /// Collect value for check duplicate when exporting data
        /// </summary>
        private Dictionary<string, int> TempExport;
        /// <summary>
        /// Collect pName
        /// </summary>
        private Dictionary<string, string> TempPName;

        private ViewResult resultForm;

        private bool isPromo;

        /// <summary>
        /// Keep selected index
        /// </summary>
        List<int> indexDgv;

        string invalidSpeed, invalidContractCode, invalidMKTCode, invalidChannel, invalidUOM,
            invalidEffDate, invalidEndDate, invalidOrder, invalidPName;

        //For move form
        int mov, movX, movY;

        Validation validate;
        ChangeFormat format;
        DgvSettings viewSetting;
        #endregion

        public MainHispeed(HomeHispeed form, OracleConnection con,
            string filename, string fileDesc, string folder, string user, string ur, bool flagPromo)
        {
            InitializeComponent();

            ConnectionProd = con;
            homeHispeed = form;
            FILENAME = filename;
            FILENAME_DESC = fileDesc;
            USER = user;
            UR_NO = ur;
            isPromo = flagPromo;
            outputPath = folder;
        }

        #region "Drop Shadow"
        private const int CS_DropShadow = 0x00020000;

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ClassStyle |= CS_DropShadow;
                return cp;
            }
        }
        #endregion

        #region "Event handle"

        private void MainHispeed_Load(object sender, EventArgs e)
        {
            listBox1.Hide();
            Cursor.Current = Cursors.WaitCursor;

            this.dataGridView1.AllowUserToAddRows = false;
            format = new ChangeFormat();
            viewSetting = new DgvSettings();

            settingDataGridView();

            homeHispeed.Hide();
            
            Cursor.Current = Cursors.Default;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void ViewHispeed_SizeChanged(object sender, EventArgs e)
        {
            int w = this.Size.Width;
            int h = this.Size.Height;

            if (listBox1.Items.Count > 0)
            {
                btnOK.Location = new Point((w - panelHome.Width) + 60, h - 124);
            }
            else
            {
                btnOK.Location = new Point((w - panelHome.Width) + 60, h - 105);
            }

            btnClose.Location = new Point(w - 23, 0);
            btnMinimize.Location = new Point(w - 88, 3);
            btnMaximize.Location = new Point(w - 53, 3);

            dataGridView1.Size = new Size(w - panelHome.Width, h - 280);
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if ((String.IsNullOrEmpty(invalidMKTCode) == false) && (String.IsNullOrEmpty(invalidSpeed) == false)
                && (String.IsNullOrEmpty(invalidContractCode) == false) && (String.IsNullOrEmpty(invalidUOM) == false)
                && (String.IsNullOrEmpty(invalidEndDate) == false) && (String.IsNullOrEmpty(invalidOrder) == false)
                && (String.IsNullOrEmpty(invalidEffDate) == false))
            {
                MessageBox.Show("Please revise data following the red row and click validate again." + "\r\n" +
                    "Please find more detail in dialog box or log file", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if ((String.IsNullOrEmpty(invalidChannel)) == false)
            {
                DialogResult result = MessageBox.Show("The channels are not found in the database." + "\r\n" +
                    "Do you want to continue?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    isExport = false;

                    execute();
                }
            }
            else
            {
                isExport = false;
                execute();
            }
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            mov = 1;
            movX = e.X;
            movY = e.Y;
        }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (mov == 1)
            {
                this.SetDesktopLocation(MousePosition.X - movX, MousePosition.Y - movY);
            }
        }

        private void panel1_MouseUp(object sender, MouseEventArgs e)
        {
            mov = 0;
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

        private void ViewHispeed_FormClosing(object sender, FormClosingEventArgs e)
        {
            string message = "";
            if (flagClose == false)
            {
                e.Cancel = true;
            }
            else
            {
                if (ConnectionTemp != null)
                {
                    try
                    {
                        if (ConnectionTemp.State != ConnectionState.Open)
                        {
                            ConnectionTemp.Open();
                        }

                        string query = "SELECT * FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = 'Hispeed' AND COMPLETE_FLAG = 'N' " +
                            "AND USERNAME = '" + USER + "'";
                        OracleCommand cmd = new OracleCommand(query, ConnectionTemp);
                        OracleDataReader reader = cmd.ExecuteReader();
                        reader.Read();
                        if (reader.HasRows)
                        {
                            reader.Close();

                            message = "Can't delete row COMPLETE_FLAG = 'N' in database." + "\r\n" + "Please manual delete data from database";

                            string qryDel = "DELETE FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = 'Hispeed' AND COMPLETE_FLAG = 'N' " +
                                "AND USERNAME = '" + USER + "'";
                            OracleCommand command = new OracleCommand(qryDel, ConnectionTemp);
                            command.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        string log = ex.Message;
                        MessageBox.Show(message + "\r\n" + "Please find more details as log file.");

                        System.IO.FileInfo fInfo = new System.IO.FileInfo(FILENAME);
                        string strFilePath = fInfo.DirectoryName + "\\System Error_" + UR_NO + ".txt";

                        using (StreamWriter writer = new StreamWriter(strFilePath, true))
                        {
                            writer.Write(log);
                        }

                        ConnectionTemp.Close();

                        Application.Exit();
                    }
                }
            }

            if (ConnectionProd.State == ConnectionState.Open)
            {
                ConnectionProd.Close();
                ConnectionProd.Dispose();
            }

            if (ConnectionTemp.State == ConnectionState.Open)
            {
                ConnectionTemp.Close();
                ConnectionTemp.Dispose();
            }

        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            this.Close();
            homeHispeed.Show();
        }

        private void pictureBoxExport_MouseHover(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.Hand;
        }

        private void pictureBoxExport_Click(object sender, EventArgs e)
        {
            if ((String.IsNullOrEmpty(invalidMKTCode) == false) && (String.IsNullOrEmpty(invalidSpeed) == false)
                && (String.IsNullOrEmpty(invalidContractCode) == false) && (String.IsNullOrEmpty(invalidUOM) == false)
                && (String.IsNullOrEmpty(invalidEndDate) == false) && (String.IsNullOrEmpty(invalidOrder) == false)
                && (String.IsNullOrEmpty(invalidEffDate) == false))
            {
                MessageBox.Show("Please revise data following red row and validate data again." + "\r\n" + "You can see more detail in log file.",
                  "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (String.IsNullOrEmpty(invalidChannel) == false)
            {
                DialogResult result = MessageBox.Show("The channels are not found in the database." + "\r\n" + "Do you want to continue?", "Confirmation"
                    , MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    isExport = true;
                    execute();
                }
            }
            else
            {
                isExport = true;
                execute();
            }
        }

        private void pictureBoxValidate_MouseHover(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.Hand;
        }

        private void pictureBoxValidate_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            dataGridView1.EndEdit();
            dataGridView1.Update();

            Verify_Data();
            dataGridView1.Refresh();

            Cursor.Current = Cursors.Default;
        }

        private void pictureBox1_Click_1(object sender, EventArgs e)
        {
            homeHispeed.Show();
            this.Hide();
        }

        private void pictureBox1_MouseHover_1(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.Hand;
        }
        #endregion

        /// <summary>
        /// Show data from Sheet Campaign Mapping
        /// </summary>
        private void ViewCampaignMapping()
        {

        }

        /// <summary>
        /// Show data from file excel package mapping (requirement)
        /// </summary>
        private void settingDataGridView()
        {
            try
            {
                TempPName = new Dictionary<string, string>();
                List<string> lstHeader = new List<string>();

                //CONNECTION71 = new OracleConnection(); 
                //connString71 = "Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.19.193.20)(PORT = 1560))" +
                //       "(CONNECT_DATA = (SID = TEST03)));User Id= TRUREF71; Password= TRUREF71;";

                //CONNECTION71.ConnectionString = connString71;
                ConnectionTemp = ConnectionProd;
                //CONNECTION71.Open();

                if (isPromo)
                {
                    //set header view
                    lstHeader.Add("MKT_Code");
                    lstHeader.Add("Speed");
                    lstHeader.Add("Sub_Profile");
                    lstHeader.Add("Price");
                    lstHeader.Add("OrderType");
                    lstHeader.Add("Channel");
                    lstHeader.Add("ModemType");
                    lstHeader.Add("DocsisType");
                    lstHeader.Add("Effective");
                    lstHeader.Add("Expire");
                    lstHeader.Add("Entry_Code");
                    lstHeader.Add("Ins_Code");

                    //Load data from file VCARE (Get P_NAME)
                    if (FILENAME_DESC != "")
                    {
                        if (System.IO.File.Exists(FILENAME_DESC))
                        {
                            string connString = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;""", FILENAME_DESC);
                            string query = string.Format("select * from [VCARE-MKT$B:D]", connString);
                            OleDbDataAdapter dtAdapter = new OleDbDataAdapter(query, connString);
                            DataSet ds = new DataSet();
                            dtAdapter.Fill(ds);
                            DataTable dt = ds.Tables[0];
                            //Remove head
                            for (int i = 0; i <= 3; i++)
                            {
                                dt.Rows.RemoveAt(i);
                                dt.AcceptChanges();
                            }

                            foreach (DataRow dr in dt.Rows)
                            {
                                string mkt = dr[1].ToString();
                                string name = dr[2].ToString();

                                if (mkt != "" && name != "")
                                {
                                    TempPName.Add(mkt, name);
                                }
                            }
                        }
                    }

                    viewSetting.SetDgv(dataGridView1, FILENAME, "HiSpeed Promotion$B2:N", lstHeader);

                    toolStripLabel1.Text = "Rows : " + dataGridView1.RowCount.ToString();

                    lstChannel = GetChannelFromDB();
                    Verify_Data();
                }
                else
                {
                    //set header view
                    lstHeader.Add("Request_Type");
                    lstHeader.Add("Campaign_Name");
                    lstHeader.Add("TOL_Package");
                    lstHeader.Add("TOL_Discount");
                    lstHeader.Add("TVS_Package");
                    lstHeader.Add("TVS_Discount");

                    viewSetting.SetDgv(dataGridView1, FILENAME, "Campaign Mapping$B:G", lstHeader);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot read data from file input." + "\r\n" +
                    "Please check the file format" + "\r\n" + "Error Detail : " + ex.Message);
            }

        }

        /// <summary>
        /// Get channel from DB
        /// </summary>
        /// <returns>List of channel in DB</returns>
        private List<string> GetChannelFromDB()
        {
            List<string> lstChannelFromDB = new List<string>();

            try
            {
                //Get all channel in DB
                string query = "SELECT DISTINCT(SALE_CHANNEL) FROM HISPEED_CHANNEL_PROMOTION";
                OracleCommand cmd = new OracleCommand(query, ConnectionProd);
                OracleDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    lstChannelFromDB.Add(reader["SALE_CHANNEL"].ToString());
                }

                reader.Close();
                reader.Dispose();
            }
            catch (Exception)
            {
                string msg = "Cannot get sale channel from database. Please click validate file again or check internet connection.";

                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                lstChannelFromDB = new List<string>();
            }

            return lstChannelFromDB;
        }

        /// <summary>
        /// Get PName (Description of package) from file excel
        /// </summary>
        private string GetPName(string mkt)
        {
            string pName = "";

            //From DB
            //string queryPName = "SELECT X.ATTRIB_04 MKT, S.NAME FROM SIEBEL.S_PROD_INT S , SIEBEL.S_PROD_INT_X  X WHERE S.ROW_ID " +
            //    " = X.ROW_ID AND X.ATTRIB_04 = '" + mkt + "'";

            string queryPName = "select * from discount_criteria_province where dp_type = '" + mkt + "'";

            OracleCommand command = new OracleCommand(queryPName, ConnectionProd);
            OracleDataReader reader = command.ExecuteReader();
            reader.Read();
            if (reader.HasRows)
            {
                pName = reader["NAME"].ToString();
                reader.Close();
            }
            else
            {
                pName = mkt;
            }

            return pName;
        }

        /// <summary>
        /// Process for verify data from file excel
        /// </summary>
        private void Verify_Data()
        {
            ChangeFormat changeFormateDT = new ChangeFormat();
            indexDgv = new List<int>();
            listBox1.Items.Clear();
            listBox1.Hide();

            string ValidateLog = "", invalidSpeed = "", invalidContractCode = "", invalidMKTCode = "",
             invalidChannel = "", invalidUOM = "", invalidEndDate = "", invalidEffDate = "";
            OracleDataReader reader = null;
            OracleCommand cmd = null;

            if (lstChannel == null)
            {
                lstChannel = GetChannelFromDB();
            }
            else
            {
                try
                {
                    flagClose = true;
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        //Clear selection
                        dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.Empty;
                        dataGridView1.Rows[i].Cells[1].Style.BackColor = Color.Empty;
                        dataGridView1.Rows[i].Cells[4].Style.BackColor = Color.Empty;
                        dataGridView1.Rows[i].Cells[5].Style.BackColor = Color.Empty;
                        dataGridView1.Rows[i].Cells[8].Style.BackColor = Color.Empty;
                        dataGridView1.Rows[i].Cells[9].Style.BackColor = Color.Empty;
                        dataGridView1.Rows[i].Cells[10].Style.BackColor = Color.Empty;
                        dataGridView1.Rows[i].Cells[11].Style.BackColor = Color.Empty;

                        dataGridView1.ClearSelection();

                        string mkt = dataGridView1.Rows[i].Cells[0].Value.ToString();
                        if (mkt.Contains(">>"))
                        {
                            mkt = mkt.Substring(0, mkt.IndexOf(">")).Trim();
                            dataGridView1.Rows[i].Cells[0].Value = mkt;
                        }

                        //-----Check speed ----
                        //Speed from MKT Code
                        string[] spMkt = mkt.Split('-');
                        string mktCode = spMkt[0].Trim();
                        string speedMkt = spMkt[1].Trim();

                        //Download speed -- > before '/'
                        string speed = dataGridView1.Rows[i].Cells[1].Value.ToString();
                        if (speed.Contains('/'))
                        {
                            string[] spSpeed = speed.Split('/');
                            string downSp = spSpeed[0].Trim();
                            string upSpeed = spSpeed[1].Trim();

                            //Keep only numeric
                            string downSpeed = Regex.Replace(downSp, "[^0-9]", "");

                            if (speedMkt != downSpeed)
                            {
                                if (downSp.Contains("G"))
                                {
                                    string cvMkt = ((Convert.ToInt32(speedMkt)) / 1000).ToString();

                                    if (cvMkt != downSpeed)
                                    {
                                        string msg = mkt + " mismatch download speed " + downSp;
                                        invalidSpeed += "\r\n" + "      " + msg;
                                        listBox1.Items.Add(msg);
                                        indexDgv.Add(i);

                                        dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.Red;
                                        dataGridView1.Rows[i].Cells[1].Style.BackColor = Color.Red;
                                    }
                                }
                                else
                                {
                                    string msg = mkt + " mismatch download speed " + downSp;
                                    invalidSpeed += "\r\n" + "      " + msg;
                                    listBox1.Items.Add(msg);
                                    indexDgv.Add(i);

                                    dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.Red;
                                    dataGridView1.Rows[i].Cells[1].Style.BackColor = Color.Red;
                                }
                            }
                        }
                        else
                        {
                            //invalidSpeed += "\r\n" + "      " + msg;
                            string msg = "Invalid speed " + speed + " --> The speed format must consist of '/'";
                            listBox1.Items.Add(msg);
                            indexDgv.Add(i);

                            dataGridView1.Rows[i].Cells[1].Style.BackColor = Color.Red;
                        }

                        //--------- Check contract code -------------
                        string entry = dataGridView1.Rows[i].Cells[10].Value.ToString().Trim();
                        string install = dataGridView1.Rows[i].Cells[11].Value.ToString().Trim();

                        string queryEnt = "SELECT * FROM TRUE9_BPT_CONTRACT WHERE ENTRY = '" + entry + "'";
                        string queryIns = "SELECT * FROM TRUE9_BPT_CONTRACT WHERE INSTALL = '" + install + "'";

                        //Entry Code
                        cmd = new OracleCommand(queryEnt, ConnectionTemp);
                        reader = cmd.ExecuteReader();
                        if (reader.HasRows == false)
                        {
                            string msg = "Entry Code " + entry + " of " + mkt + " ,Not found in table >> TRUE9_BPT_CONTRACT";
                            invalidContractCode += "\r\n" + "      " + msg;
                            listBox1.Items.Add(msg);
                            indexDgv.Add(i);

                            dataGridView1.Rows[i].Cells[10].Style.BackColor = Color.Red;

                            reader.Close();
                        }

                        //Install code
                        cmd = new OracleCommand(queryIns, ConnectionTemp);
                        reader = cmd.ExecuteReader();
                        if (reader.HasRows == false)
                        {
                            string msg = "Install Code " + install + " of " + mkt + " ,Not found in table >> TRUE9_BPT_CONTRACT";
                            invalidContractCode += "\r\n" + "      " + msg;
                            listBox1.Items.Add(msg);
                            indexDgv.Add(i);

                            dataGridView1.Rows[i].Cells[11].Style.BackColor = Color.Red;

                            reader.Close();
                        }

                        //------- Check Product Type -------

                        string prefixMKT;
                        if (mktCode.StartsWith("TRL"))
                        {
                            prefixMKT = mktCode.Substring(0, 5);
                        }
                        else
                        {
                            prefixMKT = mktCode.Substring(0, 2);
                        }

                        string query = "SELECT * FROM TRUE9_BPT_HISPEED_PRODTYPE WHERE MKT = '" + prefixMKT + "'";
                        cmd = new OracleCommand(query, ConnectionTemp);
                        reader = cmd.ExecuteReader();

                        if (reader.HasRows == false)
                        {
                            string msg = "The prefix " + prefixMKT + " of " + mkt + " ,Not found in table >> TRUE9_BPT_HISPEED_PRODTYPE";
                            invalidMKTCode += "\r\n" + "      " + msg;
                            listBox1.Items.Add(msg);
                            indexDgv.Add(i);

                            dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.Red;

                            reader.Close();
                        }

                        //-----------Check channel ----------
                        string channel = dataGridView1.Rows[i].Cells[5].Value.ToString().Trim();

                        if (String.IsNullOrEmpty(channel) == false)
                        {
                            //Check channel compare with DB
                            List<string> lstInvaid = CheckChannel(channel);

                            //Check conflict channel
                            if (channel.Contains(","))
                            {
                                string[] lst = channel.Split(',');
                                foreach (string val in lst)
                                {
                                    string upperCh = val.ToUpper();
                                    if (upperCh == "ALL" || upperCh == "DEFAULT")
                                    {
                                        //conflict channel
                                        string msg = "There are channel 'ALL' included with other channel in MKT Code " + mkt;
                                        invalidChannel += "\r\n" + "      " + msg;
                                        listBox1.Items.Add(msg);
                                        indexDgv.Add(i);

                                        dataGridView1.Rows[i].Cells[5].Style.BackColor = Color.Red;

                                        break;
                                    }

                                }
                            }

                            if (lstInvaid.Count > 0)
                            {
                                string invalid = "";
                                dataGridView1.Rows[i].Cells[5].Style.BackColor = Color.Red;
                                foreach (string val in lstInvaid)
                                {
                                    invalid += val + ", ";
                                }

                                invalid = invalid.Substring(0, invalid.Length - 2);

                                string msg = "Not found channel : " + invalid + " in database";
                                invalidChannel += "\r\n" + "      " + msg;
                                listBox1.Items.Add(msg);
                                indexDgv.Add(i);
                            }
                        }
                        else
                        {
                            string endDateF = dataGridView1.Rows[i].Cells[9].Value.ToString();
                            if (endDateF == "")
                            {
                                string msg = "MKT Code : " + mkt + " >> The channel is empty, Expire date cannot be null  [End sales]";
                                invalidEndDate += "\r\n" + "      " + msg;
                                listBox1.Items.Add(msg);
                                indexDgv.Add(i);

                                dataGridView1.Rows[i].Cells[9].Style.BackColor = Color.Red;
                            }
                        }

                        //Check order type
                        string order = dataGridView1.Rows[i].Cells[4].Value.ToString();
                        if (order.Contains("/"))
                        {
                            string msg = "MKT Code : " + mkt + " Order type contain characters '/'";
                            invalidOrder += "\r\n" + "      " + msg;
                            listBox1.Items.Add(msg);
                            indexDgv.Add(i);

                            dataGridView1.Rows[i].Cells[4].Style.BackColor = Color.Red;
                        }

                        //Get P-Name
                        string name = GetPName(mkt);
                        if (name == mkt)
                        {
                            string msg = mkt + " >> not found P-Name!!";

                            dataGridView1.Rows[i].Cells[0].Value = msg;
                            dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.Yellow;

                            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

                            if (TempPName.ContainsKey(mkt) == false)
                            {
                                TempPName.Add(mkt, mkt);
                            }
                        }
                        else
                        {
                            if (TempPName.ContainsKey(mkt) == false)
                            {
                                TempPName.Add(mkt, name);
                            }

                        }

                        //Check Date
                        string dtEff = dataGridView1.Rows[i].Cells[8].Value.ToString();
                        string dtEx = dataGridView1.Rows[i].Cells[9].Value.ToString();

                        if (String.IsNullOrEmpty(dtEff) || dtEff == "-")
                        {
                            string msg = "MKT Code : " + mkt + " >> Effective date cannot be null";
                            invalidEffDate += "\r\n" + "      " + msg;
                            listBox1.Items.Add(msg);
                            indexDgv.Add(i);

                            dataGridView1.Rows[i].Cells[8].Style.BackColor = Color.Red;
                        }
                        else
                        {
                            string effective = changeFormateDT.FormatDate(dtEff);
                            if (effective == "Invalid")
                            {
                                string msg = "MKT Code : " + mkt + " >> Effective date cannot be null";
                                invalidEffDate += "\r\n" + "      " + msg;
                                listBox1.Items.Add(msg);
                                indexDgv.Add(i);

                                dataGridView1.Rows[i].Cells[8].Style.BackColor = Color.Red;
                            }
                            else
                            {
                                DateTime date = Convert.ToDateTime(effective);
                                dataGridView1.Rows[i].Cells[8].Value = effective;

                                if (date < DateTime.Now.Date)
                                {
                                    string msg = "MKT Code : " + mkt + " >> Effective date cannot be less than sysdate";
                                    invalidEffDate += "\r\n" + "      " + msg;
                                    listBox1.Items.Add(msg);
                                    indexDgv.Add(i);

                                    dataGridView1.Rows[i].Cells[8].Style.BackColor = Color.Red;
                                }
                            }
                        }

                        string expire = changeFormateDT.FormatDate(dtEx);

                        if (String.IsNullOrEmpty(expire) == false)
                        {
                            if (expire == "Invalid")
                            {
                                string msg = "MKT Code : " + mkt + " >> The channel is empty, Expire date cannot be null  [End sales]";
                                invalidEndDate += "\r\n" + "      " + msg;
                                listBox1.Items.Add(msg);
                                indexDgv.Add(i);

                                dataGridView1.Rows[i].Cells[9].Style.BackColor = Color.Red;
                            }
                            else
                            {
                                dataGridView1.Rows[i].Cells[9].Value = expire;

                                DateTime date = Convert.ToDateTime(expire);
                                if (date < DateTime.Now.Date)
                                {
                                    string msg = "MKT Code : " + mkt + " >> Expire date cannot be less than sysdate";
                                    invalidEndDate += "\r\n" + "      " + msg;
                                    listBox1.Items.Add(msg);
                                    indexDgv.Add(i);

                                    dataGridView1.Rows[i].Cells[9].Style.BackColor = Color.Red;
                                }
                            }
                        }

                    }
                }
                catch (Exception ex)
                {
                    ValidateLog += "[System Error][" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "] " + ex.ToString() + "\r\n";
                }

                if (String.IsNullOrEmpty(invalidSpeed) == false)
                {
                    ValidateLog += "[Speed][" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "] " + invalidSpeed + "\r\n";
                }

                if (String.IsNullOrEmpty(invalidChannel) == false)
                {
                    ValidateLog += "[Channel][" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "] " + invalidChannel + "\r\n";
                }

                if (String.IsNullOrEmpty(invalidContractCode) == false)
                {
                    ValidateLog += "[Contract Code][" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "] " + invalidContractCode + "\r\n";
                }

                if (String.IsNullOrEmpty(invalidMKTCode) == false)
                {
                    ValidateLog += "[Product Type][" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "] " + invalidMKTCode + "\r\n";
                }
                if (String.IsNullOrEmpty(invalidUOM) == false)
                {
                    ValidateLog += "[UOM][" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "] " + invalidUOM + "\r\n";
                }
                if (String.IsNullOrEmpty(invalidEffDate) == false)
                {
                    ValidateLog += "[Effective][" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "] " + invalidEffDate + "\r\n";
                }
                if (String.IsNullOrEmpty(invalidEndDate) == false)
                {
                    ValidateLog += "[End Sale][" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "] " + invalidEndDate + "\r\n";
                }
                if (String.IsNullOrEmpty(invalidOrder) == false)
                {
                    ValidateLog += "[OrderType][" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "] " + invalidOrder + "\r\n";
                }

                if (String.IsNullOrEmpty(ValidateLog) == false)
                {
                    System.IO.FileInfo fInfo = new System.IO.FileInfo(FILENAME);
                    string strFilePath = fInfo.DirectoryName + "\\Log_" + UR_NO + ".txt";

                    using (StreamWriter writer = new StreamWriter(strFilePath, true))
                    {
                        writer.Write(ValidateLog);
                    }

                }
            }

            if (listBox1.Items.Count > 0)
            {
                listBox1.Show();
            }

        }

        /// <summary>
        /// Check channel from file compare with channel in database
        /// </summary>
        /// <param name="channel"></param>
        /// <returns></returns>
        private List<string> CheckChannel(string channel)
        {
            List<string> lstInvaid = new List<string>();

            if (channel.Contains(","))
            {
                string[] lstCh = channel.Split(',');

                foreach (string val in lstCh)
                {
                    val.Trim();

                    if (val.Equals("ALL", StringComparison.OrdinalIgnoreCase) == false)
                    {
                        if (lstChannel.Contains(val) == false)
                        {
                            lstInvaid.Add(val);
                        }
                    }

                }
            }
            else
            {
                if (channel.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    channel = "DEFAULT";
                }

                if (lstChannel.Contains(channel) == false)
                {
                    lstInvaid.Add(channel);
                }
            }

            return lstInvaid;
        }


        /// <summary>
        /// Main process
        /// </summary>
        private void execute()
        {
            Cursor.Current = Cursors.WaitCursor;

            ReserveID reserve = new ReserveID();

            TempExport = new Dictionary<string, int>();
            flagClose = false;
            string[] lstOrder = null;
            int id = 0;
            systemLog = "";
            int minID = 0;
            int maxID = 0;
            string lstID = "";

            bool isInProcess = reserve.checkStatus(ConnectionTemp, "Hispeed", USER, UR_NO);

            if (isInProcess == false)
            {
                try
                {
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        //Get MKT code from file
                        string mktCode = dataGridView1.Rows[i].Cells[0].Value.ToString();

                        //split >> in field mkt , Don't have p-name in file or DB
                        if (mktCode.Contains(">>"))
                        {
                            mktCode = mktCode.Substring(0, mktCode.IndexOf(">")).Trim();
                            dataGridView1.Rows[i].Cells[0].Value = mktCode;
                        }

                        string[] mkt = mktCode.Split('-');
                        string pCodeF = mkt[0].Trim();

                        //Get speed from file
                        string[] speed = dataGridView1.Rows[i].Cells[1].Value.ToString().Split('/');
                        string download = Regex.Replace(speed[0], "[^0-9]", "");
                        string upload = Regex.Replace(speed[1], "[^0-9]", "");

                        int num;
                        string downUOM;
                        string upUOM;
                        //Get UOM download speed
                        if (int.TryParse(speed[0], out num))
                        {
                            downUOM = Regex.Replace(speed[1], "[0-9]", "");
                        }
                        else
                        {
                            downUOM = Regex.Replace(speed[0], "[0-9]", "");
                        }
                        //Get UOM upload speed
                        if (int.TryParse(speed[1], out num))
                        {
                            upUOM = Regex.Replace(speed[0], "[0-9]", "");
                        }
                        else
                        {
                            upUOM = Regex.Replace(speed[1], "[0-9]", "");
                        }

                        //Get order type from file
                        string orderType = dataGridView1.Rows[i].Cells[4].Value.ToString();
                        if (orderType.Contains(","))
                        {
                            lstOrder = new string[2];
                            lstOrder = (dataGridView1.Rows[i].Cells[4].Value.ToString()).Split(',');
                        }
                        else
                        {
                            lstOrder = new string[1];
                            lstOrder[0] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                        }

                        //Get channel from file
                        string channelF = dataGridView1.Rows[i].Cells[5].Value.ToString();

                        //Get Sub-Profile from file
                        string subProfile = dataGridView1.Rows[i].Cells[2].Value.ToString();

                        if (subProfile.StartsWith("STL"))
                        {
                            subProfile = "N";
                        }

                        //Get price from file
                        double price = Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value);

                        //Get start & end date from file
                        string startDateF = format.FormatDate(dataGridView1.Rows[i].Cells[8].Value.ToString());
                        string endDateF = format.FormatDate(dataGridView1.Rows[i].Cells[9].Value.ToString());

                        //Get modem & docsis type
                        string modem = dataGridView1.Rows[i].Cells[6].Value.ToString();
                        string docsisType = dataGridView1.Rows[i].Cells[7].Value.ToString();

                        //Get contractCode from file
                        string entryCode = dataGridView1.Rows[i].Cells[10].Value.ToString();
                        string installCode = dataGridView1.Rows[i].Cells[11].Value.ToString();

                        //Check speed in table hispeed_speed
                        int speedID = GetSpeedID(speed[0], downUOM);

                        if (speedID != -1)
                        {
                            //Searching P_ID from hispeed_promotion
                            for (int index = 0; index < lstOrder.Length; index++)
                            {
                                string query = "SELECT P.P_ID, P.P_CODE, P.P_NAME, P.ORDER_TYPE, P.END_DATE, P.STATUS, S.SPEED_ID, " +
                                     "S.PRICE , S.ACTIVE_PRICE FROM HISPEED_PROMOTION P, HISPEED_SPEED_PROMOTION S WHERE P_CODE = '" +
                                     pCodeF + "' AND SPEED_ID = " + speedID + " AND P.P_ID = S.P_ID AND ORDER_TYPE = '" +
                                     lstOrder[index] + " AND STATUS IN('Active', 'Pending')";

                                OracleCommand cmd = new OracleCommand(query, ConnectionProd);
                                OracleDataReader reader = cmd.ExecuteReader();
                                List<string> lstOrderDB = new List<string>();

                                //Existing
                                if (reader.HasRows)
                                {
                                    reader.Read();

                                    string sql = "";
                                    id = Convert.ToInt32(reader["P_ID"]);
                                    string status = reader["STATUS"].ToString();
                                    string endDateDB = reader["END_DATE"].ToString();

                                    if (status == "Pending")
                                    {
                                        if (String.IsNullOrEmpty(endDateF))
                                        {
                                            sql = "UPDATE HISPEED_PROMOTION SET STATUS = 'Active' , END_DATE = null WHERE P_ID = " + id;
                                        }
                                        else
                                        {
                                            DateTime dateF = Convert.ToDateTime(endDateF);
                                            DateTime dateDb = Convert.ToDateTime(endDateDB);

                                            if (dateF >= DateTime.Now || dateF > dateDb)
                                            {
                                                sql = "UPDATE HISPEED_PROMOTION SET STATUS = 'Active' , END_DATE = TO_DATE('" + endDateF + "','dd/MM/yyyy') WHERE P_ID = " + id;
                                            }
                                        }

                                        lstID += id + ", ";
                                    }

                                    if (isExport)
                                    {
                                        if (sql != "")
                                        {
                                            Write_SQL += sql + "\r\n";
                                        }
                                    }
                                    else
                                    {
                                        if (sql != "")
                                        {
                                            Write_SQL += sql + "\r\n";

                                            //update status & end date
                                            OracleCommand command = ConnectionProd.CreateCommand();
                                            OracleTransaction transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted); ;
                                            // Assign transaction object for a pending local transaction
                                            command.Transaction = transaction;

                                            try
                                            {
                                                command.CommandText = sql;
                                                command.CommandType = CommandType.Text;

                                                command.ExecuteNonQuery();
                                                transaction.Commit();
                                            }
                                            catch (Exception e)
                                            {
                                                transaction.Rollback();
                                                systemLog += DateTime.Now.ToString() + "  Cannot update status of P_ID : " + id + " into table HISPEED_PROMOTION" + "Detail : " + e.Message.ToString() + "\r\n";
                                            }
                                        }
                                    }

                                    //Call method for check data in table hispeed_channel_promotion
                                    if (String.IsNullOrEmpty(channelF))
                                    {
                                        Channel_Promotion(id, channelF, startDateF, endDateF);
                                    }

                                    //Check price in table hispeed_speed_promotion
                                    HispeedSpeedPromotion(id, price, download, upload, upUOM, speedID, modem, docsisType);

                                }
                                else
                                {
                                    //create new id
                                    if (minID == 0)
                                    {
                                        minID = reserve.reserveID(ConnectionTemp, ConnectionProd, "Hispeed", USER, UR_NO);
                                        maxID = minID;
                                    }
                                    else
                                    {
                                        maxID = maxID + 1;
                                    }

                                    lstID += maxID + ", ";

                                    NewHispeedPromotion(maxID, mktCode, lstOrder[index], subProfile, price, channelF, modem, docsisType,
                                        entryCode, installCode, startDateF, endDateF, download, upload, upUOM);

                                    if (TempExport.ContainsKey(mktCode + "," + lstOrder[index]) == false)
                                    {
                                        TempExport.Add(mktCode + "," + lstOrder[index], maxID);
                                    }
                                }
                            }
                        }
                    }

                    if (isExport)
                    {
                        if (Write_SQL == "")
                        {
                            MessageBox.Show("The data already exists in the database." + "\r\n" +
                                "The program will not insert/update the data in database so you cannot export script.", "Data already exists",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            string fileName = "Export_Script_" + UR_NO + ".txt";
                            string[] files = Directory.GetFiles(outputPath);
                            int count = files.Count(file => { return file.Contains(fileName); });
                            string newFileName = fileName;

                            //newFileName = (count == 0) ? fileName : String.Format("{0} ({1}).txt", fileName, count);
                            if (count > 1)
                            {
                                DialogResult result = MessageBox.Show("There is already a file with the same name in this location" + "\r\n" +
                                    "Do you want to replace it?", "Replace File", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                                if (result == DialogResult.Yes)
                                {
                                    newFileName = String.Format("{0} ({1}).txt", fileName, count);

                                }
                            }

                            string newFilePath = outputPath + "\\" + newFileName;
                            using (StreamWriter writer = new StreamWriter(newFilePath, true))
                            {
                                writer.Write(Write_SQL);
                            }

                            MessageBox.Show("Your script was exported successfully");

                        }
                    }
                    else
                    {
                        lstID = lstID.Substring(0, lstID.Length - 2);

                        if (systemLog != "")
                        {
                            string strFilePath = outputPath + "\\SystemLog_" + UR_NO + ".txt";
                            using (StreamWriter writer = new StreamWriter(strFilePath, true))
                            {
                                writer.Write(systemLog);
                            }
                        }

                        resultForm = new ViewResult(this, null, ConnectionProd, ConnectionTemp, UR_NO, lstID, null, USER, "", outputPath);

                        resultForm.ExportImp(outputPath);

                        reserve.updateID(ConnectionTemp, minID.ToString(), maxID.ToString(), "Hispeed", USER, UR_NO);

                        MessageBox.Show("Already insert/update your data into database");
                    }

                    flagClose = true;
                }
                catch (Exception ex)
                {
                    flagClose = true;
                    MessageBox.Show("Failed..." + "\r\n" + ex.Message);
                }
                finally
                {
                    Cursor.Current = Cursors.Default;
                    Application.UseWaitCursor = false;

                    dataGridView1.ClearSelection();
                }
            }
        }

        private void Channel_Promotion(int id, string channel, string startDateF, string endDateF)
        {
            OracleCommand command = ConnectionProd.CreateCommand();
            OracleTransaction transaction = null;

            List<string> distinctChF = new List<string>();
            Dictionary<string, string> lstChannelDB_ID = new Dictionary<string, string>();

            if (channel.Contains(","))
            {
                string[] lstCh = channel.Split(',');
                // distinct channel from file
                List<string> distinct = lstCh.Distinct().ToList();
                distinctChF = distinct;
            }
            else
            {
                if (channel.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    channel = "DEFAULT";
                }

                distinctChF.Add(channel);
            }

            try
            {
                //Get channel from DB
                string query = "SELECT * FROM HISPEED_CHANNEL_PROMOTION WHERE P_ID = " + id;

                OracleCommand cmd = new OracleCommand(query, ConnectionProd);
                OracleDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    lstChannelDB_ID.Add(reader["SALE_CHANNEL"].ToString(), (reader["START_DATE"].ToString() + "," + reader["END_DATE"].ToString()));
                }
                ///**** check if end date (DB) = null , what is the return value
                reader.Close();
            }
            catch (Exception)
            { }

            foreach (string chF in distinctChF)
            {
                //Existing channel
                if (lstChannelDB_ID.Keys.Contains(chF))
                {
                    string[] date = lstChannelDB_ID[chF].Split(',');
                    string startCh = date[0];
                    string endCh = date[1];

                    DateTime startF = Convert.ToDateTime(startDateF);
                    DateTime startDB = Convert.ToDateTime(startCh);

                    if (startF > startDB)
                    {
                        string cmdTxt = "UPDATE HISPEED_CHANNEL_PROMOTION SET START_DATE = TO_DATE('" + startDateF + "','dd/MM/yyyy') WHERE P_ID = " + id +
                                        " AND SALE_CHANNEL = '" + chF + "'";

                        if (isExport)
                        {
                            if (cmdTxt != "")
                            {
                                Write_SQL += cmdTxt + "\r\n";
                            }
                        }
                        else
                        {
                            Write_SQL += cmdTxt + "\r\n";

                            //update start date
                            using (transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted))
                            {
                                try
                                {
                                    command.CommandText = cmdTxt;
                                    //command = new OracleCommand(cmdTxt, ConnectionProd);
                                    command.ExecuteNonQuery();

                                    transaction.Commit();
                                }
                                catch (Exception ex)
                                {
                                    transaction.Rollback();

                                    systemLog += DateTime.Now.ToString() + "  Cannot update end date of channel '" + chF + "' to P_ID : " + id + " Detail : " +
                                        ex.Message.ToString() + "\r\n";
                                }
                            }
                        }

                    }

                    if (endDateF != endCh)
                    {
                        string cmdTxt = "UPDATE HISPEED_CHANNEL_PROMOTION SET END_DATE = TO_DATE('" + endDateF + "','dd/MM/yyyy') WHERE P_ID = " + id +
                                        " AND SALE_CHANNEL = '" + chF + "'";

                        if (isExport)
                        {
                            if (cmdTxt != "")
                            {
                                Write_SQL += cmdTxt + "\r\n";
                            }
                        }
                        else
                        {
                            Write_SQL += cmdTxt + "\r\n";

                            //update end date
                            using (transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted))
                            {
                                try
                                {
                                    //command = new OracleCommand(cmdTxt, ConnectionProd, transaction);
                                    command.CommandText = cmdTxt;
                                    command.ExecuteNonQuery();

                                    transaction.Commit();
                                }
                                catch (Exception ex)
                                {
                                    transaction.Rollback();

                                    systemLog += DateTime.Now.ToString() + "  Cannot update end date of channel '" + chF + "' to P_ID : " + id + " Detail : " +
                                        ex.Message.ToString() + "\r\n";
                                }
                            }
                        }
                    }
                }
                else
                {
                    //insert new channel
                    string cmdTxt = "INSERT INTO HISPEED_CHANNEL_PROMOTION VALUES(" + id + ", '" + chF + "', TO_DATE('" + startDateF + "','dd/MM/yyyy'), " +
                        "TO_DATE('" + endDateF + "','dd/MM/yyyy'), 'S')";

                    if (isExport)
                    {
                        if (cmdTxt != "")
                        {
                            Write_SQL += cmdTxt + "\r\n";
                        }
                    }
                    else
                    {
                        Write_SQL += cmdTxt + "\r\n";

                        using (transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted))
                        {
                            try
                            {
                                //command = new OracleCommand(cmdTxt, ConnectionProd, transaction);
                                command.CommandText = cmdTxt;
                                command.ExecuteNonQuery();

                                transaction.Commit();
                            }
                            catch (Exception ex)
                            {
                                transaction.Rollback();

                                systemLog += DateTime.Now.ToString() + "  Cannot add channel : " + chF + " to P_ID : " + id + " Detail : " +
                                        ex.Message.ToString() + "\r\n";
                            }

                        }
                    }
                }
            }
        }

        private void HispeedSpeedPromotion(int id, double priceF, string download, string upload, string upUOM, int speedID, string modem, string Tdocsis)
        {
            OracleCommand command = ConnectionProd.CreateCommand();
            OracleTransaction transaction = null;

            //Get channel from DB
            string query = "SELECT * FROM HISPEED_SPEED_PROMOTION WHERE P_ID = " + id;

            OracleCommand cmd = new OracleCommand(query, ConnectionProd);
            OracleDataReader reader = cmd.ExecuteReader();
            reader.Read();
            if (reader.HasRows)
            {
                double rateDB = Convert.ToDouble(reader["PRICE"].ToString());

                if (rateDB != priceF)
                {
                    string cmdTxt = "UPDATE HISPEED_SPEED_PROMOTION SET PRICE = " + priceF + " WHERE P_ID = " + id;

                    if (isExport)
                    {
                        if (cmdTxt != "")
                        {
                            Write_SQL += cmdTxt + "\r\n" + "\r\n";
                        }
                    }
                    else
                    {
                        Write_SQL += cmdTxt + "\r\n" + "\r\n";

                        DialogResult dialog = MessageBox.Show("Do you want to revise rate in table Hispeed_Speed_Promotion?", "Rate Changed",
                            MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                        if (dialog == DialogResult.OK)
                        {
                            //update rate
                            using (transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted))
                            {
                                try
                                {
                                    //command = new OracleCommand(cmdTxt, ConnectionProd, transaction);
                                    command.CommandText = cmdTxt;
                                    command.ExecuteNonQuery();

                                    transaction.Commit();
                                }
                                catch (Exception ex)
                                {
                                    transaction.Rollback();

                                    systemLog += DateTime.Now.ToString() + "  Cannot update the new price of P_ID : " + id + " Detail : " +
                                            ex.Message.ToString() + "\r\n";
                                }

                            }
                        }
                    }

                }
            }
            else
            {
                //insert new
                string upload_K = ConvertSpeed(upload, upUOM);

                string cmdTxt = "INSERT INTO HISPEED_SPEED_PROMOTION  VALUES (" + speedID + ", " + id + ", " + priceF + ", null, 'Y', '" + download + "', '" + modem + "', " +
                             "'" + upload_K + "', '" + Tdocsis + "')";

                if (isExport)
                {
                    if (Write_SQL.Contains(cmdTxt) == false)
                    {
                        if (cmdTxt != "")
                        {
                            Write_SQL += cmdTxt + "\r\n" + "\r\n";
                        }
                    }
                }
                else
                {
                    Write_SQL += cmdTxt + "\r\n" + "\r\n";

                    using (transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted))
                    {
                        try
                        {
                            //command = new OracleCommand(cmdTxt, ConnectionProd, transaction);
                            command.CommandText = cmdTxt;
                            command.ExecuteNonQuery();

                            transaction.Commit();
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();

                            systemLog += DateTime.Now.ToString() + "  Cannot update the data of P_ID : " + id + " into table HISPEED_SPEED_PROMOTION" + "Detail : " +
                                            ex.Message.ToString() + "\r\n";
                        }

                    }
                }
            }

            reader.Close();

        }

        private void NewHispeedPromotion(int id, string mktCode, string orderType, string subProfile, double price, string channel, string modemType,
            string docsis, string entry, string install, string start, string end, string download, string upload, string uomUp)
        {
            OracleCommand command;
            OracleDataReader reader = null;
            string pName = "";
            string[] code = mktCode.Split('-');
            string pCode = code[0].Trim();
            int speedID = Convert.ToInt32(code[1].Trim());

            //set month contract
            string month = entry.Substring(12);
            month = Regex.Replace(month, "[^0-9]", "");
            month = month + "M";

            //set contract
            entry = entry.Substring(0, 10);
            install = install.Substring(0, 10);

            string prefix;
            if (mktCode.StartsWith("TRL"))
            {
                prefix = mktCode.Substring(0, 5);
            }
            else
            {
                prefix = mktCode.Substring(0, 2);
            }

            //borrow modem
            string modem;
            if (orderType == "New")
            {
                modem = "BM";
            }
            else
            {
                modem = "BM,NM";
            }

            //Get prod_type
            string queryProd = "SELECT * FROM TRUE9_BPT_HISPEED_PRODTYPE WHERE MKT = '" + prefix + "' AND ORDER_TYPE = '" + orderType + "'";

            command = new OracleCommand(queryProd, ConnectionTemp);
            reader = command.ExecuteReader();
            reader.Read();
            string prod = reader["PROD_TYPE"].ToString();
            reader.Close();

            //get pName
            pName = TempPName[mktCode];

            //insert data into hispeed promotion
            OracleCommand cmd = ConnectionProd.CreateCommand();
            OracleTransaction transaction = null;

            string cmdTxt = "INSERT INTO HISPEED_PROMOTION VALUES (" + id + ", '" + pCode + "', '" + pCode + "', '" + pName + "', '" + pName + "', '" + orderType + "', 'Active','','',0,0,'Y','Y','',0,'N','0'," +
                           "'Y','Y','N','" + prod + "', sysdate, sysdate, '" + month + "',0,'TI', TO_DATE('" + start + "','dd/mm/yyyy'), TO_DATE('" + end + "','dd/mm/yyyy'), 'M', '" + pCode + "','N','N','Y', '" +
                           entry + "', '" + install + "','" + modem + "','N','" + subProfile + "','')";

            if (isExport)
            {
                if (TempExport.ContainsKey(mktCode + "," + orderType))
                {
                    foreach (KeyValuePair<string, int> keyValuePair in TempExport)
                    {
                        string key = keyValuePair.Key;

                        if (key == (mktCode + "," + orderType))
                        {
                            id = keyValuePair.Value;
                            break;
                        }
                    }
                }
                else
                {
                    Write_SQL += cmdTxt + "\r\n";
                }
            }
            else
            {
                Write_SQL += cmdTxt + "\r\n";

                using (transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted))
                {
                    try
                    {
                        //cmd = new OracleCommand(cmdTxt, ConnectionProd, transaction);
                        cmd.CommandText = cmdTxt;
                        cmd.ExecuteNonQuery();

                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();

                        systemLog += DateTime.Now.ToString() + "  Cannot insert the data of " + id + " : " + mktCode + ", order " +
                            orderType + " into table HISPEED_PROMOTION" + " Detail : " + ex.Message.ToString() + "\r\n";
                    }
                }
            }

            //Set data into table hispeed_channel_promotion
            Channel_Promotion(id, channel, start, end);

            //Set data into table hispeed_speed_promotion
            HispeedSpeedPromotion(id, price, download, upload, uomUp, speedID, modemType, docsis);
        }

        /// <summary>
        /// Check already speed in table hispeed_speed
        /// </summary>
        /// <returns></returns>
        private int GetSpeedID(string download, string uom)
        {
            int speedID = -1;
            try
            {
                //Create an OracleCommand object using the connection object
                OracleCommand command = ConnectionProd.CreateCommand();
                OracleTransaction transaction = null;

                string downloadK = ConvertSpeed(download, uom);
                int speedF = Convert.ToInt32(Regex.Replace(download, "[^0-9]", ""));
                string speed_detail = downloadK + "K";

                string query = "SELECT * FROM HISPEED_SPEED WHERE SPEED_DESC = '" + downloadK + "'";

                OracleCommand cmd = new OracleCommand(query, ConnectionProd);
                OracleDataReader reader = cmd.ExecuteReader();
                reader.Read();

                if (reader.HasRows)
                {
                    speedID = Convert.ToInt32(reader["SPEED_ID"].ToString());
                }
                else
                {
                    using (transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted))
                    {
                        try
                        {
                            string cmdTxt = "INSERT INTO HISPEED_SPEED VALUES (" + speedF + ",'" + downloadK + "','" +
                                speed_detail + "','" + downloadK + "')";

                            //command = new OracleCommand(cmdTxt, ConnectionProd, transaction);
                            command.CommandText = cmdTxt;
                            command.ExecuteNonQuery();

                            transaction.Commit();
                        }
                        catch (Exception)
                        {
                            transaction.Rollback();
                            throw new Exception();
                        }
                    }
                    speedID = speedF;
                }

                reader.Close();
            }
            catch (Exception ex)
            {
                systemLog += DateTime.Now.ToString() + "  Cannot add new speed " + speedID + " into table HISPEED_SPEED" + "Error Detail : " +
                                            ex.Message.ToString() + "\r\n";
            }

            return speedID;
        }

        private void listBox1_MouseClick(object sender, MouseEventArgs e)
        {
            dataGridView1.ClearSelection();
            if (listBox1.SelectedItem != null)
            {
                int selected = listBox1.SelectedIndex;
                dataGridView1.Rows[indexDgv[selected]].Selected = true;
            }
        }

        /// <summary>
        /// update complete flag in table reserve
        /// </summary>
        private void UpdateCompleteFlag(int min, int max)
        {
            try
            {
                if (ConnectionTemp.State != ConnectionState.Open)
                {
                    ConnectionTemp.Open();
                }

                //Create an OracleCommand object using the connection object
                OracleCommand command = ConnectionTemp.CreateCommand();
                OracleTransaction transaction;

                // Start a local transaction
                transaction = ConnectionTemp.BeginTransaction(IsolationLevel.ReadCommitted);
                // Assign transaction object for a pending local transaction
                command.Transaction = transaction;

                try
                {
                    command.CommandText = "UPDATE TRUE9_BPT_RESERVE_ID SET COMPLETE_FLAG = 'Y', MAX_ID = '" + max + "' " +
                       "WHERE TYPE_NAME = 'Hispeed' AND UR_NO = '" + UR_NO + "' AND MIN_ID = '" + min + "'";

                    command.CommandType = CommandType.Text;

                    command.ExecuteNonQuery();
                    transaction.Commit();
                }
                catch (Exception)
                {
                    transaction.Rollback();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot update COMPLETE_FLAG to 'Y' in TRUE9_BPT_RESERVE_ID." + "\r\n" + "Please manual update flag in database" + "\r\n" + ex);

                //string msg = "Cannot update COMPLETE_FLAG to 'Y' in TRUE9_BPT_RESERVE_ID." + "\r\n" + "Please manual update flag in database" + "\r\n" + ex;
                //Color color = Color.FromArgb(204, 28, 68);

                //string FileName = string.Format("{0}Resources\\Error.png",
                //    Path.GetFullPath(Path.Combine(RunningPath, @"..\")));

                //dialogMessage = new DialogMessage(msg, color, FileName);
                //dialogMessage.ShowDialog();


                Application.Exit();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="speed"></param>
        /// <param name="uom"></param>
        /// <returns></returns>
        private string ConvertSpeed(string speed, string uom)
        {
            string convSpeed = Regex.Replace(speed, "[^0-9]", "");

            if (uom == "G")
            {
                convSpeed = (Convert.ToInt32(convSpeed) * 1024000).ToString();
            }
            else if (uom == "M")
            {
                convSpeed = (Convert.ToInt32(convSpeed) * 1024).ToString();
            }

            return convSpeed;
        }
    }

}
