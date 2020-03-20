using AutomateMappingTool;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MappingDiscount
{
    public partial class MainDiscount : Form
    {
        #region "Private field"
        private OracleConnection ConnectionProd;
        private OracleConnection ConnectionTemp;
        private HomeDiscount home;
        private ViewResult viewResult;
        private DataTable dataTableNew;
        /// <summary>
        /// Flag for new DC criteria
        /// </summary>
        private bool FLAG_NEW;
        /// <summary>
        /// Flag for existing criteria
        /// </summary>
        private bool FLAG_EXISTING;
        /// <summary>
        /// Requirement file
        /// </summary>
        private string filename;
        /// <summary>
        /// username implementer
        /// </summary>
        private string user;
        /// <summary>
        /// UR No#
        /// </summary>
        private string urNo;
        private bool Flag_Close = true;
        private string outputPath;

        private string dcCode;
        private string month;
        private string channel;
        private string mkt;
        private string orderType;
        private string speed;
        private string province;
        private string effective;
        private string expire;
        private string dcGroupID;
        private List<int> indexListbox;
        private bool isCorrect;

        bool UpdateCompleteFlag = true;
        bool CompleteFlag = true;

        private string log = "";

        private DgvSettings settingView;
        private ChangeFormat changeFormat;
        private Validation validate;
        private ReserveID reserv;

        //For move form
        int mov;
        int movX;
        int movY;
        #endregion

        public MainDiscount(HomeDiscount form, OracleConnection con, bool newDC, bool existingDC,
            string filename, string folder, string user, string UR)
        {
            InitializeComponent();

            //Get private variable
            home = form;
            ConnectionProd = con;
            FLAG_NEW = newDC;
            FLAG_EXISTING = existingDC;
            this.filename = filename;
            this.user = user;
            urNo = UR;
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

        /// <summary>
        /// When this form loading
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainDiscount_Load(object sender, EventArgs e)
        {
            home.Hide();

            listBox1.Items.Clear();
            listBox1.Hide();

            Cursor.Current = Cursors.Default;

            toolStripStatusLabel1.Text = filename;
            this.dataGridView1.AllowUserToAddRows = false;
            btnOK.Enabled = true;

            settingView = new DgvSettings();
            changeFormat = new ChangeFormat();
            reserv = new ReserveID();

            if (FLAG_NEW == true && FLAG_EXISTING == false)
            {
                labelPg.Text = "Page1/1";
                ViewNewDC();
            }
            else if (FLAG_NEW == false && FLAG_EXISTING == true)
            {
                labelPg.Text = "Page1/1";
                ViewExistingDC();
            }
            else
            {
                labelPg.Text = "Page1/2";
                btnOK.Text = "Next";
                ViewNewDC();
            }

            Cursor.Current = Cursors.Default;
        }

        /// <summary>
        /// When clicked Next button for generate new discount
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOK_Click(object sender, System.EventArgs e)
        {
            if (isCorrect)
            {
                Cursor.Current = Cursors.WaitCursor;
                //Process for Both criteria
                if (btnOK.Text == "Next")
                {
                    btnOK.Text = "Execute";
                    labelPg.Text = "Page2/2";
                    NewDiscount();

                    if (CompleteFlag)
                    {
                        ViewExistingDC();
                    }
                    else
                    {
                        btnOK.Enabled = false;
                    }
                }
                else
                {
                    if (FLAG_NEW == true && FLAG_EXISTING == false)
                    {
                        NewDiscount();

                        if (CompleteFlag == false)
                        {
                            btnOK.Enabled = false;
                        }
                    }
                    else if (FLAG_NEW == false && FLAG_EXISTING == true)
                    {
                        UpdateCompleteFlag = true;
                        ExistingDC(null);
                    }
                    else
                    {
                        UpdateCompleteFlag = true;
                        ExistingDC(dataTableNew);
                    }
                }

                Cursor.Current = Cursors.Default;
            }
            else
            {
                MessageBox.Show("There is some data an error." + "Please check detail in log file.");
            }
        }

        /// <summary>
        /// When clicked close this page
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Hide();
            home.Show();
        }

        /// <summary>
        /// Set datagridview for new DC criteria
        /// </summary>
        private void ViewNewDC()
        {
            labelTitle.Text = "New Criteria Discount";

            List<string> lstHeader = new List<string>();
            lstHeader.Add("Vcare Discount Code");
            lstHeader.Add("Criteria (Mth)");
            lstHeader.Add("Channel");
            lstHeader.Add("MKT Code");
            lstHeader.Add("Order type");
            lstHeader.Add("Product");
            lstHeader.Add("Speed");
            lstHeader.Add("Province");
            lstHeader.Add("Effective Start Date");
            lstHeader.Add("End Date");

            settingView.SetDgv(dataGridView1, filename, "Discount New Sale(SMART UI)$B3:K", lstHeader);

            validate = new Validation(dataGridView1, ConnectionProd, ConnectionTemp, outputPath);
            listBox1.Items.Clear();
            ListBox ls = validate.Verify("Disc");

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
        }

        private void ViewExistingDC()
        {
            labelTitle.Text = "Existing Criteria Discount";

            List<string> lstHeader = new List<string>();
            lstHeader.Add("DC_ID");
            lstHeader.Add("DISCOUNT_CODE");
            lstHeader.Add("DISCOUNT_DESCRIPTION");
            lstHeader.Add("MONTH_MIN");
            lstHeader.Add("MONTH_MAX");
            lstHeader.Add("START_DATE");
            lstHeader.Add("END_DATE");
            lstHeader.Add("ACTIVE");
            lstHeader.Add("SALE_CHANNEL");
            lstHeader.Add("SPEED");
            lstHeader.Add("MKT_CODE");
            lstHeader.Add("ORDER");
            lstHeader.Add("PROV");
            lstHeader.Add("PROD");
            lstHeader.Add("GROUP_ID");
            lstHeader.Add("DG_ORDER");
            lstHeader.Add("GROUPID");

            settingView.SetDgv(dataGridView1, filename, "Update Date Discount (SMART UI)$", lstHeader);

            //Hide unused column
            dataGridView1.Columns[3].Visible = false;
            dataGridView1.Columns[4].Visible = false;
            dataGridView1.Columns[7].Visible = false;
            dataGridView1.Columns[11].Visible = false;
            dataGridView1.Columns[12].Visible = false;
            dataGridView1.Columns[13].Visible = false;
            dataGridView1.Columns[15].Visible = false;
            dataGridView1.Columns[16].Visible = false;

            validate = new Validation(dataGridView1, ConnectionProd, ConnectionTemp, outputPath);
            listBox1.Items.Clear();
            ListBox ls = validate.Verify("Disc");

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
        }

        #region "Existing criteria"

        private void ExistingDC(DataTable table)
        {
            //TableResult New DC
            DataTable dataTableNew = table;
            bool HasWrongFormat = true;
            Dictionary<string, string> dicValidDate = new Dictionary<string, string>();
            DataTable dataTable = new DataTable();
            DateTime StartDB = new DateTime();
            DateTime EndDB = new DateTime();
            string startF = "";
            string endF = "";
            Flag_Close = true;

            try
            {
                //Check Date
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    string dc_id = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    string start = dataGridView1.Rows[i].Cells[5].Value.ToString();
                    string end = dataGridView1.Rows[i].Cells[6].Value.ToString();

                    //Get start date from DB
                    string query = "SELECT * FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_ID = '" + dc_id + "'";

                    OracleCommand cmd = new OracleCommand(query, ConnectionProd);
                    OracleDataReader reader = cmd.ExecuteReader();

                    if (reader.HasRows)
                    {
                        reader.Read();

                        //Get date from DB
                        string strStart = reader["DC_START_DT"].ToString();
                        string strEnd = reader["DC_END_DT"].ToString();

                        StartDB = Convert.ToDateTime(strStart);

                        if (end.Equals("-"))
                        {
                            end = String.Empty;
                        }
                        else
                        {
                            if (strEnd != "")
                            {
                                EndDB = Convert.ToDateTime(strEnd);
                            }
                        }

                        //Check Date
                        DateTime dDateStart = new DateTime();
                        DateTime dDateEnd = new DateTime();

                        //Start Date
                        if (DateTime.TryParse(start, out dDateStart))
                        {//update
                            if (dDateStart > DateTime.Today && StartDB >= DateTime.Today)
                            {
                                if (dicValidDate.ContainsKey(dc_id) == false)
                                {
                                    if (dDateStart == DateTime.Today)
                                    {
                                        start = string.Format("{0:dd/MM/yyyy HH:mm:ss}", dDateStart);
                                        startF = start;
                                    }
                                    else
                                    {
                                        start = string.Format("{0:dd/MM/yyyy}", dDateStart);
                                        startF = start;
                                    }
                                }
                                else
                                {
                                    //duplicate id in file
                                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Gray;
                                    continue;
                                }

                            }
                            else
                            {
                                startF = null;
                            }
                        }
                        else
                        {
                            dataGridView1.Rows[i].Cells[5].Style.BackColor = Color.Red;

                            startF = null;
                        }

                        //End Date
                        if (DateTime.TryParse(end, out dDateEnd))
                        {
                            if (dDateEnd > dDateStart && dDateEnd >= DateTime.Today)
                            {
                                //update
                                if (dDateEnd == DateTime.Today)
                                {
                                    end = string.Format("{0:dd/MM/yyyy HH:mm:ss}", dDateEnd);

                                    endF = end;
                                }
                                else
                                {
                                    end = string.Format("{0:dd/MM/yyyy}", dDateEnd);

                                    endF = end;
                                }

                            }

                        }
                        else
                        {
                            if (end != "")
                            {
                                HasWrongFormat = false;
                                dataGridView1.Rows[i].Cells[6].Style.BackColor = Color.Red;
                            }
                            else
                            {
                                endF = "";
                            }
                        }

                    }
                    else
                    {
                        MessageBox.Show("DC_ID : " + dc_id + " not found in database. Please check the data and try again!", "Data Not Found",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Environment.Exit(0);
                    }

                    dicValidDate[dc_id] = startF + "," + endF;

                }

                if (HasWrongFormat == false)
                {
                    MessageBox.Show("Please edit effective/end date in this file", "Wrong Format", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    DialogResult dialogResult = MessageBox.Show("Do you want to update sales period?", "Confirm",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (dialogResult == DialogResult.Yes)
                    {
                        Flag_Close = false;
                        UpdateData(dicValidDate);
                    }
                    else
                    {
                        UpdateCompleteFlag = false;
                    }

                    if (UpdateCompleteFlag)
                    {
                        string lstId = "";
                        Flag_Close = true;

                        foreach (KeyValuePair<string, string> keyValuePair in dicValidDate)
                        {
                            lstId += "'" + keyValuePair.Key + "',";
                        }

                        lstId = lstId.Substring(0, lstId.Length - 1);

                        string query = "SELECT * FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_ID in (" + lstId + ")";

                        OracleDataAdapter adapter = new OracleDataAdapter(query, ConnectionProd);
                        adapter.Fill(dataTable);

                        DataTable[] lstTable = new DataTable[2];

                        //Both Criteria
                        if (FLAG_NEW == true && FLAG_EXISTING == true)
                        {
                            lstTable[0] = dataTableNew;
                            lstTable[1] = dataTable;

                            viewResult = new ViewResult(this, lstTable, ConnectionProd, ConnectionTemp, urNo, lstId, null, user, log, outputPath);
                            viewResult.Show();
                        }
                        else
                        {
                            lstTable[1] = dataTable;
                            viewResult = new ViewResult(this, lstTable, ConnectionProd, null, urNo, lstId, null, user, null, outputPath);
                            viewResult.Show();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Cursor.Current = Cursors.Default;
                Flag_Close = true;
                MessageBox.Show("Found Problem." + "\r\n" + e.Message, "Error"
                    , MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateData(Dictionary<string, string> ValidDate)
        {
            //Create an OracleCommand object using the connection object
            OracleCommand command = ConnectionProd.CreateCommand();
            OracleTransaction transaction = null;
            try
            {
                // Start a local transaction
                transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted);
                // Assign transaction object for a pending local transaction
                command.Transaction = transaction;

                foreach (KeyValuePair<string, string> valuePair in ValidDate)
                {
                    string id = valuePair.Key;
                    string[] lstDate = valuePair.Value.Split(',');
                    string effective = lstDate[0];
                    string end = lstDate[1];
                    string sqlCommand;

                    if (String.IsNullOrEmpty(effective))
                    {
                        sqlCommand = "UPDATE DISCOUNT_CRITERIA_MAPPING SET DC_END_DT = TO_DATE('" + end + "', 'dd/mm/yyyy') " +
                            " WHERE DC_ID = '" + id + "'";
                    }
                    else
                    {
                        sqlCommand = "UPDATE DISCOUNT_CRITERIA_MAPPING SET DC_START_DT = TO_DATE('" + effective + "','dd/mm/yyyy')," +
                        "DC_END_DT = TO_DATE('" + end + "', 'dd/mm/yyyy') WHERE DC_ID = '" + id + "'";
                    }

                    command.CommandText = sqlCommand;
                    command.CommandType = CommandType.Text;

                    command.ExecuteNonQuery();
                }

                transaction.Commit();

            }
            catch (Exception e)
            {
                transaction.Rollback();

                Cursor.Current = Cursors.Default;
                MessageBox.Show("Cannot update data into the database" + "\r\n" + "Please Check and try again or use script in excel for manual insert" + "\r\n" +
                        e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateCompleteFlag = false;
            }

        }

        #endregion


        /// <summary>
        /// Process for check user login & update flag in table true9_bpt_reserve_id
        /// </summary>
        private void NewDiscount()
        {
            //ConnectionTemp = new OracleConnection();
            //string connString = "Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.19.193.20)(PORT = 1560))" +
            //       "(CONNECT_DATA = (SID = TEST03)));User Id= TRUREF71; Password= TRUREF71;";

            //ConnectionTemp.ConnectionString = connString;

            try
            {
                //ConnectionTemp.Open();

                ConnectionTemp = ConnectionProd;

                bool hasUserInProcess = reserv.CheckStatus(ConnectionTemp, "Disc", user, urNo);

                if (hasUserInProcess == false)
                {
                    int dcID = reserv.GetMaxID(ConnectionTemp, ConnectionProd, "Disc", user, urNo);
                    ExecuteNewDC(dcID);
                }
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show("Failed..." + "\r\n" + ex.Message);

                this.Close();
                home.Show();
            }
        }

        /// <summary>
        /// Main process for New discount criteria
        /// </summary>
        /// <param name="minID"></param>
        private void ExecuteNewDC(int minID)
        {
            CompleteFlag = true;
            string strMinID = "DC" + string.Format("{0:00000000}", minID);

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

            Dictionary<string, string> dicGroupID = null;
            string existID = "";

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                bool hasID = false;

                dcCode = dataGridView1.Rows[i].Cells[0].Value.ToString().Trim();
                month = dataGridView1.Rows[i].Cells[1].Value.ToString();
                channel = dataGridView1.Rows[i].Cells[2].Value.ToString();
                mkt = dataGridView1.Rows[i].Cells[3].Value.ToString();
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

                orderType = dataGridView1.Rows[i].Cells[4].Value.ToString();
                string getSpeed = dataGridView1.Rows[i].Cells[6].Value.ToString();
                province = dataGridView1.Rows[i].Cells[7].Value.ToString().ToUpper();
                effective = dataGridView1.Rows[i].Cells[8].Value.ToString();
                expire = dataGridView1.Rows[i].Cells[9].Value.ToString();
                if (expire.Equals("-"))
                {
                    expire = String.Empty;
                }

                List<Dictionary<string, string>> lstRangeM = validate.FormatGroupMonth(month);
                Dictionary<string, string> groupMonth = new Dictionary<string, string>();
                List<string> lstChannel = new List<string>();
                string[] lstorder = null;
                string[] lstprov = null;

                //Cut each Order_Type
                if (orderType.Contains(","))
                {
                    lstorder = orderType.Split(',');
                }
                else
                {
                    lstorder = new string[] { orderType };

                    string upper = lstorder[0].ToUpper();
                    if (upper.StartsWith("ALL"))
                    {
                        lstorder[0] = "ALL";
                    }
                }

                //Cut each Channel
                if (channel.Contains(","))
                {
                    string[] lstCh = channel.Split(',');
                    List<string> distinct = lstCh.Distinct().ToList();
                    lstChannel = distinct;
                }
                else
                {
                    lstChannel = new List<string>() { channel };

                    string upper = lstChannel[0].ToUpper();
                    if (upper.Equals("ALL") || upper.Equals("DEFAULT"))
                    {
                        lstChannel[0] = "ALL";
                    }
                }

                //Speed
                speed = changeFormat.FormatSpeed(getSpeed);

                //Cut each Province
                if (province.Contains(","))
                {
                    lstprov = province.Split(',');
                }
                else
                {
                    lstprov = new string[] { province };
                }

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

                for (int num = 0; num < lstRangeM.Count; num++)
                {
                    groupMonth = lstRangeM[num];
                    string minMonth = groupMonth["min" + num];
                    string maxMonth = groupMonth["max" + num];
                    string GroupName = dcCode + minMonth + maxMonth;
                    dcGroupID = dicGroupID[GroupName];

                    for (int j = 0; j < lstChannel.Count; j++)
                    {
                        for (int k = 0; k < lstorder.Length; k++)
                        {
                            for (int p = 0; p < lstprov.Length; p++)
                            {
                                string ch = lstChannel[j].Trim();
                                string ord = lstorder[k].Trim();
                                string prov = lstprov[p].Trim();

                                string strQuery = "SELECT * FROM (SELECT * FROM (SELECT DC_ID, DG_DISCOUNT DISCOUNT_CODE, DG_DISCOUNT_DESC  DISCOUNT_DESCRIPTION,DG_MONTH_MIN MONTH_MIN" +
                                                  ", DG_MONTH_MAX MONTH_MAX, to_char(trunc(DC_START_DT), 'dd/mm/yyyy')  START_DATE ,to_char(trunc(DC_END_DT), 'dd/mm/yyyy')   END_DATE" +
                                                  ", DG_ACTIVE_FLAG ACITVE_FLAG, SALE_CHANNEL, SPEED, MARKETING_CODE, ORDER_TYPE, PROVINCE, PRODUCT, DG_GROUPID, DG_ORDER_BY, DC_GROUPID " +
                                                  "FROM(SELECT * FROM(" +
                                                  "SELECT DC1.DC_START_DT, DC1.DC_END_DT, DC1.DC_ID, DC1.DC_GROUPID, NVL(PRODUCT, 'ALL') PRODUCT, NVL(SPEED, 'ALL') SPEED, NVL(MARKETING_CODE, 'ALL')MARKETING_CODE," +
                                                  "NVL(PROVINCE, 'ALL')PROVINCE, NVL(SALE_CHANNEL, 'ALL')SALE_CHANNEL, NVL(ORDER_TYPE, 'ALL')ORDER_TYPE FROM(" +
                                                  "SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PRODUCT', DC_VALUE, 'ALL') PRODUCT FROM DISCOUNT_CRITERIA_MAPPING" +
                                                  " WHERE DC_TYPE = 'PRODUCT'  AND DC_ACTIVE_FLAG = 'Y')  DC1,(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'SPEED', DC_VALUE, 'ALL') SPEED " +
                                                  "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'SPEED' AND DC_ACTIVE_FLAG = 'Y')  DC2,(" +
                                                  "SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'MARKETING_CODE', DC_VALUE, 'ALL') MARKETING_CODE FROM DISCOUNT_CRITERIA_MAPPING " +
                                                  "WHERE DC_TYPE = 'MARKETING_CODE' AND DC_ACTIVE_FLAG = 'Y') DC3, (" +
                                                  "SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PROVINCE', DC_VALUE, 'ALL') PROVINCE " +
                                                  "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PROVINCE' AND DC_ACTIVE_FLAG = 'Y'   ) DC4, (" +
                                                  "SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'SALE_CHANNEL', DC_VALUE, 'ALL') SALE_CHANNEL " +
                                                  "FROM DISCOUNT_CRITERIA_MAPPING " +
                                                  "WHERE DC_TYPE = 'SALE_CHANNEL' AND DC_ACTIVE_FLAG = 'Y'  ) DC5, (" +
                                                  "SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'ORDER_TYPE', DC_VALUE, 'ALL') ORDER_TYPE " +
                                                  "FROM DISCOUNT_CRITERIA_MAPPING " +
                                                  "WHERE DC_TYPE = 'ORDER_TYPE' AND DC_ACTIVE_FLAG = 'Y'  ) DC6" +
                                                  " WHERE DC1.DC_ID = DC2.DC_ID(+) " +
                                                  "AND DC1.DC_ID = DC3.DC_ID(+) " +
                                                  "AND DC1.DC_ID = DC4.DC_ID(+) " +
                                                  "AND DC1.DC_ID = DC5.DC_ID(+) " +
                                                  "AND DC1.DC_ID = DC6.DC_ID(+) " +
                                                  ")  ) DISCOUNT, DISCOUNT_CRITERIA_GROUP WHERE DISCOUNT.DC_GROUPID = DISCOUNT_CRITERIA_GROUP.DG_GROUPID AND DISCOUNT_CRITERIA_GROUP.DG_ACTIVE_FLAG = 'Y' " +
                                                  " ORDER BY DG_ORDER_BY,DG_MONTH_MIN , DG_DISCOUNT,DC_START_DT)) WHERE DG_GROUPID ='" + dcGroupID + "' AND SALE_CHANNEL = '" +
                                                     ch + "' AND MARKETING_CODE = '" + mkt + "' AND ORDER_TYPE = '" + ord + "' AND SPEED = '" + speed + "' AND PROVINCE = '" + prov + "'";

                                OracleCommand cmd = new OracleCommand(strQuery, ConnectionProd);
                                OracleDataReader reader = cmd.ExecuteReader();
                                reader.Read();
                                if (reader.HasRows)
                                {
                                    hasID = true;
                                    string idEx = reader["DC_ID"].ToString();
                                    existID += "DC_ID : " + idEx + ", DC Code : " + dcCode + ", MKT : " + mkt + ", Order : " + ord + ", Month " +
                                        minMonth + ":" + maxMonth + ", Channel : " + ch + ", Speed : " + speed + ", Province : " + prov + "\r\n";
                                }
                                else
                                {
                                    hasID = false;
                                }

                                for (int m = 1; m <= 6; m++)
                                {
                                    object[] obj = new object[7];

                                    if (hasID == false)
                                    {
                                        obj[0] = "DC" + string.Format("{0:00000000}", minID);
                                    }
                                    else
                                    {
                                        obj[0] = "";
                                    }

                                    obj[3] = dcGroupID;
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
                                            obj[2] = "ALL";
                                            break;
                                        case 4:
                                            obj[1] = "PROVINCE";
                                            obj[2] = prov;
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
                                    minID += 1;
                                }

                            }
                        }
                    }

                }

            }

            minID = minID - 1;
            string max_id = "DC" + string.Format("{0:00000000}", minID);

            if (FLAG_NEW == true && FLAG_EXISTING == true)
            {
                dataTableNew = dtTableView;
            }
            else
            {
                if (existID != "")
                {
                    log += "Already exists data in database" + "\r\n" + existID + "\r\n";
                }

                DataTable[] lstTable = new DataTable[2];
                lstTable[0] = dtTableView;

                viewResult = new ViewResult(this, lstTable, ConnectionProd, ConnectionTemp, urNo, strMinID,
                    max_id, user, log, outputPath);

                viewResult.ShowDialog();
            }
        }

        /// <summary>
        /// Delete flag when close form immediately
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ViewDCData_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Flag_Close == false)
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

                        string query = "SELECT * FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = 'Disc' AND COMPLETE_FLAG = 'N'";

                        OracleCommand cmd = new OracleCommand(query, ConnectionTemp);
                        OracleDataReader reader = cmd.ExecuteReader();
                        reader.Read();
                        if (reader.HasRows)
                        {
                            reader.Close();

                            string qryDel = "DELETE FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = 'Disc' AND COMPLETE_FLAG = 'N'";
                            OracleCommand command = new OracleCommand(qryDel, ConnectionTemp);
                            //command.CommandText = "DELETE FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = 'Disc' AND COMPLETE_FLAG = 'N'";
                            command.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        string log = ex.Message;
                        MessageBox.Show("Can't delete row in field COMPLETE_FLAG = 'N' ... Please manual delete!" + "\r\n" + ex);

                        ConnectionTemp.Close();
                        ConnectionProd.Close();

                        Environment.Exit(0);
                    }
                }
            }

        }

        #region "Event Handle"
        private void ViewDCData_SizeChanged(object sender, EventArgs e)
        {
            int w = this.Size.Width;
            int h = this.Size.Height;

            labelPg.Location = new Point(w - 105, 115);
            //btnOK.Location = new Point(w - 130, h - 100);

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
        #endregion

        private void pictureBoxValidate_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            dataGridView1.EndEdit();
            dataGridView1.Update();

            validate = new Validation(dataGridView1, ConnectionProd, ConnectionTemp, outputPath);
            listBox1.Items.Clear();
            ListBox lst = validate.Verify("Disc");

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

        private void pictureBoxValidate_MouseHover(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.Hand;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Hide();
            home.Show();
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
