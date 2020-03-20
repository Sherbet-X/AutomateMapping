using AutomateMappingTool;
using System;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace MappingDiscount
{
    public partial class ViewResult : Form
    {
        #region "Private Field"
        private DataTable[] lstTableResult;
        private OracleConnection ConnectionProd;
        private OracleConnection ConnectionTemp;
        private string UR_NO;
        private string Min_ID;
        private string Max_ID;
        private bool Flag_close = true;
        private bool isExisting = false;
        private bool isBoth = false;
        private string minID_New; //for both
        private string maxID_New; // for both
        private string USER;
        private string connString71;
        private string LogFile;
        private Size size;
        private Form main;
        private string outputPath;
        private ReserveID reserve;

        //For move form
        int mov;
        int movX;
        int movY;

        //For maximize form
        int maximize = 1;
        #endregion

        public ViewResult(Form form, DataTable[] lstTable, OracleConnection conn, OracleConnection conn71,
            string ur, string min, string max, string user, string log, string path)
        {
            InitializeComponent();
            lstTableResult = lstTable;
            ConnectionProd = conn;
            ConnectionTemp = conn71;
            UR_NO = ur;
            Min_ID = min;
            Max_ID = max;
            USER = user;
            LogFile = log;
            main = form;
            size = form.Size;
            outputPath = path;
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

        private void ViewResult_Load(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            main.Hide();
            int screenW = SystemInformation.VirtualScreen.Width;
            int screenH = SystemInformation.VirtualScreen.Height;

            if (screenW - size.Width <= 250 || screenH - size.Height <= 50)
            {
                this.WindowState = FormWindowState.Maximized;
            }
            else
            {
                this.Size = size;
            }

            connString71 = "Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.19.193.20)(PORT = 1560))" +
                   "(CONNECT_DATA = (SID = TEST03)));User Id= TRUREF71; Password= TRUREF71;";

            btnInsert.Enabled = true;

            reserve = new ReserveID();

            //lstTableResult[0] = New
            //lstTableResult[1] = Existing
            //Both
            if (lstTableResult[0] != null && lstTableResult[1] != null)
            {
                isBoth = true;

                DataTable dtTableNew = lstTableResult[0];
                DataTable dtTableEx = lstTableResult[1];

                DataTable AllTable = new DataTable();

                AllTable = dtTableNew.Copy();
                AllTable.Merge(dtTableEx, true, MissingSchemaAction.Ignore);

                dataGridViewResult.DataSource = AllTable;

                toolStripStatusLabel1.Text = "Sum : " + dataGridViewResult.RowCount.ToString() + " rows";

            }
            else
            {
                if (lstTableResult[0] != null)
                {
                    //NewDC
                    DataTable dtTableNew = lstTableResult[0];
                    dataGridViewResult.DataSource = dtTableNew;
                    toolStripStatusLabel1.Text = "Sum : " + dataGridViewResult.RowCount.ToString() + " rows";
                }
                else
                {
                    //Existing
                    DataTable dtTableEx = lstTableResult[1];
                    btnExport.Enabled = false;
                    isExisting = true;
                    dataGridViewResult.DataSource = dtTableEx;
                    toolStripStatusLabel1.Text = "Sum : " + dataGridViewResult.RowCount.ToString() + " rows";
                }
            }

            HilightRows();

            Cursor.Current = Cursors.Default;
        }

        private void HilightRows()
        {
            dataGridViewResult.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridViewResult.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(189, 10, 59);
            dataGridViewResult.ColumnHeadersHeight = 25;
            dataGridViewResult.EnableHeadersVisualStyles = false;

            if (lstTableResult[0] != null)
            {
                for (int i = 0; i < lstTableResult[0].Rows.Count; i++)
                {
                    string id = dataGridViewResult.Rows[i].Cells["DC_ID"].ToString();

                    dataGridViewResult.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(199, 245, 245);

                    if (id == "")
                    {
                        dataGridViewResult.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                    }
                }
            }

            if (lstTableResult[1] != null)
            {
                if (isBoth)
                {
                    int row = lstTableResult[0].Rows.Count;
                    for (int j = row; j < dataGridViewResult.Rows.Count; j++)
                    {
                        dataGridViewResult.Rows[j].DefaultCellStyle.BackColor = Color.FromArgb(240, 214, 123);
                    }
                }
                else
                {
                    for (int j = 0; j < lstTableResult[1].Rows.Count; j++)
                    {
                        dataGridViewResult.Rows[j].DefaultCellStyle.BackColor = Color.FromArgb(240, 214, 123);
                    }
                }

            }

        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            ExportScript(outputPath);

            if (Max_ID.StartsWith("DC"))
            {
                reserve.UpdateID(ConnectionTemp, Min_ID, Max_ID, "Disc", USER, UR_NO);
            }
            else if (Max_ID.StartsWith("VAS"))
            {
                reserve.UpdateID(ConnectionTemp, Min_ID, Max_ID, "VAS", USER, UR_NO);
            }

            string msg = "Your script was exported successfully" + "\r\n" + "Min ID: " + Min_ID + "\r\n" + "\r\n" + "Max ID: " + Max_ID;

            MessageBox.Show(msg, "Successfully", MessageBoxButtons.OK, MessageBoxIcon.Information);

            Cursor.Current = Cursors.Default;

        }

        /// <summary>
        /// When clicked Insert button from insert data to database
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnInsert_Click(object sender, EventArgs e)
        {
            if (isExisting)
            {
                Cursor.Current = Cursors.WaitCursor;
                ExportImp(outputPath);

                string msg = "Already update your data into database";

                MessageBox.Show(msg, "Successfully", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Cursor.Current = Cursors.Default;
            }
            else
            {
                DialogResult result = MessageBox.Show("Do you want to insert data to database?", "Confirmation",
                        MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                if (result == DialogResult.OK)
                {
                    Flag_close = false;
                    Application.UseWaitCursor = true;
                    Cursor.Current = Cursors.WaitCursor;
                    labelProgress.Visible = true;

                    if (backgroundWorker1.IsBusy != true)
                    {
                        backgroundWorker1.RunWorkerAsync();
                        btnInsert.Enabled = false;
                        btnExport.Enabled = false;

                    }

                }
            }

            Flag_close = true;

        }

        private void InsertData()
        {
            //Create an OracleCommand object using the connection object
            OracleCommand command = ConnectionProd.CreateCommand();
            OracleTransaction transaction;
            string id = "";
            int rowCount;

            // Start a local transaction
            transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted);
            // Assign transaction object for a pending local transaction
            command.Transaction = transaction;

            try
            {
                if (isBoth)
                {
                    rowCount = lstTableResult[0].Rows.Count;
                    minID_New = dataGridViewResult.Rows[0].Cells[0].Value.ToString();
                    maxID_New = dataGridViewResult.Rows[rowCount - 1].Cells[0].Value.ToString();
                }
                else
                {
                    rowCount = dataGridViewResult.Rows.Count;
                }

                //Insert data to database
                for (int i = 0; i < rowCount; i++)
                {
                    id = dataGridViewResult.Rows[i].Cells[0].Value.ToString();

                    if (id != "")
                    {
                        string type = dataGridViewResult.Rows[i].Cells[1].Value.ToString();
                        string value = dataGridViewResult.Rows[i].Cells[2].Value.ToString();
                        string group = dataGridViewResult.Rows[i].Cells[3].Value.ToString();
                        string flag = dataGridViewResult.Rows[i].Cells[6].Value.ToString();

                        string start = dataGridViewResult.Rows[i].Cells[4].Value.ToString();
                        string end = dataGridViewResult.Rows[i].Cells[5].Value.ToString();

                        if (start.Contains("-") || start == "null")
                        {
                            start = "";
                        }

                        if (end.Contains("-") || end == "null")
                        {
                            end = "";
                        }

                        command.CommandText = "insert into DISCOUNT_CRITERIA_MAPPING values ('" + id + "','" + type + "','" + value + "','" + group +
                          "',to_date('" + start + "','dd/mm/yyyy')," + "to_date('" + end + "', 'dd/mm/yyyy'),'" + flag + "',sysdate,'" + USER + "',null,null)";

                        command.CommandType = CommandType.Text;
                        command.ExecuteNonQuery();
                    }

                }

                transaction.Commit();
            }
            catch (Exception exception)
            {
                Flag_close = true;
                transaction.Rollback();

                Cursor.Current = Cursors.Default;

                MessageBox.Show("Cannot insert data to database" + "\r\n" + "Please Check and try again or Use script in excel insert" + "\r\n" +
                    exception.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateCompleteFlag()
        {
            try
            {
                if (ConnectionTemp.State != ConnectionState.Open)
                {
                    try
                    {
                        ConnectionTemp.Open();
                    }
                    catch
                    {
                        ConnectionTemp.ConnectionString = connString71;
                        ConnectionTemp.Open();
                    }
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
                    if (Max_ID.Contains("DC"))
                    {
                        command.CommandText = "UPDATE TRUE9_BPT_RESERVE_ID SET COMPLETE_FLAG = 'Y', MAX_ID = '" + Max_ID + "' " +
                        "WHERE TYPE_NAME = 'Disc' AND UR_NO = '" + UR_NO + "' AND MIN_ID = '" + Min_ID + "'";
                    }
                    else
                    {
                        command.CommandText = "UPDATE TRUE9_BPT_RESERVE_ID SET COMPLETE_FLAG = 'Y', MAX_ID = '" + Max_ID + "' " +
                        "WHERE TYPE_NAME = 'VAS' AND UR_NO = '" + UR_NO + "' AND MIN_ID = '" + Min_ID + "'";
                    }

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
                Cursor.Current = Cursors.Default;
                MessageBox.Show("Please manual update flag" + "\r\n" + "Cannot update COMPLETE_FLAG to 'Y'" + "\r\n" + ex.Message);
            }
        }

        public void ExportImp(string path)
        {
            //load excel and create a new workbook
            var excelApp = new Excel.Application();
            Excel.Workbook workbook = null;
            string name = "";

            try
            {
                //Open connection
                if (ConnectionProd.State != ConnectionState.Open)
                {
                    ConnectionProd.Open();
                }

                if (Min_ID.Contains("DC"))
                {
                    name = UR_NO + "_Disc_Criteria";
                }
                else if (Min_ID.StartsWith("VAS"))
                {
                    name = UR_NO + "_VAS_Criteria";
                }
                else
                {
                    name = UR_NO + "_Hispeed_Criteria";
                }

                string dirPath = path;
                string fileName = name + ".xlsx";
                string[] files = Directory.GetFiles(dirPath);
                int count = files.Count(file => { return file.Contains(fileName); });
                string newFileName = fileName;

                if (count > 0)
                {
                    DialogResult result = MessageBox.Show("There is already a file with the same name in this location" + "\r\n" +
                        "Do you want to replace it?", "Replace File", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                    if (result == DialogResult.Yes)
                    {
                        newFileName = String.Format("{0} ({1}).xlsx", fileName, count);
                    }
                }

                string newFilePath = path + "\\" + newFileName;

                string lstID = "";
                if (isBoth)
                {
                    lstID = Min_ID;
                    Min_ID = minID_New;
                    Max_ID = maxID_New;
                }
                else if (isExisting)
                {
                    lstID = Min_ID;
                }
                else
                {
                    lstID = Min_ID;
                }

                //Script
                string queryEX = "SELECT DC_ID, DG_DISCOUNT DISCOUNT_CODE, DG_DISCOUNT_DESC  DISCOUNT_DESCRIPTION,DG_MONTH_MIN MONTH_MIN" +
                   ", DG_MONTH_MAX MONTH_MAX, to_char(trunc(DC_START_DT), 'dd/mm/yyyy')  START_DATE ,to_char(trunc(DC_END_DT), 'dd/mm/yyyy')   END_DATE" +
                    ", DG_ACTIVE_FLAG ACITVE_FLAG, SALE_CHANNEL, SPEED, MARKETING_CODE, ORDER_TYPE, PROVINCE, PRODUCT, DG_GROUPID " +
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
                    "AND DC_ID in (" + lstID + ")" +
                    " ORDER BY DG_ORDER_BY,DG_MONTH_MIN , DG_DISCOUNT,DC_START_DT";

                string queryNew = "SELECT DC_ID, DG_DISCOUNT DISCOUNT_CODE, DG_DISCOUNT_DESC  DISCOUNT_DESCRIPTION,DG_MONTH_MIN MONTH_MIN" +
                       ", DG_MONTH_MAX MONTH_MAX, to_char(trunc(DC_START_DT), 'dd/mm/yyyy')  START_DATE ,to_char(trunc(DC_END_DT), 'dd/mm/yyyy')   END_DATE" +
                        ", DG_ACTIVE_FLAG ACITVE_FLAG, SALE_CHANNEL, SPEED, MARKETING_CODE, ORDER_TYPE, PROVINCE, PRODUCT, DG_GROUPID " +
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
                        "AND DC_ID between '" + Min_ID + "' and '" + Max_ID + "'" +
                        " ORDER BY DG_ORDER_BY,DG_MONTH_MIN , DG_DISCOUNT,DC_START_DT";


                string queryVas = "SELECT DC_ID, VAS_CODE, VAS_NAME, VAS_TYPE, VAS_STATUS, VAS_RULE, to_char(trunc(DC_START_DT), 'dd/mm/yyyy')  START_DATE" +
                         ",to_char(trunc(DC_END_DT), 'dd/mm/yyyy') END_DATE, SALE_CHANNEL, SPEED, MARKETING_CODE ,ORDER_TYPE, PROVINCE, PRODUCT" +
                         ",VAS_PRICE, VAS_CHANNEL, PARENT_VAS_CODE" +
                        " FROM(SELECT * FROM (SELECT DC1.DC_START_DT, DC1.DC_END_DT, DC1.DC_ID, DC1.DC_GROUPID, NVL(PRODUCT, 'ALL') PRODUCT, " +
                        "NVL(SPEED, 'ALL') SPEED, NVL(MARKETING_CODE, 'ALL')MARKETING_CODE, NVL(PROVINCE, 'ALL')PROVINCE, NVL(SALE_CHANNEL, 'ALL')SALE_CHANNEL, " +
                        "NVL(ORDER_TYPE, 'ALL')ORDER_TYPE FROM (SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PRODUCT', DC_VALUE, 'ALL') PRODUCT " +
                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PRODUCT'  AND DC_ACTIVE_FLAG = 'Y')  DC1," +
                        "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'SPEED', DC_VALUE, 'ALL') SPEED " +
                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'SPEED' AND DC_ACTIVE_FLAG = 'Y')  DC2," +
                        "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'MARKETING_CODE', DC_VALUE, 'ALL') MARKETING_CODE " +
                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'MARKETING_CODE' AND DC_ACTIVE_FLAG = 'Y') DC3," +
                        "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PROVINCE', DC_VALUE, 'ALL') PROVINCE " +
                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PROVINCE' AND DC_ACTIVE_FLAG = 'Y') DC4," +
                        "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'SALE_CHANNEL', DC_VALUE, 'ALL') SALE_CHANNEL " +
                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'SALE_CHANNEL' AND DC_ACTIVE_FLAG = 'Y') DC5, " +
                        "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'ORDER_TYPE', DC_VALUE, 'ALL') ORDER_TYPE " +
                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'ORDER_TYPE' AND DC_ACTIVE_FLAG = 'Y') DC6 " +
                        "WHERE DC1.DC_ID = DC2.DC_ID(+) " +
                        "AND DC1.DC_ID = DC3.DC_ID(+) " +
                        "AND DC1.DC_ID = DC4.DC_ID(+) " +
                        "AND DC1.DC_ID = DC5.DC_ID(+) " +
                        "AND DC1.DC_ID = DC6.DC_ID(+))) DISCOUNT_CRITERIA_MAPPING, VAS_PRODUCT WHERE DISCOUNT_CRITERIA_MAPPING.DC_GROUPID = VAS_PRODUCT.VAS_CODE " +
                        "AND upper(VAS_PRODUCT.VAS_STATUS) = 'ACTIVE' AND VAS_CODE LIKE 'VAS%' AND PROD_CATG_CD IS NULL " +
                        "AND DC_ID BETWEEN '" + Min_ID + "' AND '" + Max_ID + "' ORDER BY PARENT_VAS_CODE, VAS_CODE, VAS_RULE ";

                string queryHispeed = "SELECT P.P_ID,P.P_CODE,P.P_NAME,P.ORDER_TYPE,P.STATUS,P.PRODTYPE,P.BUNDLE_CAMPAIGN,C.SALE_CHANNEL,C.START_DATE, " +
                        "C.END_DATE,S.PRICE,S.SPEED_ID DOWNLOAD_SPEED, S.UPLOAD_SPEED,S.MODEM_TYPE,S.DOCSIS_TYPE FROM HISPEED_PROMOTION P, HISPEED_CHANNEL_PROMOTION C, " +
                        "HISPEED_SPEED_PROMOTION S WHERE P.P_ID = C.P_ID AND P.P_ID = S.P_ID AND P.P_ID IN (" + lstID + ") ORDER BY P_ID";

                string campaign = "SELECT * FROM CAMPAIGN_MAPPING WHERE CAMPAIGN_NAME IN (" + lstID + ")";

                //set properties
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                workbook = excelApp.Workbooks.Add(Type.Missing);
                Excel.Worksheet sheet1 = workbook.ActiveSheet as Excel.Worksheet;
                sheet1.Name = "New Criteria";

                Excel.Worksheet sheet2 = workbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing) as Excel.Worksheet;
                sheet2.Name = "Existing Criteria";

                if (isBoth)
                {
                    //New Criteria
                    Write_Data(sheet1, queryNew);

                    //Existing Criteria  
                    Write_Data(sheet2, queryEX);
                }
                else if (isExisting)
                {
                    sheet2.Delete();
                    sheet1.Name = "Existing_Criteria";
                    sheet1.Activate();

                    Write_Data(sheet1, queryEX);
                }
                else
                {
                    string cmdTxt = "";
                    sheet2.Delete();

                    if (Min_ID.StartsWith("DC"))
                    {
                        cmdTxt = queryNew;
                    }
                    else if (Min_ID.StartsWith("VAS"))
                    {
                        cmdTxt = queryVas;
                    }
                    else if (Min_ID.StartsWith("200"))
                    {
                        cmdTxt = queryHispeed;
                        sheet1.Name = "Hispeed Promotion";
                    }
                    else
                    {
                        cmdTxt = campaign;
                        sheet1.Name = "Campaign Mapping";
                    }

                    Write_Data(sheet1, cmdTxt);
                }

                workbook.SaveAs(newFilePath, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Excel.XlSaveAsAccessMode.xlNoChange,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception)
            {
                //throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n" + ex.Message);
            }
            finally
            {
                Cursor.Current = Cursors.Default;

                workbook.Close();
                excelApp.Quit();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }

        /// <summary>
        /// Write data for Imp to excel file
        /// </summary>
        /// <param name="sheet"> Excel worksheet</param>
        /// <param name="query"> Query String</param>
        private void Write_Data(Excel.Worksheet sheet, string query)
        {
            OracleDataAdapter adapter = new OracleDataAdapter(query, ConnectionProd);
            DataTable dt = new DataTable();
            adapter.Fill(dt);

            //Set column heading
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                sheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
            }
            sheet.get_Range("A1", "O1").Interior.Color = Excel.XlRgbColor.rgbSkyBlue;
            sheet.get_Range("A1", "O1").Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;

            //Write data
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (j == 5 || j == 6)
                    {
                        string date = dt.Rows[i][j].ToString();
                        DateTime dDate;
                        if (DateTime.TryParse(date, out dDate))
                        {
                            date = string.Format("{0:dd/MMM/yyyy}", dDate);
                            sheet.Cells[i + 2, j + 1] = date;
                        }
                        else
                        {
                            sheet.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
                        }
                    }
                    else
                    {
                        sheet.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
                    }
                }
            }
            adapter.Dispose();
        }

        /// <summary>
        /// Function for export script insert (only new discount)
        /// </summary>
        /// <param name="path"></param>
        private void ExportScript(string path)
        {
            string name = UR_NO + "_Export_Script";
            string now = DateTime.Now.ToString("dd/MMM/yyyy hh:mm:ss");

            DataTable tableNew = lstTableResult[0];

            //load excel and create a new workbook
            var excelApp = new Excel.Application();

            //set properties
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            excelApp.Workbooks.Add(Type.Missing);

            //single worksheet
            Excel._Worksheet workSheet = excelApp.ActiveSheet;

            //Set column heading
            for (int i = 0; i < tableNew.Columns.Count; i++)
            {
                workSheet.Cells[1, i + 1] = tableNew.Columns[i].ColumnName;
            }

            workSheet.Cells[1, 8] = "CREATED";
            workSheet.Cells[1, 9] = "CREATED_BY";

            int rowCount = tableNew.Rows.Count;

            //Write data into row
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < tableNew.Columns.Count; j++)
                {
                    string id = tableNew.Rows[i][0].ToString();

                    if (id != "")
                    {
                        if (j == 4 || j == 5)
                        {
                            DateTime dDate;
                            string date = tableNew.Rows[i][j].ToString();
                            if (DateTime.TryParse(date, out dDate))
                            {
                                date = string.Format("{0:dd/MMM/yyyy}", dDate);
                                workSheet.Cells[i + 2, j + 1] = date;
                            }
                        }
                        else
                        {
                            workSheet.Cells[i + 2, j + 1] = tableNew.Rows[i][j];
                        }

                        workSheet.Cells[i + 2, 8] = now;
                        workSheet.Cells[i + 2, 9] = USER;

                        string type = tableNew.Rows[i][1].ToString();
                        string value = tableNew.Rows[i][2].ToString();
                        string group = tableNew.Rows[i][3].ToString();
                        string flag = tableNew.Rows[i][6].ToString();

                        string start = tableNew.Rows[i][4].ToString();
                        string end = tableNew.Rows[i][5].ToString();

                        if (start.Contains("-") || start == "null")
                        {
                            start = "";
                        }

                        if (end.Contains("-") || end == "null")
                        {
                            end = "";
                        }

                        workSheet.Cells[i + 2, 11] = "insert into DISCOUNT_CRITERIA_MAPPING values ('" + id + "','" + type + "','" + value + "','" + group +
                     "',to_date('" + start + "','dd/mm/yyyy')," + "to_date('" + end + "', 'dd/mm/yyyy'),'" + flag + "',sysdate,'" + USER + "',null,null);";
                    }

                }
            }

            try
            {
                string dirPath = path;
                string fileName = name + ".xlsx";
                string[] files = Directory.GetFiles(dirPath);
                int count = files.Count(file => { return file.Contains(fileName); });
                string newFileName = fileName;

                //newFileName = (count == 0) ? fileName : String.Format("{0} ({1}).txt", fileName, count);
                if (count > 0)
                {
                    DialogResult result = MessageBox.Show("There is already a file with the same name in this location" + "\r\n" +
                        "Do you want to replace it?", "Replace File", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                    if (result == DialogResult.Yes)
                    {
                        newFileName = String.Format("{0} ({1}).txt", fileName, count);
                    }
                }

                string newFilePath = path + "\\" + newFileName;

                workSheet.SaveAs(newFilePath, Excel.XlFileFormat.xlWorkbookDefault);

            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n" + ex.Message);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
                excelApp.Workbooks.Close();
                excelApp.Quit();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }

        private void ViewResult_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Flag_close == false)
            {
                e.Cancel = true;
            }
            else
            {
                try
                {
                    if (ConnectionTemp != null)
                    {
                        if (ConnectionTemp.State != ConnectionState.Open)
                        {
                            try
                            {
                                ConnectionTemp.Open();
                            }
                            catch
                            {
                                ConnectionTemp.ConnectionString = connString71;
                                ConnectionTemp.Open();
                            }
                        }

                        string type;
                        if (Max_ID.StartsWith("DC"))
                        {
                            type = "Disc";
                        }
                        else
                        {
                            type = "VAS";
                        }

                        string query = "SELECT * FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = '" + type + "' AND COMPLETE_FLAG = 'N'";

                        OracleCommand cmd = new OracleCommand(query, ConnectionTemp);
                        OracleDataReader reader = cmd.ExecuteReader();
                        reader.Read();
                        if (reader.HasRows)
                        {
                            OracleCommand command = new OracleCommand();
                            command.CommandText = "DELETE FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = '" + type + "' AND COMPLETE_FLAG = 'N'";
                            command.ExecuteNonQuery();
                        }

                        reader.Close();
                    }
                }
                catch (Exception)
                {
                    Flag_close = true;
                }
            }
        }

        private void labelClose_Click(object sender, EventArgs e)
        {
            main.Show();
            this.Close();
            MessageBox.Show("Exist!!");
            //comment
        }

        private void ViewResult_SizeChanged(object sender, EventArgs e)
        {
            int w = this.Size.Width;
            int h = this.Size.Height;

            btnExport.Location = new Point(w - 155, h - 110);
            btnInsert.Location = new Point(w - 294, h - 110);

            labelClose.Location = new Point(w - 23, 0);
            btnMinimize.Location = new Point(w - 88, 3);
            btnMaximize.Location = new Point(w - 53, 3);

            dataGridViewResult.Size = new Size(w, h - 280);
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
            if (maximize == 1)
            {
                this.WindowState = FormWindowState.Maximized;
                maximize = 0;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;
                this.Size = new Size(1038, 749);
                maximize = 1;
            }

        }

        private void btnMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            InsertData();

            ExportImp(outputPath);
            //ExportScript(outputPath);

            //Update complete flag & max ID
            if (Max_ID.StartsWith("DC"))
            {
                reserve.UpdateID(ConnectionTemp, Min_ID, Max_ID, "Disc", USER, UR_NO);
            }
            else if (Max_ID.StartsWith("VAS"))
            {
                reserve.UpdateID(ConnectionTemp, Min_ID, Max_ID, "VAS", USER, UR_NO);
            }

        }

        private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            labelProgress.Text = e.ProgressPercentage.ToString();
            //labelProgress.BeginInvoke(new Action(() =>
            //{
            //    labelProgress.Text = e.ProgressPercentage.ToString();
            //}));
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            labelProgress.Visible = false;
            btnExport.Enabled = true;
            btnInsert.Enabled = true;

            string msg = "";
            if (LogFile != "")
            {
                string newMsg = "There is no data that can be inserted as detailed below" + "\r\n" + LogFile;
                string strFilePath = outputPath + "\\log_" + UR_NO + ".txt";
                using (StreamWriter writer = new StreamWriter(strFilePath, true))
                {
                    writer.Write(LogFile);
                }

                msg = "Please check log." + "\r\n" + "There is no data that can be inserted.";
            }
            else
            {
                msg = "Already insert your data into database" + "\r\n" + "\r\n" + "Min ID: " + Min_ID + "\r\n" + "Max ID: " + Max_ID;
            }

            Application.UseWaitCursor = false;
            Cursor.Current = Cursors.Default;

            MessageBox.Show(msg);

        }
    }
}

