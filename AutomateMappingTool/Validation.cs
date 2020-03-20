using System;
using System.Collections.Generic;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutomateMappingTool
{
    class Validation
    {
        DataGridView dataGridView;
        OracleConnection ConnectionProd;
        OracleConnection ConnectionTemp;
        OracleCommand cmd;
        OracleDataReader reader;

        string outputPath, type, message = null, suffixMkt = null, code = null, month = null, channel = null, mkt = null, order = null,
            province = null, effective = null, expire = null, entry = null, install = null, speed = null,
            downSpeed = null, upSpeed = null, uom = null;

        private List<string> lstProv = new List<string>();
        private List<string> lstEff = new List<string>();
        private List<string> lstExp = new List<string>();
        private List<string> lstMonth = new List<string>();
        private List<string> lstUOM = new List<string>();
        private List<string> lstEntry = new List<string>();
        private List<string> lstInstall = new List<string>();
        private List<string> lstProdType = new List<string>();
        private List<string> lstChannelFromDB = new List<string>();
        private List<string> lstChannel = new List<string>();
        private List<string> lstNewChannel = new List<string>();
        private List<string> lstMktCode = new List<string>();
        private List<string> lstOrder = new List<string>();

        private ListBox listBox = new ListBox();

        public Validation(DataGridView dgv, OracleConnection connProd, OracleConnection connTemp, string output)
        {
            this.dataGridView = dgv;
            this.ConnectionProd = connProd;
            this.ConnectionTemp = connTemp;
            this.outputPath = output;
        }
        public List<int> IndexDgv { get; set; } = new List<int>();
        public List<string> InvalidSpeed { get; set; } = new List<string>();

        public List<string> NewChannel
        {
            get
            {
                return lstNewChannel;
            }
            set
            {
                lstNewChannel = value;
            }
        }

        /// <summary>
        /// Process for validate data of requirement file
        /// </summary>
        /// <returns></returns>
        public ListBox Verify(string t)
        {
            type = t;

            InitialValue();

            //Get all channel from DB
            if(lstChannelFromDB.Count <= 0)
            {
                lstChannelFromDB = GetChannelFromDB();
            }

            for (int i = 0; i < dataGridView.RowCount; i++)
            {
                if (type.Equals("VAS"))
                {
                    code = dataGridView.Rows[i].Cells[0].Value.ToString().Trim();
                    channel = dataGridView.Rows[i].Cells[1].Value.ToString().Trim();
                    mkt = dataGridView.Rows[i].Cells[2].Value.ToString().ToUpper();
                    order = dataGridView.Rows[i].Cells[3].Value.ToString();
                    string speed = dataGridView.Rows[i].Cells[5].Value.ToString().ToUpper().Trim();
                    province = dataGridView.Rows[i].Cells[6].Value.ToString().ToUpper();
                    effective = dataGridView.Rows[i].Cells[7].Value.ToString().Trim();
                    expire = dataGridView.Rows[i].Cells[8].Value.ToString().Trim();

                    string uom = GetUOM(speed);
                    speed = Regex.Replace(speed, "[^0-9]", "");

                    if (uom == "-1")
                    {
                        //error not found uom
                        lstUOM.Add("VAS Code : "+code+", MKT Code : " + mkt + ", Speed : " + speed + " not found UOM.");
                        listBox.Items.Add("This speed not found UOM.");
                        IndexDgv.Add(i);

                        hilightRow(type, "speed", i);
                    }
                    else
                    {
                        if (VerifySpeed(i, speed, uom) == -1)
                        {
                            break;
                        }
                    }

                    VerifyProvince(i);
                    VerifyDate(i);

                    if (String.IsNullOrEmpty(channel) == false)
                    {
                        VerifyChannel(i, channel);
                    }
                    else
                    {
                        lstChannel.Add("Vas Code : "+code+", MKT Code : " + mkt + ", Speed : "+speed+" not found sale channel in file(.xlsx)");
                        string msg = "Sale channel is empty";
                        listBox.Items.Add(msg);
                        IndexDgv.Add(i);

                        hilightRow(type, "channel", i);
                    }
                }
                else if (type.Equals("Disc"))
                {
                    code = dataGridView.Rows[i].Cells[0].Value.ToString().Trim();
                    month = dataGridView.Rows[i].Cells[1].Value.ToString();
                    channel = dataGridView.Rows[i].Cells[2].Value.ToString().Trim();
                    mkt = dataGridView.Rows[i].Cells[3].Value.ToString().ToUpper();
                    order = dataGridView.Rows[i].Cells[4].Value.ToString();
                    string speed = dataGridView.Rows[i].Cells[6].Value.ToString().ToUpper().Trim();
                    province = dataGridView.Rows[i].Cells[7].Value.ToString().ToUpper();
                    effective = dataGridView.Rows[i].Cells[8].Value.ToString();
                    expire = dataGridView.Rows[i].Cells[9].Value.ToString();

                    if(mkt.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                    {
                        if(speed.Equals("ALL", StringComparison.OrdinalIgnoreCase) == false)
                        {
                            string uom = GetUOM(speed);
                            speed = Regex.Replace(speed, "[^0-9]", "");

                            if (uom == "-1")
                            {
                                //error not found uom
                                lstUOM.Add("Discount Code : " + code + ", Month : " + month + ", MKT Code : " + mkt +
                                    ", Speed : " + speed + " not found UOM.");
                                listBox.Items.Add("This speed not found UOM.");
                                IndexDgv.Add(i);

                                hilightRow(type, "speed", i);
                            }
                        }     
                    }
                    else
                    {
                        string uom = GetUOM(speed);
                        speed = Regex.Replace(speed, "[^0-9]", "");

                        if (uom == "-1")
                        {
                            //error not found uom
                            lstUOM.Add("Discount Code : " + code + ", Month : " + month + ", MKT Code : " + mkt +
                                ", Speed : " + speed + " not found UOM.");
                            listBox.Items.Add("This speed not found UOM.");
                            IndexDgv.Add(i);

                            hilightRow(type, "speed", i);
                        }
                        else
                        {
                            if (VerifySpeed(i, speed, uom) == -1)
                            {
                                break;
                            }
                        }
                    }
                   
                    VerifyGroupMonth(i);
                    VerifyProvince(i);
                    VerifyDate(i);

                    if (String.IsNullOrEmpty(channel) == false)
                    {
                        VerifyChannel(i, channel);
                    }
                    else
                    {
                        lstChannel.Add("Discount Code : " + code + ", MKT Code : " + mkt + ", Month : "+month+
                            ", Speed : " + speed + " not found sale channel in file(.xlsx)");
                        string msg = "Sale channel is empty";
                        listBox.Items.Add(msg);
                        IndexDgv.Add(i);

                        hilightRow(type, "channel", i);

                    }
                }
                else//Hispeed
                {
                    mkt = dataGridView.Rows[i].Cells[0].Value.ToString().ToUpper();
                    string speed = dataGridView.Rows[i].Cells[1].Value.ToString().ToUpper().Trim();
                    string order = dataGridView.Rows[i].Cells[4].Value.ToString().Trim();

                    //speed
                    if (speed.Contains('/'))
                    {
                        string[] spSpeed = speed.Split('/');
                        string downSpeed = spSpeed[0].Trim();
                        string upSpeed = spSpeed[1].Trim();

                        string uom = GetUOM(downSpeed);
                        downSpeed = Regex.Replace(downSpeed, "[^0-9]", "");

                        if (uom  == "-1")
                        {
                            uom = GetUOM(upSpeed);

                            if (uom == "-1")
                            {
                                //error not found uom
                                lstUOM.Add("MKT Code : " + mkt + ", Speed : " + speed + " not found UOM.");
                                listBox.Items.Add("This speed not found UOM.");
                                IndexDgv.Add(i);

                                hilightRow(type, "speed", i);
                            }
                            else
                            {
                                if(VerifySpeed(i, downSpeed, uom) == -1)
                                {
                                    break;
                                }
                            }
                        }
                        else
                        {
                            if (VerifySpeed(i, downSpeed, uom) == -1)
                            {
                                break;
                            }
                        }
                    }
                    else
                    {
                        InvalidSpeed.Add("MKT Code : "+mkt+" Speed : "+speed+" is not supported format speed.");
                        listBox.Items.Add("This speed format is not supported.");
                        IndexDgv.Add(i);

                        hilightRow(type, "speed", i);
                    }

                    //Order type
                    VerifyOrderType(i, order);




                    channel = dataGridView.Rows[i].Cells[5].Value.ToString().Trim();
                    channel = Regex.Replace(channel, "ALL", "DEFAULT", RegexOptions.IgnoreCase);



                    effective = dataGridView.Rows[i].Cells[8].Value.ToString();
                    expire = dataGridView.Rows[i].Cells[9].Value.ToString();
                    entry = dataGridView.Rows[i].Cells[10].Value.ToString().Trim();
                    install = dataGridView.Rows[i].Cells[11].Value.ToString().Trim();


                    if (String.IsNullOrEmpty(channel) == false)
                    {
                        VerifyChannel(i, channel);
                    }
                    else
                    {
                        if (String.IsNullOrEmpty(expire))
                        {
                            lstChannel.Add("MKT Code : " + mkt + ", Speed : " + speed + " not found sale channel in file(.xlsx) "+
                                "and expire date is null");
                            string msg = "Sale channel is empty and expire date is null";
                            listBox.Items.Add(msg);
                            IndexDgv.Add(i);

                            hilightRow(type, "channel", i);
                        }
                    }

                 

                }
            }

            if (InvalidSpeed.Count > 0)
            {
                message += "[" + DateTime.Now.ToString() + "] >>Mismatch speed between suffix MKT Code and download speed<<" + "\r\n";

                foreach (string msg in InvalidSpeed)
                {
                    message += "    " + msg + "\r\n";
                }
            }
            
            if (lstProv.Count > 0)
            {
                message += "[" + DateTime.Now.ToString() + "] " + "\r\n";

                foreach (string msg in lstProv)
                {
                    message += "    " + msg + "\r\n";
                }
            }

            if (lstEff.Count > 0)
            {
                message += "[" + DateTime.Now.ToString() + "] >>Wrong effective date format or effective is null<<" + "\r\n";

                foreach (string msg in lstEff)
                {
                    message += "    " + msg + "\r\n";
                }
            }

            if (lstExp.Count > 0)
            {
                message += "[" + DateTime.Now.ToString() + "] >>Wrong expire date format<<" + "\r\n";

                foreach (string msg in lstExp)
                {
                    message += "    " + msg + "\r\n";
                }
            }

            if (lstMonth.Count > 0)
            {
                message += "[" + DateTime.Now.ToString() + "] >>Invalid GroupID<<" + "\r\n";

                foreach (string msg in lstMonth)
                {
                    message += "    " + msg + "\r\n";
                }
            }

            if (message != "")
            {
                message += "\r\n" + "***********************************************" + "\r\n";

                string strFilePath = outputPath + "\\Validation_Log" + ".txt";
                using (StreamWriter writer = new StreamWriter(strFilePath, true))
                {
                    writer.Write(message);
                }
            }
            return listBox;
        }

        private int VerifySpeed(int index, string speed, string uom)
        {
            string msg;
            int status = 0;

            if (mkt.Contains("-"))
            {
                string[] lstmkt = mkt.Split('-');
                string suffixMkt = lstmkt[1].Trim();
                int speedID;

                if (suffixMkt.EndsWith("G"))
                {
                    suffixMkt = suffixMkt.Substring(0, suffixMkt.Length - 1);
                    suffixMkt = (Convert.ToInt32(suffixMkt) * 1000).ToString();
                }

                if (int.TryParse(suffixMkt, out _))
                {
                    speedID = Convert.ToInt32(suffixMkt);

                    string query = "SELECT SPEED_ID, SPEED_DESC FROM HISPEED_SPEED WHERE SPEED_ID = " + speedID;

                    cmd = new OracleCommand(query, ConnectionProd);
                    reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        reader.Read();
                        int speed_desc = Convert.ToInt32(reader["SPEED_DESC"]);

                        if (uom == "G")
                        {
                            speed_desc = speed_desc / 1024000;
                        }
                        else if (uom == "M")
                        {
                            speed_desc = speed_desc / 1024;
                        }

                        suffixMkt = Convert.ToString(speed_desc);

                        if (speed.Equals(suffixMkt) == false)
                        {
                            if (type == "Disc")
                            {
                                msg = code + " Month : " + month + ", MKT Code : " + mkt + " mismatch speed between suffix MKTCode and "
                                    + "download speed " + speed;
                            }
                            else
                            {
                                msg = code + " MKT Code : " + mkt + " mismatch speed between suffix MKTCode and "
                                    + "download speed " + speed;
                            }

                            InvalidSpeed.Add(msg);
                            listBox.Items.Add("Mismatch speed between suffix MKTCode "+mkt+" and speed "+speed);
                            IndexDgv.Add(index);

                            hilightRow(type, "speed", index);
                        }
                    }
                    else
                    {
                        string message = "Do you want to insert new speed " + speed + "?" + "\r\n" +
                            "SPEED_ID: " + speedID + " SPEED_DESC: " + (speedID * 1024) + " SPEED_DETAIL: " +
                            (speedID * 1024) + "K SPEED_LOCAL: " + (speedID * 1024) + "\r\n" + "\r\n" +
                            "       YES = insert new speed" +
                            "       NO = Do not insert new speed" +
                            "       Cancel = Stop process";

                        DialogResult dialogResult = MessageBox.Show(message, "Confirmation", MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Question);

                        if (dialogResult == DialogResult.Yes)
                        {
                            //insert new speed
                            OracleCommand command = ConnectionProd.CreateCommand();
                            OracleTransaction transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted);
                            command.Transaction = transaction;

                            try
                            {
                                string sql = "INSERT INTO HISPEED_SPEED VALUES (" + speedID + ",'" + (speedID * 1024) + "','" +
                                    (speedID * 1024) + "K" + "','" + (speedID * 1024) + "')";

                                command.CommandText = sql;
                                command.ExecuteNonQuery();

                                transaction.Commit();
                            }
                            catch (Exception)
                            {
                                transaction.Rollback();
                                MessageBox.Show("Cannot insert the new speed." + "\r\n" + "Please try again!!", "Error",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                                status = -1;
                            }
                        }
                        else if (dialogResult == DialogResult.No)
                        {
                            if (type == "Disc")
                            {
                                msg = code + " Month : " + month + ", MKT Code : " + mkt + " not found speed_id "
                                    + speedID + " in database[HISPEED_SPEED]";
                            }
                            else
                            {
                                msg = code + " MKT Code : " + mkt + " not found speed_id " + speedID + " in database[HISPEED_SPEED]";
                            }

                            InvalidSpeed.Add(msg);
                            listBox.Items.Add("Not found speed_id : " + speedID + " in database[HISPEED_SPEED]");
                            IndexDgv.Add(index);

                            hilightRow(type, "speed", index);
                        }
                        else
                        {
                            //stop process
                            status = -1;
                        }
                               
                    }

                    reader.Close();
                }
                else
                {
                    if (type == "Disc")
                    {
                        msg = code + " Month : " + month + ", MKT Code : " + mkt + " is invalid.";
                    }
                    else
                    {
                        msg = code + " MKT Code : " + mkt + " is invalid.";

                        if (msg.StartsWith(""))
                        {
                            msg.TrimStart();
                        }
                    }

                    InvalidSpeed.Add(msg);
                    listBox.Items.Add("This MKT Code is invalid");
                    IndexDgv.Add(index);

                    hilightRow(type, "mkt", index);
                }
            }
            else
            {
                if (type == "Disc")
                {
                    msg = code + " Month : " + month + ", MKT Code : " + mkt + " is invalid.";
                }
                else
                {
                    msg = code + " MKT Code : " + mkt + " is invalid.";

                    if(msg.StartsWith(""))
                    {
                        msg.TrimStart();
                    }
                }

                InvalidSpeed.Add(msg);
                listBox.Items.Add("This MKT Code is invalid");
                IndexDgv.Add(index);

                hilightRow(type, "mkt", index);
            }

            return status;
        }
        private void VerifyOrderType(int index, string order)
        {
            char[] toChar = order.ToCharArray();
            char[] specialChar = new char[toChar.Length];
            int count = 0;

            for (int i = 0; i < toChar.Length; i++)
            {
                if (!Char.IsLetterOrDigit(toChar[i]))
                {
                    specialChar[count] = toChar[i];
                    count++;
                }

                Array.Resize(ref specialChar, count);
            }

            if(count > 0)
            {
                if (specialChar.Contains(',') == false)
                {
                    lstOrder.Add("MKT Code : " + mkt + " order type doesn't contain characters ','");
                    listBox.Items.Add("This order type doesn't contain ','");
                    IndexDgv.Add(index);

                    hilightRow(type, "order", index);
                }
            }
        }

        private bool VerifyUOM(int index)
        {
            int num;
            bool hasUOM;


            if ((int.TryParse(downSpeed, out num)) && (int.TryParse(upSpeed, out num)))
            {
                string msg = mkt + ", Speed : " + "Don't have UOM of Speed";
                lstUOM.Add(msg);
                listBox.Items.Add(msg);
                IndexDgv.Add(index);

                hilightRow(type, "speed", index);

                hasUOM = false;
            }
            else
            {
                hasUOM = true;
            }

            return hasUOM;
        }
        private string GetUOM(string speed)
        {
            string uom = null;

            if (int.TryParse(speed, out _) == false)
            {
                uom = Regex.Replace(speed, "[0-9]", "");
            }
            else
            {
                uom = "-1";
            }

            return uom;
        }

        private void VerifyProvince(int index)
        {
            //Check Province
            if (province.Contains(","))
            {
                if (province.Contains("ALL"))
                {
                    string msg = code + ", " + month + " month, MKTCode : " + mkt + ", Order : " + order +
                        ", province : " + province + " >>> Conflict province";
                    hilightRow(type, "province", index);
                    lstProv.Add(msg);
                    listBox.Items.Add(msg);
                    IndexDgv.Add(index);
                }
                else
                {
                    string[] lstProvince = province.Split(',');

                    foreach (string prov in lstProvince)
                    {
                        bool hasRow = GetProvince(prov);
                        if (hasRow == false)
                        {
                            string msg = code + ", " + month + " month, MKTCode : " + mkt + ", Order : " + order + ", province : " + province;
                            lstProv.Add(msg);
                            listBox.Items.Add(msg);
                            IndexDgv.Add(index);

                            hilightRow(type, "province", index);
                        }
                    }
                }
            }
            else
            {
                bool hasRow;
                if (province.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    province = "ALL";
                    hasRow = true;
                }
                else
                {
                    hasRow = GetProvince(province);
                }

                if (hasRow == false)
                {
                    string msg = code + ", " + month + " month, MKTCode : " + mkt + ", Order : " + order + ", province : " + province;
                    lstProv.Add(msg);
                    listBox.Items.Add(msg);
                    IndexDgv.Add(index);

                    hilightRow(type, "province", index);
                }
            }
        }

        //Check province in DB
        private bool GetProvince(string province)
        {
            string queryProv = "SELECT * FROM DISCOUNT_CRITERIA_PROVINCE WHERE DP_TYPE = '" + province.Trim() + "'";
            bool hasRow = true;
            try
            {
                cmd = new OracleCommand(queryProv, ConnectionProd);
                reader = cmd.ExecuteReader();
                reader.Read();
                if (reader.HasRows == false)
                {
                    hasRow = false;
                }

                reader.Close();
            }
            catch (Exception ex)
            {
                string log = "[" + DateTime.Now.ToString() + "]" + "Cannot searching province from database(DISCOUNT_CRITERIA_PROVINCE)." +
                    "\r\n" + "Error message : " + ex.Message;

                hasRow = false;
            }

            return hasRow;
        }

        private void VerifyDate(int index)
        {
            ChangeFormat chgFormat = new ChangeFormat();
            //Check Date
            string dateEff = chgFormat.FormatDate(effective);

            if (dateEff == "Invalid")
            {
                hilightRow(type, "effective", index);

                string msg = code + ", " + month + " month, MKTCode : " + mkt;
                lstEff.Add(msg);

                listBox.Items.Add(msg);
                IndexDgv.Add(index);

            }

            if (expire.Equals("-"))
            {
                expire = String.Empty;
            }

            if (String.IsNullOrEmpty(expire) == false)
            {
                string date = chgFormat.FormatDate(expire);
                if (date == "Invalid")
                {
                    hilightRow(type, "expire", index);

                    string msg = code + ", " + month + " month, MKTCode : " + mkt;
                    lstExp.Add(msg);
                    listBox.Items.Add(msg);
                    IndexDgv.Add(index);

                }
            }
        }

        private void VerifyGroupMonth(int index)
        {
            List<Dictionary<string, string>> lstRangeM = FormatGroupMonth(month);
            Dictionary<string, string> dicresultM = new Dictionary<string, string>();
            Dictionary<string, string> dicGroupID = new Dictionary<string, string>();

            for (int num = 0; num < lstRangeM.Count; num++)
            {
                dicresultM = lstRangeM[num];
                string minMonth = dicresultM["min" + num];
                string maxMonth = dicresultM["max" + num];

                //Check GroupID
                string queryGroup = "SELECT * FROM discount_criteria_group where DG_DISCOUNT = '" + code + "' and dg_month_min = " +
                    minMonth + " and dg_month_max = " + maxMonth;

                try
                {
                    cmd = new OracleCommand(queryGroup, ConnectionProd);
                    reader = cmd.ExecuteReader();
                    reader.Read();
                    if (reader.HasRows)
                    {
                        string distinct = code + minMonth + maxMonth;
                        string dgGroup = reader["DG_GROUPID"].ToString();

                        if (dicGroupID.ContainsKey(distinct) == false)
                        {
                            dicGroupID.Add(distinct, dgGroup);
                        }
                    }
                    else
                    {
                        hilightRow(type, "month", index);

                        string msg = code + ", month : " + minMonth + "-" + maxMonth + ", MKTCode : " + mkt;
                        lstMonth.Add(msg);
                        listBox.Items.Add(msg);
                        IndexDgv.Add(index);
                    }

                    reader.Close();
                }
                catch (Exception ex)
                {
                    string log = "[" + DateTime.Now.ToString() + "]" + "Cannot searching discount group(DG_GROUPID) from database(DISCOUNT_CRITERIA_GROUP)." +
                        "\r\n" + "Error message : " + ex.Message;
                }

            }
        }

        public List<Dictionary<string, string>> FormatGroupMonth(string month)
        {
            List<Dictionary<string, string>> lstRangeM = new List<Dictionary<string, string>>();
            Dictionary<string, string> dic;

            if (month.Equals("ตลอดอายุการใช้งาน"))
            {
                month = "-1";
            }

            if (month.Contains(","))
            {
                string[] lstMonth = month.Split(',');

                for (int j = 0; j < lstMonth.Length; j++)
                {
                    dic = new Dictionary<string, string>();

                    if (lstMonth[j].StartsWith("-1"))
                    {
                        if (lstMonth[j].Length > 2)
                        {
                            lstMonth[j] = lstMonth[j].Substring(2);

                            string[] lstRange = lstMonth[j].Split('-');
                            dic.Add("min" + j, "-1");
                            dic.Add("max" + j, lstRange[1]);
                            lstRangeM.Add(dic);
                        }
                        else
                        {
                            dic.Add("min" + j, lstMonth[j]);
                            dic.Add("max" + j, lstMonth[j]);
                            lstRangeM.Add(dic);
                        }
                    }
                    else
                    {
                        if (lstMonth[j].Contains("-"))
                        {
                            string[] lstRange = lstMonth[j].Split('-');
                            dic.Add("min" + j, lstRange[0]);
                            dic.Add("max" + j, lstRange[1]);
                            lstRangeM.Add(dic);
                        }
                        else
                        {
                            dic.Add("min" + j, lstMonth[j]);
                            dic.Add("max" + j, lstMonth[j]);
                            lstRangeM.Add(dic);
                        }
                    }
                }
            }
            else
            {
                dic = new Dictionary<string, string>();

                if (month.StartsWith("-1"))
                {
                    if (month.Length > 2)
                    {
                        month = month.Substring(2);
                        string[] lstRange = month.Split('-');

                        dic.Add("min0", "-1");
                        dic.Add("max0", lstRange[1]);
                    }
                    else
                    {
                        dic.Add("min0", month);
                        dic.Add("max0", month);
                    }
                }
                else
                {
                    if (month.Contains("-"))
                    {
                        string[] lstRange = month.Split('-');
                        dic.Add("min0", lstRange[0]);
                        dic.Add("max0", lstRange[1]);
                    }
                    else
                    {
                        dic.Add("min0", month);
                        dic.Add("max0", month);
                    }
                }

                lstRangeM.Add(dic);
            }

            return lstRangeM;
        }

        private void VerifyContract(int index)
        {
            string queryEnt = "SELECT * FROM TRUE9_BPT_CONTRACT WHERE ENTRY = '" + entry + "'";
            string queryIns = "SELECT * FROM TRUE9_BPT_CONTRACT WHERE INSTALL = '" + install + "'";

            //Entry Code
            cmd = new OracleCommand(queryEnt, ConnectionTemp);
            reader = cmd.ExecuteReader();
            if (reader.HasRows == false)
            {
                string msg = "Entry Code " + entry + " of " + mkt + " ,Not found in table >> TRUE9_BPT_CONTRACT";
                lstEntry.Add(msg);
                listBox.Items.Add(msg);
                IndexDgv.Add(index);

                hilightRow(type, "entry", index);

                reader.Close();
            }

            //Install code
            cmd = new OracleCommand(queryIns, ConnectionTemp);
            reader = cmd.ExecuteReader();
            if (reader.HasRows == false)
            {
                string msg = "Install Code " + install + " of " + mkt + " ,Not found in table >> TRUE9_BPT_CONTRACT";

                lstInstall.Add(msg);
                listBox.Items.Add(msg);
                IndexDgv.Add(index);

                hilightRow(type, "install", index);

                reader.Close();
            }
        }

        private void VerifyProdType(int index, string prod)
        {
            string prefixMKT;
            if (mkt.StartsWith("TRL"))
            {
                prefixMKT = mkt.Substring(0, 5);
            }
            else
            {
                prefixMKT = mkt.Substring(0, 2);
            }

            string query = "SELECT * FROM TRUE9_BPT_HISPEED_PRODTYPE WHERE MKT = '" + prefixMKT + "'";
            cmd = new OracleCommand(query, ConnectionTemp);
            reader = cmd.ExecuteReader();

            if (reader.HasRows == false)
            {
                string msg = "The prefix " + prefixMKT + " of " + mkt + " ,Not found in table >> TRUE9_BPT_HISPEED_PRODTYPE";
                lstProdType.Add(msg);
                listBox.Items.Add(msg);
                IndexDgv.Add(index);

                hilightRow(type, "mkt", index);

                reader.Close();
            }
        }

        /// <summary>
        /// Check channel from file compare with channel in database
        /// </summary>
        /// <param name="channel"></param>
        /// <returns></returns>
        private void VerifyChannel(int index, string channel)
        {
            string newChannel = "";

            if (channel.Contains(","))
            {
                string[] lstCh = channel.Split(',');

                foreach (string val in lstCh)
                {
                    val.Trim();

                    if (val.Equals("ALL", StringComparison.OrdinalIgnoreCase) ||
                        val.Equals("DEFAULT", StringComparison.OrdinalIgnoreCase))
                    {
                        //Conflict channel
                        lstChannel.Add("MKT Code : " + mkt + " found conflicting sale channel between 'ALL' and '" + val + "'");
        
                        string msg = "Sale Channel : " + val + " included with channel 'ALL/DEFAULT' in same MKT code";
                        listBox.Items.Add(msg);
                        IndexDgv.Add(index);

                        hilightRow(type, "channel", index);
                    }
                    else
                    {
                        if (lstChannelFromDB.Contains(val) == false)
                        {
                            lstChannel.Add("MKT Code : " + mkt + " not found sale channel " + val + " in database [HISPEED_CHANNEL_PROMOTION]");

                            string msg = "Sale Channel : " + val + " Not found in database.";
                            listBox.Items.Add(msg);
                            IndexDgv.Add(index);

                            hilightRow(type, "channel", index);

                            newChannel += val + ",";
                        }
                    }
                }
            }
            else
            {
                if (channel.Equals("ALL", StringComparison.OrdinalIgnoreCase) == false)
                {
                    if (lstChannelFromDB.Contains(channel) == false)
                    {
                        lstChannel.Add("MKT Code : "+mkt+" Not found sale channel "+channel+ " in database [HISPEED_CHANNEL_PROMOTION]");

                        string msg = "Sale Channel : " + channel + " Not found in database.";
                        listBox.Items.Add(msg);
                        IndexDgv.Add(index);

                        hilightRow(type, "channel", index);

                        newChannel += channel + ",";
                    }
                }
            }

            if(newChannel.StartsWith(","))
            {
                newChannel = newChannel.Substring(1);
            }

            lstNewChannel.Add(mkt + "|" + speed + "|" + dataGridView.Rows[index].Cells[2].Value + "|" + dataGridView.Rows[index].Cells[3].Value + "|" +
                order + "|" + newChannel + "|" + dataGridView.Rows[index].Cells[6].Value + "|" + dataGridView.Rows[index].Cells[7].Value + "|" +
                dataGridView.Rows[index].Cells[8].Value + "|" + effective + "|" + expire + "|" + entry + "|" + install);
        }

        /// <summary>
        /// Get channel from DB
        /// </summary>
        /// <returns>List of channel in DB</returns>
        public List<string> GetChannelFromDB()
        {
            List<string> list = new List<string>();

            try
            {
                //Get all channel in DB
                string query = "SELECT DISTINCT(SALE_CHANNEL) FROM HISPEED_CHANNEL_PROMOTION";
                cmd = new OracleCommand(query, ConnectionProd);
                reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    list.Add(reader["SALE_CHANNEL"].ToString());
                }

                reader.Close();
            }
            catch (Exception)
            {
                string msg = "Cannot get sale channel from database. Please click validate file again or check internet connection.";

                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                list = new List<string>();
            }

            return list;
        }

        /// <summary>
        /// Clear selected row 
        /// </summary>
        private void InitialValue()
        {
            InvalidSpeed.Clear();
            lstProv.Clear();
            lstEff.Clear();
            lstExp.Clear();
            lstMonth.Clear();
            lstUOM.Clear();
            lstEntry.Clear();
            lstProdType.Clear();
            lstInstall.Clear();
            lstChannel.Clear();
            lstNewChannel.Clear();
            listBox.Items.Clear();
            IndexDgv = new List<int>();

            //Clear selection
            dataGridView.ClearSelection();
            for (int i = 0; i < dataGridView.RowCount; i++)
            {
                for (int j = 0; j < dataGridView.ColumnCount; j++)
                {
                    dataGridView.Rows[i].Cells[j].Style.BackColor = Color.Empty;
                }
            }
        }

        /// <summary>
        /// Hilight mistake row
        /// </summary>
        /// <param name="indexRow"></param>
        /// <param name="indexCol"></param>
        private void hilightRow(string type, string key, int indexRow)
        {
            Dictionary<string, int> indexDisc = new Dictionary<string, int>
            { {"month",1}, {"channel",2 },{"mkt",3},{"order",4},{"speed",6},{"province",7},{"effective",8},{"expire",9} };

            Dictionary<string, int> indexVas = new Dictionary<string, int>
            {{"channel",1 },{"mkt",2},{"order",3},{"speed",5},{"province",6},{"effective",7},{"expire",8} };

            Dictionary<string, int> indexHis = new Dictionary<string, int>
            {{"mkt",0 },{"speed",1},{"order",4},{"channel",5},{"effective",8},{"expire",9},{"entry",10}, {"install",11} };

            if (type.Equals("VAS"))
            {
                int indexCol = indexVas[key];
                dataGridView.Rows[indexRow].Cells[indexCol].Style.BackColor = Color.Red;
            }
            else if (type.Equals("Disc"))
            {
                int indexCol = indexDisc[key];
                dataGridView.Rows[indexRow].Cells[indexCol].Style.BackColor = Color.Red;
            }
            else
            {
                int indexCol = indexDisc[key];
                dataGridView.Rows[indexRow].Cells[indexCol].Style.BackColor = Color.Red;
            }

        }
    }
}
