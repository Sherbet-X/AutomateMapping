 private void Verify_Data()
        {
            ChangeFormat changeFormateDT = new ChangeFormat();
            indexDgv = new List<int>();
            listBox1.Items.Clear();
            listBox1.Hide();

            string ValidateLog = "", invalidSpeed = "", invalidContractCode = "", invalidMKTCode = "",
             invalidChannel = "", invalidUOM = "", invalidEndDate = "", invalidPName = "", invalidEffDate = "";
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
                            string effective = changeFormateDT.formatDate(dtEff);
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

                        string expire = changeFormateDT.formatDate(dtEx);

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


/////////////////
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