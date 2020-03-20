using System;
using System.Collections.Generic;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutomateMappingTool
{
    class ReserveID
    {

        public bool CheckStatus(OracleConnection ConnectionTemp, string type, string implementer, string urNo)
        {
            bool isInProcess = false;
            OracleCommand cmd = null;

            try
            {
                string query = "SELECT * FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = '" + type + "' AND COMPLETE_FLAG = 'N'";

                cmd = new OracleCommand(query, ConnectionTemp);
                OracleDataReader reader = cmd.ExecuteReader();
                reader.Read();
                if (reader.HasRows)
                {
                    string user = reader["USERNAME"].ToString();
                    string typeName = reader["TYPE_NAME"].ToString();

                    if (user == implementer && type == typeName)
                    {
                        string qryDel = "DELETE FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = '" + type + "' AND COMPLETE_FLAG = 'N'";
                        OracleCommand command = new OracleCommand(qryDel, ConnectionTemp);

                        command.ExecuteNonQuery();

                        isInProcess = false;
                        cmd = ConnectionTemp.CreateCommand();

                        using (OracleTransaction transaction = ConnectionTemp.BeginTransaction(IsolationLevel.ReadCommitted))
                        {
                            cmd.Transaction = transaction;

                            try
                            {
                                cmd.CommandText = "INSERT INTO TRUE9_BPT_RESERVE_ID VALUES('" + type + "', 'N', '0', '0', '" + urNo + "', '" + implementer + "', sysdate)";

                                cmd.CommandType = CommandType.Text;

                                cmd.ExecuteNonQuery();
                                transaction.Commit();
                            }
                            catch (Exception ex)
                            {
                                transaction.Rollback();
                                throw new Exception(ex.Message);
                            }
                        }
                    }
                    else
                    {
                        isInProcess = true;
                        MessageBox.Show("UserName : " + user + " is in the process of inserting." + "\r\n" + "Please try again later");
                    }
                }
                else
                {
                    isInProcess = false;
                    cmd = ConnectionTemp.CreateCommand();

                    using (OracleTransaction transaction = ConnectionTemp.BeginTransaction(IsolationLevel.ReadCommitted))
                    {
                        cmd.Transaction = transaction;

                        try
                        {
                            cmd.CommandText = "INSERT INTO TRUE9_BPT_RESERVE_ID VALUES('" + type + "', 'N', '0', '0', '" + urNo + "', '" + implementer + "', sysdate)";

                            cmd.CommandType = CommandType.Text;

                            cmd.ExecuteNonQuery();
                            transaction.Commit();
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            throw new Exception(ex.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot reserve UR into table[TRUE9_BPT_RESERVE_ID] " + "\r\n" + "Error Detail : " + ex.Message);
                ConnectionTemp.Close();
                Environment.Exit(0);
            }

            return isInProcess;
        }

        public int GetMaxID(OracleConnection ConnectionTemp, OracleConnection ConnectionProd, 
            string type, string implementer, string urNo)
        {
            OracleCommand cmd = null;
            int minID = 0;
            int max = 0;
            string prefixID = "";
            string col = "";
            string table = "";

            if (type == "Hispeed")
            {
                prefixID = "20";
                col = "P_ID";
                table = "HISPEED_PROMOTION";
            }
            else if (type == "Disc")
            {
                prefixID = "DC";
                col = "DC_ID";
                table = "DISCOUNT_CRITERIA_MAPPING";
            }
            else
            {
                prefixID = "VAS";
                col = "DC_ID";
                table = "DISCOUNT_CRITERIA_MAPPING";
            }

            string queryMax = "SELECT MAX(" + col + ") FROM " + table + " WHERE " + col + " LIKE '" + prefixID + "%'";
            string queryMax_reserve = "SELECT MAX(MAX_ID) FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = '" + type + "'";

            cmd = new OracleCommand(queryMax, ConnectionProd);
            OracleDataReader readerMax = cmd.ExecuteReader();
            readerMax.Read();

            OracleCommand cmd1 = new OracleCommand(queryMax_reserve, ConnectionTemp);
            OracleDataReader dataReader = cmd.ExecuteReader();
            dataReader.Read();

            if (type == "Hispeed")
            {
                minID = Convert.ToInt32(readerMax[0]) + 1;
                max = Convert.ToInt32(dataReader[0]);
            }
            else
            {
                string minid = Convert.ToString(readerMax[0]).Substring(prefixID.Length);
                string maxid = Convert.ToString(dataReader[0]).Substring(prefixID.Length);
                minID = Convert.ToInt32(minid) + 1;
                max = Convert.ToInt32(maxid);
            }

            if (minID < max)
            {
                MessageBox.Show("There is a conflict ID between production and reserve table[TRUE9_BPT_RESERVE_ID]" + "\r\n"
                    + "Please review and confirm the information");

                string qryDel = "DELETE FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = '" + type + "' AND COMPLETE_FLAG = 'N'";
                OracleCommand command = new OracleCommand(qryDel, ConnectionTemp);
                command.ExecuteNonQuery();

                ConnectionProd.Close();
                ConnectionTemp.Close();

                Environment.Exit(0);
            }

            return minID;
        }

        public void UpdateID(OracleConnection ConnectionTemp, string minID, string maxID, string type, string implementer, string urNO)
        {
            OracleCommand cmd = ConnectionTemp.CreateCommand();

            using (OracleTransaction transaction = ConnectionTemp.BeginTransaction(IsolationLevel.ReadCommitted))
            {
                cmd.Transaction = transaction;

                try
                {
                    cmd.CommandText = "UPDATE TRUE9_BPT_RESERVE_ID SET COMPLETE_FLAG = 'Y', MIN_ID = '" + minID + "', MAX_ID = '" +
                        maxID + "' WHERE TYPE_NAME = '" + type + "' AND UR_NO = '" + urNO + "' AND USERNAME = '" + implementer + "'";

                    cmd.CommandType = CommandType.Text;

                    cmd.ExecuteNonQuery();
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    MessageBox.Show("Cannot reserve MinID into table[TRUE9_BPT_RESERVE_ID]" + "\r\n" +
                        "Please check error message and manual reserve ID into table[TRUE9_BPT_RESERVE_ID]" + "Error Detail : " + ex.Message,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
