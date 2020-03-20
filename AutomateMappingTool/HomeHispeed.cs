using System;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace MappingDiscount
{
    public partial class HomeHispeed : Form
    {
        private OracleConnection ConnectionProd;
        private string implementer;
        private string FILENAME = "";
        private string FILENAME_DESC = "";
        private string outputPath;
        private bool isPromo = true;
        private string folder;

        Excel.Application xlApp;

        MainHispeed mainHispeed;
        FormLogin formLogin;

        //For move form
        int mov, movX, movY;

        [Obsolete]
        public HomeHispeed(OracleConnection con, string user)
        {
            InitializeComponent();

            ConnectionProd = con;
            implementer = user;

            formLogin = new FormLogin();
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

        private void CoverHispeed_Load(object sender, System.EventArgs e)
        {
            txtImp.Text = implementer;

            if (FILENAME == "")
            {
                txtInput.Clear();
            }

            if (FILENAME_DESC == "")
            {
                txtInputDesc.Clear();
            }
        }

        private void btnNext_Click(object sender, System.EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (txtUrNo.Text == "")
            {
                MessageBox.Show("Please input Ur.NO#");
            }
            else if (txtImp.Text == "")
            {
                MessageBox.Show("Please input Implementer");
            }
            else if (txtInput.Text == "")
            {
                MessageBox.Show("Please input requirement file");
            }
            else if (txtOutput.Text == "")
            {
                MessageBox.Show("Please select output path");
            }
            else
            {
                mainHispeed = new MainHispeed(this, ConnectionProd, FILENAME, FILENAME_DESC, folder,
                    implementer, txtUrNo.Text.Trim(), isPromo);

                mainHispeed.Show();
            }

            Cursor.Current = Cursors.Default;
        } 

        /// <summary>
        /// Open file requirment
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOpen_Click(object sender, System.EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            FILENAME = "";
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FILENAME = openFileDialog1.FileName;
                txtInput.Text = FILENAME;
            }

            Cursor.Current = Cursors.Default;
        }

        /// <summary>
        /// Open file for set MKT code
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOpenDesc_Click(object sender, System.EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            FILENAME_DESC = "";
            openFileDialog2.Filter = "Excel Files|*.xls;*.xlsx";

            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                FILENAME_DESC = openFileDialog2.FileName;
                txtInputDesc.Text = FILENAME_DESC;
            }

            Cursor.Current = Cursors.Default;
        }

        private void panel5_MouseDown(object sender, MouseEventArgs e)
        {
            mov = 1;
            movX = e.X;
            movY = e.Y;
        }

        private void panel5_MouseMove(object sender, MouseEventArgs e)
        {
            if (mov == 1)
            {
                this.SetDesktopLocation(MousePosition.X - movX, MousePosition.Y - movY);
            }
        }

        private void panel5_MouseUp(object sender, MouseEventArgs e)
        {
            mov = 0;
        }

        private void btnMapping_MouseClick(object sender, MouseEventArgs e)
        {
            btnMapping.FlatStyle = FlatStyle.Flat;
            btnMapping.FlatAppearance.BorderColor = Color.White;
            btnMapping.FlatAppearance.BorderSize = 1;
            btnMapping.BackColor = Color.FromArgb(202, 1, 68);
            btnMapping.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 0, 73);

            btnCampaign.BackColor = Color.FromArgb(224, 2, 107);
            btnCampaign.FlatAppearance.BorderSize = 0;
        }

        private void btnCampaign_MouseClick(object sender, MouseEventArgs e)
        {
            btnCampaign.FlatStyle = FlatStyle.Flat;
            btnCampaign.FlatAppearance.BorderColor = Color.White;
            btnCampaign.FlatAppearance.BorderSize = 1;
            btnCampaign.BackColor = Color.FromArgb(202, 1, 68);
            btnCampaign.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 0, 73);

            btnMapping.BackColor = Color.FromArgb(224, 2, 107);
            btnMapping.FlatAppearance.BorderSize = 0;
        }

        private void labelClose_Click(object sender, System.EventArgs e)
        {
            this.Close();
            formLogin.Show();
        }

        private void btnBack_Click(object sender, System.EventArgs e)
        {
            this.Close();
            formLogin.Show();
        }

        private void btnCampaign_Click(object sender, System.EventArgs e)
        {
            isPromo = false;
        }
        private void btnMapping_Click(object sender, System.EventArgs e)
        {
            isPromo = true;
        }

        private void CampaignMapping(string path)
        {
            Cursor.Current = Cursors.WaitCursor;
            string log = "";
            string logSystem = "";
            xlApp = new Microsoft.Office.Interop.Excel.Application();

            try
            {
                string request = "";
                string campaign = "";
                string TOLPack = "";
                string TOLDisc = "";
                Nullable<double> TVS = null;
                Nullable<double> TVSDisc = null;
                string lstTOLPack = "";

                //Read data from file input
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(FILENAME);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets["Campaign Mapping"];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                for (int index = 3; index <= xlRange.Rows.Count; index++)
                {
                    double result;
                    double result2;

                    request = xlWorksheet.Range["B" + index].Value2;
                    campaign = xlWorksheet.Range["C" + index].Value2;
                    TOLPack = xlWorksheet.Range["D" + index].Value2;
                    TOLDisc = xlWorksheet.Range["E" + index].Value2;

                    if (double.TryParse(Convert.ToString(xlWorksheet.Range["F" + index].Value2), out result))
                    {
                        TVS = result;
                    }
                    else
                    {
                        TVS = null;
                    }

                    if (double.TryParse(Convert.ToString(xlWorksheet.Range["G" + index].Value2), out result2))
                    {
                        TVSDisc = result2;
                    }
                    else
                    {
                        TVSDisc = null;
                    }

                    if (request != "" && campaign != "" && TVS != null)
                    {
                        if (request == "Insert")
                        {
                            //check data exists
                            string query = "SELECT * FROM CAMPAIGN_MAPPING WHERE CAMPAIGN_NAME = '" + request + "'AND TOL_PACKAGE = '" + TOLPack + "' AND TVS_PACKAGE = '" +
                                TVS + "' AND STATUS = 'A'";

                            OracleCommand cmd1 = new OracleCommand(query, ConnectionProd);
                            OracleDataReader reader = cmd1.ExecuteReader();

                            if (reader.HasRows)
                            {
                                MessageBox.Show("The data already exists in the database.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                            else
                            {
                                //Insert data
                                OracleCommand command = ConnectionProd.CreateCommand();

                                OracleTransaction transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted);
                                command.Transaction = transaction;

                                try
                                {
                                    command.CommandText = "INSERT INTO CAMPAIGN_MAPPING (CAMPAIGN_NAME,TOL_PACKAGE,TOL_DISCOUNT,TVS_PACKAGE,TVS_DISCOUNT,STATUS) VALUES ('" +
                                        campaign + "','" + TOLPack + "','" + TOLDisc + "','" + TVS + "','" + TVSDisc + "','A')";

                                    command.CommandType = CommandType.Text;

                                    command.ExecuteNonQuery();
                                    transaction.Commit();

                                    lstTOLPack += "'" + TOLPack + "',";
                                }
                                catch (Exception ex)
                                {
                                    transaction.Rollback();
                                    //Console.WriteLine("No record was inserted into the database table.");
                                    log += "Campaign : " + campaign + ", Package: " + TOLPack + "\r\n";
                                    logSystem += ex.Message + "\r\n";
                                }
                            }

                            reader.Close();
                            reader.Dispose();
                        }
                        else
                        {
                            //update Inactive
                            OracleCommand command = ConnectionProd.CreateCommand();

                            OracleTransaction transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted);
                            command.Transaction = transaction;

                            try
                            {
                                command.CommandText = "UPDATE CAMPAIGN_MAPPING SET STATUS = 'I' WHERE CAMPAIGN_NAME = '" + campaign + "' AND TOL_PACKAGE = '" +
                                    TOLPack + "' AND TVS_PACKAGE = '" + TVS + "'";

                                command.CommandType = CommandType.Text;

                                command.ExecuteNonQuery();
                                transaction.Commit();

                                lstTOLPack += "'" + TOLPack + "',";
                            }
                            catch (Exception ex)
                            {
                                transaction.Rollback();
                                //Console.WriteLine("No record was inserted into the database table.");
                                log += "Campaign : " + campaign + ", Package: " + TOLPack + "\r\n";
                                logSystem += ex.Message + "\r\n";
                            }

                        }

                    }
                }

                if (logSystem != "" || log != "")
                {
                    MessageBox.Show("Cannot insert campaign_mapping!" + "\r\n" + log + "\r\n" + "\r\n" + "\r\n" + "System Error" + "\r\n" + logSystem);
                }
                else
                {
                    ViewResult resultForm = new ViewResult(this, null, ConnectionProd, null, txtUrNo.Text.Trim(), lstTOLPack, null, implementer, "", outputPath);
                    resultForm.ExportImp(path);
                    MessageBox.Show("Successfully");
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Found Problem" + "\r\n" + e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                xlApp.Quit();

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Cursor.Current = Cursors.Default;
            }
        }

        private void btnFolder_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                folder = folderBrowserDialog1.SelectedPath;

                txtOutput.Text = folder;
            }
            Cursor.Current = Cursors.Default;
        }
    }
}
