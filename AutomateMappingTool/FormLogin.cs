using System;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.Windows.Forms;

namespace MappingDiscount
{
    public partial class FormLogin : Form
    {
        private OracleConnection ConnectionProd;
        private bool FLAG_DISCOUNT = true;
        private bool FLAG_HISPEED = false;
        private bool FLAG_VAS = false;

        public FormLogin()
        {
            InitializeComponent();
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

        #region "Event Handle"

        /// <summary>
        /// when this form loding
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormLogin_Load(object sender, EventArgs e)
        {
            FLAG_DISCOUNT = true;
            FLAG_VAS = false;
            FLAG_HISPEED = false;

            btnLogin.Enabled = true;
        }

        /// <summary>
        /// When clicked discount button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDiscount_Click(object sender, EventArgs e)
        {
            FLAG_DISCOUNT = true;
            FLAG_HISPEED = false;
            FLAG_VAS = false;
        }

        /// <summary>
        /// when clicked hispeed button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnHispeed_Click(object sender, EventArgs e)
        {
            FLAG_HISPEED = true;
            FLAG_DISCOUNT = false;
            FLAG_VAS = false;
        }

        /// <summary>
        /// when clicked vas button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnVas_Click(object sender, EventArgs e)
        {
            FLAG_VAS = true;
            FLAG_DISCOUNT = false;
            FLAG_HISPEED = false;
        }

        private void btnHispeed_MouseClick(object sender, MouseEventArgs e)
        {
            btnHispeed.FlatStyle = FlatStyle.Flat;
            btnHispeed.FlatAppearance.BorderColor = Color.White;
            btnHispeed.FlatAppearance.BorderSize = 1;
            btnHispeed.BackColor = Color.FromArgb(202, 1, 68);
            btnHispeed.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 0, 73);

            btnDiscount.BackColor = Color.FromArgb(224, 2, 107);
            btnDiscount.FlatAppearance.BorderSize = 0;

            btnVas.BackColor = Color.FromArgb(224, 2, 107);
            btnVas.FlatAppearance.BorderSize = 0;
        }

        private void btnDiscount_MouseClick(object sender, MouseEventArgs e)
        {
            btnDiscount.FlatStyle = FlatStyle.Flat;
            btnDiscount.FlatAppearance.BorderColor = Color.White;
            btnDiscount.FlatAppearance.BorderSize = 1;
            btnDiscount.BackColor = Color.FromArgb(202, 1, 68);
            btnDiscount.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 0, 73);

            btnHispeed.BackColor = Color.FromArgb(224, 2, 107);
            btnHispeed.FlatAppearance.BorderSize = 0;

            btnVas.BackColor = Color.FromArgb(224, 2, 107);
            btnVas.FlatAppearance.BorderSize = 0;
        }

        private void btnVas_MouseClick(object sender, MouseEventArgs e)
        {
            btnVas.FlatStyle = FlatStyle.Flat;
            btnVas.FlatAppearance.BorderColor = Color.White;
            btnVas.FlatAppearance.BorderSize = 1;
            btnVas.BackColor = Color.FromArgb(202, 1, 68);
            btnVas.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 0, 73);

            btnHispeed.BackColor = Color.FromArgb(224, 2, 107);
            btnHispeed.FlatAppearance.BorderSize = 0;

            btnDiscount.BackColor = Color.FromArgb(224, 2, 107);
            btnDiscount.FlatAppearance.BorderSize = 0;
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            if (txtUser.Text == "" || txtPassword.Text == "")
            {
                MessageBox.Show("Please input UserName / Password");
            }
            else
            {
                ConnectDB();
            }
        }

        private void labelClose_Click(object sender, EventArgs e)
        {
            if (ConnectionProd != null)
            {
                if (ConnectionProd.State == ConnectionState.Open)
                {
                    ConnectionProd.Close();
                    ConnectionProd.Dispose();
                }
            }

            Application.Exit();
        }

        #endregion

        private void ConnectDB()
        {
            Cursor.Current = Cursors.WaitCursor;

            string user = txtUser.Text;
            string password = txtPassword.Text;

            try
            {
                ConnectionProd = new OracleConnection();

                string connString = "Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.19.193.20)(PORT = 1560))" +
                    "(CONNECT_DATA = (SID = TEST03)));User Id=" + user + "; Password=" + password + "; Min Pool Size=10; Max Pool Size =20";

                //string connString = @"Data Source= (DESCRIPTION =(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = 150.4.2.2)(PORT = 1521)) )" +
                //   "(CONNECT_DATA =(SERVICE_NAME = TAPRD)));User ID=" + user + ";Password=" + password + ";";

                ConnectionProd.ConnectionString = connString;
                ConnectionProd.Open();

                if (ConnectionProd.State == ConnectionState.Open)
                {
                    btnLogin.Enabled = false;
                    if (FLAG_DISCOUNT == true)
                    {
                        HomeDiscount coverDiscount = new HomeDiscount(ConnectionProd, user);
                        this.Hide();

                        coverDiscount.ShowDialog();
                    }
                    else if (FLAG_HISPEED == true)
                    {
                        HomeHispeed coverHispeed = new HomeHispeed(ConnectionProd, user);
                        this.Hide();

                        coverHispeed.ShowDialog();
                    }
                    else
                    {
                        HomeVas coverVas = new HomeVas(ConnectionProd, user);

                        this.Hide();
                        coverVas.Show();
                    }
                }
                else
                {
                    btnLogin.Enabled = true;
                    DialogResult result = MessageBox.Show("Please try again!!" + "\r\n" + "Cannot connect to database.",
                   "Warning", MessageBoxButtons.OKCancel);
                    if (result == DialogResult.Cancel)
                    {
                        Application.Exit();
                    }
                }

            }
            catch (Exception ex)
            {
                DialogResult result = MessageBox.Show("Please try again!! " + "\r\n" + "Connection database failed" + "\r\n" + ex.Message,
                    "Confirmation", MessageBoxButtons.OKCancel);

                if (result == DialogResult.Cancel)
                {
                    ConnectionProd.Close();
                    ConnectionProd.Dispose();

                    Application.Exit();
                }
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

    }
}
