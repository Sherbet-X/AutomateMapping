using System;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.Windows.Forms;

namespace MappingDiscount
{
    public partial class HomeDiscount : Form
    {
        MainDiscount mainDiscount;
        FormLogin formLogin;

        private OracleConnection ConnectionProd;
        private string implementer;
        private bool FLAG_NEW = true;
        private bool FLAG_EXISTING = false;
        private string fileName;
        private string folder;

        //For move form
        int mov;
        int movX;
        int movY;

        public HomeDiscount(OracleConnection con, string user)
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

        private void HomeDiscount_Load(object sender, EventArgs e)
        {
            txtImp.Text = implementer;
            FLAG_NEW = true;
            FLAG_EXISTING = false;
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (txtUrNo.Text == "")
            {
                MessageBox.Show("Please input UR No.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (txtImp.Text == "")
            {
                MessageBox.Show("Please input implementer name.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (txtInput.Text == "")
            {
                MessageBox.Show("Please select requirement file.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (txtOutput.Text == "")
            {
                MessageBox.Show("Please select output path.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                if (FLAG_NEW == true && FLAG_EXISTING == false)
                {
                    mainDiscount = new MainDiscount(this, ConnectionProd, true, false, fileName, folder,
                        txtImp.Text.Trim(), txtUrNo.Text.Trim());
                }
                else if (FLAG_NEW == false && FLAG_EXISTING == true)
                {
                    mainDiscount = new MainDiscount(this, ConnectionProd, false, true, fileName, folder,
                        txtImp.Text.Trim(), txtUrNo.Text.Trim());
                }
                else
                {
                    mainDiscount = new MainDiscount(this, ConnectionProd, true, true, fileName, folder,
                        txtImp.Text.Trim(), txtUrNo.Text.Trim());
                }

                mainDiscount.Show();
            }

            Cursor.Current = Cursors.Default;
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            fileName = "";
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                txtInput.Text = fileName;
            }

            Cursor.Current = Cursors.Default;
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            FLAG_NEW = true;
            FLAG_EXISTING = false;
        }

        private void btnExisting_Click(object sender, EventArgs e)
        {
            FLAG_NEW = false;
            FLAG_EXISTING = true;
        }

        private void btnBoth_Click(object sender, EventArgs e)
        {
            FLAG_NEW = true;
            FLAG_EXISTING = true;
        }

        private void btnNew_MouseClick(object sender, MouseEventArgs e)
        {
            btnNew.FlatStyle = FlatStyle.Flat;
            btnNew.FlatAppearance.BorderColor = Color.White;
            btnNew.FlatAppearance.BorderSize = 1;
            btnNew.BackColor = Color.FromArgb(202, 1, 68);
            btnNew.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 0, 73);

            btnExisting.BackColor = Color.FromArgb(224, 2, 107);
            btnExisting.FlatAppearance.BorderSize = 0;

            btnBoth.BackColor = Color.FromArgb(224, 2, 107);
            btnBoth.FlatAppearance.BorderSize = 0;
        }

        private void btnExisting_MouseClick(object sender, MouseEventArgs e)
        {
            btnExisting.FlatStyle = FlatStyle.Flat;
            btnExisting.FlatAppearance.BorderColor = Color.White;
            btnExisting.FlatAppearance.BorderSize = 1;
            btnExisting.BackColor = Color.FromArgb(202, 1, 68);
            btnExisting.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 0, 73);

            btnNew.BackColor = Color.FromArgb(224, 2, 107);
            btnNew.FlatAppearance.BorderSize = 0;

            btnBoth.BackColor = Color.FromArgb(224, 2, 107);
            btnBoth.FlatAppearance.BorderSize = 0;
        }

        private void btnBoth_MouseClick(object sender, MouseEventArgs e)
        {
            btnBoth.FlatStyle = FlatStyle.Flat;
            btnBoth.FlatAppearance.BorderColor = Color.White;
            btnBoth.FlatAppearance.BorderSize = 1;
            btnBoth.BackColor = Color.FromArgb(202, 1, 68);
            btnBoth.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 0, 73);

            btnNew.BackColor = Color.FromArgb(224, 2, 107);
            btnNew.FlatAppearance.BorderSize = 0;

            btnExisting.BackColor = Color.FromArgb(224, 2, 107);
            btnExisting.FlatAppearance.BorderSize = 0;
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

        private void labelClose_Click(object sender, EventArgs e)
        {
            if (ConnectionProd.State == System.Data.ConnectionState.Open)
            {
                ConnectionProd.Close();
            }
            Application.Exit();
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            this.Close();
            formLogin.Show();
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
