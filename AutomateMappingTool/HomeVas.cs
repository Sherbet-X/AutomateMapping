using System;
using Oracle.ManagedDataAccess.Client;
using System.Windows.Forms;

namespace MappingDiscount
{
    public partial class HomeVas : Form
    {
        private OracleConnection ConnectionProd;
        private string implementer;
        private string filename;
        private string folder;
        FormLogin formLogin;

        //For move form
        int mov;
        int movX;
        int movY;

        public HomeVas(OracleConnection con, string user)
        {
            InitializeComponent();

            ConnectionProd = con;
            implementer = user;
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

        private void CoverVas_Load(object sender, EventArgs e)
        {
            txtImp.Text = implementer;
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            if (txtUrNo.Text == "")
            {
                MessageBox.Show("Please input UR_NO#");
            }
            else if (txtImp.Text == "")
            {
                MessageBox.Show("Please input implementer");
            }
            else if (txtInput.Text == "")
            {
                MessageBox.Show("Please input requirement file");
            }
            else if (txtOutput.Text == "")
            {
                MessageBox.Show("Please select output path.");
            }
            else
            {
                MainVas mainVas = new MainVas(this, ConnectionProd, filename, folder,
                         implementer, txtUrNo.Text.Trim());

                mainVas.Show();
            }
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            filename = "";
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filename = openFileDialog1.FileName;
                txtInput.Text = filename;
            }

            Cursor.Current = Cursors.Default;
        }

        private void labelClose_Click(object sender, EventArgs e)
        {
            ConnectionProd.Close();
            Application.Exit();
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

        private void btnBack_Click(object sender, EventArgs e)
        {
            this.Close();
            formLogin.Show();
        }

        private void btnOut_Click(object sender, EventArgs e)
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
