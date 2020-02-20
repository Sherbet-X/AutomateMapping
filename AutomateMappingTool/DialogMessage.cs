using System;
using System.Drawing;
using System.Windows.Forms;

namespace MappingDiscount
{
    public partial class DialogMessage : Form
    {
        string message;
        string fileImg;
        Color colorPanel;

        public DialogMessage(string msg, Color color, string file)
        {
            Application.UseWaitCursor = true;
            Cursor.Current = Cursors.WaitCursor;

            InitializeComponent();

            message = msg;
            colorPanel = color;
            fileImg = file;

            Application.UseWaitCursor = false;
            Cursor.Current = Cursors.Default;
        }

        private void DialogMessage_Load(object sender, EventArgs e)
        {
            Application.UseWaitCursor = true;
            Cursor.Current = Cursors.WaitCursor;

            labelMessage.Text = message;
            panelColor.BackColor = colorPanel;
            Bitmap bmp = new Bitmap(fileImg);
            pictureBox1.Image = bmp;
            btnOK.BackColor = colorPanel;
            this.Refresh();

            Application.UseWaitCursor = false;
            Cursor.Current = Cursors.Default;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
