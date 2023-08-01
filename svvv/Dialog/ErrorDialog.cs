using System;
using System.IO;
using System.Windows.Forms;

namespace svvv.Dialog
{
    public partial class ErrorDialog : Form
    {
        static Common c = new Common();
        public ErrorDialog()
        {
            InitializeComponent();
        }

        public ErrorDialog(Exception ex) : this()
        {
            lblMessage.Text = ex.GetBaseException().Message;
            txtDetail.Text = ex.StackTrace;
        }

        private void ToggleDetail()
        {
            if (txtDetail.Height == 0)
            {
                lblCopy.Visible = true;
                txtDetail.Height = 100;
            }
            else
            {
                lblCopy.Visible = false;
                txtDetail.Height = 0;
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ErrorDialog2_Load(object sender, EventArgs e)
        {
            ToggleDetail();
        }

        private void btnDetail_Click(object sender, EventArgs e)
        {
            ToggleDetail();
        }

        private void lblCopy_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Clipboard.SetText(txtDetail.Text);
        }

        private void lblLog_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //var logPath = Logger.GetLogPath();
            //if (File.Exists(logPath))
            //    c.Open(logPath);
            //else
            //    c.ShowMessage("ไม่พบไฟล์ Log");
        }
    }
}
