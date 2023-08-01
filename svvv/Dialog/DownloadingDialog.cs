using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace svvv.Dialog
{
    public partial class DownloadDialog : Form
    {
        WebClient client;
        string Url, SavePath;

        public DownloadDialog(string url, string savePath)
        {
            InitializeComponent();

            this.Url = url;
            this.SavePath = savePath;
        }

        private void DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            if (e.TotalBytesToReceive == -1)
            {
                lblProgress.Text = $@"Downloaded {e.BytesReceived/1000000f:#,0.00} MB";
                progressBar1.Style = ProgressBarStyle.Marquee;
            }
            else
            {
                lblProgress.Text = $@"{e.BytesReceived/1000000f:#,0.00} MB of {e.TotalBytesToReceive/1000000f:#,0.00} MB ({e.ProgressPercentage}%)";
                progressBar1.Value = e.ProgressPercentage;
            }
        }

        private void DownloadComplete(object sender, AsyncCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                this.DialogResult = DialogResult.Cancel;
            }
            else if (e.Error != null)
            {
                this.DialogResult = DialogResult.Abort;
            }
            else
            {
                this.DialogResult = DialogResult.OK;
            }

            this.Close();
        }

        private void DownloadDialog_Load(object sender, EventArgs e)
        {
            client = new WebClient();
            client.DownloadFileCompleted += DownloadComplete;
            client.DownloadProgressChanged += DownloadProgressChanged;

            var fi = new FileInfo(SavePath);
            if (fi.Exists)
                fi.Delete();

            if (!fi.Directory.Exists)
                fi.Directory.Create();

            client.DownloadFileAsync(new Uri(Url), SavePath);

            //client.DownloadFile(new Uri(Url), SavePath);
            //this.DialogResult = DialogResult.OK;
            //this.Close();
    }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            var confirm=MessageBox.Show(this,"Stop download?", "Confirm", MessageBoxButtons.YesNo,MessageBoxIcon.Question);
            if(confirm==DialogResult.Yes)
            {
                client.CancelAsync();
            }
        }
    }
}
