using svvv;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace svvv.Dialog
{
    public partial class ProcessingDialog : Form
    {
        public bool ShowError=true;

        private Action Worker { get; set; }
        private Func<string> WorkerString { get; set; }
        public Exception Error { get; private set; }

        public string Message { get; private set; }

        public ProcessingDialog(string Title = "Processing...")
        {
            InitializeComponent();

            this.Text = Title;
        }

        public ProcessingDialog(Action worker, string Title = "Processing...")
        {
            InitializeComponent();

            if (worker == null)
                throw new ArgumentNullException();

            this.Text = Title;
            Worker = worker;
        }

        public void SetWorkerString(Func<string> worker)
        {
            WorkerString = worker;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            //Task.Factory.StartNew(Worker).ContinueWith(t => { this.Close(); }, TaskScheduler.FromCurrentSynchronizationContext());
            bw.RunWorkerAsync();
        }

        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            using (Process p = Process.GetCurrentProcess())
                p.PriorityClass = ProcessPriorityClass.Normal;

            this.Close();
        }

        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            using (Process p = Process.GetCurrentProcess())
                p.PriorityClass = ProcessPriorityClass.High;

            if (!bw.CancellationPending)
            {
                try
                {
                    if (Worker != null)
                        Worker.Invoke();

                    if (WorkerString != null)
                        Message = WorkerString.Invoke();

                    this.DialogResult = DialogResult.OK;
                }
                catch (Exception ex)
                {
                    if(ShowError)
                        Message = ex.GetBaseException().Message;

                    Error = ex;
                }
            }
        }

        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            var result=MessageBox.Show($@"Your current task will be stop.{Environment.NewLine}Are your sure?", "Confirm Stop", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if(result==DialogResult.Yes)
            {
                this.Text = "Stoping";
                bw.CancelAsync();
            }
        }

    }
}
