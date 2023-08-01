using RG35XX_Cheat.Classes;
using svvv;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RG35XX_Cheat
{
    public partial class Form1 : Form
    {

        Cheats c = new Cheats();
        private Common cc = new Common();

        public Form1()
        {
            InitializeComponent();
            InitialScreen();
        }

        private void InitialScreen()
        {
            txtOutput.Text = c.outputPath;
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            c.outputPath = txtOutput.Text.Trim();
            if(chkClean.Checked)
            {
                cc.Processing(c.CleanOutputDirectory, false, "Cleaning output directory");
            }
            c.CopyCheats();
            cc.Processing(c.CustomMapping, false, "Processing : Custom Mapping");
            
            cc.ShowMessage("Complete");
        }

        private void btnBrows_Click(object sender, EventArgs e)
        {
            cc.SelectFolderTextBox(txtOutput, txtOutput.Text);
        }
    }
}
