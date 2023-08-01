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

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            c.CopyCheats();

        }

        private void btnCustomMapping_Click(object sender, EventArgs e)
        {
            c.CustomMapping();
        }

    }
}
