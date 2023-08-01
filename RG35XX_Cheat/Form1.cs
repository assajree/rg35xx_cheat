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
        private Common c = new Common();
        private CoreMapping _currentMapping;
        private string retroarchCheatPath = @"D:\Game\libretro-database\cht";
        private string outputPath = Path.Combine(Application.StartupPath, "output");

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CopyCheats();

        }

        private void CopyCheats()
        {            
            // read mapping from json
            var mappingPath = Path.Combine(Application.StartupPath, "mapping.json");
            var mapping = c.FromJsonFile<List<CoreMapping>>(mappingPath);

            // loop cht file in directory
            foreach (var m in mapping)
            {
                _currentMapping = m;

                var path = Path.Combine(retroarchCheatPath, m.DirectoryName);
                var d = new DirectoryInfo(path);
                c.Processing(ProcessCheats, false, $@"Processing : {d.Name}");
            }

            c.ShowMessage("Complete");
        }

        private void ProcessCheats()
        {
            var m = _currentMapping;

            // create core directory
            var corePath = Path.Combine(outputPath, m.CoreName);
            if (!Directory.Exists(corePath))
            {
                Directory.CreateDirectory(corePath);
            }

            // read all .cht file in source directory
            var path = Path.Combine(retroarchCheatPath, m.DirectoryName);
            var cheats = Directory.GetFiles(path, "*.cht", SearchOption.TopDirectoryOnly);
            foreach (var c in cheats)
            {
                var cheatFile = new FileInfo(c);
                var destFile = new FileInfo(Path.Combine(corePath, cheatFile.Name));
                if (!destFile.Exists)
                {
                    cheatFile.CopyTo(destFile.FullName);
                }

            }
        }

        private void MakeMapping()
        {
            // read all subfolder
            DirectoryInfo directory = new DirectoryInfo(@"D:\Game\libretro-database\cht");
            DirectoryInfo[] directories = directory.GetDirectories();

            // build object
            var result = new List<CoreMapping>();
            foreach (DirectoryInfo folder in directories)
            {
                var mapping = new CoreMapping()
                {
                    DirectoryName = folder.Name,
                    CoreName = folder.Name
                };
                result.Add(mapping);
            }

            // write to json
            var jsonPath = Path.Combine(@"D:\Desktop", "mappimg.json");
            c.WriteJson(result, jsonPath);
        }
    }
}
