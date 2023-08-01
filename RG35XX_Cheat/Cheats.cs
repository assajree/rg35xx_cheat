using OfficeOpenXml;
using RG35XX_Cheat.Classes;
using svvv;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RG35XX_Cheat
{
    public class Cheats
    {
        private Common c = new Common();
        private CoreMapping _currentMapping;
        private string retroarchCheatPath = @"D:\Game\libretro-database\cht";
        private string outputPath = Path.Combine(Application.StartupPath, "output");


        public void CopyCheats()
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

        public void ProcessCheats()
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

        public void MakeMapping()
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

        public void CustomMapping()
        {
            // read custom mapping
            List<CustomMapping> mapping = ReadCustomMapping();

            // copy cheat 
            foreach (var m in mapping)
            {
                // skip empty cheat name
                if (string.IsNullOrWhiteSpace(m.CheatName))
                    continue;

                var cheatPath = Path.Combine(outputPath, m.CoreName, m.CheatName + ".cht");
                var destPath = Path.Combine(outputPath, m.CoreName, m.RomName + ".cht");
                var cheatFile = new FileInfo(cheatPath);
                var destFile = new FileInfo(destPath);

                // check cheat exists and destination file name not exists
                if (cheatFile.Exists && !destFile.Exists)
                {
                    cheatFile.CopyTo(destFile.FullName);
                }
            }

            c.ShowMessage("Complete");
            
        }

        private List<CustomMapping> ReadCustomMapping()
        {
            int ROW_START = 2;
            int COL_ROM_NAME = 1;
            int COL_CHEAT_NAME = 2;

            List<CustomMapping> result = new List<CustomMapping>();
            var mappingPath = Path.Combine(Application.StartupPath, "custom_mapping.xlsx");
            using (ExcelPackage package = new ExcelPackage(new FileInfo(mappingPath)))
            {
                // loop all sheet
                foreach (var sht in package.Workbook.Worksheets)
                {
                    var coreName = sht.Name;

                    // loop rows
                    var currentRow = ROW_START;
                    var romName = sht.Cells[currentRow, COL_ROM_NAME].Text;
                    while (!String.IsNullOrWhiteSpace(romName))
                    {
                        var cheatName = sht.Cells[currentRow, COL_CHEAT_NAME].Text;
                        //if (!String.IsNullOrWhiteSpace(cheatName))
                        //{
                            var data = new CustomMapping()
                            {
                                CoreName = coreName,
                                RomName = romName,
                                CheatName = cheatName
                            };
                            result.Add(data);
                        //}

                        // move to next row
                        currentRow++;
                        romName = sht.Cells[currentRow, COL_ROM_NAME].Text;
                    }
                }

            }

            return result;
        }
    }
}
