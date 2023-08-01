using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TheWitcher3Thai
{
    public class Updater
    {
        public void DownloadMod()
        {
            new WebClient().DownloadFileAsync(new Uri("https://dl.dropbox.com/s/iyn4vn4eiq4oegc/Thaimods.zip?dl=0"), string.Concat(Application.StartupPath, "\\datas\\thaimod.zip"));
        }

        public string GetLastVersion()
        {
            var request = WebRequest.Create("http://dl.dropbox.com/s/kj962og72ffu3yv/version.ini?dl=0");
            var stream = new StreamReader(request.GetResponse().GetResponseStream());
            var lastVersion = stream.ReadToEnd().ToString();

            return lastVersion;
        }

        public string GetCurrentVersion()
        {
            return File.ReadAllText(string.Concat(Application.StartupPath, "\\datas\\version.ini"));
        }

        public void CheckVersion()
        {
            //if (Operators.CompareString(this.TextBox1.Text, string.Empty, false) != 0)
            //{
            string lastVersion = GetLastVersion();
            string currentVersion = GetCurrentVersion();



            if (currentVersion == lastVersion)
            {
                //this.Label2.Text = "ภาษาไทยเป็นเวอร์ชั่นล่าสุด เข้าเกมได้ตามปกติ";
                //this.Label2.ForeColor = Color.Green;
                //this.Button7.Enabled = true;
            }
            else
            {
                //this.Label2.Text = "มีเวอร์ชั่นใหม่กว่ากรุณาอัพเดทก่อนเข้าเกม";
                //this.Label2.ForeColor = Color.Red;
                //this.Button2.Enabled = true;
            }

            //if (Directory.Exists(string.Concat(Application.StartupPath, "\\datas\\backup\\content")))
            //{
            //    this.Button6.Enabled = true;
            //}
            //else
            //{
            //    this.Button5.Enabled = true;
            //}

            //}
        }

        public void ReadVersion()
        {
            //this.Button7.Hide();
            string currentVersion = GetCurrentVersion();
            //if (Operators.CompareString(this.TextBox1.Text, string.Empty, false) == 0)
            //{
            //    this.Label2.Text = "กรุณากดค้นหาเพื่อระบุตำแหน่งที่ลงเกม";
            //    this.Button4.Enabled = false;
            //    this.Button2.Enabled = false;
            //    this.Button5.Enabled = false;
            //    this.Button6.Enabled = false;
            //    this.Button7.Enabled = false;
            if (currentVersion.Length < 8)
            {
                //this.Label3.Text = "Last update: Unknown";
                //this.Label3.ForeColor = Color.Red;
            }
            else if (currentVersion.Length > 9)
            {
                //string[] strArrays = new string[] { currentVersion.Substring(8, 2), currentVersion.Substring(5, 2), currentVersion.Substring(0, 4) };
                //string str1 = string.Join(".", strArrays);
                //this.Label3.Text = string.Concat("Last update: ", str1);
                //this.Label3.ForeColor = Color.Red;
            }
            //}
        }

        public void CopyDirectory(string SourcePath, string DestinationPath)
        {
            //Now Create all of the directories
            foreach (string dirPath in Directory.GetDirectories(SourcePath, "*", SearchOption.AllDirectories))
                Directory.CreateDirectory(dirPath.Replace(SourcePath, DestinationPath));

            //Copy all the files & Replaces any files with the same name
            foreach (string newPath in Directory.GetFiles(SourcePath, "*.*", SearchOption.AllDirectories))
                File.Copy(newPath, newPath.Replace(SourcePath, DestinationPath), true);
        }



        public void Install(string DestinationPath)
        {
            CopyDirectory(string.Concat(Application.StartupPath, "\\datas\\thaimod\\modkuntoonw3thai_mod"), string.Concat(DestinationPath, "\\mods"));
            //ProgressBar progressBar1 = this.ProgressBar1;
            //progressBar1.Value = checked(progressBar1.Value + 7);

            CopyDirectory(string.Concat(Application.StartupPath, "\\datas\\thaimod\\modkuntoonw3thai_modfile"), string.Concat(DestinationPath, "\\content"));
            //this.ProgressBar1.Value = 90;

            CopyDirectory(string.Concat(Application.StartupPath, "\\datas\\thaimod\\modKuntoonW3thai_modfileDLC"), string.Concat(DestinationPath, "\\dlc"));
            //this.ProgressBar1.Value = 98;


            MessageBox.Show("การติดตั้งภาษาไทยเรียบร้อย", "", MessageBoxButtons.OK);
            //this.ProgressBar1.Value = 0;
            //this.Button4.Enabled = false;


            string versionPath = string.Concat(Application.StartupPath, "\\datas\\version.ini");
            var version = GetLastVersion();
            File.Delete(string.Concat(Application.StartupPath, "\\datas\\version.ini"));


            StreamWriter streamWriter = new StreamWriter(versionPath, true);
            streamWriter.WriteLine(version);
            streamWriter.Close();

            string str1 = string.Concat(Application.StartupPath, "\\datas\\thaimod\\");
            (new DirectoryInfo(str1)).Attributes = FileAttributes.Normal;
            Directory.Delete(string.Concat(Application.StartupPath, "\\datas\\thaimod\\"), true);

            //this.Label2.Text = "เป็นเวอร์ชั่นล่าสุดแล้ว เข้าเกมได้ตามปกติ";
            //this.Label2.ForeColor = Color.Green;
            //this.Button7.Enabled = true;

            string str2 = GetCurrentVersion();
            if (str2.ToString().Length > 9)
            {
                //string[] strArrays = new string[] { str2.Substring(8, 2), str2.Substring(5, 2), str2.Substring(0, 4) };
                //string str3 = string.Join(".", strArrays);
                //this.Label3.Text = string.Concat("Last update: ", str3);
            }
        }

        public void UpdateProgram()
        {
            //if (this.ProgressBar1.Value == this.ProgressBar1.Maximum)
            //{
            //this.Timer1.Stop();
            //try
            //{
            ZipFile.ExtractToDirectory(string.Concat(Application.StartupPath, "\\downloads\\w3files.zip"), string.Concat(Application.StartupPath, "\\downloads\\w3files\\"));
            //}
            //catch (Exception exception)
            //{
            //    ProjectData.SetProjectError(exception);
            //    ProjectData.ClearProjectError();
            //}

            File.Delete(string.Concat(Application.StartupPath, "\\downloads\\w3files.zip"));
            File.Copy(string.Concat(Application.StartupPath, "\\downloads\\w3files\\AUpdate.exe"), string.Concat(Application.StartupPath, "\\AUpdate.exe"), true);
            //Process.Start(string.Concat(Application.StartupPath, "\\AUpdate.exe"));
            //ProjectData.EndApp();
            //}
        }


        public void Backup(string gameDirectory)
        {
            //try
            //{
            File.Copy(string.Concat(gameDirectory, "\\content\\content0\\en.w3strings"), string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content0\\en.w3strings"), true);
            //ProgressBar progressBar1 = this.ProgressBar1;
            //progressBar1.Value = checked(progressBar1.Value + 7);

            File.Copy(string.Concat(gameDirectory, "\\content\\content1\\en.w3strings"), string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content1\\en.w3strings"), true);
            //ProgressBar value = this.ProgressBar1;
            //value.Value = checked(value.Value + 7);

            File.Copy(string.Concat(gameDirectory, "\\content\\content2\\en.w3strings"), string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content2\\en.w3strings"), true);
            //ProgressBar progressBar = this.ProgressBar1;
            //progressBar.Value = checked(progressBar.Value + 7);

            File.Copy(string.Concat(gameDirectory, "\\content\\content3\\en.w3strings"), string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content3\\en.w3strings"), true);
            //ProgressBar progressBar11 = this.ProgressBar1;
            //progressBar11.Value = checked(progressBar11.Value + 7);

            File.Copy(string.Concat(gameDirectory, "\\content\\content4\\en.w3strings"), string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content4\\en.w3strings"), true);
            //ProgressBar value1 = this.ProgressBar1;
            //value1.Value = checked(value1.Value + 7);

            File.Copy(string.Concat(gameDirectory, "\\content\\content5\\en.w3strings"), string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content5\\en.w3strings"), true);
            //ProgressBar progressBar12 = this.ProgressBar1;
            //progressBar12.Value = checked(progressBar12.Value + 7);

            File.Copy(string.Concat(gameDirectory, "\\content\\content6\\en.w3strings"), string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content6\\en.w3strings"), true);
            //ProgressBar value2 = this.ProgressBar1;
            //value2.Value = checked(value2.Value + 7);

            File.Copy(string.Concat(gameDirectory, "\\content\\content7\\en.w3strings"), string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content7\\en.w3strings"), true);
            //ProgressBar progressBar2 = this.ProgressBar1;
            //progressBar2.Value = checked(progressBar2.Value + 7);

            File.Copy(string.Concat(gameDirectory, "\\content\\content8\\en.w3strings"), string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content8\\en.w3strings"), true);
            //ProgressBar progressBar13 = this.ProgressBar1;
            //progressBar13.Value = checked(progressBar13.Value + 7);

            File.Copy(string.Concat(gameDirectory, "\\content\\content9\\en.w3strings"), string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content9\\en.w3strings"), true);
            //ProgressBar value3 = this.ProgressBar1;
            //value3.Value = checked(value3.Value + 7);

            File.Copy(string.Concat(gameDirectory, "\\content\\content10\\en.w3strings"), string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content10\\en.w3strings"), true);
            //ProgressBar progressBar3 = this.ProgressBar1;
            //progressBar3.Value = checked(progressBar3.Value + 7);

            File.Copy(string.Concat(gameDirectory, "\\content\\content11\\en.w3strings"), string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content11\\en.w3strings"), true);
            //ProgressBar progressBar14 = this.ProgressBar1;
            //progressBar14.Value = checked(progressBar14.Value + 7);

            File.Copy(string.Concat(gameDirectory, "\\content\\content12\\en.w3strings"), string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content12\\en.w3strings"), true);
            //ProgressBar value4 = this.ProgressBar1;
            //value4.Value = checked(value4.Value + 2);

            File.Copy(string.Concat(gameDirectory, "\\dlc\\bob\\content\\en.w3strings"), string.Concat(Application.StartupPath, "\\datas\\backup\\dlc\\bob\\content\\en.w3strings"), true);
            //ProgressBar progressBar4 = this.ProgressBar1;
            //progressBar4.Value = checked(progressBar4.Value + 2);

            File.Copy(string.Concat(gameDirectory, "\\dlc\\EP1\\content\\en.w3strings"), string.Concat(Application.StartupPath, "\\datas\\backup\\dlc\\ep1\\content\\en.w3strings"), true);
            //ProgressBar progressBar15 = this.ProgressBar1;
            //progressBar15.Value = checked(progressBar15.Value + 3);

            //this.Button5.Enabled = false;
            //this.Button6.Enabled = true;
            //this.ProgressBar1.Value = 0;
            MessageBox.Show("สำรองไฟล์เสร็จสิ้น", "", MessageBoxButtons.OK);
            //}
            //catch (Exception exception)
            //{
            //    ProjectData.SetProjectError(exception);
            //    ProjectData.ClearProjectError();
            //}
        }

        public void Restore(string gameDirectory)
        {
            //try
            //{
            File.Delete(string.Concat(gameDirectory, "\\content\\content0\\en.w3strings"));
            //ProgressBar progressBar1 = this.ProgressBar1;
            //progressBar1.Value = checked(progressBar1.Value + 3);

            File.Delete(string.Concat(gameDirectory, "\\content\\content1\\en.w3strings"));
            //ProgressBar value = this.ProgressBar1;
            //value.Value = checked(value.Value + 3);

            File.Delete(string.Concat(gameDirectory, "\\content\\content2\\en.w3strings"));
            //ProgressBar progressBar = this.ProgressBar1;
            //progressBar.Value = checked(progressBar.Value + 3);

            File.Delete(string.Concat(gameDirectory, "\\content\\content3\\en.w3strings"));
            //ProgressBar progressBar11 = this.ProgressBar1;
            //progressBar11.Value = checked(progressBar11.Value + 3);

            File.Delete(string.Concat(gameDirectory, "\\content\\content4\\en.w3strings"));
            //ProgressBar value1 = this.ProgressBar1;
            //value1.Value = checked(value1.Value + 3);

            File.Delete(string.Concat(gameDirectory, "\\content\\content5\\en.w3strings"));
            //ProgressBar progressBar12 = this.ProgressBar1;
            //progressBar12.Value = checked(progressBar12.Value + 3);

            File.Delete(string.Concat(gameDirectory, "\\content\\content6\\en.w3strings"));
            //ProgressBar value2 = this.ProgressBar1;
            //value2.Value = checked(value2.Value + 3);

            File.Delete(string.Concat(gameDirectory, "\\content\\content7\\en.w3strings"));
            //ProgressBar progressBar2 = this.ProgressBar1;
            //progressBar2.Value = checked(progressBar2.Value + 3);

            File.Delete(string.Concat(gameDirectory, "\\content\\content8\\en.w3strings"));
            //ProgressBar progressBar13 = this.ProgressBar1;
            //progressBar13.Value = checked(progressBar13.Value + 3);

            File.Delete(string.Concat(gameDirectory, "\\content\\content9\\en.w3strings"));
            //ProgressBar value3 = this.ProgressBar1;
            //value3.Value = checked(value3.Value + 3);

            File.Delete(string.Concat(gameDirectory, "\\content\\content10\\en.w3strings"));
            //ProgressBar progressBar3 = this.ProgressBar1;
            //progressBar3.Value = checked(progressBar3.Value + 3);

            File.Delete(string.Concat(gameDirectory, "\\content\\content11\\en.w3strings"));
            //ProgressBar progressBar14 = this.ProgressBar1;
            //progressBar14.Value = checked(progressBar14.Value + 3);

            File.Delete(string.Concat(gameDirectory, "\\content\\content12\\en.w3strings"));
            //ProgressBar value4 = this.ProgressBar1;
            //value4.Value = checked(value4.Value + 3);

            File.Copy(string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content0\\en.w3strings"), string.Concat(gameDirectory, "\\content\\content0\\en.w3strings"), true);
            //ProgressBar progressBar4 = this.ProgressBar1;
            //progressBar4.Value = checked(progressBar4.Value + 3);

            File.Copy(string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content1\\en.w3strings"), string.Concat(gameDirectory, "\\content\\content1\\en.w3strings"), true);
            //ProgressBar progressBar15 = this.ProgressBar1;
            //progressBar15.Value = checked(progressBar15.Value + 3);

            File.Copy(string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content2\\en.w3strings"), string.Concat(gameDirectory, "\\content\\content2\\en.w3strings"), true);
            //ProgressBar value5 = this.ProgressBar1;
            //value5.Value = checked(value5.Value + 3);

            File.Copy(string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content3\\en.w3strings"), string.Concat(gameDirectory, "\\content\\content3\\en.w3strings"), true);
            //ProgressBar progressBar5 = this.ProgressBar1;
            //progressBar5.Value = checked(progressBar5.Value + 3);

            File.Copy(string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content4\\en.w3strings"), string.Concat(gameDirectory, "\\content\\content4\\en.w3strings"), true);
            //ProgressBar progressBar16 = this.ProgressBar1;
            //progressBar16.Value = checked(progressBar16.Value + 3);

            File.Copy(string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content5\\en.w3strings"), string.Concat(gameDirectory, "\\content\\content5\\en.w3strings"), true);
            //ProgressBar value6 = this.ProgressBar1;
            //value6.Value = checked(value6.Value + 3);

            File.Copy(string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content6\\en.w3strings"), string.Concat(gameDirectory, "\\content\\content6\\en.w3strings"), true);
            //ProgressBar progressBar6 = this.ProgressBar1;
            //progressBar6.Value = checked(progressBar6.Value + 3);

            File.Copy(string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content7\\en.w3strings"), string.Concat(gameDirectory, "\\content\\content7\\en.w3strings"), true);
            //ProgressBar progressBar17 = this.ProgressBar1;
            //progressBar17.Value = checked(progressBar17.Value + 3);

            File.Copy(string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content8\\en.w3strings"), string.Concat(gameDirectory, "\\content\\content8\\en.w3strings"), true);
            //ProgressBar value7 = this.ProgressBar1;
            //value7.Value = checked(value7.Value + 3);

            File.Copy(string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content9\\en.w3strings"), string.Concat(gameDirectory, "\\content\\content9\\en.w3strings"), true);
            //ProgressBar progressBar7 = this.ProgressBar1;
            //progressBar7.Value = checked(progressBar7.Value + 3);

            File.Copy(string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content10\\en.w3strings"), string.Concat(gameDirectory, "\\content\\content10\\en.w3strings"), true);
            //ProgressBar progressBar18 = this.ProgressBar1;
            //progressBar18.Value = checked(progressBar18.Value + 3);

            File.Copy(string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content11\\en.w3strings"), string.Concat(gameDirectory, "\\content\\content11\\en.w3strings"), true);
            //ProgressBar value8 = this.ProgressBar1;
            //value8.Value = checked(value8.Value + 3);

            File.Copy(string.Concat(Application.StartupPath, "\\datas\\backup\\content\\content12\\en.w3strings"), string.Concat(gameDirectory, "\\content\\content12\\en.w3strings"), true);
            //ProgressBar progressBar8 = this.ProgressBar1;
            //progressBar8.Value = checked(progressBar8.Value + 1);

            File.Copy(string.Concat(Application.StartupPath, "\\datas\\backup\\dlc\\bob\\content\\en.w3strings"), string.Concat(gameDirectory, "\\dlc\\bob\\content\\en.w3strings"), true);
            //ProgressBar progressBar19 = this.ProgressBar1;
            //progressBar19.Value = checked(progressBar19.Value + 1);

            File.Copy(string.Concat(Application.StartupPath, "\\datas\\backup\\dlc\\ep1\\content\\en.w3strings"), string.Concat(gameDirectory, "\\dlc\\ep1\\content\\en.w3strings"), true);
            //ProgressBar value9 = this.ProgressBar1;
            //value9.Value = checked(value9.Value + 1);

            Directory.Delete(string.Concat(Application.StartupPath, "\\datas\\backup"), true);
            //this.Button6.Enabled = false;
            //this.Button5.Enabled = false;
            File.Delete(string.Concat(Application.StartupPath, "\\datas\\version.ini"));

            //this.ProgressBar1.Value = 0;

            MessageBox.Show("คืนค่าสำรองไฟล์เสร็จสิ้น", "", MessageBoxButtons.OK);
            //}
            //catch (Exception exception)
            //{
            //    ProjectData.SetProjectError(exception);
            //    ProjectData.ClearProjectError();
            //}
        }

        public void ExtractMod()
        {
            //if (this.ProgressBar1.Value == this.ProgressBar1.Maximum)
            //{
            //    this.Timer2.Stop();
            //    try
            //    {
            ZipFile.ExtractToDirectory(string.Concat(Application.StartupPath, "\\datas\\thaimod.zip"), string.Concat(Application.StartupPath, "\\datas\\thaimod\\"));
            //}
            //catch (Exception exception)
            //{
            //    ProjectData.SetProjectError(exception);
            //    ProjectData.ClearProjectError();
            //}
            File.Delete(string.Concat(Application.StartupPath, "\\datas\\thaimod.zip"));
            string str = string.Concat(Application.StartupPath, "\\datas\\thaimod\\");
            (new DirectoryInfo(str)).Attributes = FileAttributes.Hidden;

            //this.ProgressBar1.Value = 0;
            //this.Label4.Text = "";
            //this.Button2.Enabled = false;
            //this.Label2.Text = "กดปุ่มติดตั้งภาษาไทย(หากต้องการสำรองไฟล์ให้กดสำรองไฟล์ก่อน)";
            //this.Label2.ForeColor = Color.Blue;
            //this.Button4.Enabled = true;
        }


        #region Update Program

        //public void DownloadProgram()
        //{
        //    this.Button2.Hide();
        //    this.Button1.Hide();
        //    this.Launcher = new WebClient();
        //    this.Launcher.DownloadFileAsync(new Uri("https://dl.dropbox.com/s/zwf5mi5jh5igvw7/witcher3thai.zip?dl=0"), string.Concat(Application.StartupPath, "\\downloads\\w3files.zip"));
        //    this.Timer1.Start();
        //}

        //public void UpdateProgram()
        //{
        //    if (!Directory.Exists(string.Concat(Application.StartupPath, "\\datas\\")))
        //    {
        //        Directory.CreateDirectory(string.Concat(Application.StartupPath, "\\datas\\"));
        //        string str = string.Concat(Application.StartupPath, "\\datas\\version.ini");
        //        StreamWriter streamWriter = new StreamWriter(str, true);
        //        streamWriter.WriteLine("0.0.0");
        //        streamWriter.Close();
        //    }
        //    else if (Directory.Exists(string.Concat(Application.StartupPath, "\\datas\\")) & !File.Exists(string.Concat(Application.StartupPath, "\\datas\\version.ini")))
        //    {
        //        string str1 = string.Concat(Application.StartupPath, "\\datas\\version.ini");
        //        StreamWriter streamWriter1 = new StreamWriter(str1, true);
        //        streamWriter1.WriteLine("0.0.0");
        //        streamWriter1.Close();
        //    }
        //    this.ProgressBar1.Increment(3);
        //    if (this.ProgressBar1.Value == this.ProgressBar1.Maximum)
        //    {
        //        this.Timer1.Enabled = false;
        //        if (this.Launcher.DownloadString("https://dl.dropbox.com/s/ddji9hkyb5b0war/launcher.ini?dl=0").Contains(Application.ProductVersion))
        //        {
        //            base.Hide();
        //            if (Directory.Exists(string.Concat(Application.StartupPath, "\\downloads\\")))
        //            {
        //                Directory.Delete(string.Concat(Application.StartupPath, "\\downloads\\"), DeleteDirectoryOption.DeleteAllContents);
        //            }
        //            if (File.Exists(string.Concat(Application.StartupPath, "\\AUpdate.exe")))
        //            {
        //                File.Delete(string.Concat(Application.StartupPath, "\\AUpdate.exe"));
        //            }
        //            Form item = Application.OpenForms["Form1"];
        //            if (item == null)
        //            {
        //                item = new Form1();
        //            }
        //            item.Show();
        //        }
        //        else if (Interaction.MsgBox("กรุณาอัพเดทโปรแกรม", MsgBoxStyle.YesNo, null) != MsgBoxResult.Yes)
        //        {
        //            ProjectData.EndApp();
        //        }
        //        else
        //        {
        //            base.Hide();
        //            if (!Directory.Exists(string.Concat(Application.StartupPath, "\\downloads\\")))
        //            {
        //                Directory.CreateDirectory(string.Concat(Application.StartupPath, "\\downloads\\"));
        //            }
        //            else if (Directory.Exists(string.Concat(Application.StartupPath, "\\downloads\\")))
        //            {
        //                Directory.Delete(string.Concat(Application.StartupPath, "\\downloads\\"), DeleteDirectoryOption.DeleteAllContents);
        //                Directory.CreateDirectory(string.Concat(Application.StartupPath, "\\downloads\\"));
        //            }
        //            Form form3 = Application.OpenForms["Form3"];
        //            if (form3 == null)
        //            {
        //                form3 = new Form3();
        //            }
        //            form3.Show();
        //        }
        //    }
        //}

        #endregion


    }
}
