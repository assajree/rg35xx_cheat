using Microsoft.WindowsAPICodePack.Dialogs;
using Newtonsoft.Json;
using svvv.Classes;
using svvv.Dialog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace svvv
{
    public class Common
    {

        public void OpenGoogleSheet(string fileId)
        {
            Open($@"https://docs.google.com/spreadsheets/d/{fileId}/edit");
        }


        #region Dialog

        public void SelectCsvTextBox(TextBox txt, string defaultPath = null)
        {
            var path = SelectCsv(txt.Text, defaultPath);
            if (path != null)
                txt.Text = path;
        }

        public string SelectCsv(string initialPath, string defaultPath = null)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                //InitialDirectory = @"D:\",
                Title = "Browse Csv Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "csv",
                Filter = "csv files (*.csv)|*.csv",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (String.IsNullOrWhiteSpace(initialPath))
                openFileDialog1.InitialDirectory = defaultPath ?? Application.StartupPath;
            else
                openFileDialog1.InitialDirectory = initialPath;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog1.FileName;
            }
            else
            {
                return null;
            }
        }

        public void SaveXlsxTextBox(TextBox txt, string defaultPath = null)
        {
            var path = SaveXlsx(txt.Text, defaultPath);
            if (path != null)
            {
                txt.Text = path;
            }
        }

        public string SaveXlsx(string initialPath, string defaultPath = null)
        {
            var fi = new FileInfo(initialPath);

            SaveFileDialog dlg = new SaveFileDialog()
            {
                //Title = "Save text Files",
                //CheckFileExists = false,
                //CheckPathExists = false,
                FilterIndex = 2,
                RestoreDirectory = true,
                FileName = fi.Name,

                DefaultExt = "xlsx",
                Filter = "Exccel files (*.xlsx)|*.xlsx"
            };


            if (String.IsNullOrWhiteSpace(initialPath))
                dlg.InitialDirectory = defaultPath ?? Application.StartupPath;
            else
                dlg.InitialDirectory = fi.DirectoryName;



            if (dlg.ShowDialog() == DialogResult.OK)
            {
                return dlg.FileName;
            }
            else
            {
                return null;
            }

        }

        public void SelectXlsxTextBox(TextBox txt, string defaultPath = null)
        {
            var path = SelectXlsx(txt.Text, defaultPath);
            if (path != null)
                txt.Text = path;
        }

        public void SelectJsonTextBox(TextBox txt, string defaultPath = null)
        {
            var path = SelectFile("Json", "json", txt.Text, defaultPath);
            if (path != null)
                txt.Text = path;
        }

        public string SelectXlsx(string initialPath, string defaultPath = null)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                //InitialDirectory = @"D:\",
                Title = "Browse Excel Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xlsx",
                Filter = "Exccel files (*.xlsx)|*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (String.IsNullOrWhiteSpace(initialPath))
                openFileDialog1.InitialDirectory = defaultPath ?? Application.StartupPath;
            else
                openFileDialog1.InitialDirectory = initialPath;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog1.FileName;
            }
            else
            {
                return null;
            }
        }

        public string SelectFile(string fileType, string fileExtension, string initialPath, string defaultPath = null)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                //InitialDirectory = @"D:\",
                Title = $@"Browse {fileType} Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = fileExtension,
                Filter = $@"{fileType} files (*.{fileExtension})|*.{fileExtension}",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (String.IsNullOrWhiteSpace(initialPath))
                openFileDialog1.InitialDirectory = defaultPath ?? Application.StartupPath;
            else
                openFileDialog1.InitialDirectory = initialPath;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog1.FileName;
            }
            else
            {
                return null;
            }
        }

        public void SelectFolderTextBox(TextBox txt, string defaultPath = null)
        {
            var path = SelectFolder(txt.Text, defaultPath);
            if (path != null)
                txt.Text = path;
        }

        public string SelectFolder(string initialPath, string defaultPath = null)
        {
            CommonOpenFileDialog dlg = new CommonOpenFileDialog();
            dlg.IsFolderPicker = true;

            if (String.IsNullOrWhiteSpace(initialPath))
                dlg.InitialDirectory = defaultPath ?? Application.StartupPath;
            else
                dlg.InitialDirectory = initialPath;


            if (dlg.ShowDialog() == CommonFileDialogResult.Ok)
            {
                return dlg.FileName;
            }
            else
            {
                return null;
            }
        }

        public void DeleteFile(string filePath)
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }


        public void ShowErrorMessage(string message, string caption = "Error")
        {
            MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public void ShowMessage(string message, string caption = "")
        {
            MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public bool ShowConfirm(string message, string caption = "")
        {
            var result = MessageBox.Show(message, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            return result == DialogResult.Yes;
        }


        public bool ShowConfirmWarning(string message, string caption = "")
        {
            var result = MessageBox.Show(message, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            return result == DialogResult.Yes;
        }

        #endregion

        #region Utility

        public void CallConsoleApp(string appPath, string args)
        {
            using (Process p = new Process())
            {
                p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                p.StartInfo.FileName = appPath;
                p.StartInfo.Arguments = args;

                // Redirect the output stream of the child process.
                //p.StartInfo.UseShellExecute = false;
                //p.StartInfo.RedirectStandardOutput = true;


                p.Start();

                // Do not wait for the child process to exit before
                // reading to the end of its redirected stream.
                // p.WaitForExit();
                // Read the output stream first and then wait.
                //string output = p.StandardOutput.ReadToEnd();

                p.WaitForExit();
            }
        }

        public static void DeleteDirectory(string target_dir)
        {
            if (!Directory.Exists(target_dir))
                return;

            string[] files = Directory.GetFiles(target_dir);
            string[] dirs = Directory.GetDirectories(target_dir);

            foreach (string file in files)
            {
                File.SetAttributes(file, FileAttributes.Normal);
                File.Delete(file);
            }

            foreach (string dir in dirs)
            {
                DeleteDirectory(dir);
            }

            Directory.Delete(target_dir, true);
        }


        public string ReadGoogleFileContent(string fileId)
        {
            string url = GetGoogleDownloadUrl(fileId);
            var content = ReadWebFileContent(url);
            return content;
        }

        public string ReadWebFileContent(string url)
        {
            var client = new WebClient();
            var data = client.DownloadData(url);
            var stream = new StreamReader(new MemoryStream(data));
            var content = stream.ReadToEnd().ToString();

            return content;
        }

        public string ReadLocalVersion(string modPath)
        {
            string versionPath = Path.Combine(modPath, "version.ini");

            if (!File.Exists(versionPath))
                return "N/A";

            var lines = File.ReadLines(versionPath).ToList();
            return lines.FirstOrDefault() ?? "N/A";
        }

        public string ReadVersion(string versionPath)
        {
            if (!File.Exists(versionPath))
                return "N/A";

            var lines = File.ReadLines(versionPath).ToList();
            return lines.FirstOrDefault() ?? "N/A";
        }     

        public string GetGoogleDownloadUrl(string id)
        {
            return $@"https://drive.google.com/uc?export=download&id={id}";
        }

        public string GetGoogleSheetDownloadUrl(string id)
        {
            return $@"https://docs.google.com/spreadsheets/d/{id}/export?format=xlsx";
        }

        public bool DownloadGoogleFile(string id, string saveToPath)
        {
            string url = GetGoogleDownloadUrl(id);
            using (var dlg = new DownloadDialog(url, saveToPath))
            {
                var result = dlg.ShowDialog();

                if (result == DialogResult.OK)
                    return true;
                else
                    return false;
            }
        }

      

        public DialogResult DownloadFile(string url, string saveToPath)
        {
            using (var dlg = new DownloadDialog(url, saveToPath))
            {
                var result = dlg.ShowDialog();
                return result;
            }
        }

        public void ExtractFile(string zipPath, string outputPath)
        {
            if (Directory.Exists(outputPath))
                DeleteDirectory(outputPath);
            else
                Directory.CreateDirectory(outputPath);

            ZipFile.ExtractToDirectory(zipPath, outputPath);
        }
       

        private void CreateFileDirectory(string filePath)
        {
            var fi = new FileInfo(filePath);
            if (!fi.Directory.Exists)
                fi.Directory.Create();
        }

        public void CopyDirectory(string SourcePath, string DestinationPath, List<string> skips = null)
        {
            if (skips == null)
                skips = new List<string>();

            // create all directories
            foreach (string dirPath in Directory.GetDirectories(SourcePath, "*", SearchOption.AllDirectories))
                Directory.CreateDirectory(dirPath.Replace(SourcePath, DestinationPath));

            // copy file and replace
            foreach (string path in Directory.GetFiles(SourcePath, "*.*", SearchOption.AllDirectories))
            {
                if (skips.Contains(path))
                    continue;
                var fi = new FileInfo(path.Replace(SourcePath, DestinationPath));

                if (fi.Exists)
                    fi.Delete();
                else if (!fi.Directory.Exists)
                    fi.Directory.Create();

                File.Copy(path, fi.FullName, true);
            }
        }

        #endregion

      
        public void Open(string path)
        {
            try
            {
                Process.Start(path);
            }
            catch (Exception ex)
            {
                ShowErrorMessage(ex.GetBaseException().Message);
            }
        }


        public DialogResult Processing(Action worker, string title = "Processing", string completeMessage = "Complete")
        {
            return Processing(worker, true, title, completeMessage);
        }

        public DialogResult Processing(Action worker, bool showCompleteMessage, string title = "Processing", string completeMessage = "Complete")
        {
            using (var dlg = new ProcessingDialog(worker, title))
            {
                dlg.ShowDialog();
                if (dlg.DialogResult == DialogResult.OK)
                {
                    if (showCompleteMessage)
                        ShowMessage(completeMessage);
                }
                else
                {
                    ShowErrorMessage(dlg.Error);
                }

                return dlg.DialogResult;
            }
        }

        public void ShowErrorMessage(Exception ex, string caption = "Error")
        {


            //var b = ex.GetBaseException();
            //var message = b.Message;

            //if (!(ex is KnowException))
            //    message += "\n\n" + b.StackTrace.Trim();

            //MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (ex is KnowException)
            {
                var error = ex as KnowException;
                MessageBox.Show(error.Text, error.Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                var dlg = new ErrorDialog(ex);
                dlg.ShowDialog();
            }

        }

        public void ProcessingString(Func<string> worker, string title = "Processing", string completeMessage = "Complete")
        {
            ProcessingString(worker, title, true, completeMessage);
        }

        public string ProcessingStringSilence(Func<string> worker, string title = "Processing", bool showCompleteMessage = true, string completeMessage = "Complete")
        {
            string result = null;
            using (var dlg = new ProcessingDialog(title))
            {
                dlg.ShowError = false;
                dlg.SetWorkerString(worker);
                dlg.ShowDialog();
                if (dlg.DialogResult == DialogResult.OK)
                {
                    result = dlg.Message;

                    if (showCompleteMessage)
                        ShowMessage(result, completeMessage);
                }
            }

            return result;
        }

        public string ProcessingString(Func<string> worker, string title = "Processing", bool showCompleteMessage = true, string completeMessage = "Complete")
        {
            string result = null;
            using (var dlg = new ProcessingDialog(title))
            {
                dlg.SetWorkerString(worker);
                dlg.ShowDialog();
                if (dlg.DialogResult == DialogResult.OK)
                {
                    result = dlg.Message;

                    if (showCompleteMessage)
                        ShowMessage(result, completeMessage);
                }
                else
                {
                    ShowErrorMessage(dlg.Message);
                }
            }

            return result;
        }


        public void OpenFileDirectory(string filePath)
        {
            var fi = new FileInfo(filePath);
            if (fi.Exists)
                Open(fi.Directory.FullName);

        }


        private void ReplaceAll(string path, string textToReplace, string replaceWith)
        {
            if (File.Exists(path))
            {
                string text = File.ReadAllText(path);
                text = text.Replace(textToReplace, replaceWith);
                File.WriteAllText(path, text);
            }
        }


        public void CopyFile(string sourcePath, string targetPath)
        {
            if (!File.Exists(sourcePath))
                return;

            var fi = new FileInfo(targetPath);
            if (fi.Exists)
                fi.Delete();
            else if (!fi.Directory.Exists)
                fi.Directory.Create();

            File.Copy(sourcePath, targetPath);
        }



        public void WriteJson(string json, string path)
        {
            var dir = Path.GetDirectoryName(path);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            File.WriteAllText(path, json);
        }

        public string ToJson(object obj)
        {
            var setting = new JsonSerializerSettings()
            {
                NullValueHandling = NullValueHandling.Ignore,
                MissingMemberHandling = MissingMemberHandling.Ignore
            };
            var json = JsonConvert.SerializeObject(obj, setting);
            return json;
        }

        public string ReadAllText(string path)
        {
            try
            {
                return File.ReadAllText(path);
            }
            catch (Exception)
            {
                return null;
            }
        }

        public string ReadUrl(string url, int tryCount)
        {
            try
            {
                var client = new WebClient();
                var data = client.DownloadData(url);
                var stream = new StreamReader(new MemoryStream(data));

                var result = stream.ReadToEnd().ToString();
                return result;
            }
            catch (Exception ex)
            {
                //tryCount++;
                //if (tryCount >= 2)
                //    return null;
                //else
                //    return ReadUrl(url, tryCount);

                return null;
            }
        }

    }
}
