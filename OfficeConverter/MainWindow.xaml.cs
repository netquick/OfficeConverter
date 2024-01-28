using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Word = Microsoft.Office.Interop.Word;
using Forms = System.Windows.Forms;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Threading;
using System.Windows.Media.TextFormatting;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace OfficeConverter
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //Definiere Background workder für die Konvertierung
        private BackgroundWorker backgroundWorker;

        List<string> combinedFiles = new List<string>();
        List<string> convertedFiles = new List<string>();
       

        bool doSubfolders = false;
        bool doReplace = false;
        bool doWord = false;
        bool doExcel = false;
        bool doPPoint = false;


        public MainWindow()
        {
            InitializeComponent();
            setLangEN();
            chkWord.IsChecked = true;
            chkExcel.IsChecked = true;      
            chkPowerpoint.IsChecked = true;
            cmbLang.Items.Add("EN");
            cmbLang.Items.Add("DE");
            cmbLang.SelectedIndex = 0;
            lstDestFiles.ItemsSource = convertedFiles;

            backgroundWorker = new BackgroundWorker();
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.DoWork += BackgroundWorker_DoWork;
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
            lblState.Content = "Ready";

        }
        private void btnConvert_Click(object sender, RoutedEventArgs e)
        {
            string folderPath = txtSourceFolder.Text;
            if (cmbLang.SelectedIndex == 0)
            {
                lblState.Content = "Conversion in progress";
            }
            if (cmbLang.SelectedIndex == 1)
            {
                lblState.Content = "Konvertierung läuft";
            }
            
            doSubfolders = (bool)chkSubfolders.IsChecked;
            doReplace = (bool)chkReplace.IsChecked;
            doWord = (bool)chkWord.IsChecked;
            doExcel = (bool)chkExcel.IsChecked;
            doPPoint = (bool)chkPowerpoint.IsChecked;

            // Check if the background worker is not already running
            if (!backgroundWorker.IsBusy)
            {
                // Start the existing background worker
                backgroundWorker.RunWorkerAsync(folderPath);
            }
            else
            {
                // The worker is already busy, handle accordingly
                System.Windows.MessageBox.Show("Conversion is already in progress. Please wait for the current operation to finish.");
            }
        }
        private async void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {

            // Clear the list of converted files at the beginning of each conversion
            convertedFiles.Clear();
            string folderPath = e.Argument as string;
            // Disable buttons during conversion
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                UpdateButtonStates(false);
            });

            if (Directory.Exists(folderPath))
            {
                try
                {
                    await SearchAndConvertDocs(folderPath);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"Error: {ex.Message}");
                }
            }
            else
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    System.Windows.MessageBox.Show("Ungültiger Ordnerpfad.");
                });
            }
            // Enable buttons after conversion completion
            UpdateButtonStates(true);
        }


        private async Task SearchAndConvertDocs(string folderPath)
        {
            string[] docFiles = null;
            string[] xlsFiles = null;
            string[] pptFiles = null;

            if (doWord)
            {
                docFiles = Directory.GetFiles(folderPath, "*.doc");
            }
            if (doExcel)
            {
                xlsFiles = Directory.GetFiles(folderPath, "*.xls");
            }
            if (doPPoint)
            {
                pptFiles = Directory.GetFiles(folderPath, "*.ppt");
            }

            //List<string> currentFolderFiles = new List<string>();
            // Check for null before adding to combinedFiles
            if (docFiles != null)
            {
                combinedFiles.AddRange(docFiles);
            }
            if (xlsFiles != null)
            {
                combinedFiles.AddRange(xlsFiles);
            }
            if (pptFiles != null)
            {
                combinedFiles.AddRange(pptFiles);
            }

            Console.WriteLine($"Processing files in folder: {folderPath}");


            // Create a copy of the collection to avoid modification during iteration
            List<string> snapshot = new List<string>(combinedFiles);

            // Iterate over currentFolderFiles and start the conversion asynchronously
            await Task.Run(async () =>
            {
                foreach (var docFile in snapshot)
                {
                    Console.WriteLine($"Converting file: {docFile}");

                    try
                    {
                        await ConvertFileToNewFormatAsync(docFile);
                        Console.WriteLine($"DisplayCombinedFiles called");
                        DisplayCombinedFiles();
                    }
                    catch (Exception ex)
                    {
                        // Handle exception (log, display error message, etc.)
                        Console.WriteLine($"Error converting {docFile}: {ex.Message}");
                    }

                    Console.WriteLine($"Task completed for file: {docFile}");
                }
            });

            // Recursively process subfolders
            if (doSubfolders)
            {
                string[] subfolders = Directory.GetDirectories(folderPath);
                foreach (var subfolder in subfolders)
                {
                    // Await the completion of the recursive call
                    await SearchAndConvertDocs(subfolder);
                }
            }
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                lblState.Content = "Background work completed!";
            });
            DisplayCombinedFiles();

        }








        private async Task ConvertFileToNewFormatAsync(string filePath)
        {
            await Task.Run(() =>
            {
                try
                {
                    // Determine the file type based on the extension
                    string extension = System.IO.Path.GetExtension(filePath);

                    switch (extension.ToLowerInvariant())
                    {
                        case ".doc":
                            ConvertDocToDocx(filePath, doSubfolders, doReplace);
                            break;

                        case ".xls":
                            ConvertXlsToXlsx(filePath, doSubfolders, doReplace);
                            break;

                        case ".ppt":
                            ConvertPptToPptx(filePath, doSubfolders, doReplace);
                            break;

                        default:
                            // Handle other file types or show an error message
                            Console.WriteLine($"Unsupported file type: {filePath}");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    // Handle exceptions during conversion
                    Console.WriteLine($"Error converting {filePath}: {ex.Message}");
                }
            });

            // Update UI on the main thread
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                combinedFiles.Remove(filePath);
                convertedFiles.Add(filePath);

                DisplayCombinedFiles();
            });
        }
        private void ConvertXlsToXlsx(string xlsFile, bool doSubfolders, bool doReplace)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            try
            {
                Excel.Workbook workbook = excelApp.Workbooks.Open(xlsFile);

                string targetFolderPath = "";

                // Use Dispatcher.Invoke to execute code on the UI thread
                Dispatcher.Invoke(() =>
                {
                    targetFolderPath = GetTargetFolderPath(doReplace, doSubfolders, xlsFile);
                });

                // Ensure the target folder exists
                if (!Directory.Exists(targetFolderPath))
                {
                    Directory.CreateDirectory(targetFolderPath);
                }

                // Construct the new path for the .xlsx file
                string newXlsxPath = Path.Combine(targetFolderPath, Path.ChangeExtension(Path.GetFileName(xlsFile), ".xlsx"));
                workbook.SaveAs(newXlsxPath, Excel.XlFileFormat.xlOpenXMLWorkbook);
                workbook.Close();
            }
            finally
            {
                // Quit Excel and release resources
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);

                // Ensure Excel processes are terminated
                KillProcess("EXCEL");
            }
        }


        private void ConvertPptToPptx(string pptFile, bool doSubfolders, bool doReplace)
        {
            PowerPoint.Application pptApp = new PowerPoint.Application();
            pptApp.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;

            // Set PowerPoint application visibility to false
            pptApp.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

            try
            {
                PowerPoint.Presentation presentation = pptApp.Presentations.Open(pptFile);

                string targetFolderPath = "";

                // Use Dispatcher.Invoke to execute code on the UI thread
                Dispatcher.Invoke(() =>
                {
                    targetFolderPath = GetTargetFolderPath(doReplace, doSubfolders, pptFile);
                });

                // Ensure the target folder exists
                if (!Directory.Exists(targetFolderPath))
                {
                    Directory.CreateDirectory(targetFolderPath);
                }

                // Construct the new path for the .pptx file
                string newPptxPath = Path.Combine(targetFolderPath, Path.ChangeExtension(Path.GetFileName(pptFile), ".pptx"));
                presentation.SaveAs(newPptxPath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                presentation.Close();
            }
            finally
            {
                // Set PowerPoint application visibility back to true before quitting
                pptApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;

                // Quit PowerPoint
                pptApp.Quit();
            }
        }

        private void ConvertDocToDocx(string docFile, bool doSubfolders, bool doReplace)
        {
            Word.Application wordApp = new Word.Application();
            wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

            try
            {
                Word.Document doc = wordApp.Documents.Open(docFile);

                string targetFolderPath = "";

                // Use Dispatcher.Invoke to execute code on the UI thread
                Dispatcher.Invoke(() =>
                {
                    targetFolderPath = GetTargetFolderPath(doReplace, doSubfolders, docFile);
                });

                // Ensure the target folder exists
                if (!Directory.Exists(targetFolderPath))
                {
                    Directory.CreateDirectory(targetFolderPath);
                }

                // Construct the new path for the .docx file
                string newDocxPath = Path.Combine(targetFolderPath, Path.ChangeExtension(Path.GetFileName(docFile), ".docx"));
                doc.SaveAs2(newDocxPath, Word.WdSaveFormat.wdFormatXMLDocument);
                doc.Close();
            }
            finally
            {
                // Quit Word and release resources
                wordApp.Quit();
                Marshal.ReleaseComObject(wordApp);

                // Ensure Word processes are terminated
                KillProcess("WINWORD");
            }
        }

        private void KillProcess(string processName)
        {
            try
            {
                foreach (var process in Process.GetProcessesByName(processName))
                {
                    process.Kill();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error terminating {processName} processes: {ex.Message}");
            }
        }


        // Helper method to determine the target folder path
        private string GetTargetFolderPath(bool doReplace, bool doSubfolders, string docFilePath)
        {
            string targetFolder;

            if (doReplace)
            {
                // If doReplace is true, use the original folder of the document file
                targetFolder = Path.GetDirectoryName(docFilePath);
            }
            else
            {
                // If doReplace is false, use the folder defined in txtDestFolder
                targetFolder = txtDestFolder.Text.TrimEnd('\\'); // Ensure no trailing backslash

                // If doSubfolders is true, adjust the target folder based on relative path
                if (doSubfolders)
                {
                    string originalFolderPath = txtSourceFolder.Text.TrimEnd('\\'); // Ensure no trailing backslash
                    string relativePath = GetRelativePath(docFilePath, originalFolderPath);

                    // Combine the target folder with the modified relative path
                    targetFolder = Path.Combine(targetFolder, relativePath);

                    // Ensure the target folder does not include the source folder name
                    string sourceFolderName = Path.GetFileName(originalFolderPath);

                    // Remove the source folder name from the target path
                    targetFolder = targetFolder.Replace(sourceFolderName, "").TrimEnd('\\');

                    // Replace the original folder path with the destination folder path
                    targetFolder = targetFolder.Replace(originalFolderPath, txtDestFolder.Text.TrimEnd('\\'));
                }
            }

            return targetFolder;
        }










        private string GetRelativePath(string fullPath, string basePath)
        {
            Uri baseUri = new Uri(basePath + (basePath.EndsWith("\\") ? "" : "\\"));
            Uri fullUri = new Uri(fullPath);

            Uri relativeUri = baseUri.MakeRelativeUri(fullUri);
            string relativePath = Uri.UnescapeDataString(relativeUri.ToString());

            // Replace forward slashes with backslashes
            relativePath = relativePath.Replace('/', '\\');

            // Remove the filename from the relative path
            relativePath = Path.GetDirectoryName(relativePath);

            return relativePath;
        }







        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Update the UI with the progress value
            //progressBar.Value = e.ProgressPercentage;
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // Perform any additional tasks after the background work is completed
            // Enable buttons after conversion completion
 
        }


        private void setLangEN()
        {
            grpFolders.Header = "Folders";
            lblSouceFolder.Content = "Source Folder";
            lblDestFolder.Content = "Destination Folder";
            btnDestFolder.Content = "Browse";
            btnSourceFolder.Content = "Browse";
            chkReplace.Content = "Replace files (Preserve folder structure in subfolders)";
            chkSubfolders.Content = "incl. Subfolders";
            grpFiles.Header = "Files";
            grpSourceFiles.Header = "Queue";
            grpDestFiles.Header = "Completed";
            btnConvert.Content = "Convert";
            btnDelete.Content = "Delete Files";
            btnExport.Content = "Export list";

        }
        private void setLangDE()
        {
            grpFolders.Header = "Verzeichnisse";
            lblSouceFolder.Content = "Quellordner";
            lblDestFolder.Content = "Zielordner";
            btnDestFolder.Content = "Suchen";
            btnSourceFolder.Content = "Suchen";
            chkReplace.Content = "Ersetze Dateien (erhalte die Ordnerstruktur für Unterordner)";
            chkSubfolders.Content = "Unterordner mit einbeziehen";
            grpFiles.Header = "Dateien";
            grpSourceFiles.Header = "Warteschlange";
            grpDestFiles.Header = "Fertiggestellt";
            btnConvert.Content = "Konvertieren";
            btnDelete.Content = "Dateien löschen";
            btnExport.Content = "Liste exportieren";
        }

        private void txtDestFolder_TextChanged(object sender, TextChangedEventArgs e)
        {

        }



        private void btnDestFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var folderBrowserDialog = new Forms.FolderBrowserDialog())
            {
                Forms.DialogResult result = folderBrowserDialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowserDialog.SelectedPath))
                {
                    txtDestFolder.Text = folderBrowserDialog.SelectedPath;
                }
            }
        }


        private void DisplayCombinedFiles()
        {
            if (combinedFiles != null)
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    // Assuming lstSourceFiles and lstDestFiles are the names of your WPF ListBox controls
                    lstSourceFiles.ItemsSource = combinedFiles;
                    lstSourceFiles.Items.Refresh();
                    lstDestFiles.Items.Refresh();  // Refresh the ListBox to reflect changes
                });
            }
        }




        private void btnSourceFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var folderBrowserDialog = new Forms.FolderBrowserDialog())
            {
                Forms.DialogResult result = folderBrowserDialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowserDialog.SelectedPath))
                {
                    txtSourceFolder.Text = folderBrowserDialog.SelectedPath;
                    if (txtDestFolder.Text == "")
                    {
                        txtDestFolder.Text = folderBrowserDialog.SelectedPath;
                    }
                }
            }
        }

        private void cmbLang_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (cmbLang.SelectedIndex)
            {
                case 0: 
                    setLangEN();
                    break;
                case 1:
                    setLangDE();
                    break;
            }
        }
        private void chkReplace_Clicked(object sender, RoutedEventArgs e)
        {
            if (chkReplace.IsChecked == true)
            {
                lblDestFolder.IsEnabled = false;
                txtDestFolder.IsEnabled = false;
                btnDestFolder.IsEnabled = false;
                doReplace = true;

            }
            else
            {
                lblDestFolder.IsEnabled = true;
                txtDestFolder.IsEnabled = true;
                btnDestFolder.IsEnabled = true;
                doReplace= false;   
            }
        }
        private void chkSubfolders_Clicked(object sender, RoutedEventArgs e)
        {
            if (chkSubfolders.IsChecked == true)
            {
                doSubfolders = true;
            }
            else { doSubfolders = false; }

        }
        private void chkSubfolders_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void chkReplace_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void ExportConvertedFilesToFile(string filePath)
        {
            try
            {
                // Write the contents of the convertedFiles list to a text file
                File.WriteAllLines(filePath, convertedFiles);

                System.Windows.MessageBox.Show($"Export successful. File saved at: {filePath}");
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error exporting converted files: {ex.Message}");
            }
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            // Use a SaveFileDialog to let the user choose the export file location
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                DefaultExt = "txt"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                // Call the ExportConvertedFilesToFile method with the selected file path
                ExportConvertedFilesToFile(saveFileDialog.FileName);
            }
        }

        private async void DeleteConvertedFilesAsync()
        {
            // Show a confirmation dialog
            MessageBoxResult result = System.Windows.MessageBox.Show(
                "Are you sure you want to delete the converted files?",
                "Confirmation",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                await Task.Run(() =>
                {
                    try
                    {
                        foreach (var filePath in convertedFiles)
                        {
                            if (File.Exists(filePath))
                            {
                                File.Delete(filePath);
                            }
                        }

                        System.Windows.MessageBox.Show("Deletion successful.");

                        // Clear the convertedFiles list
                        convertedFiles.Clear();

                        // Refresh the ListBox to reflect changes
                        DisplayCombinedFiles();
                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show($"Error deleting converted files: {ex.Message}");
                    }
                });
            }
        }


        // Button click event
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            // Call the DeleteConvertedFilesAsync method
            DeleteConvertedFilesAsync();
        }
        private void UpdateButtonStates(bool isEnabled)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                btnConvert.IsEnabled = isEnabled;
                btnExport.IsEnabled = isEnabled;
                btnDelete.IsEnabled = isEnabled;
            });
        }




    }
}
