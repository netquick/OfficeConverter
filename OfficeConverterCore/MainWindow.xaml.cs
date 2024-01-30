using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeConverter
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //Definiere Background worker für die Konvertierung
        private BackgroundWorker backgroundWorker;
        private CancellationTokenSource cancellationTokenSource;

        //Definiter listen für die Anzeige im GUI
        List<string> combinedFiles = new List<string>();
        List<string> convertedFiles = new List<string>();

        private List<string> logEntries = new List<string>();
        string msgConversionInProgress = "Conversion in progress";
        string msgConversionComplete = "Conversion complete";

        //Globalvariables
        bool doSubfolders = false;
        bool doReplace = false;
        bool doWord = false;
        bool doExcel = false;
        bool doPPoint = false;
        bool doWordTmpl = false;
        bool doExcelTmpl = false; 
        string errorFolderEmpty = "";

        public MainWindow()
        {

            InitializeComponent();
            System.Diagnostics.PresentationTraceSources.SetTraceLevel(lstSourceFiles.ItemContainerGenerator, System.Diagnostics.PresentationTraceLevel.High);

            setLangEN();
            chkWord.IsChecked = true;
            chkExcel.IsChecked = true;      
            chkPowerpoint.IsChecked = true;
            cmbLang.Items.Add("EN");
            cmbLang.Items.Add("DE");
            cmbLang.Items.Add("FR");
            cmbLang.Items.Add("IT");
            cmbLang.Items.Add("BN");
            cmbLang.SelectedIndex = 0;
            lstDestFiles.ItemsSource = convertedFiles;

            backgroundWorker = new BackgroundWorker();
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.DoWork += BackgroundWorker_DoWork;
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
            lblState.Content = "Ready";
            cancellationTokenSource = new CancellationTokenSource();
        }
        private void UpdateLog(string logEntry)
        {
            // Use Dispatcher.Invoke to update UI elements from the UI thread
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                logEntries.Add(logEntry);

                // Update the ListBox with log entries
                lstLog.ItemsSource = logEntries;
                lstLog.Items.Refresh();

                lstLog.ScrollIntoView(logEntry);
                lstLog.HorizontalContentAlignment = HorizontalAlignment.Right;
            });
        }



        //Konvertier-Button
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
            doWordTmpl = (bool)chkWordTmpl.IsChecked; 
            doExcelTmpl = (bool)chkExcelTmpl.IsChecked;
            logEntries.Clear();
            lstLog.Items.Refresh();
            
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

        //Background Worker
        private async void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            // Clear the list of converted files at the beginning of each conversion
            convertedFiles.Clear();
            string folderPath = e.Argument as string;

            // Use Dispatcher.Invoke to access UI elements from the UI thread
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                // Check if txtDestFolder is empty and doReplace is false
                if (string.IsNullOrWhiteSpace(txtDestFolder.Text) && !doReplace)
                {
                    // Show a warning message in a MessageBox
                    System.Windows.MessageBox.Show(errorFolderEmpty, "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);

                    // Cancel the task
                    e.Cancel = true;
                    cancellationTokenSource.Cancel();

                    return;
                }

                // Disable buttons during conversion
                UpdateButtonStates(false);

                // Initialize the CancellationTokenSource
                cancellationTokenSource = new CancellationTokenSource();
            });

            try
            {
                // Pass the cancellation token to SearchAndConvertDocs
                await SearchAndConvertDocs(folderPath, cancellationTokenSource.Token);
            }
            catch (OperationCanceledException)
            {
                // Handle cancellation if needed
            }
            finally
            {
                // Use Dispatcher.Invoke to update UI elements from the UI thread
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    // Clear the UI-bound collection
                    combinedFiles.Clear();
                    // Add the contents of the combinedFiles list to the UI-bound collection
                    combinedFiles.ForEach(file => lstSourceFiles.Items.Add(file));
                    // Enable buttons after conversion completion
                    UpdateButtonStates(true);
                    grpSourceFiles.Header = "Queue";
                    lblState.Content = msgConversionComplete;
                });
            }
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


        // Iterate over currentFolderFiles and start the conversion asynchronously
        private async Task SearchAndConvertDocs(string folderPath, CancellationToken cancellationToken)
        {
            string[] docFiles = null;
            string[] xlsFiles = null;
            string[] pptFiles = null;
            string[] dotFiles = null;
            string[] xltFiles = null;

            int wordFilesCount = 0;
            int excelFilesCount = 0;
            int powerpointFilesCount = 0;
            int wordTemplateFilesCount = 0;
            int excelTemplateFilesCount = 0;

            if (doWord)
            {
                docFiles = Directory.GetFiles(folderPath, "*.doc");
                wordFilesCount = docFiles.Length;
                UpdateLog($"Found {wordFilesCount} Word files (*.doc) in folder {folderPath}");
            }
            if (doExcel)
            {
                xlsFiles = Directory.GetFiles(folderPath, "*.xls");
                excelFilesCount = xlsFiles.Length;
                UpdateLog($"Found {excelFilesCount} Excel files (*.xls) in folder {folderPath}");
            }
            if (doPPoint)
            {
                pptFiles = Directory.GetFiles(folderPath, "*.ppt");
                powerpointFilesCount = pptFiles.Length;
                UpdateLog($"Found {powerpointFilesCount} PowerPoint files (*.ppt) in folder {folderPath}");
            }
            if (doWordTmpl)
            {
                dotFiles = Directory.GetFiles(folderPath, "*.dot");
                wordTemplateFilesCount = dotFiles.Length;
                UpdateLog($"Found {wordTemplateFilesCount} Word template files (*.dot) in folder {folderPath}");
            }
            if (doExcelTmpl)
            {
                xltFiles = Directory.GetFiles(folderPath, "*.xlt");
                excelTemplateFilesCount = xltFiles.Length;
                UpdateLog($"Found {excelTemplateFilesCount} Excel template files (*.xlt) in folder {folderPath}");
            }

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
            if (dotFiles != null)
            {
                combinedFiles.AddRange(dotFiles);
            }
            if (xltFiles != null)
            {
                combinedFiles.AddRange(xltFiles);
            }

            Console.WriteLine($"Processing files in folder: {folderPath}");
            UpdateLog($"Processing files in folder: {folderPath}");

            /// Create a copy of the collection to avoid modification during iteration
            List<string> snapshot = new List<string>(combinedFiles);

            // Check for cancellation after creating the snapshot
            cancellationToken.ThrowIfCancellationRequested();


            try
            {
                // Iterate over currentFolderFiles and start the conversion asynchronously
                foreach (var docFile in snapshot)
                {
                    // Check for cancellation before each iteration
                    cancellationToken.ThrowIfCancellationRequested();

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

                // Check for cancellation before displaying the completion message
                cancellationToken.ThrowIfCancellationRequested();

                // Clear the list of combined files if the operation was cancelled
                if (cancellationToken.IsCancellationRequested)
                {
                    combinedFiles.Clear();
                }

            }
            catch (OperationCanceledException)
            {
                // Handle cancellation if needed
            }

            // Recursively process subfolders
            if (doSubfolders)
            {
                string[] subfolders = Directory.GetDirectories(folderPath);
                foreach (var subfolder in subfolders)
                {
                    // Pass the cancellation token to the recursive call
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        string headerName = Path.GetFileName(subfolder);
                        grpSourceFiles.Header = headerName;

                    });
                    await SearchAndConvertDocs(subfolder, cancellationToken);

                    // Check for cancellation after processing each subfolder
                    if (cancellationToken.IsCancellationRequested)
                    {
                        // Stop processing if cancellation is requested
                        break;
                    }
                }
            }
            // Use Dispatcher.Invoke to update UI elements from the UI thread
            //System.Windows.Application.Current.Dispatcher.Invoke(() =>
            //{
            //    lblState.Content = "Background work completed!";
            //});


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
                        case ".dot":
                            ConvertDocToDotx(filePath, doSubfolders, doReplace);
                            break;
                        case ".xlt":
                            ConvertXltToXltx(filePath, doSubfolders, doReplace);
                            break;


                        default:
                            // Handle other file types or show an error message
                            Console.WriteLine($"Unsupported file type: {filePath}");
                            break;

                    }

                    string logEntry = $"Converted file: {filePath}";
                    Console.WriteLine(logEntry);
                    UpdateLog(logEntry);

                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        combinedFiles.Remove(filePath);
                        if (!convertedFiles.Contains(filePath))
                        {
                            convertedFiles.Add(filePath);
                            DisplayCombinedFiles();
                        }
                        DisplayCombinedFiles();
                    });

                }
                catch (Exception ex)
                {
                    // Handle exceptions during conversion
                    string logEntry = $"Error converting {filePath}: {ex.Message}";
                    Console.WriteLine(logEntry);
                    UpdateLog(logEntry);
                }
            });

            // Update UI on the main thread
 
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
            catch (Exception ex)
            {
                // Handle or log the exception
                Console.WriteLine($"Error converting {xlsFile} to .dotx: {ex.Message}");
                string logEntry = $"Error converting {xlsFile} to .dotx: {ex.Message}";
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    logEntries.Add(logEntry);
                    // Update the ListBox with log entries
                    lstLog.ItemsSource = logEntries;
                    lstLog.Items.Refresh();
                    lstLog.ScrollIntoView(logEntry);
                });
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
        private void ConvertXltToXltx(string xltFile, bool doSubfolders, bool doReplace)
        {
            try
            {
                Excel.Application excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;

                try
                {
                    Excel.Workbook workbook = excelApp.Workbooks.Open(xltFile);

                    string targetFolderPath = "";

                    // Use Dispatcher.Invoke to execute code on the UI thread
                    Dispatcher.Invoke(() =>
                    {
                        targetFolderPath = GetTargetFolderPath(doReplace, doSubfolders, xltFile);
                    });

                    // Ensure the target folder exists
                    if (!Directory.Exists(targetFolderPath))
                    {
                        Directory.CreateDirectory(targetFolderPath);
                    }

                    // Construct the new path for the .xltx file
                    string newXltxPath = Path.Combine(targetFolderPath, Path.ChangeExtension(Path.GetFileName(xltFile), ".xltx"));
                    workbook.SaveAs(newXltxPath, Excel.XlFileFormat.xlOpenXMLTemplate);
                    workbook.Close();
                }
                catch (Exception ex)
                {
                    // Handle or log the exception
                    Console.WriteLine($"Error converting {xltFile} to .dotx: {ex.Message}");
                    string logEntry = $"Error converting {xltFile} to .dotx: {ex.Message}";
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        logEntries.Add(logEntry);
                        // Update the ListBox with log entries
                        lstLog.ItemsSource = logEntries;
                        lstLog.Items.Refresh();
                        lstLog.ScrollIntoView(logEntry);
                    });
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
            catch (Exception ex)
            {
                // Handle exceptions during conversion
                Console.WriteLine($"Error converting {xltFile}: {ex.Message}");
            }

            // Update UI on the main thread
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                combinedFiles.Remove(xltFile);
                convertedFiles.Add(xltFile);

                DisplayCombinedFiles();
            });
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
            catch (Exception ex)
            {
                // Handle or log the exception
                Console.WriteLine($"Error converting {pptFile} to .dotx: {ex.Message}");
                string logEntry = $"Error converting {pptFile} to .dotx: {ex.Message}";
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    logEntries.Add(logEntry);
                    // Update the ListBox with log entries
                    lstLog.ItemsSource = logEntries;
                    lstLog.Items.Refresh();
                    lstLog.ScrollIntoView(logEntry);
                });
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
            catch (Exception ex)
            {
                // Handle or log the exception
                Console.WriteLine($"Error converting {docFile} to .dotx: {ex.Message}");
                string logEntry = $"Error converting {docFile} to .dotx: {ex.Message}";
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    logEntries.Add(logEntry);
                    // Update the ListBox with log entries
                    lstLog.ItemsSource = logEntries;
                    lstLog.Items.Refresh();
                    lstLog.ScrollIntoView(logEntry);
                });
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
        private void ConvertDocToDotx(string dotFile, bool doSubfolders, bool doReplace)
        {
            Word.Application wordApp = new Word.Application();
            wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

            try
            {
                Word.Document doc = wordApp.Documents.Open(dotFile);

                string targetFolderPath = "";

                // Use Dispatcher.Invoke to execute code on the UI thread
                Dispatcher.Invoke(() =>
                {
                    targetFolderPath = GetTargetFolderPath(doReplace, doSubfolders, dotFile);
                });

                // Ensure the target folder exists
                if (!Directory.Exists(targetFolderPath))
                {
                    Directory.CreateDirectory(targetFolderPath);
                }

                // Construct the new path for the .dotx file
                string newDotxPath = Path.Combine(targetFolderPath, Path.ChangeExtension(Path.GetFileName(dotFile), ".dotx"));
                doc.SaveAs2(newDotxPath, Word.WdSaveFormat.wdFormatXMLTemplate);
                doc.Close();
            }
            catch (Exception ex)
            {
                // Handle or log the exception
                Console.WriteLine($"Error converting {dotFile} to .dotx: {ex.Message}");
                string logEntry = $"Error converting {dotFile} to .dotx: {ex.Message}";
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    logEntries.Add(logEntry);
                    // Update the ListBox with log entries
                    lstLog.ItemsSource = logEntries;
                    lstLog.Items.Refresh();
                    lstLog.ScrollIntoView(logEntry);
                });
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
                    targetFolder = targetFolder.TrimEnd('\\');

                    // Replace the original folder path with the destination folder path
                   //targetFolder = targetFolder.Replace(originalFolderPath, txtDestFolder.Text.TrimEnd('\\'));
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
            errorFolderEmpty = "Destination folder is required when 'Replace files' is not selected.";
            btnExportLog.Content = "Save Log";
            msgConversionInProgress = "Conversion in progress";
            msgConversionComplete = "Conversion complete";
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
            errorFolderEmpty = "Zielordner darf nicht leer sein, wenn 'Ersetze Dateien' nicht gewählt wurde.";
            btnExportLog.Content = "Log sichern";
            msgConversionInProgress = "Konvertierung läuft";
            msgConversionComplete = "Konvertierung abgeschlossen";
        }
        private void setLangFR()
        {
            // Update labels, buttons, and headers in the "Folders" section
            grpFolders.Header = "Dossiers";
            lblSouceFolder.Content = "Dossier source";
            lblDestFolder.Content = "Dossier de destination";
            btnDestFolder.Content = "Parcourir";
            btnSourceFolder.Content = "Parcourir";
            chkReplace.Content = "Remplacer les fichiers (Préserver la structure des sous-dossiers)";
            chkSubfolders.Content = "Inclure les sous-dossiers";

            // Update labels and headers in the "Files" section
            grpFiles.Header = "Fichiers";
            grpSourceFiles.Header = "File d'attente";
            grpDestFiles.Header = "Terminé";

            // Update button labels in various sections
            btnConvert.Content = "Convertir";
            btnDelete.Content = "Supprimer les fichiers";
            btnExport.Content = "Exporter la liste";
            btnExportLog.Content = "Enregistrer le journal";

            // Set error message for an empty destination folder
            errorFolderEmpty = "Le dossier de destination est requis lorsque 'Remplacer les fichiers' n'est pas sélectionné.";

            // Set messages for conversion progress and completion
            msgConversionInProgress = "Conversion en cours";
            msgConversionComplete = "Conversion terminée";
        }
        private void setLangIT()
        {
            // Update labels, buttons, and headers in the "Folders" section
            grpFolders.Header = "Cartelle";
            lblSouceFolder.Content = "Cartella di origine";
            lblDestFolder.Content = "Cartella di destinazione";
            btnDestFolder.Content = "Sfoglia";
            btnSourceFolder.Content = "Sfoglia";
            chkReplace.Content = "Sostituisci i file (Preserva la struttura delle sottocartelle)";
            chkSubfolders.Content = "Includi sottocartelle";

            // Update labels and headers in the "Files" section
            grpFiles.Header = "File";
            grpSourceFiles.Header = "Coda";
            grpDestFiles.Header = "Completato";

            // Update button labels in various sections
            btnConvert.Content = "Converti";
            btnDelete.Content = "Elimina i file";
            btnExport.Content = "Esporta lista";
            btnExportLog.Content = "Salva il registro";

            // Set error message for an empty destination folder
            errorFolderEmpty = "La cartella di destinazione è richiesta quando 'Sostituisci i file' non è selezionato.";

            // Set messages for conversion progress and completion
            msgConversionInProgress = "Conversione in corso";
            msgConversionComplete = "Conversione completata";
        }

        private void setLangBN()
        {
            // Update labels, buttons, and headers in the "Folders" section
            grpFolders.Header = "ফোল্ডার";
            lblSouceFolder.Content = "উৎস ফোল্ডার";
            lblDestFolder.Content = "গন্তব্য ফোল্ডার";
            btnDestFolder.Content = "ব্রাউজ";
            btnSourceFolder.Content = "ব্রাউজ";
            chkReplace.Content = "ফাইল প্রতিস্থাপন (সাবফোল্ডারে ফোল্ডার কাঠামো সংরক্ষণ করুন)";
            chkSubfolders.Content = "সাবফোল্ডারগুলি অন্তর্ভুক্ত করুন";

            // Update labels and headers in the "Files" section
            grpFiles.Header = "ফাইলগুলি";
            grpSourceFiles.Header = "কিউ";
            grpDestFiles.Header = "সম্পূর্ণ";

            // Update button labels in various sections
            btnConvert.Content = "কনভার্ট";
            btnDelete.Content = "ফাইলগুলি মুছুন";
            btnExport.Content = "তালিকা রপ্তানি করুন";
            btnExportLog.Content = "লগ সংরক্ষণ করুন";

            // Set error message for an empty destination folder
            errorFolderEmpty = "ফাইলগুলি নির্বাচন করার সময় গন্তব্য ফোল্ডার প্রয়োজন।";

            // Set messages for conversion progress and completion
            msgConversionInProgress = "কনভার্ট চলছে";
            msgConversionComplete = "কনভার্ট সম্পূর্ণ";
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
                case 2:
                    setLangFR();
                    break;

                case 3:
                    setLangIT();
                    break;
                case 4:
                    setLangBN();
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
        private void ExportConvertedLogFilesToFile(string filePath)
        {
            try
            {
                // Write the contents of the convertedFiles list to a text file
                File.WriteAllLines(filePath, logEntries);

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
        private void btnExportLog_Click(object sender, RoutedEventArgs e)
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
                ExportConvertedLogFilesToFile(saveFileDialog.FileName);
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
