

using OfficeOpenXml;
using static System.Net.Mime.MediaTypeNames;
using System.Reflection.Emit;
using Label = System.Windows.Forms.Label;
using Application = System.Windows.Forms.Application;
using OfficeOpenXml.Table;


namespace XmlToExcel;

/// <summary>
/// Represents the main form for the Ada XML to Excel Converter application.
/// </summary>
public class MainForm : Form
{
    private TextBox txtDataPath;
    private TextBox txtExcelFile;
    private Button btnBrowseDataPath;
    private Button btnBrowseExcelFile;
    private Button btnProcess;
    private CheckBox chkDeleteFiles;
    private CheckBox chkRenameFiles;
    private ProgressBar progressBar;
    private Label lblStatus;
    private readonly string configFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
    private ConfigData ConfigData = new ConfigData();
    private TextBox txtLog;

    /// <summary>
    /// Initializes a new instance of the <see cref="MainForm"/> class.
    /// </summary>
    public MainForm()
    {
        Text = "Ada XML to Excel Converter";
        Width = 800;
        Height = 600;
        Icon = new Icon("YanivRabl.ico");

        Label lblDataPath = new Label { Text = "Data Path:", Left = 10, Top = 20, Width = 80 };
        txtDataPath = new TextBox { Left = 100, Top = 20, Width = 350 };
        btnBrowseDataPath = new Button { Text = "Browse", Left = 460, Top = 18 };
        btnBrowseDataPath.Click += (s, e) => BrowseFolder(txtDataPath);

        Label lblExcelFile = new Label { Text = "Excel File:", Left = 10, Top = 60, Width = 80 };
        txtExcelFile = new TextBox { Left = 100, Top = 60, Width = 350 };
        btnBrowseExcelFile = new Button { Text = "Browse", Left = 460, Top = 58 };
        btnBrowseExcelFile.Click += (s, e) => BrowseFile(txtExcelFile);

        chkDeleteFiles = new CheckBox { Text = "Delete XML files after processing", Left = 100, Top = 100, Width = 250 };
        chkRenameFiles = new CheckBox { Text = "Rename Files", Left = 100, Top = 130, Width = 250 };

        btnProcess = new Button { Text = "Process", Left = 250, Top = 160 };
        btnProcess.Click += (s, e) => ProcessXmlToExcel();

        progressBar = new ProgressBar { Left = 100, Top = 200, Width = 460, Visible = false };
        lblStatus = new Label { Left = 10, Top = 230, Width = 550 };
        txtLog = new TextBox { Left = 10, Top = 250, Width = 760, Height = 200, Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true };

        Controls.Add(lblDataPath);
        Controls.Add(txtDataPath);
        Controls.Add(btnBrowseDataPath);
        Controls.Add(lblExcelFile);
        Controls.Add(txtExcelFile);
        Controls.Add(btnBrowseExcelFile);
        Controls.Add(chkDeleteFiles);
        Controls.Add(chkRenameFiles);
        Controls.Add(btnProcess);
        Controls.Add(progressBar);
        Controls.Add(lblStatus);
        Controls.Add(txtLog);

        LoadSettings();
    }

    /// <summary>
    /// Clears the log text box.
    /// </summary>
    private void ClearLog()
    {
        if (txtLog.InvokeRequired)
        {
            txtLog.Invoke(new Action(() => txtLog.Clear()));
        }
        else
        {
            txtLog.Clear();
        }
    }

    /// <summary>
    /// Logs a message to the log text box.
    /// </summary>
    /// <param name="message">The message to log.</param>
    private void Log(string message)
    {
        if (txtLog.InvokeRequired)
        {
            txtLog.Invoke(new Action(() => txtLog.AppendText(message + Environment.NewLine)));
        }
        else
        {
            txtLog.AppendText(message + Environment.NewLine);
        }
    }

    /// <summary>
    /// Saves the current settings to the configuration file.
    /// </summary>
    /// <param name="isFromLoad">Indicates whether the settings are being saved during the load process.</param>
    private void SaveSettings(bool isFromLoad)
    {
        if (!isFromLoad)
        {
            ConfigData.DataPath = txtDataPath.Text;
            ConfigData.ExcelFile = txtExcelFile.Text;
            ConfigData.DeleteFiles = chkDeleteFiles.Checked;
            ConfigData.RenameFiles = chkRenameFiles.Checked;
        }
        File.WriteAllText(configFilePath, System.Text.Json.JsonSerializer.Serialize(ConfigData, new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
    }

    /// <summary>
    /// Loads the settings from the configuration file.
    /// </summary>
    private void LoadSettings()
    {
        if (File.Exists(configFilePath))
        {
            var json = File.ReadAllText(configFilePath);
            var settings = System.Text.Json.JsonSerializer.Deserialize<ConfigData>(json);
            if (settings != null)
            {
                ConfigData.DataPath = settings.DataPath;
                ConfigData.ExcelFile = settings.ExcelFile;
                ConfigData.DeleteFiles = settings.DeleteFiles;
                ConfigData.RenameFiles = settings.RenameFiles;
            }
        }
        SaveSettings(true);
        txtDataPath.Text = ConfigData.DataPath;
        txtExcelFile.Text = ConfigData.ExcelFile;
        chkDeleteFiles.Checked = ConfigData.DeleteFiles;
        chkRenameFiles.Checked = ConfigData.RenameFiles;
    }

    /// <summary>
    /// Opens a folder browser dialog to select a folder and sets the selected path to the specified text box.
    /// </summary>
    /// <param name="textBox">The text box to set the selected path.</param>
    private void BrowseFolder(TextBox textBox)
    {
        using (FolderBrowserDialog dialog = new FolderBrowserDialog())
        {
            if (dialog.ShowDialog() == DialogResult.OK)
                textBox.Text = dialog.SelectedPath;
        }
    }

    /// <summary>
    /// Opens a file save dialog to select a file and sets the selected path to the specified text box.
    /// </summary>
    /// <param name="textBox">The text box to set the selected path.</param>
    private void BrowseFile(TextBox textBox)
    {
        using (SaveFileDialog dialog = new SaveFileDialog { Filter = "Excel Files|*.xlsx" })
        {
            if (dialog.ShowDialog() == DialogResult.OK)
                textBox.Text = dialog.FileName;
        }
    }

    /// <summary>
    /// Processes the XML files in the specified data path and generates an Excel file.
    /// </summary>
    private void ProcessXmlToExcel()
    {
        string dataPath = txtDataPath.Text.Trim();
        string excelFile = txtExcelFile.Text.Trim();

        if (string.IsNullOrWhiteSpace(dataPath) || string.IsNullOrWhiteSpace(excelFile))
        {
            MessageBox.Show("Please select both Data Path and Excel File", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (!Directory.Exists(dataPath))
        {
            MessageBox.Show("Invalid Data Path", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        SaveSettings(false);

        progressBar.Visible = true;
        lblStatus.Text = "Processing...";
        ClearLog();
        Log("Processing started...");
        Application.DoEvents();

        var xmlFiles = Directory.GetFiles(dataPath, "index*.xml", SearchOption.AllDirectories);
        var gimlaTypes = new SortedSet<GimlaType>();
        var docTypes = new SortedSet<DocumentType>();

        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        using (var package = File.Exists(excelFile) ? new ExcelPackage(new FileInfo(excelFile)) : new ExcelPackage())
        {
            var ws1 = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "GimlaTypes") ?? package.Workbook.Worksheets.Add("GimlaTypes");
            var ws2 = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "DocTypes") ?? package.Workbook.Worksheets.Add("DocTypes");
            var ws3 = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "Gimla2Doc") ?? package.Workbook.Worksheets.Add("Gimla2Doc");

            for (int row = 2; row <= (ws1.Dimension?.End.Row ?? 1); row++)
            {
                var gimlaType = new GimlaType
                {
                    Code = ws1.Cells[row, 1].GetValue<int>(),
                    Description = ws1.Cells[row, 2].GetValue<string>()
                };
                gimlaTypes.Add(gimlaType);
            }

            for (int row = 2; row <= (ws2.Dimension?.End.Row ?? 1); row++)
            {
                var docType = new DocumentType
                {
                    Code = ws2.Cells[row, 1].GetValue<int>(),
                    Description = ws2.Cells[row, 2].GetValue<string>()
                };
                docTypes.Add(docType);
            }

            for (int row = 2; row <= (ws3.Dimension?.End.Row ?? 1); row++)
            {
                var gimlaToDocument = new GimlaToDocument
                {
                    GimlaCode = ws3.Cells[row, 1].GetValue<int>(),
                    GimlaDescription = ws3.Cells[row, 2].GetValue<string>(),
                    DocType = ws3.Cells[row, 3].GetValue<int>(),
                    DocDescription = ws3.Cells[row, 4].GetValue<string>()
                };
                var gimlaType = gimlaTypes.FirstOrDefault(x => x.Code == gimlaToDocument.GimlaCode);
                if (gimlaType != null)
                {
                    var docType = docTypes.FirstOrDefault(x => x.Code == gimlaToDocument.DocType);
                    if (docType != null)
                    {
                        gimlaType.DocumentTypes.Add(docType);
                    }
                }
            }

            foreach (var file in xmlFiles)
            {
                Log($"Processing file: {file}");
                AdaDocumentSet adaDocumentSet = new AdaDocumentSet(file);
                foreach (var ada in adaDocumentSet)
                {
                    var gimlaType = gimlaTypes.FirstOrDefault(g => g.Code == ada.GimlaCode);
                    if (gimlaType == null)
                    {
                        gimlaType = new GimlaType { Code = ada.GimlaCode, Description = ada.GimlaDescription };
                        gimlaTypes.Add(gimlaType);
                    }

                    var docType = new DocumentType { Code = ada.DocumentType, Description = ada.DocumentTypeDescription };
                    docTypes.Add(docType);

                    gimlaType.DocumentTypes.Add(docType);
                }
                if (chkRenameFiles.Checked)
                {
                    try
                    {
                        RenameDocumentFiles(adaDocumentSet);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error renaming files: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Log($"Error renaming files for {file}:{Environment.NewLine}{ex}");
                        return;
                    }
                }
                if (chkDeleteFiles.Checked)
                {
                    Log($"Deleting file: {file}");
                    File.Delete(file);
                }
            }
            Log("Finished processing files.");
            LoadToWorksheet(ws1, gimlaTypes);
            LoadToWorksheet(ws2, docTypes);
            var gimlaToDocuments = new SortedSet<GimlaToDocument>();
            foreach (var gimlaType in gimlaTypes)
            {
                foreach (var gimlaToDocument in gimlaType.GetGimlaToDocuments())
                {
                    gimlaToDocuments.Add(gimlaToDocument);
                }
            }
            LoadToWorksheet(ws3, gimlaToDocuments, "Gimla2Doc");
            while (File.Exists(excelFile) && IsFileLocked(excelFile))
            {
                var result = MessageBox.Show($"The file '{excelFile}' is currently open. Please close it and press OK to retry or Cancel to abort.",
                                             "File in Use", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (result == DialogResult.Cancel)
                {
                    Log("Operation aborted. File was in use.");
                    progressBar.Visible = false;
                    lblStatus.Text = "Operation aborted.";
                    return;
                }
            }
            while (true)
            {
                try
                {
                    package.SaveAs(new FileInfo(excelFile));
                    break;
                }
                catch (Exception ex)
                {
                    var result = MessageBox.Show($"The file '{excelFile}' is currently open. Please close it and press OK to retry or Cancel to abort.", "File in Use", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (result == DialogResult.Cancel)
                    {
                        Log("Operation aborted. File was in use.");
                        progressBar.Visible = false;
                        lblStatus.Text = "Operation aborted.";
                        return;
                    }
                }
            }
        }

        progressBar.Visible = false;
        lblStatus.Text = "Excel file created successfully!";
    }

    /// <summary>
    /// Checks if the specified file is locked.
    /// </summary>
    /// <param name="filePath">The path to the file to check.</param>
    /// <returns>true if the file is locked; otherwise, false.</returns>
    private bool IsFileLocked(string filePath)
    {
        try
        {
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                return false;
            }
        }
        catch (IOException)
        {
            return true;
        }
    }

    /// <summary>
    /// Loads data into the specified worksheet.
    /// </summary>
    /// <typeparam name="T">The type of data to load.</typeparam>
    /// <param name="ws">The worksheet to load data into.</param>
    /// <param name="data">The data to load.</param>
    /// <param name="tableName">The name of the table.</param>
    private void LoadToWorksheet<T>(ExcelWorksheet ws, IEnumerable<T> data, string tableName = null)
    {
        ws.Cells.Clear();
        var properties = typeof(T).GetProperties().Where(p => p.GetCustomAttribute<BrowsableAttribute>()?.Browsable != false).ToArray();
        var headers = properties.Select(p => p.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName ?? p.Name).ToArray();

        for (int i = 0; i < headers.Length; i++)
        {
            ws.Cells[1, i + 1].Value = headers[i]; // Set custom headers names
        }

        var dataToLoad = data.Select(item => properties.Select(p => p.GetValue(item)).ToArray()).ToArray();
        ws.Cells[2, 1].LoadFromArrays(dataToLoad); // Load data without auto headers
        var range = ws.Cells[1, 1, ws.Dimension.End.Row, ws.Dimension.End.Column];
        if (string.IsNullOrEmpty(tableName))
        {
            tableName = ws.Name;
        }

        var existingTable = ws.Tables.FirstOrDefault(t => t.Name == tableName);
        if (existingTable != null)
        {
            ws.Tables.Delete(existingTable);
        }

        var table = ws.Tables.Add(range, tableName);
        table.TableStyle = TableStyles.Medium2;
        ws.Cells.AutoFitColumns(10, 35);
    }

    /// <summary>
    /// Renames the document files based on their metadata.
    /// </summary>
    /// <param name="adaSet">The set of Ada documents containing the metadata.</param>
    private void RenameDocumentFiles(AdaDocumentSet adaSet)
    {
        int renamedFiles = 0;
        var folder = Path.GetDirectoryName(adaSet.SourceFilePath);
        Log($"Start renaming files in folder: {folder}");
        foreach (var ada in adaSet)
        {
            var originalFile = ada.GetMatchingFileInPath(folder);
            string newFilePath = ada.GetRenamedFileInPath(folder);
            if (newFilePath != null)
            {
                Log($"Found renamed file: {newFilePath}");
                renamedFiles++;
                if (originalFile != null)
                {
                    Log($"Deleting file: {originalFile}");
                    File.Delete(originalFile);
                }
                continue;
            }
            if (originalFile == null)
            {
                Log($"File not found \"{ada.DocumentAdaId}.*\" in {folder}");
                continue;
            }

            string extension = Path.GetExtension(originalFile);
            newFilePath = Path.Combine(folder, $"{ada.NewFileName}{extension}");

            if (!File.Exists(newFilePath))
            {
                Log($"Renaming file \"{Path.GetFileName(originalFile)}\" to \"{Path.GetFileName(newFilePath)}\"");
                File.Move(originalFile, newFilePath);
                renamedFiles++;
            }
        }
        Log($"Completed renaming {renamedFiles} files in folder: {folder}");
    }

    /// <summary>
    /// The main entry point for the application.
    /// </summary>
    [STAThread]
    public static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}
