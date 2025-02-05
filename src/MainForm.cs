

using OfficeOpenXml;
using static System.Net.Mime.MediaTypeNames;
using System.Reflection.Emit;
using Label = System.Windows.Forms.Label;
using Application = System.Windows.Forms.Application;
using OfficeOpenXml.Table;


namespace XmlToExcel;

public class MainForm : Form
{
    private TextBox txtDataPath;
    private TextBox txtExcelFile;
    private Button btnBrowseDataPath;
    private Button btnBrowseExcelFile;
    private Button btnProcess;
    private CheckBox chkDeleteFiles;
    private ProgressBar progressBar;
    private Label lblStatus;
    private readonly string configFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
    private ConfigData ConfigData = new ConfigData();

    public MainForm()
    {
        Text = "Ada XML to Excel Converter";
        Width = 600;
        Height = 300;
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

        btnProcess = new Button { Text = "Process", Left = 250, Top = 130 };
        btnProcess.Click += (s, e) => ProcessXmlToExcel();

        progressBar = new ProgressBar { Left = 100, Top = 170, Width = 460, Visible = false };
        lblStatus = new Label { Left = 10, Top = 200, Width = 550 };

        Controls.Add(lblDataPath);
        Controls.Add(txtDataPath);
        Controls.Add(btnBrowseDataPath);
        Controls.Add(lblExcelFile);
        Controls.Add(txtExcelFile);
        Controls.Add(btnBrowseExcelFile);
        Controls.Add(chkDeleteFiles);
        Controls.Add(btnProcess);
        Controls.Add(progressBar);
        Controls.Add(lblStatus);

        LoadSettings();
    }

    private void SaveSettings(bool isFromLoad)
    {
        if(!isFromLoad)
        {
            ConfigData.DataPath = txtDataPath.Text;
            ConfigData.ExcelFile = txtExcelFile.Text;
            ConfigData.DeleteFiles = chkDeleteFiles.Checked;
        }
        File.WriteAllText(configFilePath, System.Text.Json.JsonSerializer.Serialize(ConfigData, new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
    }

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
            }
        }
        SaveSettings(true);
        txtDataPath.Text = ConfigData.DataPath;
        txtExcelFile.Text = ConfigData.ExcelFile;
        chkDeleteFiles.Checked = ConfigData.DeleteFiles;
    }

    private void BrowseFolder(TextBox textBox)
    {
        using (FolderBrowserDialog dialog = new FolderBrowserDialog())
        {
            if (dialog.ShowDialog() == DialogResult.OK)
                textBox.Text = dialog.SelectedPath;
        }
    }

    private void BrowseFile(TextBox textBox)
    {
        using (SaveFileDialog dialog = new SaveFileDialog { Filter = "Excel Files|*.xlsx" })
        {
            if (dialog.ShowDialog() == DialogResult.OK)
                textBox.Text = dialog.FileName;
        }
    }

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

            for (int row = 2; row<=(ws3.Dimension?.End.Row ?? 1); row++)
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
                XDocument doc = XDocument.Load(file);
                foreach (var ada in doc.Descendants("Ada"))
                {
                    var gimlaCode = (int)ada.Element("gimla_code");
                    var gimlaType = gimlaTypes.FirstOrDefault(g => g.Code == gimlaCode);
                    if (gimlaType == null)
                    {
                        gimlaType = new GimlaType { Code = gimlaCode, Description = (string)ada.Element("gimal_desc") };
                        gimlaTypes.Add(gimlaType);
                    }

                    var docType = new DocumentType { Code = (int)ada.Element("doc_type"), Description = (string)ada.Element("doc_type_desc") };
                    docTypes.Add(docType);


                    gimlaType.DocumentTypes.Add(docType);
                }
                if (chkDeleteFiles.Checked)
                {
                    File.Delete(file);
                }
            }

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
            package.SaveAs(new FileInfo(excelFile));
        }

        progressBar.Visible = false;
        lblStatus.Text = "Excel file created successfully!";
    }

    private void LoadToWorksheet<T>(ExcelWorksheet ws, IEnumerable<T> data, string tableName = null)
    {
        ws.Cells.Clear();
        var properties = typeof(T).GetProperties().Where(p=> p.GetCustomAttribute<BrowsableAttribute>()?.Browsable != false).ToArray();
        var headers = properties.Select(p => p.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName ?? p.Name).ToArray();

        for (int i = 0; i < headers.Length; i++)
        {
            ws.Cells[1, i + 1].Value = headers[i]; // Set custom headers names
        }

        var dataToLoad = data.Select(item=>properties.Select(p => p.GetValue(item)).ToArray()).ToArray();
        ws.Cells[2, 1].LoadFromArrays(dataToLoad); // Load data without auto headers
        var range = ws.Cells[1, 1, ws.Dimension.End.Row, ws.Dimension.End.Column];
        if(string.IsNullOrEmpty(tableName))
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

    [STAThread]
    public static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}