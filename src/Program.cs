using OfficeOpenXml;
using OfficeOpenXml.Table;
using System.Text.RegularExpressions;

internal partial class Program
{
    private static void Main(string[] args)
    {
        string dataPath = string.Empty;
        string excelFile = string.Empty;
        var regex = new Regex(@"(?:-DataPath=""*(?<DataPath>[\w \\\:\.]+)""* ){1}(?:-ExcelFile=""*(?<ExcelFile>[\w \\\:\.]+\.xlsx)""*){1}");

        var match = regex.Match(string.Join(" ", args));
        if (match.Success)
        {
            dataPath = match.Groups["DataPath"].Value;
            excelFile = match.Groups["ExcelFile"].Value;
            if (string.IsNullOrWhiteSpace(dataPath) || string.IsNullOrWhiteSpace(excelFile))
            {
                Console.WriteLine("Usage: xml_to_excel.exe DataPath=<DataPath> ExcelFile=<ExcelFile>");
                return;
            }
        }
        else
        {
            Console.WriteLine("Usage: xml_to_excel.exe DataPath=<DataPath> ExcelFile=<ExcelFile>");
            return;
        }

        Console.WriteLine($"DataPath: {dataPath}");
        Console.WriteLine($"ExcelFile: {excelFile}");

        if (!Directory.Exists(dataPath))
        {
            Console.WriteLine($"Invalid DataPath: {dataPath}");
            return;
        }

        var xmlFiles = Directory.GetFiles(dataPath, "index*.xml");
        var gimlaTypes = new SortedSet<GimlaType>();
        var docTypes = new SortedSet<DocumentType>();
        var gimla2Doc = new SortedSet<GimlaToDocument>();

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = File.Exists(excelFile) ? new ExcelPackage(new FileInfo(excelFile)) : new ExcelPackage())
        {
            var ws1 = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "GimlaTypes") ?? package.Workbook.Worksheets.Add("GimlaTypes");
            var ws2 = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "DocTypes") ?? package.Workbook.Worksheets.Add("DocTypes");
            var ws3 = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "Gimla2Doc") ?? package.Workbook.Worksheets.Add("Gimla2Doc");

            if (ws1.Dimension != null)
            {
                foreach (var row in ws1.Cells[2, 1, ws1.Dimension.End.Row, 2])
                {
                    gimlaTypes.Add(new GimlaType
                    {
                        Code = row.Offset(0, 0).GetValue<int>(),
                        Description = row.Offset(0, 1).GetValue<string>()
                    });
                }
            }

            if (ws2.Dimension != null)
            {
                foreach (var row in ws2.Cells[2, 1, ws2.Dimension.End.Row, 2])
                {
                    docTypes.Add(new DocumentType
                    {
                        Code = row.Offset(0, 0).GetValue<int>(),
                        Description = row.Offset(0, 1).GetValue<string>()
                    });
                }
            }

            if (ws3.Dimension != null)
            {
                foreach (var row in ws3.Cells[2, 1, ws3.Dimension.End.Row, 2])
                {
                    gimla2Doc.Add(new GimlaToDocument
                    {
                        GimlaCode = row.Offset(0, 0).GetValue<int>(),
                        DocType = row.Offset(0, 1).GetValue<int>()
                    });
                }
            }

            foreach (var file in xmlFiles)
            {
                XDocument doc = XDocument.Load(file);
                foreach (var ada in doc.Descendants("Ada"))
                {
                    var gimlaType = new GimlaType
                    {
                        Code = (int)ada.Element("gimla_code"),
                        Description = (string)ada.Element("gimal_desc")
                    };
                    gimlaTypes.Add(gimlaType);

                    var docType = new DocumentType
                    {
                        Code = (int)ada.Element("doc_type"),
                        Description = (string)ada.Element("doc_type_desc")
                    };
                    docTypes.Add(docType);

                    gimla2Doc.Add(new GimlaToDocument { GimlaCode = gimlaType.Code, DocType = docType.Code });
                }
            }

            ws1.Cells[1, 1].LoadFromCollection(gimlaTypes, true);
            ws2.Cells[1, 1].LoadFromCollection(docTypes, true);
            ws3.Cells[1, 1].LoadFromCollection(gimla2Doc, true);

            foreach (var ws in new[] { ws1, ws2, ws3 })
            {
                var tableRange = ws.Cells[1, 1, ws.Dimension.End.Row, ws.Dimension.End.Column];
                var table = ws.Tables.Add(tableRange, ws.Name + "Table");
                table.ShowHeader = true;
                table.TableStyle = TableStyles.Medium2;
                ws.Cells.AutoFitColumns(0, 35);
            }

            package.SaveAs(new FileInfo(excelFile));
        }

        Console.WriteLine("Excel file updated successfully.");
    }
}
