using CsvHelper;
using iText.Kernel.Pdf;
using loadfile.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.Diagnostics;
using System.Formats.Asn1;
using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;
using System.Text;

using iText.Layout;
using iText.Layout.Element;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Canvas.Parser;
using System.Globalization;

namespace loadfile.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly TestDbContext _context = new TestDbContext();
        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;

        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult GetFiles()
        {
            return View();

        }

        [HttpPost]
        public IActionResult GetFiles(List<IFormFile> files)
        {
            foreach (var file in files)
            {
                if (file.Length > 0)
                {
                    var filePath = Path.GetTempFileName();

                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        file.CopyTo(stream);
                    }

                    ProcessUploadedFile(filePath, file.FileName);
                    
                }
            }

            return View();
        }

        private void ProcessUploadedFile(string filePath, string fileName)
        {
            if (fileName.EndsWith(".xlsx"))
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets.First();
                    var tableName = Path.GetFileNameWithoutExtension(fileName);
                    CreateTable(tableName, worksheet);

                    InsertData(tableName, worksheet);
                }
            }
            else if (fileName.EndsWith(".csv"))
            {
                ProcessCsvFile(filePath, fileName);
            }
            else if (fileName.EndsWith(".pdf"))
            {

                ProcessPdfFile(filePath, fileName);


            }
            else
            {
                ViewBag.error = "This type of file not allowed .Only(csv,pdf,xlsx)";
            }
        }
        private void ProcessCsvFile(string filePath, string fileName)
        {
            using (var reader = new StreamReader(filePath))
            {
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var records = csv.GetRecords<dynamic>().ToList(); // Adjust the dynamic type based on your CSV structure

                    if (records.Any())
                    {
                        var columns = ((IDictionary<string, object>)records[0]).Keys.ToList();
                        var csvData = new List<List<string>> { columns };

                        foreach (var record in records)
                        {
                            var row = new List<string>();
                            foreach (var column in columns)
                            {
                                var value = ((IDictionary<string, object>)record)[column]?.ToString();
                                row.Add(value);
                            }
                            csvData.Add(row);
                        }

                        using (var package = new ExcelPackage())
                        {
                            var excelWorksheet = CreateExcelWorksheetFromCsv(package, csvData);

                            var tableName = Path.GetFileNameWithoutExtension(fileName);
                            CreateTable(tableName, excelWorksheet);
                            InsertData(tableName, excelWorksheet);
                        }
                    }
                }
            }
        }

        private ExcelWorksheet CreateExcelWorksheetFromCsv(ExcelPackage package, List<List<string>> csvData)
        {
            if (csvData.Count > 0)
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Populate the Excel worksheet with the CSV data
                for (var row = 0; row < csvData.Count; row++)
                {
                    for (var col = 0; col < csvData[row].Count; col++)
                    {
                        worksheet.Cells[row + 1, col + 1].Value = csvData[row][col];
                    }
                }

                return worksheet;
            }

            return null;
        }

        private string DeterminePrimaryKeyColumn(ExcelWorksheet worksheet, string[] columnNames)
        {

            return columnNames.FirstOrDefault();
        }

        private void CreateTable(string tableName, ExcelWorksheet worksheet)
        {
            // Assuming the first row contains column names
            var columnNames = worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns].Select(cell => cell.Text).ToArray();

            // Create a dictionary to map column names to indices
            var columnIndexMap = columnNames.Select((name, index) => (name, index)).ToDictionary(pair => pair.name, pair => pair.index);

            var primaryKeyColumn = DeterminePrimaryKeyColumn(worksheet, columnNames);
                
            // Determine data types for each column
            var columnDataTypes = columnNames.Select(name => DetermineDataType(worksheet.Cells[2, columnIndexMap[name] + 1].Text));

            // Escape identifiers with square brackets
            tableName = $"[{tableName}]";

            primaryKeyColumn = $"[{primaryKeyColumn}]";
            var escapedColumnNames = columnNames.Select(name => $"[{name}]");

            var columnsWithTypes = escapedColumnNames.Zip(columnDataTypes, (name, type) => $"{name} {type}");
            var createTableSql = $"CREATE TABLE {tableName} ({string.Join(", ", columnsWithTypes)}, PRIMARY KEY ({primaryKeyColumn}));";

            _context.Database.ExecuteSqlRaw(createTableSql);
        }


        private void InsertData(string tableName, ExcelWorksheet worksheet)
        {
            // Assuming the first row contains column names
            var columnNames = worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns].Select(cell => cell.Text).ToArray();

            for (int row = 2; row <= worksheet.Dimension.Rows; row++)
            {
                var rowData = new Dictionary<string, object>();

                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                {
                    var columnName = columnNames[col - 1];
                    var cellValue = worksheet.Cells[row, col].Text;
                    rowData.Add(columnName, cellValue);
                }

                try
                {
                    // Escape the table name with square brackets
                    var escapedTableName = $"[{tableName}]";

                    // Insert the data into the dynamically created table
                    var columns = string.Join(", ", rowData.Keys.Select(name => $"[{name}]"));
                    var values = string.Join(", ", rowData.Values.Select(value => $"'{value}'"));
                    _context.Database.ExecuteSqlRaw($"INSERT INTO {escapedTableName} ({columns}) VALUES ({values});");
                }
                catch (Exception ex)
                {
                    ViewBag.Message = "An error occurred while inserting data.";
                }
            }
        }




        private List<string> DetermineDataTypes(ExcelWorksheet worksheet, string[] columnNames)
        {
            var dataTypes = new List<string>();

            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                var columnName = columnNames[col - 1];

                // Assume that the first non-empty cell in the column determines the data type
                var firstDataCell = worksheet.Cells.FirstOrDefault(c => c.Start.Column == col && c.Start.Row > 1 && !string.IsNullOrEmpty(c.Text));

                // Use a method to dynamically determine the data type
                var dataType = DetermineDataType(firstDataCell?.Text);
                dataTypes.Add(dataType);
            }

            return dataTypes;
        }

        private string DetermineDataType(string cellValue)
        {
            if (int.TryParse(cellValue, out _))
            {
                return "INT";
            }
            else if (decimal.TryParse(cellValue, out _))
            {
                return "DECIMAL(18,2)";
            }
            else if (DateTime.TryParse(cellValue, out _))
            {
                return "DATETIME";
            }
            else
            {
                return "NVARCHAR(MAX)";
            }
        }
        private void ProcessPdfFile(string filePath, string fileName)
        {
            using (var package = new ExcelPackage())
            {
                var pdfText = ExtractTextFromPdf(filePath);
                var excelData = ConvertPdfTextToExcelData(pdfText);
                var excelWorksheet = CreateExcelWorksheet(package, excelData);

                if (excelWorksheet != null)
                {
                    var tableName = Path.GetFileNameWithoutExtension(fileName);
                    CreateTable(tableName, excelWorksheet);
                    InsertData(tableName, excelWorksheet);
                }
                else
                {
                    ViewBag.Message = "An error occurred while processing the PDF file.";
                }
            }
        }
        private List<List<string>> ConvertPdfTextToExcelData(string pdfText)
        {
            var rows = pdfText.Split('\n')
                            .Select(row => row.Split((char[])null, StringSplitOptions.RemoveEmptyEntries)
                                               .Select(cell => cell.Trim())
                                               .ToList())
                            .ToList();

            // Ensure that each row has the same number of columns
            var maxColumns = rows.Max(row => row.Count);
            foreach (var row in rows)
            {
                while (row.Count < maxColumns)
                {
                    // If a row has fewer columns, add empty strings to make them equal
                    row.Add(string.Empty);
                }
            }

            return rows;
        }



        private string ExtractTextFromPdf(string filePath)
        {
            var text = new StringBuilder();

            using (var pdfReader = new PdfReader(filePath))
            {
                using (var pdfDocument = new PdfDocument(pdfReader))
                {
                    for (var page = 1; page <= pdfDocument.GetNumberOfPages(); page++)
                    {
                        var currentPage = pdfDocument.GetPage(page);
                        var strategy = new SimpleTextExtractionStrategy();
                        text.Append(PdfTextExtractor.GetTextFromPage(currentPage, strategy));
                    }
                }
            }

            return text.ToString();
        }


        private ExcelWorksheet CreateExcelWorksheet(ExcelPackage package, List<List<string>> excelData)
        {
            if (excelData.Count > 0)
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Populate the Excel worksheet with the data
                for (var row = 0; row < excelData.Count; row++)
                {
                    for (var col = 0; col < excelData[row].Count; col++)
                    {
                        worksheet.Cells[row + 1, col + 1].Value = excelData[row][col];
                    }
                }

                return worksheet;
            }

            return null;
        }

    }
}
