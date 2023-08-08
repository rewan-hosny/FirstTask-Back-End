using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using static System.Net.Mime.MediaTypeNames;
using System.Text;
using FirstTask.models;
using Microsoft.AspNetCore.Hosting;
using OfficeOpenXml;
using System.Text.Json;
using System.Security.Claims;

namespace FirstTask.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FileController : ControllerBase
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        private readonly ExcelPackage _excelPackage;
        public FileController(IWebHostEnvironment webHostEnvironment , ExcelPackage excelPackage)
        {
            _webHostEnvironment = webHostEnvironment;
            _excelPackage = excelPackage;

        }

        [HttpPost("uploadd")]
        public async Task<IActionResult> UploadCsv([FromServices] IWebHostEnvironment webHostEnvironment)
        {
            try
            {
                if (Request.Form.Files.Count == 0)
                {
                    return BadRequest(new { error = "File not found or empty." });
                }
                IFormFile file = Request.Form.Files[0];

                if (file == null || file.Length == 0)
                    return BadRequest(new { error = "File not found or empty." });

                if (!file.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                    return BadRequest(new { error = "Invalid file format. Only .csv files are allowed." });



                // Parse CSV data
                var (headers, rows) = await ParseCsvFile(file);
                if (headers.Count == 0 || rows.Count == 0)
                    return BadRequest(new { error = "CSV data is empty." });

                // Create Excel file and return a download link
                var excelFileName = "data.xlsx";
             
                var excelFilePath = GetExcelFilePath(webHostEnvironment,excelFileName);


                // Generate Excel file
                GenerateExcelFile(headers, rows, excelFilePath);

                // Construct the download link
                var downloadLink = Url.Content("~/uploads/" + excelFileName);

                var dataFile = new DataCsv
                {
                    Headers = headers,
                    Rows = rows
                };
              
                return Ok(new ResponseModel { DataFile = dataFile, ExcelDownloadLink = downloadLink });

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());

                return StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }
    
        private async Task<(List<string> Headers, List<List<string>> Rows)> ParseCsvFile(IFormFile file)
        {
            List<List<string>> rows = new List<List<string>>();
            List<string> headers = new List<string>();

            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                stream.Position = 0;

                using (var reader = new StreamReader(stream, Encoding.UTF8))
                {
                    bool isFirstLine = true;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(',');

                        if (isFirstLine)
                        {
                            headers.AddRange(values);
                            isFirstLine = false;
                        }
                        else
                        {
                            rows.Add(new List<string>(values));
                        }
                    }
                }
            }

            return (headers, rows);
        }
        private string GetExcelFilePath(IWebHostEnvironment webHostEnvironment, string excelFileName)
        {
            var excelDirectory = Path.Combine(webHostEnvironment.WebRootPath, "uploads");
            Directory.CreateDirectory(excelDirectory);
            return Path.Combine(excelDirectory, excelFileName);
        }


        private void GenerateExcelFile(List<string> headers, List<List<string>> rows, string excelFilePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Data");

                // Set headers
                for (int col = 1; col <= headers.Count; col++)
                {
                    worksheet.Cells[1, col].Value = headers[col - 1];
                }

                // Set rows
                for (int row = 2; row <= rows.Count + 1; row++)
                {
                    for (int col = 1; col <= rows[row - 2].Count; col++)
                    {
                        worksheet.Cells[row, col].Value = rows[row - 2][col - 1];
                    }
                }

                // Save Excel 
                package.SaveAs(new FileInfo(excelFilePath));
            }
        }

    }

}
