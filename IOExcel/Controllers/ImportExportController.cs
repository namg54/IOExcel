using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Text;

namespace IOExcel.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ImportExportController : ControllerBase
    {
        private readonly Microsoft.AspNetCore.Hosting.IHostingEnvironment _hostingEnvironment;

        public ImportExportController(Microsoft.AspNetCore.Hosting.IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }

        ///<summary>  
        /// This Method Send My Data to Demo Excel.
        /// <summary> 

        [HttpGet]
        [Route("Export")]
        public string Export()
        {
            string swebRootFolder = _hostingEnvironment.ContentRootPath;
            string sFileName = @"demo.xlsx";
            string Url = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, sFileName);
            FileInfo file = new FileInfo(Path.Combine(swebRootFolder, sFileName));
            if (file.Exists)
            {
                file.Delete();
                file=new FileInfo(Path.Combine(swebRootFolder, sFileName));
            }
            using (ExcelPackage package = new ExcelPackage(file))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Employee");
                //First add the headers
                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "Name";
                worksheet.Cells[1, 3].Value = "Gender";
                worksheet.Cells[1, 4].Value = "Salary (in $)";

                //Add values
                worksheet.Cells["A2"].Value = 1000;
                worksheet.Cells["B2"].Value = "Morteza Khorsand";
                worksheet.Cells["C2"].Value = "M";
                worksheet.Cells["D2"].Value = 10000;

                worksheet.Cells["A3"].Value = 1001;
                worksheet.Cells["B3"].Value = "Setayesh Sepehri";
                worksheet.Cells["C3"].Value = "F";
                worksheet.Cells["D3"].Value = 11000;

                worksheet.Cells["A4"].Value = 1002;
                worksheet.Cells["B4"].Value = "Nasrin Mahboob";
                worksheet.Cells["C4"].Value = "F";
                worksheet.Cells["D4"].Value = 12000;

                package.Save(); //Save the workbook.
            }
            return Url;
        }

        /// <summary> 
        /// This Method Read and Show ExcelData.
        /// <summary> 
        [HttpGet]
        [Route("Import")]
        
        public string Import()
        {
            string sWebRootFolder = _hostingEnvironment.ContentRootPath;
            string sFileName = @"demo.xlsx";
            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            try
            {
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    StringBuilder sb = new StringBuilder();
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                    int rowCount = worksheet.Dimension.Rows;
                    int ColCount = worksheet.Dimension.Columns;
                    bool bHeaderRow = true;
                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int col = 1; col <= ColCount; col++)
                        {
                            if (bHeaderRow)
                            {
                                sb.Append(worksheet.Cells[row, col].Value.ToString() + "\t");
                            }
                            else
                            {
                                sb.Append(worksheet.Cells[row, col].Value.ToString() + "\t");
                            }
                        }
                        sb.Append(Environment.NewLine);
                    }
                    return sb.ToString();
                }
            }
            catch (Exception ex)
            {
                return "Some error occured while importing." + ex.Message;
            }
        }



    }
}
