using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Syncfusion.Licensing;
using Syncfusion.Lic;
using Syncfusion.XlsIO;
using Microsoft.AspNetCore.Hosting;
using OfficeOpenXml;
using NPOI.XSSF.UserModel;
using GemBox.Spreadsheet;

namespace WebProject.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class DataReadingController : ControllerBase
    {

        private readonly IWebHostEnvironment _env;
        public DataReadingController(IWebHostEnvironment env)
        {
            _env = env;
        }
        [HttpPost]
        [Route("exceldatareader")]
        public IActionResult GetReadingFromExceldatareader()
        {
            try
            {
                IExcelDataReader reader = null;
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                string FilePath = _env.ContentRootPath + "\\" + "Bulk Records.xlsx";
                int content = 0;
                using var filestream = new FileStream("Bulk Records.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                if (Path.GetExtension(FilePath).Equals(".xls"))
                    reader = ExcelReaderFactory.CreateBinaryReader(filestream);
                else if (Path.GetExtension(FilePath).Equals(".xlsx"))
                    reader = ExcelReaderFactory.CreateOpenXmlReader(filestream);
                if (reader != null)
                {
                    content = reader.FieldCount;
                    while (reader.Read())
                    {
                        object i = reader.GetValue(0);
                    }
                    //var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    //{
                    //    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    //    {
                    //        UseHeaderRow = true
                    //    }
                    //});
                    //https://discoverdot.net/projects/excel-data-reader
                    //DataTable s = result.Tables[0];
                }
                return Ok(content);
            }
            catch (Exception ex)
            {
                return Ok(ex.Message + " " + ex.StackTrace.ToString());
            }

        }
        //epplus and npo
        [HttpPost]
        [Route("syncfusion")]
        public IActionResult GetReadingFromSyncfusion()
        {
            try
            {
                ExcelEngine excelEngine = new ExcelEngine();
                string FilePath = _env.ContentRootPath + "\\" + "Bulk Records.xlsx";
                using (var stream = new FileStream("Bulk Records.xlsx", FileMode.Open, FileAccess.Read))
                {
                    var workbook = excelEngine.Excel.Workbooks.Open(stream);
                    return Ok(workbook.MaxColumnCount);
                }
            }
            catch (Exception ex)
            {

                return Ok(ex.Message + " " + ex.StackTrace.ToString());
            }

        }
        [HttpPost]
        [Route("epplus")]
        public IActionResult GetReadingFromEPPlus()
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                using (var pck = new OfficeOpenXml.ExcelPackage())
                {
                    using (var stream = new FileStream("Bulk Records.xlsx", FileMode.Open, FileAccess.Read))
                    {

                        pck.Load(stream);
                    }
                    var ws = pck.Workbook.Worksheets.First();
                    return Ok(ws.Workbook.Worksheets.Count);
                }
            }
            catch (Exception ex)
            {

                return Ok(ex.Message + " " + ex.StackTrace.ToString());
            }

        }
        [HttpPost]
        [Route("npoi")]
        public IActionResult GetReadingFromnpoi()
        {
            try
            {

                using (var stream = new FileStream("Bulk Records.xlsx", FileMode.Open, FileAccess.Read))
                {
                    XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
                    return Ok(xssWorkbook.Count);
                }
            }
            catch (Exception ex)
            {
                return Ok(ex.Message + " " + ex.StackTrace.ToString());
            }
        }
        [HttpPost]
        [Route("gemboxspreadsheet")]
        public IActionResult GetReadingFromGemBoxSpreadsheet()
        {
            try
            {

                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                // Load Excel workbook from file's path.
                ExcelFile workbook = ExcelFile.Load("Bulk Records.xlsx");
                return Ok();
            }
            catch (Exception ex)
            {
                return Ok(ex.Message + " " + ex.StackTrace.ToString());
            }
        }
        [HttpPost]
        [Route("oledb")]
        public IActionResult GetReadingFromOLEDB()
        {
            try
            {

                return Ok();
            }
            catch (Exception ex)
            {
                return Ok(ex.Message + " " + ex.StackTrace.ToString());
            }
        }
    }
}
