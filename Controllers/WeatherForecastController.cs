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
                }
                return Ok(content);
            }
            catch (Exception ex)
            {
                return Ok(ex.Message + " " + ex.StackTrace.ToString());
            }
            
        }
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
    }
}
