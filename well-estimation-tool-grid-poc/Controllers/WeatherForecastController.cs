using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using well_estimation_tool_grid_poc.Helpers;

namespace well_estimation_tool_grid_poc.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
        "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
    };
        private ExcelHelpers excelHelpers;

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
            excelHelpers = new ExcelHelpers();
        }

        [HttpGet(Name = "GetWeatherForecast")]
        public IEnumerable<WeatherForecast> Get()
        {
            //excelHelpers.BeginExcelWorkFlow();
            //excelHelpers.ReadExcel_Dummy2();
            // excelHelpers.ReadExcel_Dummy1();

            ExcelOpenXMLHelper excelOpenXMLHelper = new ExcelOpenXMLHelper();
            // excelOpenXMLHelper.TestOpenXMLExcel();// this is working but lacks feature of the datatable
            // excelOpenXMLHelper.TestOpenXMLExcelInDataTable(); // this is working but lacks model to map the data from datatables
            //excelOpenXMLHelper.CallExportDataSet();

            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateTime.Now.AddDays(index),
                TemperatureC = Random.Shared.Next(-20, 55),
                Summary = Summaries[Random.Shared.Next(Summaries.Length)]
            })
            .ToArray();
        }
            
    }
}