using Microsoft.AspNetCore.Mvc;
using AppERP.Services;

namespace AppERP.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly ExcelService _excelService;

        public ExcelController()
        {
            // Asegúrate de que el filePath apunte al archivo Excel correcto
            var filePath = "C:\\Users\\bruno\\Desktop\\ORACULO\\oraculo2\\AppERP\\data.xlsx";
            _excelService = new ExcelService(filePath);
        }

        [HttpGet]
        public IActionResult Get()
        {
            var (data, sheetNames) = _excelService.ReadExcel();
            return Ok(new { Data = data, SheetNames = sheetNames });
        }
    }
}
