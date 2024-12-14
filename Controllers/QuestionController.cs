using Microsoft.AspNetCore.Mvc;
using AppERP.Services;

namespace AppERP.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class QuestionController : ControllerBase
    {
        private readonly ExcelService _excelService;

        public QuestionController()
        {
            var filePath = "C:\\Users\\bruno\\Desktop\\ORACULO\\oraculo2\\AppERP\\data.xlsx";
            _excelService = new ExcelService(filePath);
        }

        [HttpPost]
        public IActionResult Post([FromBody] QuestionRequest request)
        {
            var answer = _excelService.AnswerQuestion(request.Question);
            return Ok(new { Answer = answer });
        }
    }

    public class QuestionRequest
    {
        public string Question { get; set; }
    }
}
