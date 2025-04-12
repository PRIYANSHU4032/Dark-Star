using Microsoft.AspNetCore.Mvc;
using MoonDancer.Extracters;

namespace MoonDancer.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class SheetSyncerController : Controller
    {
        private readonly ExcelTableExtractor _excelTableExtractor;
        public SheetSyncerController(ExcelTableExtractor excelTableExtractor)
        {
            _excelTableExtractor = excelTableExtractor;
        }

        [HttpPost("SheetSyncer")]
        public IActionResult ExtractTables([FromQuery] string filePath,  [FromQuery] string module, [FromQuery] string submodule, [FromQuery] string parent_id, [FromQuery] string pivotcolumn,[FromQuery] string sheetname = null)
        {
            var result = _excelTableExtractor.ExtractTables(filePath, module, submodule,parent_id,pivotcolumn, sheetname);
            if (result)
            {
                return Ok("Sheet synchronization successfully.");
            }
            else
            {
                return BadRequest("Sheet synchronization failed.");
            }
        }
    }
}
