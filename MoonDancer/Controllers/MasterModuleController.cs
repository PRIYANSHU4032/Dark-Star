using Microsoft.AspNetCore.Mvc;
using MoonDancer.Extracters;

namespace MoonDancer.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class MasterModuleController : Controller
    {
        private readonly MasterModuleManager _masterModuleManager;

        public MasterModuleController(MasterModuleManager MasterModuleManager)
        {
            _masterModuleManager = MasterModuleManager;
        }


        [HttpPost("Module-Master")]
        public IActionResult ModuleSyncer([FromQuery] string excelpath)
        {
            if (string.IsNullOrWhiteSpace(excelpath))
            {
                return BadRequest("searchString and pivotColumn are required.");
            }

            var result = _masterModuleManager.MasterBPSyncer(excelpath);
            if (result)
            {
                return Ok("Process synchronization started successfully.");
            }
            else
            {
                return StatusCode(500, $"An error occurred");
            }  
        }
    }
}
