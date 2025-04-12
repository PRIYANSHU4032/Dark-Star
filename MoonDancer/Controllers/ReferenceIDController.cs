using Microsoft.AspNetCore.Mvc;
using MoonDancer.Extracters;

namespace MoonDancer.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ReferenceIDController : Controller
    {
        private readonly ReferenceIDsManager _referenceIDsManager;

        public ReferenceIDController(ReferenceIDsManager ReferenceIDsManager)
        {
            _referenceIDsManager = ReferenceIDsManager;
        }

        [HttpPost("Refernece_ID-Syncer")]
        public IActionResult ModuleSyncer([FromQuery] string excelpath)
        {
            if (string.IsNullOrWhiteSpace(excelpath))
            {
                return BadRequest("searchString and pivotColumn are required.");
            }

            var result = _referenceIDsManager.referenceidExtracter(excelpath);
            if (result)
            {
                return Ok("Process synchronization  successfully.");
            }
            else
            {
                return StatusCode(500, $"An error occurred");
            }
        }


    }
}
