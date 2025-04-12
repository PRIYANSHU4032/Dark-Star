using Microsoft.AspNetCore.Mvc;
using MoonDancer.Extracters;

namespace MoonDancer.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ProcessSyncController : ControllerBase
    {
        private readonly ProcessSyncManager _processSyncManager;

        public ProcessSyncController(ProcessSyncManager processSyncManager)
        {
            _processSyncManager = processSyncManager;
        }

        [HttpPost("sync")]
        public IActionResult SyncProcess([FromQuery] string searchString, [FromQuery] string pivotColumn, [FromQuery] string module, [FromQuery] string submodule,[FromQuery] ProcessTypee processType, [FromQuery] string parent_id = null)
        {
            if (string.IsNullOrWhiteSpace(searchString) || string.IsNullOrWhiteSpace(pivotColumn) || string.IsNullOrWhiteSpace(module) || string.IsNullOrWhiteSpace(submodule))
            {
                return BadRequest("searchString and pivotColumn are required.");
            }

            try
            {
                _processSyncManager.ProcessSync(searchString, pivotColumn,module,submodule, processType.ToString());
                return Ok("Process synchronization started successfully.");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"An error occurred: {ex.Message}");
            }
        }
    }
}
