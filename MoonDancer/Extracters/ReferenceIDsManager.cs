using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace MoonDancer.Extracters
{
    public class ReferenceIDsManager
    {


        private readonly string _maintainerPath;
        private readonly string _processSyncJsonPath;
        private readonly string _workdayfolderpath;
        private readonly string _discoveryfolderpath;
        private readonly string _excelpath;

        public ReferenceIDsManager(IConfiguration configuration)
        {
            _maintainerPath = configuration["AppSettings:Maintainer"];
            _processSyncJsonPath = Path.Combine(_maintainerPath, "ProcessSyncJson.json");
            _discoveryfolderpath = Path.Combine(_maintainerPath, "Discovery Processes Configurations");
            _workdayfolderpath = Path.Combine(_discoveryfolderpath, "Workday");
            _excelpath = configuration["AppSettings:Excel_path"];
        }


        public  bool referenceidExtracter(string excelpath)
        {
            Console.Clear();
            Logo.showLogo();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelWorksheet ws;
            using (var package = new ExcelPackage(new FileInfo(excelpath)))
            {
                ws = package.Workbook.Worksheets["Status"];
                if (ws == null)
                {
                    throw new Exception($"\nWorksheet status not found in the Excel file.");
                }

                for (int i = 2; i <= ws.Dimension.Rows; i++)
                {
                    var businessprocess = (string)ws.Cells[i, 1].Value;
                    var task = (string)ws.Cells[i, 2].Value;
                    var status = (string)ws.Cells[i, 4].Value;
                    if(status == "done")
                    {
                        refernceidExtractor_inner(businessprocess, task, excelpath);
                    }
                }
            }
            return true;
        }

        private void refernceidExtractor_inner(string businessprocess, string task, string excelpath)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            List<string> reference_ids = new List<string>();

            using (var package = new ExcelPackage(new FileInfo(excelpath)))
            {
                ExcelWorksheet ws = package.Workbook.Worksheets[task.Trim()];
                if (ws == null)
                {
                    throw new Exception($"\nWorksheet {task} not found in the Excel file.");
                }

                if (ws.Dimension == null)
                {
                    throw new Exception($"\nWorksheet {task} is empty.");
                }

                for (int i = 2; i <= ws.Dimension.Rows; i++)
                {
                    var ids = ws.Cells[i, 2].Value?.ToString();
                    if (!string.IsNullOrEmpty(ids))
                    {
                        reference_ids.Add(ids);
                    }
                }
            }

            var distinctReferenceIds = reference_ids.Distinct().ToList();

            string workdayFolderPath = Path.Combine(_workdayfolderpath, businessprocess.Trim());
            if (!Directory.Exists(workdayFolderPath))
            {
                Console.WriteLine($"\nFolder not found in Workday. Creating folder: {workdayFolderPath}");
                Directory.CreateDirectory(workdayFolderPath);
            }
            string configJsonPath = Path.Combine(workdayFolderPath, "ConfigurationDataJson.json");

            JArray configJson;
            if (File.Exists(configJsonPath))
            {
                string existingJson = File.ReadAllText(configJsonPath);
                configJson = JArray.Parse(existingJson);
            }
            else
            {
                configJson = new JArray();
            }
            foreach (var itm in distinctReferenceIds)
            {
                if (!configJson.Any(j => j["referenceId"]?.ToString() == itm))
                {
                    configJson.Add(new JObject
                    {
                        ["referenceId"] = itm,
                        ["fieldName"] = new JArray()
                    });
                }
            }
            File.WriteAllText(configJsonPath, JsonConvert.SerializeObject(configJson, Formatting.Indented));
            Console.WriteLine($"\nConfig.json updated at: {configJsonPath}");
        }


    }
}
