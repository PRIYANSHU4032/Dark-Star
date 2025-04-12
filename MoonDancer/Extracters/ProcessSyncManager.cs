using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using OfficeOpenXml;
using MoonDancer.DTOs;

namespace MoonDancer.Extracters
{
    public class ProcessSyncManager
    {
        private readonly string _maintainerPath;
        private readonly string _processSyncJsonPath;
        private readonly string _workdayfolderpath;
        private readonly string _discoveryfolderpath;
        private readonly string _excelpath;

        public ProcessSyncManager(IConfiguration configuration)
        {
            _maintainerPath = configuration["AppSettings:Maintainer"];
            _processSyncJsonPath = Path.Combine(_maintainerPath, "ProcessSyncJson.json");
            _discoveryfolderpath = Path.Combine(_maintainerPath, "Discovery Processes Configurations");
            _workdayfolderpath = Path.Combine(_discoveryfolderpath, "Workday");
            _excelpath = configuration["AppSettings:Excel_path"];
        }

        public void ProcessSync(string searchString, string pivotColumn,string module,string submodule, string processType , string prent_id = null)
        {
            if (!string.IsNullOrEmpty(searchString))
            {
                string jsonContent = File.Exists(_processSyncJsonPath) ? File.ReadAllText(_processSyncJsonPath) : "[]";
                var processSyncList = JsonConvert.DeserializeObject<List<JObject>>(jsonContent) ?? new List<JObject>();

                string processSyncJsonPath = _processSyncJsonPath;
                string excelFilePath = "~Resourses\\AB_BusinessProcess.xlsx";


                JArray processSyncArray = File.Exists(processSyncJsonPath)
                    ? JArray.Parse(File.ReadAllText(processSyncJsonPath))
                    : new JArray();


                var matchedEntry = processSyncArray.FirstOrDefault(entry =>
                            entry["db_name"]?.ToString().Equals(searchString, StringComparison.OrdinalIgnoreCase) == true &&
                            entry["application"]?.ToString() == "Workday");


                if (matchedEntry != null)
                {
                    Console.WriteLine($"Entry found in processSync.json: {searchString}");
                }
                else
                {
                    Console.WriteLine($"Entry not found. Creating new entry in processSync.json...");
                    var newEntry = new JObject
                    {
                        ["name"] = searchString,
                        ["db_name"] = searchString.ToLower(),
                        ["application"] = "Workday",
                        ["submodule_name"] = submodule,
                        ["submodule_name_abbreviation"] = submodule,
                        ["module_name"] = module,
                        ["module_name_abbreviation"] = module,
                        ["pivot_columns"] = new JArray(pivotColumn),
                        ["workday_specific_process_details"] = new JObject
                        {
                            ["process_defined_by"] = "Workday",
                            ["workday_definition_id"] = GetWorkdayDefinitionId(_excelpath, searchString.Trim()),
                            ["workday_transaction_id"] = "",
                            ["process_type"] = processType
                        },
                        ["isExecutbleProcess"] = true,
                        ["parentID"] = prent_id
                    };

                    processSyncArray.Add(newEntry);
                    File.WriteAllText(processSyncJsonPath, JsonConvert.SerializeObject(processSyncArray, Formatting.Indented));
                }

                string workdayFolderPath = Path.Combine(_workdayfolderpath, searchString.Trim());
                if (!Directory.Exists(workdayFolderPath) && !string.IsNullOrEmpty(pivotColumn) && processType != ProcessTypee.Task.ToString())
                {
                    Console.WriteLine($"Folder not found in Workday. Creating folder: {workdayFolderPath}");
                    Directory.CreateDirectory(workdayFolderPath);

                    string configJsonPath = Path.Combine(workdayFolderPath, "ConfigurationDataJson.json");
                    var configJson = new JArray
                {
                    new JObject
                    {
                        ["referenceId"] = pivotColumn,
                        ["fieldName"] = new JArray()
                    }
                };

                    File.WriteAllText(configJsonPath, JsonConvert.SerializeObject(configJson, Formatting.Indented));
                    Console.WriteLine($"Config.json created at: {configJsonPath}");
                }
                else if (!string.IsNullOrEmpty(pivotColumn)&& Directory.Exists(workdayFolderPath))
                {
                    Console.WriteLine($"Folder already exists: {workdayFolderPath}");
                    string configJsonPath = Path.Combine(workdayFolderPath, "ConfigurationDataJson.json");
                    string existingJson = string.Empty;
                    List<ConfigDTO> configData;
                    if (File.Exists(configJsonPath))
                    {
                        existingJson = File.ReadAllText(configJsonPath);
                        configData = JsonConvert.DeserializeObject<List<ConfigDTO>>(existingJson) ?? new List<ConfigDTO>();
                        if (!configData.Any(c => c.referenceId == pivotColumn))
                        {
                            configData.Add(new ConfigDTO
                            {
                                referenceId = pivotColumn,
                                fieldName = new List<string>()
                            });


                            File.WriteAllText(configJsonPath, JsonConvert.SerializeObject(configData, Formatting.Indented));
                            Console.WriteLine($"Updated JSON at: {configJsonPath}");
                        }
                    }
                    else
                    {
                        string filePath = Path.Combine(workdayFolderPath, "ConfigurationDataJson.json");
                        var configJson = new JArray
                            {
                                new JObject
                                {
                                    ["referenceId"] = pivotColumn,
                                    ["fieldName"] = new JArray()
                                }
                            };

                        File.WriteAllText(configJsonPath, JsonConvert.SerializeObject(configJson, Formatting.Indented));
                        Console.WriteLine($"Config.json created at: {configJsonPath}");
                    }
                    
                }
            }
            
        }
        //catch (Exception ex)
        //{
        //    Console.WriteLine($"Error: {ex.Message}");
        //}
    


        public string GetWorkdayDefinitionId(string excelPath, string searchValue)
        {
            if (!File.Exists(excelPath))
            {
                Console.WriteLine("Excel file not found.");
                return "";
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                var ws = package.Workbook.Worksheets.FirstOrDefault();
                if (ws == null)
                {
                    Console.WriteLine("No worksheets found in the Excel file.");
                    return "";
                }

                for (int row = 2; row <= ws.Dimension.Rows; row++)
                {
                    string columnA = ws.Cells[row, 1].Value?.ToString().Trim();
                    string columnB = ws.Cells[row, 2].Value?.ToString().Trim();

                    if (!string.IsNullOrEmpty(columnA) && columnA.Equals(searchValue, StringComparison.OrdinalIgnoreCase))
                    {
                        return columnB; 
                    }
                }
            }
            return "";
        }
    }
}
