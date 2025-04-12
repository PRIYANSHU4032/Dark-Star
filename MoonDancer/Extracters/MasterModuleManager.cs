using MoonDancer.DTOs;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

namespace MoonDancer.Extracters
{
    public class MasterModuleManager
    {
        private readonly string _maintainerPath;
        private readonly string _processSyncJsonPath;
        private readonly string _workdayfolderpath;
        private readonly string _discoveryfolderpath;
        private readonly string _excelpath;
        private readonly ProcessSyncManager _processsyncmanager;


        public MasterModuleManager(IConfiguration configuration, ProcessSyncManager processSyncManager)
        {
            _maintainerPath = configuration["AppSettings:Maintainer"];
            _processSyncJsonPath = Path.Combine(_maintainerPath, "ProcessSyncJson.json");
            _discoveryfolderpath = Path.Combine(_maintainerPath, "Discovery Processes Configurations");
            _workdayfolderpath = Path.Combine(_discoveryfolderpath, "Workday");
            _excelpath = configuration["AppSettings:Excel_path"];
            _processsyncmanager = processSyncManager;
        }

        public bool MasterBPSyncer(string excelpath)
        {
            Console.Clear();
            Logo.showLogo();
            string jsonContent = File.Exists(_processSyncJsonPath) ? File.ReadAllText(_processSyncJsonPath) : "[]";
            var processSyncList = JsonConvert.DeserializeObject<List<JObject>>(jsonContent) ?? new List<JObject>();

            List<ProcessSyncDTO> newEntries = ReadExcelAndGenerateJson(excelpath);

            foreach (var newEntry in newEntries)
            {

                var existingJsonObj = processSyncList.FirstOrDefault(p =>
                    p["db_name"]?.ToString() == newEntry.db_name && p["application"]?.ToString() == "Workday");

                if (existingJsonObj != null)
                {

                    existingJsonObj["name"] = newEntry.name;
                    existingJsonObj["module_name"] = newEntry.module_name;
                    existingJsonObj["module_name_abbreviation"] = newEntry.module_name_abbreviation;
                    existingJsonObj["isExecutbleProcess"] = newEntry.isExecutbleProcess;
                    existingJsonObj["parentID"] = newEntry.parentID;


                    if (existingJsonObj.ContainsKey("submodule_name"))
                    {
                        existingJsonObj["submodule_name"] = newEntry.submodule_name;
                    }

                    if (existingJsonObj.ContainsKey("submodule_name_abbreviation"))
                    {
                        existingJsonObj["submodule_name_abbreviation"] = newEntry.submodule_name_abbreviation;
                    }

                    if (existingJsonObj.ContainsKey("pivot_columns"))
                    {
                        existingJsonObj["pivot_columns"] = JToken.FromObject(newEntry.pivot_columns);
                    }

                    if (existingJsonObj.ContainsKey("workday_specific_process_details"))
                    {
                        existingJsonObj["workday_specific_process_details"] = JToken.FromObject(newEntry.workday_specific_process_details);
                    }
                }
                //else
                //{

                //    var newJsonObj = JObject.FromObject(newEntry);
                //    processSyncList.Add(newJsonObj);
                //}
            }

            File.WriteAllText(_processSyncJsonPath, JsonConvert.SerializeObject(processSyncList, Newtonsoft.Json.Formatting.Indented));
            return true;
        }




        private  List<ProcessSyncDTO> ReadExcelAndGenerateJson(string excelPath)
        {
            List<ProcessSyncDTO> entries = new List<ProcessSyncDTO>();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    List<string> reference_ids = new List<string>();
                    var pivots = new List<string>();
                    pivots = new List<string> { worksheet.Cells[row, 5].Text.Trim() };
                    if(pivots.Contains("") && worksheet.Cells[row, 7].Text.Trim() == "BusinessProcess")
                    {
                        reference_ids = ((string)worksheet.Cells[row, 6].Value).Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
                        reference_ids = (from itm in reference_ids select itm.Trim()).ToList();
                        pivots = new List<string>();
                    }

                    var entry = new ProcessSyncDTO
                    {
                        name = worksheet.Cells[row, 1].Text.Trim(),
                        db_name = worksheet.Cells[row, 1].Text.Trim(),
                        application = "Workday",
                        submodule_name = worksheet.Cells[row, 3].Text.Trim(),
                        submodule_name_abbreviation = worksheet.Cells[row, 3].Text.Trim(),
                        module_name = worksheet.Cells[row, 2].Text.Trim(),
                        module_name_abbreviation = worksheet.Cells[row, 2].Text.Trim(),
                        pivot_columns = pivots,
                        workday_specific_process_details = new WorkdayDetails
                        {
                            process_defined_by = "Workday",
                            workday_definition_id = _processsyncmanager.GetWorkdayDefinitionId(_excelpath, worksheet.Cells[row, 1].Text.Trim()),
                            workday_transaction_id = worksheet.Cells[row, 10] != null ? worksheet.Cells[row, 9].Text.Trim() : string.Empty,
                            process_type = worksheet.Cells[row, 7].Text.Trim(),
                        },
                        isExecutbleProcess = true,
                        parentID = worksheet.Cells[row, 8].Text.Trim()
                    };

                    entries.Add(entry);
                    //if(entry.workday_specific_process_details.process_type == "BusinessProcess" && entry.pivot_columns[0] != "")
                    //{
                    //    string folderPath = Path.Combine(_workdayfolderpath, entry.db_name);
                    //    string jsonpath = Path.Combine(folderPath, "ConfigurationDataJson.json");
                    //    if (Directory.Exists(folderPath))
                    //    {
                    //        List<ConfigDTO> configData;
                    //        string existingJson = string.Empty;
                    //        string filePath = Path.Combine(folderPath, "ConfigurationDataJson.json");
                    //        if (File.Exists(filePath))
                    //        {
                    //            existingJson = File.ReadAllText(filePath);
                                
                    //        }
                    //        configData = JsonConvert.DeserializeObject<List<ConfigDTO>>(existingJson) ?? new List<ConfigDTO>();

                    //        foreach (var item in reference_ids)
                    //        {
                    //            if (!configData.Any(c => c.referenceId == item))
                    //            {
                    //                configData.Add(new ConfigDTO
                    //                {
                    //                    referenceId = item,
                    //                    fieldName = new List<string>()
                    //                });


                    //                File.WriteAllText(filePath, JsonConvert.SerializeObject(configData, Formatting.Indented));
                    //                Console.WriteLine($"Updated JSON at: {filePath}");
                    //            }
                    //        }


                    //        File.WriteAllText(filePath, JsonConvert.SerializeObject(configData, Formatting.Indented));
                    //    }
                    //    else
                    //    {
                    //        Directory.CreateDirectory(folderPath);
                    //        string filePath = Path.Combine(folderPath, "ConfigurationDataJson.json");
                    //        var configData = new List<ConfigDTO>();
                    //        foreach (var item in reference_ids)
                    //        {
                    //            configData.Add(new ConfigDTO
                    //            {
                    //                referenceId = item,
                    //                fieldName = new List<string>()
                    //            });                               
                    //        }
                    //        File.WriteAllText(filePath, JsonConvert.SerializeObject(configData, Formatting.Indented));
                    //    }

                    //}
                    
                }

            }
            return entries;
        }
    }
}
