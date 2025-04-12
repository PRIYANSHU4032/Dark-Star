using MoonDancer.DTOs;
using System.Security.Cryptography;
using System.IO;
using System.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using OfficeOpenXml;

using System.ComponentModel;
using Newtonsoft.Json;
using OfficeOpenXml.Style;
using System.Drawing;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;

namespace MoonDancer.Extracters
{
    public class ExcelTableExtractor
    {


        private readonly ProcessSyncManager _processSyncManager;
        private readonly string _maintainerPath;
        private readonly string _processSyncJsonPath;
        private readonly string _workdayfolderpath;
        private readonly string _discoveryfolderpath;

        public ExcelTableExtractor(ProcessSyncManager processSyncManager, IConfiguration configuration)
        {
            _processSyncManager = processSyncManager;
            _maintainerPath = configuration["AppSettings:Maintainer"];
            _processSyncJsonPath = Path.Combine(_maintainerPath, "ProcessSyncJson.json");
            _discoveryfolderpath = Path.Combine(_maintainerPath, "Discovery Processes Configurations");
            _workdayfolderpath = Path.Combine(_discoveryfolderpath, "Workday");
        }
        private static string GetMD5Hash(string input)
        {
            using (MD5 md5 = MD5.Create())
            {
                byte[] hashBytes = md5.ComputeHash(System.Text.Encoding.UTF8.GetBytes(input));
                return BitConverter.ToString(hashBytes).Replace("-", "").ToLower();
            }
        }

        public bool ExtractTables(string filePath,string module,string submodule,string parent_id ,string pivotcolumn ,string sheetname = null)
        {
            Console.Clear();
            Logo.showLogo();
            var processModulesWithActivities = new Dictionary<String, ProcessModulesWithActivities>();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            bool isDublicatePresent = false;

            List<string> business_process = new List<string>();
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet ws;
                if (!string.IsNullOrEmpty(sheetname))
                {
                    ws = package.Workbook.Worksheets[sheetname];
                    if (ws == null)
                    {
                        throw new Exception($"Worksheet '{sheetname}' not found in the Excel file.");
                    }
                }
                else
                {
                    ws = package.Workbook.Worksheets.Where(p => p.Hidden == eWorkSheetHidden.Visible).First();
                    var wss = package.Workbook.Worksheets.First();
                }
                

                for (int i = 2; i <= ws.Dimension.Rows; i++)
                {
                    var myNode = new ProcessModulesWithActivities();

                    var scenarioName = (string)ws.Cells[i, 2].Value;
                    var pivotColName = (string)ws.Cells[i, 10].Value;
                    myNode.name = scenarioName;
                    myNode.group_name = "Standard";


                    var preReqs = new List<string>();

                    var preReqData = (string)ws.Cells[i, 5].Value;
                    var desc = "";
                    desc = (string)ws.Cells[i, 14].Value;
                    if (!String.IsNullOrEmpty(desc))
                    {
                        myNode.description = desc;
                    }

                    if (!String.IsNullOrEmpty(preReqData))
                    {
                        preReqs = ((string)ws.Cells[i, 5].Value).Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
                        preReqs = (from itm in preReqs select itm.Trim().ToLower()).ToList();
                    }

                    var myInitStep = (string)ws.Cells[i, 8].Value;

                    if (myInitStep != null)
                    {
                        myInitStep = myInitStep.Trim().ToLower();
                    }


                    var myBP = (string)ws.Cells[i, 4].Value;

                    var myTask = (string)ws.Cells[i, 3].Value;

                    var processType = ProcessTypee.BusinessProcess.ToString();

                    //if(myBP == null && myTask != null)
                    //{
                    //    if(myInitStep != null && myInitStep != "")
                    //    {
                    //        if(myInitStep.ToLower().Trim() != myTask.ToLower().Trim())
                    //        {
                    //            myBP = myTask;
                    //            processType = ProcessTypee.Task.ToString();
                    //        }
                    //    }
                    //    else
                    //    {
                    //        myBP = myTask;
                    //        processType = ProcessTypee.Task.ToString();
                    //    }
                        
                    //}
                    String myBP_Lower = null;
                    if (myBP != null)
                    {
                        myBP_Lower = myBP.Trim().ToLower();
                        business_process.Add(myBP);
                        _processSyncManager.ProcessSync(myBP.Trim(), pivotColName, module, submodule, processType, parent_id);
                    }
                    if (myInitStep != null)
                    {
                        var scenarioPath = new List<string>();
                        scenarioPath.AddRange(preReqs);

                        scenarioPath.Add(myInitStep);

                        var activitesSTR = $"['{string.Join("', '", scenarioPath).Replace("\"", "'")}']";
                        var scenario_hash = GetMD5Hash(activitesSTR).ToLower();
                        myNode.scenario_hash = scenario_hash;
                        myNode.scenario_path = new List<strpss>();
                        myNode.pivot_column = pivotColName;
                        foreach (var item in preReqs)
                        {
                            business_process.Add(item);
                            _processSyncManager.ProcessSync(item.Trim(), pivotColName, module, submodule, processType);
                            myNode.scenario_path.Add(new strpss() { business_process_name = item, initiator_component_name = null });
                        }

                        myNode.scenario_path.Add(new strpss() { business_process_name = myBP_Lower, initiator_component_name = myInitStep });



                        if (!processModulesWithActivities.ContainsKey(scenario_hash) && !String.IsNullOrEmpty(myNode.name))
                        {
                            processModulesWithActivities.Add(scenario_hash, myNode);
                        }
                        else
                        {
                            Console.WriteLine($"\n The error ouccure at {i}");
                            ColorRow(filePath, i, Color.Red, ws);
                            isDublicatePresent = true;
                        }

                    }
                    else
                    {
                        var scenarioPath = new List<string>();
                        scenarioPath.AddRange(preReqs);

                        scenarioPath.Add(myBP_Lower);
                        business_process.Add(myBP);
                        if (!string.IsNullOrEmpty(myBP))
                        {
                            _processSyncManager.ProcessSync(myBP.Trim(), pivotColName, module, submodule, processType);
                        }
                        

                        var activitesSTR = $"['{string.Join("', '", scenarioPath).Replace("\"", "'")}']";
                        var scenario_hash = GetMD5Hash(activitesSTR).ToLower();
                        myNode.scenario_hash = scenario_hash;
                        myNode.scenario_path = new List<strpss>();
                        myNode.pivot_column = pivotColName;
                        foreach (var item in preReqs)
                        {
                            business_process.Add(item);
                            _processSyncManager.ProcessSync(item.Trim(), pivotColName, module, submodule, processType, parent_id);
                            myNode.scenario_path.Add(new strpss() { business_process_name = item, initiator_component_name = null });
                        }

                        myNode.scenario_path.Add(new strpss() { business_process_name = myBP_Lower, initiator_component_name = null });

                        if (!processModulesWithActivities.ContainsKey(scenario_hash) && !String.IsNullOrEmpty(myNode.name))
                        {
                            processModulesWithActivities.Add(scenario_hash, myNode);
                        }
                        else
                        {
                            Console.WriteLine($"\n The error ouccure at {i}");
                            ColorRow(filePath, i, Color.Red, ws);
                            isDublicatePresent = true;
                        }

                    }
                }
            }

            //if (isDublicatePresent)
            //{
            //    throw new Exception($"Your Excel sheet contain some dublicate entries either component name or business process or pre-req, Please Fix first in red rows");
            //}

            var reverseDic = new Dictionary<string, ProcessModulesWithActivities>();

            foreach (var item in processModulesWithActivities.Reverse())
            {
                reverseDic.Add(item.Key, item.Value);
            }

            var json = JsonConvert.SerializeObject(reverseDic);
            json = JsonConvert.SerializeObject(processModulesWithActivities);
            ReverseJsonOrderAndSave(json,parent_id);
            MakeEntryIfNot(parent_id, module, submodule, pivotcolumn);
            SetParent_ID(business_process, parent_id);
            return true;

        }

        private static void ColorRow(string filePath, int rowIndex, Color color, ExcelWorksheet ws)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;


            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {

                if (rowIndex < 1 || rowIndex > ws.Dimension.Rows)
                {
                    throw new ArgumentException($"Invalid row index: {rowIndex}. It must be between 1 and {ws.Dimension.Rows}");
                }


                int totalColumns = ws.Dimension.Columns;


                using (ExcelRange rng = ws.Cells[rowIndex, 1, rowIndex, totalColumns])
                {
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(color);
                }


                package.Save();
            }
        }

        private void ReverseJsonOrderAndSave(string json,string parent)
        {
           
            var jsonObject = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);

            
            var reversedData = new Dictionary<string, object>(jsonObject.Reverse());

            
            string reversedJson = JsonConvert.SerializeObject(reversedData, Formatting.Indented);
            var folder = _workdayfolderpath + "/" + parent;
            Directory.CreateDirectory(folder);
            string fullFilePath = Path.Combine(folder, $"ScenariosMasterData.json");

            File.WriteAllText(fullFilePath, reversedJson);
            

            Console.WriteLine($"Reversed JSON saved to {fullFilePath}");
        }

        private void SetParent_ID(List<string> bps, string parent)
        {
  
            string jsonContent = File.ReadAllText(_processSyncJsonPath);
            JArray processSyncArray = JArray.Parse(jsonContent);
            bps = bps.Distinct().ToList();

            foreach (var itm in bps)
            {
                var matchedEntry = processSyncArray.FirstOrDefault(entry =>
                    entry["db_name"]?.ToString().Equals(itm, StringComparison.OrdinalIgnoreCase) == true &&
                    entry["application"]?.ToString() == "Workday");

                if (matchedEntry != null) 
                {
                    if (matchedEntry["parentID"] == null || matchedEntry["parentID"].Type == JTokenType.Null)
                    {
                        matchedEntry["parentID"] = parent;
                    }
                    else
                    {
                        string parents = matchedEntry["parentID"].ToString();
                        List<string> resultList = new List<string>(parents.Split(',')); 

                        if (!resultList.Contains(parent))
                        {
                            resultList.Add(parent);
                        }

                        matchedEntry["parentID"] = string.Join(",", resultList);
                    }
                }
            }
            File.WriteAllText(_processSyncJsonPath, processSyncArray.ToString(Formatting.Indented));
            Console.WriteLine("Updated JSON saved successfully!");
        }


        private void MakeEntryIfNot(string parent,string module,string submodule,string pivotcolumn)
        {
            string jsonContent = File.Exists(_processSyncJsonPath) ? File.ReadAllText(_processSyncJsonPath) : "[]";
            var processSyncList = JsonConvert.DeserializeObject<List<JObject>>(jsonContent) ?? new List<JObject>();
            JArray processSyncArray = File.Exists(_processSyncJsonPath)
                ? JArray.Parse(File.ReadAllText(_processSyncJsonPath))
                : new JArray();
            var matchedEntry = processSyncArray.FirstOrDefault(entry =>
                entry["db_name"]?.ToString().Equals(parent, StringComparison.OrdinalIgnoreCase) == true &&
                entry["application"]?.ToString() == "Workday" &&
                entry["workday_specific_process_details"]?["process_defined_by"]?.ToString() == "Opkey"
                                );


            if (matchedEntry != null)
            {
                Console.WriteLine($"Entry found in processSync.json: {parent}");
            }
            else
            {
                Console.WriteLine($"Entry not found. Creating new entry in processSync.json...");
                var newEntry = new JObject
                {
                    ["name"] = parent,
                    ["db_name"] = parent,
                    ["application"] = "Workday",
                    ["submodule_name"] = submodule,
                    ["submodule_name_abbreviation"] = submodule,
                    ["module_name"] = module,
                    ["module_name_abbreviation"] = module,
                    ["pivot_columns"] = new JArray(pivotcolumn),
                    ["workday_specific_process_details"] = new JObject
                    {
                        ["process_defined_by"] = "Opkey",
                        ["workday_definition_id"] = "",
                        ["workday_transaction_id"] = "",
                        ["process_type"] = "BusinessProcess"
                    },
                    ["isExecutbleProcess"] = true,
                    ["parentID"] = null
                };

                processSyncArray.Add(newEntry);
                File.WriteAllText(_processSyncJsonPath, JsonConvert.SerializeObject(processSyncArray, Formatting.Indented));



            }
        }
    }
    public enum ProcessTypee
    {
        BusinessProcess,
        Task
    }
}
