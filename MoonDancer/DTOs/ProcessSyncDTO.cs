namespace MoonDancer.DTOs
{
    public class ProcessSyncDTO
    {
        public string name { get; set; }
        public string db_name { get; set; }
        public string application { get; set; }
        public string submodule_name { get; set; }
        public string submodule_name_abbreviation { get; set; }
        public string module_name { get; set; }
        public string module_name_abbreviation { get; set; }
        public List<string> pivot_columns { get; set; }
        public WorkdayDetails workday_specific_process_details { get; set; }
        public bool isExecutbleProcess { get; set; }
        public string parentID { get; set; }
    }
}
