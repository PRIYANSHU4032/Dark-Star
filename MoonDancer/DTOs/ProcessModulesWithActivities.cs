namespace MoonDancer.DTOs
{
    public class ProcessModulesWithActivities
    {
        public string name;
        public string group_name;
        public List<strpss> scenario_path;
        public string scenario_hash;
        public string pivot_column;
        public string description { get; set; }
    }

    public class strpss
    {
        public string initiator_component_name;
        public string business_process_name;
    }
}
