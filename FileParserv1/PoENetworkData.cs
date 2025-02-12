namespace FileParserv1;
public class PowerOverEthernetTask
{
    public string? SerialNumber { get; set; }
    public Dictionary<string, string>? SummaryData { get; set; }
    public List<NetworkTask> NetworkTasks { get; set; }
    public bool HasPassed { get; set; }

    public PowerOverEthernetTask()
    {
        NetworkTasks = new List<NetworkTask>();
    }
    public class Summary
    {
        public string? Name { get; set; }
        public string? Value { get; set; }
    }

}

public class NetworkTask
{
    public string? Name { get; set; }
    public string? Status { get; set; }
    public List<TaskSection> TaskSections { get; set; }

    public NetworkTask()
    {
        TaskSections = new List<TaskSection>();
    }
    public class TaskSection
    {
        public string? Name { get; set; }
        public string? Value { get; set; }
        public bool IsDataSet { get; set; }
    }
}

