namespace Disarm_Server.Models
{
    public class GraphRequest
    {
        public string CentralNode  { get; set; }
        public List<GraphObject> Objects { get; set; }
    }

    public class GraphObject
    {
        public string Name { get; set; }
        public string Type { get; set; }
    }
}