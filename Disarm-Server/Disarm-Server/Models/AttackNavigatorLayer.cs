using System.Text.Json.Serialization;

namespace Disarm_Server.Models;

public class AttackNavigatorLayer
{
    [JsonPropertyName("name")]
    public string Name { get; set; }

    [JsonPropertyName("versions")]
    public Versions Versions { get; set; }

    [JsonPropertyName("domain")]
    public string Domain { get; set; }

    [JsonPropertyName("description")]
    public string Description { get; set; }

    [JsonPropertyName("filters")]
    public Filters Filters { get; set; }

    [JsonPropertyName("sorting")]
    public int Sorting { get; set; }

    [JsonPropertyName("layout")]
    public Layout Layout { get; set; }

    [JsonPropertyName("hideDisabled")]
    public bool HideDisabled { get; set; }

    [JsonPropertyName("techniques")]
    public List<Technique> Techniques { get; set; }

    [JsonPropertyName("gradient")]
    public Gradient Gradient { get; set; }

    [JsonPropertyName("legendItems")]
    public List<object> LegendItems { get; set; }

    [JsonPropertyName("metadata")]
    public List<object> Metadata { get; set; }

    [JsonPropertyName("links")]
    public List<object> Links { get; set; }

    [JsonPropertyName("showTacticRowBackground")]
    public bool ShowTacticRowBackground { get; set; }

    [JsonPropertyName("tacticRowBackground")]
    public string TacticRowBackground { get; set; }

    [JsonPropertyName("selectTechniquesAcrossTactics")]
    public bool SelectTechniquesAcrossTactics { get; set; }

    [JsonPropertyName("selectSubtechniquesWithParent")]
    public bool SelectSubtechniquesWithParent { get; set; }

    // Constructor to easily create an instance with the provided JSON values
    public AttackNavigatorLayer()
    {
        Name = "layer";
        Versions = new Versions
        {
            Attack = "1",
            Navigator = "4.8.2",
            Layer = "4.4"
        };
        Domain = "DISARM";
        Description = "";
        Filters = new Filters
        {
            Platforms = ["Windows", "Linux", "Mac"]
        };
        Sorting = 0;
        Layout = new Layout
        {
            LayoutType = "side",
            AggregateFunction = "average",
            ShowID = false,
            ShowName = true,
            ShowAggregateScores = false,
            CountUnscored = false
        };
        HideDisabled = false;
        Techniques =
        [
            new Technique
            {
                TechniqueID = "T0100",
                Tactic = "establish-legitimacy",
                Score = 1,
                Color = "#e60d0d",
                Comment = "",
                Enabled = true,
                Metadata = [],
                Links = [],
                ShowSubtechniques = false
            }
        ];
        Gradient = new Gradient
        {
            Colors = new List<string> { "#ff6666ff", "#ffe766ff", "#8ec843ff" },
            MinValue = 0,
            MaxValue = 100
        };
        LegendItems = [];
        Metadata = [];
        Links = [];
        ShowTacticRowBackground = false;
        TacticRowBackground = "#dddddd";
        SelectTechniquesAcrossTactics = true;
        SelectSubtechniquesWithParent = false;
    }
}

public class Versions
{
    [JsonPropertyName("attack")]
    public string Attack { get; set; }

    [JsonPropertyName("navigator")]
    public string Navigator { get; set; }

    [JsonPropertyName("layer")]
    public string Layer { get; set; }
}

public class Filters
{
    [JsonPropertyName("platforms")]
    public List<string> Platforms { get; set; }
}

public class Layout
{
    [JsonPropertyName("layout")]
    public string LayoutType { get; set; }

    [JsonPropertyName("aggregateFunction")]
    public string AggregateFunction { get; set; }

    [JsonPropertyName("showID")]
    public bool ShowID { get; set; }

    [JsonPropertyName("showName")]
    public bool ShowName { get; set; }

    [JsonPropertyName("showAggregateScores")]
    public bool ShowAggregateScores { get; set; }

    [JsonPropertyName("countUnscored")]
    public bool CountUnscored { get; set; }
}

public class Technique
{
    [JsonPropertyName("techniqueID")]
    public string TechniqueID { get; set; }

    [JsonPropertyName("tactic")]
    public string Tactic { get; set; }

    [JsonPropertyName("score")]
    public int Score { get; set; }

    [JsonPropertyName("color")]
    public string Color { get; set; }

    [JsonPropertyName("comment")]
    public string Comment { get; set; }

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }

    [JsonPropertyName("metadata")]
    public List<object> Metadata { get; set; }

    [JsonPropertyName("links")]
    public List<object> Links { get; set; }

    [JsonPropertyName("showSubtechniques")]
    public bool ShowSubtechniques { get; set; }
}

public class Gradient
{
    [JsonPropertyName("colors")]
    public List<string> Colors { get; set; }

    [JsonPropertyName("minValue")]
    public int MinValue { get; set; }

    [JsonPropertyName("maxValue")]
    public int MaxValue { get; set; }
}