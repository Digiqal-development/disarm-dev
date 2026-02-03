using Disarm_Server;
using Disarm_Server.Models;
using Disarm_Server.Services;
using System.Diagnostics;
using System.Text.Json;
using System.Web;

var builder = WebApplication.CreateBuilder(args);
var codes = new Dictionary<string, string>();
var path = Path.Combine(Directory.GetCurrentDirectory(), "disarmData2.json");
var disarmDataJson = await File.ReadAllTextAsync(path);
var disarmTechniqueNames = JsonSerializer.Deserialize<List<DisarmTechniqueName>>(disarmDataJson);
var wrapper = new DisarmTechniqueNameWrapper { Techniques = disarmTechniqueNames! };

builder.Services.AddSingleton(wrapper);
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddScoped<IAttackNavigatorService, AttackNavigatorService>();

builder.Services.AddCors();


var app = builder.Build();

app.UseCors(builder => builder
.AllowAnyOrigin()
.AllowAnyMethod()
.AllowAnyHeader()
);

app.UseDeveloperExceptionPage();
app.UseSwagger();
app.UseSwaggerUI();

app.MapGet("/", () => "This is the Disarm foundation server");

app.MapGet("/techniques", () =>
{
    return Results.File(Directory.GetCurrentDirectory() + "/techniques.json", "application/json");
});


app.MapGet("/tags", () =>
{
    return Results.File(Directory.GetCurrentDirectory() + "/tags.json", "application/json");
});

app.MapGet("/json/{code}", (IAttackNavigatorService service, string code) =>
{
    if (!codes.TryGetValue(code, out var values))
    {
        return Results.NotFound();
    }
    var parsedIds = HttpUtility.UrlDecode(values);
    var ids = parsedIds.Split(",", StringSplitOptions.TrimEntries).ToArray();
    var data = service.GetNavigatorLayerBasedOnTechniqueIds(ids!);
    
    var json = JsonSerializer.Serialize(data);
    return Results.Content(json, "application/json");
});


app.MapPost("/clauses", async (ToDo input) =>
{

    var proc = new Process
    {
        StartInfo = new ProcessStartInfo
        {
            FileName = "/app/consoleapp/DisarmPythonResultGenerator",
            Arguments = "\"" + input.Sentence + "\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = false,
            CreateNoWindow = true
        }
    };

    proc.Start();
    await proc.WaitForExitAsync();

    var output = await proc.StandardOutput.ReadToEndAsync();
    input.Result = output;
    return input;
});

app.MapPost("/code-store", (CreateCodeRequest request) =>
{
    codes.Add(request.Code, request.Ids);
    return Results.Ok();
});

app.Run();