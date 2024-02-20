using Disarm_Server;
using Microsoft.AspNetCore.Hosting.Server;
using Python.Runtime;
using System.Diagnostics;
using static System.Runtime.InteropServices.JavaScript.JSType;

var builder = WebApplication.CreateBuilder(args);


builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

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



app.MapPost("/clauses", async (ToDo input) =>
{

    var proc = new Process
    {
        StartInfo = new ProcessStartInfo
        {
            FileName = "C:\\inetpub\\disarm-result-generator\\DisarmPythonResultGenerator.exe",
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



app.Run();