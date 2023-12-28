using Disarm_Server;
using Microsoft.AspNetCore.Hosting.Server;
using System.Diagnostics;
using static System.Runtime.InteropServices.JavaScript.JSType;

var builder = WebApplication.CreateBuilder(args);


// Add services to the container.
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
    var process = new System.Diagnostics.Process
    {
        StartInfo = new System.Diagnostics.ProcessStartInfo
        {
            FileName = Directory.GetCurrentDirectory() + "/python.exe",
            Arguments = Directory.GetCurrentDirectory() + "/clause.py " + "\"" + input.Sentence + "\"",
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        }
    };
    process.Start();



    input.Result = await process.StandardOutput.ReadToEndAsync();
    process.WaitForExit();

    return input;



});



app.Run();