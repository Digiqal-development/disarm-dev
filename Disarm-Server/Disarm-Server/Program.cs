using Disarm_Server;
using Microsoft.AspNetCore.Hosting.Server;
using Python.Runtime;
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
 
    Runtime.PythonDLL = @"C:\Python311\python311.dll";
    var pathToVirtualEnv = @"C:\Users\esmak\Desktop\add_in\disarm-dev\Disarm-Server\Disarm-Server\pyvenv.cfg";

    var path = Environment.GetEnvironmentVariable("PATH").TrimEnd(';');
    path = string.IsNullOrEmpty(path) ? pathToVirtualEnv : path + ";" + pathToVirtualEnv;
    Environment.SetEnvironmentVariable("PATH", path, EnvironmentVariableTarget.Process);
    Environment.SetEnvironmentVariable("PATH", pathToVirtualEnv, EnvironmentVariableTarget.Process);
    Environment.SetEnvironmentVariable("PYTHONPATH", $"{pathToVirtualEnv}\\Lib\\site-packages;{pathToVirtualEnv}\\Lib", EnvironmentVariableTarget.Process);

    PythonEngine.Initialize();

    PythonEngine.PythonHome = pathToVirtualEnv;
    PythonEngine.PythonPath = Environment.GetEnvironmentVariable("PYTHONPATH", EnvironmentVariableTarget.Process);

    using (Py.GIL())
    {
        
        var fromFile = Py.Import(Path.GetFileNameWithoutExtension(@"C:\Users\esmak\Desktop\add_in\disarm-dev\Disarm-Server\Disarm-Server\clause.py"));
        var result = fromFile.InvokeMethod("main_part", Py.kw("sentence", input.Sentence));
        input.Result = result.ToString();
        

    }

   

    return input;



});



app.Run();