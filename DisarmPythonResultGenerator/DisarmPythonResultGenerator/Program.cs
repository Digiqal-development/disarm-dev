
using Python.Runtime;

Runtime.PythonDLL = @"C:\Users\Administrator\AppData\Local\Programs\Python\Python311\python311.dll";
var pathToVirtualEnv = @"C:\inetpub\disarm-result-generator\venv";

var path = Environment.GetEnvironmentVariable("PATH").TrimEnd(';');
path = string.IsNullOrEmpty(path) ? pathToVirtualEnv : path + ";" + pathToVirtualEnv;
Environment.SetEnvironmentVariable("PATH", path, EnvironmentVariableTarget.Process);
Environment.SetEnvironmentVariable("PATH", pathToVirtualEnv, EnvironmentVariableTarget.Process);
Environment.SetEnvironmentVariable("PYTHONPATH", $"{pathToVirtualEnv}\\Lib\\site-packages;{pathToVirtualEnv}\\Lib", EnvironmentVariableTarget.Process);

PythonEngine.Initialize();
PythonEngine.BeginAllowThreads();
PythonEngine.PythonHome = pathToVirtualEnv;
PythonEngine.PythonPath = Environment.GetEnvironmentVariable("PYTHONPATH", EnvironmentVariableTarget.Process);

using (Py.GIL())
{
    var fromFile = Py.Import(Path.GetFileNameWithoutExtension(@"C:\inetpub\disarm-result-generator\clause.py"));
    var result = fromFile.InvokeMethod("main_part", Py.kw("sentence", args[0]));
    Console.WriteLine(result.ToString());
}