var prog = "C:/Program Files/7-Zip/7zG.exe";
var dist = "%windir%\Temp"
var args = WScript.Arguments;
var shell = new ActiveXObject("WScript.shell");
var fso = new ActiveXObject('Scripting.FileSystemObject');

var s = args(0);
var b = fso.GetBaseName(s);
var w = dist + b;
var x = "\""+ prog +"\"" + " x " + "\""+ s +"\"" + " " + "\""+ "-o" + w +"\"";
WScript.Echo(x);
shell.Run(x, 1, true);
var y = "explorer " + fso.GetAbsolutePathName(w);
// WScript.Echo(y);
shell.Run(y);
