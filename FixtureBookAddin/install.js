var arguments;
if(WScript.Arguments(0) === "/u") {
  arguments = "FixtureBookAddin.dll /unregister"
} else {
  arguments = "FixtureBookAddin.dll /codebase /tlb";
}

var wsh = WScript.CreateObject("WScript.Shell");
var sh = WScript.CreateObject("Shell.Application");
sh.shellExecute(wsh.ExpandEnvironmentStrings("%WinDir%") + "\\Microsoft.NET\\Framework\\v4.0.30319\\RegAsm.exe", arguments, "", "runas");
