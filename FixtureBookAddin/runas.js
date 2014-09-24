var wsh = WScript.CreateObject("WScript.Shell");
var sh = WScript.CreateObject("Shell.Application");
sh.shellExecute("cmd.exe", "/k " + wsh.CurrentDirectory + "\\" + WScript.Arguments(0), "", "runas");
