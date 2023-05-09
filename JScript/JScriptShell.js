if (WScript.Arguments.Named("host") == "scriptcontrol") {
    while (true) {
		try {
        with(new ActiveXObject("MSScriptControl.ScriptControl")) { //This must be run in 32-bit CScript
            Language = "JScript";
            AllowUI = true;
            try {
                WScript.StdOut.Write(">>> ");
                WScript.StdOut.WriteLine(AddCode(WScript.StdIn.ReadLine()));
            } catch (e) {
                WScript.StdErr.WriteLine((e.number + 0x100000000).toString(16) + " " + e.name + ": " + e.description);
            }
        }
		} catch(e) {
			WScript.StdErr.WriteLine("Can't create the \"MSScriptControl.ScriptControl\" object: This script must be run in 32-bit CScript to use ScriptControl (usually located at \"%windir%\\SysWOW64\\cscript.exe\")");
			WScript.Quit();
		}
    }
} else { //eval
    while (true) {
        try {
            WScript.StdOut.Write(">>> ");
            WScript.StdOut.WriteLine(eval(WScript.StdIn.ReadLine()));
        } catch (e) {
            WScript.StdErr.WriteLine((e.number + 0x100000000).toString(16) + " " + e.name + ": " + e.description);
        }
    }
}
