# Igor Command Line  
This information is for Windows programmers only. You can call Igor Pro from a Windows batch file or even from Igor itself using  
ExecuteScriptText using this operation-like syntax:  
Igor.exe [/I /Q /X /Y /N /Automation]  [pathToFileOrCommands ] [pathToFile ] ...  
Igor.exe [/I /N /Automation]  [pathToFileOrCommands ] [pathToFile ] ...  
Igor.exe [/I /Q /X /Automation]  "commands "  
Igor.exe /SN=num /KEY="key " /NAME="name " [/ORG="org " /QUIT]
## Parameters
The usual parameter to Igor.exe is a file which Igor opens. It is recommended that both the path to Igor.exe and the path to the file parameter be enclosed in quotes:
"C:\Program Files\WaveMetrics\Igor Pro Folder\Igor.exe" "C:\Igor Files\exp.pxp"
Multiple files can be opened by appending the path to the file(s) with an intervening space:
"C:\Program Files\WaveMetrics\Igor Pro Folder\Igor.exe" "C:\Dir\exp.pxp" "C:\Dir\exp.dat"
With the /X flag, only one parameter is allowed and is interpreted as Igor commands:
"C:\Program Files\WaveMetrics\Igor Pro Folder\Igor.exe" /X "Make/O data=x;Display data"
The /SN, /KEY, and/NAME flags must all be used to successfully register Igor Pro. The optional /ORG parameter defaults to "".
## Flags
Note:	The / symbol can be replaced with a - symbol (ActiveX Automation uses a -Automation parameter when calling Igor Pro).
- /Automation	Used automatically (along with /I) by the operating system when launching Igor Pro as an Automation Server. The command parameter that the OS sends is defined in the Registry, put there by the Igor installer or by the user, by merging an IgorProCOM.reg file. This flag isn't intended for use in batch files or ExecuteScriptText. For more details on ActiveX Automation Automation Server, see Automation Server Overview.
	It can, however be used from the command line ora batch fileto communicate with other programs in combination with /X. The /Automation flag keeps Igor windows hidden, which may be useful when calling Igor Pro from a web server CGI program
- /I 	Launches a new "instance" of Igor that will open the file or execute the commands. Pressing Ctrl while launching Igor is the same as using the /I flag.
	Without /I, files are opened and commands are executed by any Igor.exe that is currently running. If a parameter is an experiment file, the currently open experiment is closed before opening the new one (see the /Y and /N flags).
- /KEY="key "	Specifies the license activation key. Use a value of the form: 
	/KEY="ABCD-EFGH-IJKL-MNOP-QR"
	Do not omit the quotes, or it will fail.
- /N	Forces the current experiment to be closed without saving if any of the file parameters are an experiment file.
	To save a currently open experiment, use:
	Igor.exe /X "SaveExperiment"
- /NAME="name "	Defines the name of the licensed user(s). Cannot be "".
- /ORG="org "	Specifies the optional name of the licensed organization. Default is "". Because Windows interprets the & symbol to mean "underline the next character when displayed in a dialog window," use && to display one & character in the About Igor dialog.
- /Q	Doesn't show
	Command line: /X "Make/O data=s;Display data"
	in Igor's history window when using /X or /SN, etc.
- /QUIT	Quits Igor Pro after entering license information when used with /SN, /KEY, and /NAME. Otherwise /QUIT is ignored.
	To quit Igor Pro, use:
	Igor.exe /X "Quit/N"
- /SN=num	Specifies the license serial number.
- /X	Executes the commands in the first (and only allowed) parameter. Use semicolons to separate commands.
## Details
As of Igor 6.2, if a copy of Igor.exe is already running and if Igor.exe is launched again without /X, /SN or any path to a file, a new instance of Igor.exe is started.
Previous to Igor 6.2, launching Igor under those conditions would only activate the already-running instance of Igor.exe.
This means that double-clicking the Igor.exe icon will start another instance of Igor.exe, but double-clicking an experiment file will still open that file in the frontmost instance of Igor (or start up Igor if it isn't running), as it always has.
Example
This function launches another instance of Igor to open an experiment file:
```
	Function LaunchAnotherIgor(expPath)
		String expPath			// Full Windows path to experiment file
									// e.g.,  "C:\\Igor Files\\Experiment.pxp"

		String quote = "\""		// String containing a double-quote

		// Get path to Igor Pro folder in Macintosh file format.
		PathInfo Igor				// Stores output in S_path.

		// Get path to Igor in Windows format.
		String igorPath= ParseFilePath(5, S_path, "\\", 0, 0) + "Igor.exe"

		String scriptText = quote + igorPath + quote + " /I " + quote + expPath + quote
		ExecuteScriptText scriptText
	End
```
These batch file commands register Igor Pro with the given (fictional) serial number and license activation key:
Igor.exe /SN=1234567/KEY="ABCD-EFGH-IJKL-MNOP-QR"/NAME="Me" /ORG="You && Me, Inc." /QUIT
