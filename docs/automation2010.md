#	Automation Server Overview
ActiveX Automation (henceforth called "Automation") is Microsoft's system for allowing one program, the client,  to call another program, the server, to obtain services. Automation is built on Microsoft's Component Object Model (COM) system, a lower-level protocol for inter-application communication. Although parts of COM have been supported on other platforms, it is fully supported only on Windows.

The Windows version of Igor Pro can play the role of an Automation server. This makes Igor's functionality available to clients written in Visual Basic, Visual C++, Java, Windows Script Host as well as in various other programming and scripting environments.

We have tested Igor Pro Server with Visual Basic 6 through Visual Basic 2010, Visual Basic for Applications 2000 (i.e., from Excel 2000 and Word 2000) and with Visual C++ 6 through Visual C++ 2010. Although Automation is a "standard", it is not uniformly implemented by all programming environments, especially when it comes to passing arrays of data. As a result, some of the methods that Igor Pro exposes may not work correctly in some programming environments.

Igor Pro as an Automation server can be used by a custom program as a number-crunching engine and as a graphics engine. It can be called from Word and Excel macros to automate the transfer of data and/or graphics.

As of this writing, Igor Pro does not have the ability to function as an Automation client but there are workarounds. See ActiveX Automation.

#	Setup
When you install Igor Pro on Windows, the Igor installer does the following things relative to Automation:

Writes information to the Windows registry that describes Igor's Automation Server capabilities. Windows uses this information to launch Igor Pro as an Automation server when requested by a client program to do so, and to manage interactions between the client and Igor.

Installs the IgorPro.tlb file in the Igor Pro Folder. This is a "type library" file which contains a description of the methods that a client program can call in a binary format. Type libraries are used by Visual Basic and other scripting programs to determine what methods Igor Pro Server supports. C++ programs use the file IgorProTypeLib.h which contains the same information in the form of a C++ header file.

Installs the "Windows Automation" folder inside "Igor Pro Folder:Miscellaneous". The "Windows Automation" folder contains a sample Visual Basic program, a sample Visual C++ program, a sample Windows Script Host script file and other sample files.

#	Background Information
If you plan to write a Visual Basic program or to use another scripting system such as JavaScript, you may be able to do so without knowing much about Automation and COM, which is the software technology on which Automation is built. On the other hand, if you plan to write a C++ program, you will need to have at least a basic understanding of COM and Automation.

There are plenty of articles about COM and Automation on the World Wide Web. One thing that is hard to find is a good overview that explains the relationship of various Microsoft technologies such as COM, OLE, ActiveX, ActiveX Controls, Automation, Visual Basic, VBScript, VBA, .NET and so on.

Briefly, COM is the low-level protocol and system support that allows one program to call another's functions, passing parameters and receiving results. OLE originally meant "Object Linking and Embedding" but evolved to mean any technology built on top of COM. Microsoft then ditched the term OLE in favor of ActiveX, which still means any technology built on top of COM. Automation is the technology built on top of COM that allows one program to access another's specific functionality. Visual Basic allows you to write simple programs that utilize Automation and hides most of the messy details. VBScript is a simplified form of Visual Basic that you can run from a web page or from Windows Script Host. VBA means "Visual Basic for Applications" and is the form of Visual Basic that is built into most Microsoft and some other application programs. And .NET is a vague term that covers a wide range of Microsoft technologies, so wide that no one can say what ".NET" means.

In terms of books, the standard reference on Automation is the Automation Programmer's Reference from Microsoft Press. A good introduction to COM is Inside COM by Dale Rogerson, also published by Microsoft Press.

#	Running ActiveX Automation in Newer Windows Versions
ActiveX Automation runs smoothly in old Windows operating systems such as Windows XP, but the increased security in recent versions of Windows interferes with ActiveX Automation. Specifically, you must either turn User Account Control off or run both the client and the server (Igor) using administrative privileges. Also, even if you run the client as administrator, it can not launch Igor automatically so you have to do that manually.
To turn User Account Control off, use the operating system's User Account Control Settings control panel and then reboot. This is not recommended as it affects all programs.
A better approach is to run the client and server (Igor) as administrator. Here is how this is done in Windows 7 using the IgorClient.exe example program as an example. For these instructions we will use the IgorClient.exe located at Automation Server Examples\CPlusPlus VC 2010\IgorClient.exe.
1. Locate IgorClient.exe in the Windows desktop.
2. Right-click the IgorClient.exe icon and choose Properties to display the Properties window.
3. Click the Compatibility tab.
4. Check the Run as Administrator checkbox.
5. Close the Properties window.
6. Locate the Igor.exe icon in the Windows desktop.
7. Right-click the Igor.exe icon and choose Properties to display the Properties window.
8. Click the Compatibility tab.
9. Check the Run as Administrator checkbox.
10. Close the Properties window.
11. Launch Igor.exe.
12. Launch IgorClient.exe.
13. Choose Connect to Igor Pro Server.
On this last step, if the IgorClient program or Igor itself is running with less than administrator privileges, you will get an error. Also, if Igor is not running at all, because you did not do step 11, you will get an error. The order of launching the programs does not matter.
Running Visual Basic on Newer Windows Versions
There are also issues when running under Visual Basic. You must launch Igor manually, you must run Visual Studio as administrator and you must configure the project so that the client program runs as administrator. Here are the steps for doing this in Visual Basic 2008 and 2010:
1. Run Igor as administrator, as described above.
2. Run Visual Studio as administrator.
3. Open IgorProServerDemo.sln in Visual Basic.
4. In the Solution Explorer window, right-click IgorProServerDemo and choose Properties.
5. Click the Application tab.
6. Click View UAC Settings (VB 2008) or View Windows Settings (VB 2010).
7. Change this:
	<requestedExecutionLevel  level="asInvoker" uiAccess="false" />
to this:
	<requestedExecutionLevel  level="requireAdministrator" uiAccess="false" />
This changes the file app.manifest so that your program runs as administrator.
Running Visual C++ on Newer Windows Versions
There are also issues when running under Visual C++. You must launch Igor manually, you must run Visual Studio as administrator and you must configure the project so that the client program runs as administrator. Here are the steps for doing this in Visual C++ 2008 and 2010:
1. Run Igor as administrator, as described above.
2. Run Visual Studio as administrator.
3. Open IgorClient.sln in Visual Studio.
4. In the Solution Explorer window, right-click IgorClient and choose Properties.
5. From the Configuration popup menu, choose All Configurations.
6. In the lefthand pane, click Linker and then Manifest File.
7. Under UAC Execution Level, choose requireAdministrator.
8. Click OK.
#	Getting Started
A good way to get started is to play with the samples provided by WaveMetrics. Then read the rest of this help file and any other background material. Once you have done that, you're ready to start your own program.

The samples are provided in a zip archive named "Automation Server Examples.zip". Unzip this to create the "Automation Server Examples" folder.

Open the "Automation Server Examples" folder and notice that projects are provided for a variety of Basic, C++, and other environments. Identify the project appropriate for your environment.

This section describes the main samples provided with Igor.

Visual Basic
In Visual Basic 6, double-click the IgorProServerDemo.vbp file to open the VB6 sample project.

For any other version of Visual Basic, double-click the IgorProServerDemo.sln file to open the sample project.

If you are running on Windows VISTA or later, follow the instructions under Running Visual Basic on Newer Windows Versions.

Press F5 to compile and run the program.

On Windows XP, if Igor Pro is already running, it connects with that instance of Igor Pro. If not it launches a new, invisible instance, which may take a while so be patient. (You can make Igor launch faster by removing any unneeded Igor help files and Igor extensions.)

On newer Windows versions, you must manually launch Igor as administrator as described under Running ActiveX Automation in Newer Windows Versions.

The Igor Pro Server Demo program then displays a panel with buttons. If Igor Pro was not already running then it will be invisible. Click the Make Igor Visible button.

Click the Send Message To History button and notice that a message is sent from the Visual Basic program to Igor's history area.

Click the "Make jack; Display jack" button. This creates a wave and graphs it.

Click the "Add One To Jack". This adds one to each point in the jack wave and thereby causes the graph to update.

Click the Test Matrix Access button. This creates a small matrix wave and puts some data into it. Click the Test Matrix Access button again. This time it adds one to each data point in the matrix but does not create a new table.

Click the Test Complex Access button. This creates a small complex wave and puts some data into it. Click the Test Complex Access button again. This time it adds one to each data point in the complex wave but does not create a new table.

Click the Adhoc Test button. This runs a variety of test routines that exercise many of Igor Pro Server's methods. In the course of doing so it creates a number of waves and variables which you can see in Igor's Data Browser.

Click the close box on the Visual Basic panel to quit the program. Igor Pro remains running. In a real client-server setup Igor Pro would have been invisible and quitting the Visual Basic program would have caused Igor to quit also.

Now examine some of the source code in the Visual Basic program. You might find it useful to single-step through some of it. For more information see Visual Basic Example below.

Visual Basic For Applications
If you have access to Microsoft Word 2000 or later you can try the simple Word demonstration provided with Igor. Double-click the file IgorExamples.doc which will open the Word file. Follow the instructions in that file.

Visual C++
In Visual C++ 6, double-click the IgorProServerDemo.dsw file to open the VC6 sample project.

For any other version of Visual C++, double-click the IgorProServerDemo.sln file to open the sample project.

If you are running on Windows VISTA or later, follow the instructions under Running Visual C++ on Newer Windows Versions.

Press F5 to compile and run it. The IgorClient program then window with a menu.

From the menu choose Connect To Igor.

On Windows XP, if Igor Pro is already running, it connects with that instance of Igor Pro. If not it launches a new, invisible instance, which may take a while so be patient. (You can make Igor launch faster by removing any unneeded Igor help files and Igor extensions.)

On newer Windows versions, you must manually launch Igor as administrator as described under Running ActiveX Automation in Newer Windows Versions.

If Igor Pro was not already running then it will be invisible. Choose Make Igor Visible from the menu.

Choose Send Message To History from the menu and notice that a message is sent from the IgorClient program to Igor's history area.

Click the "Make jack; Display jack" from the menu. This creates a wave and graphs it.

Click the "Add One To Jack". This adds one to each point in the jack wave and thereby causes the graph to update.

Click the Test Matrix Access from the menu. This creates a small matrix wave and puts some data into it. Click the Test Matrix Access from the menu again. This time it adds one to each data point in the matrix but does not create a new table.

Click the Test Complex Access from the menu. This creates a small complex wave and puts some data into it. Click the Test Complex Access from the menu again. This time it adds one to each data point in the complex wave but does not create a new table.

Click the Adhoc Test from the menu. This runs a variety of test routines that exercise many of Igor Pro Server's methods. In the course of doing so it creates a number of waves and variables which you can see in Igor's Data Browser.

Click the close box on the IgorClient window to quit the program. Igor Pro remains running. In a real client-server setup Igor Pro would have been invisible and quitting the IgorClient program would have caused Igor to quit also.

Now examine some of the source code in the IgorClient program. You might find it useful to single-step through some of it. For more information see Visual C++ Example below.

#	Techniques
This section describes the main techiques that are available to communicate with Igor Pro acting as an Automation server.

Connecting To Igor
The client program has to first establish a connection to Igor. It can either establish a connection to an already-running instance of Igor or it can create a new instance. Typically during development you will want to test and revise your client program without having to launch Igor each time you revise, so you will want to connect to an already-running instance. Once your client program is finished, you may want it to launch a new instance of Igor so that it will work correctly whether Igor is already running or not. However, as noted under Running ActiveX Automation in Newer Windows Versions, with recent version of Windows, it is required that you launch Igor manually.

This Visual Basic subroutine shows how to connect to an already-running instance of Igor or create a new instance if no instance is already running.

```objc
Function ConnectToIgorProServer(allowConnectingToAlreadyRunningInstance As Boolean) As IgorPro.Application
    Dim app As IgorPro.Application

    If allowConnectingToAlreadyRunningInstance = True Then
        ' Try to connect with running Igor.
        On Error Resume Next
        Set app = GetObject(, "IgorPro.Application")
        If Err.Number = 429 Then
            ' The following works in XP but not in recent versions of Windows.
            ' Could not connect so try to create new Igor application.
            Set app = CreateObject("IgorPro.Application")
        End If
    Else
       ' The following works in XP but not in recent versions of Windows.
         Set app = CreateObject("IgorPro.Application")
       ' Set IgorApp = New IgorPro.Application             ' This also works.
    End If

    Set ConnectToIgorProServer = app
End Function
```

ConnectToIgorProServer would be called like this:

```
Dim IgorApp As IgorPro.Application  ' Early binding
IgorApp = ConnectToIgorProServer(1)
```

Here is the same routine in C++:

```c
	static HRESULT
	ConnectToIgorProServer(IApplication** pApplicationPtr)
	{
		LPUNKNOWN punk;
		IApplication FAR* pApplication;
		HRESULT hr = NOERROR;

		*pApplicationPtr = NULL;

		//	Try to connect to an already-running instance of Igor.
		hr = GetActiveObject(CLSID_Application, NULL, &punk);
		if (FAILED(hr)) {
			// The following works in XP but not in recent versions of Windows.
			// Create new Igor Pro instance.
			hr = CoCreateInstance(CLSID_Application, NULL, CLSCTX_SERVER, IID_IUnknown, (void**)&punk);
			if (FAILED(hr))
				return hr;
		}

		// Get Igor's IApplication interface.
		hr = punk->QueryInterface(IID_IApplication, (void**)&pApplication);
		punk->Release();
		if (FAILED(hr))
			return hr;

		*pApplicationPtr = pApplication;
		return NOERROR;
	}
```

For some unknown reason, when the client program causes Igor to be launched, the client program becomes deactivated (loses focus) even though Igor is invisible at that point. We have been unable to find a way to prevent this from happening.

## Interacting With Igor
The main methods for interacting with Igor are:
	Using the Execute or Execute2 methods (IApplication interface) to send commands to Igor.
	Using the IDataFolder, IWave and IVariable interfaces to get or set Igor data.
	Using the fprintf operation to return information other than data folder, wave or variable values.

Using Execute or Execute2 To Send Commands To Igor
Execute and Execute2 are part of the IApplication interface. They allow you to submit commands to Igor for execution, much like Igor's built-in Execute operation.

Execute has just one parameter - the command to be executed.

Execute2 provides more control and information. It returns both a numeric error code and a string error message. It also returns any output sent to Igor's history area during the execution of the command. And it returns results generated by the fprintf operation (details below).

Using Methods To Get or Set Igor Data
Igor Pro provides methods for creating and deleting data folders, waves, numeric variables and string variables. It also provides methods for getting and setting wave, numeric variable and string variable data.

The IDataFolder interface provides methods for dealing with a specific data folder. The IDataFolders interface provides methods for creating, deleting and interating over data folders.

The IWave interface provides methods for dealing with a specific wave. The IWaves interface provides methods for creating, deleting and interating over waves.

The IVariable interface provides methods for dealing with a specific variable. The IVariables interface provides methods for creating, deleting and interating over variables.

These interfaces and their methods are described in detail below. See Interface Hierarchy.

Using fprintf To Get Information From Igor
You can use the fprintf operation to obtain from Igor information that is not available through a method call. Using the special refNum value of zero tells Igor to return the information printed by fprintf through the results parameter of the Execute2 method. Here is a Visual Basic example. It obtains the name of the current Igor experiment by making Igor execute the IgorInfo function.

    Dim cmd As String
    Dim errorCode As Long
    Dim errorMsg As String
    Dim history As String
    Dim results As String

    cmd = "fprintf 0, ""%s"", IgorInfo(1)"
    IgorApp.Execute2 0, 0, cmd, errorCode, errorMsg, history, results
    Debug.WriteLine results  ' Print name of current experiment.  ' Use Debug.Print in VB 6

Exporting Graphics From Igor
There are no methods specifically for exporting graphics from Igor. You must use the Execute or Execute2 methods to make Igor export graphics to the clipboard or to a file.

For example, here is code from the Word 2000 demo:

Private Sub InsertGraph_Click() ' Inserts top Igor graph into Word as EMF picture
    Dim cmd As String
    Dim result As Integer

    Dim doc As Word.Document
    Set doc = Documents(1)

    ' Export graph as EMF to a file.
    Dim filePath As String
    Dim sep As String
    Dim quote As String
    sep = Application.PathSeparator
    quote = Chr(34)

    Dim useClipboard As Boolean
    useClipboard = 1

    If useClipboard Then
        ' This causes Igor to write the picture to the clipboard.
        ' This will not work if the client and server are on different machines.
        cmd = "SavePICT/O/E=-2 as " & quote & "Clipboard" & quote
    Else
        ' This makes filePath point to a file in the folder containing the Word document.
        ' This has a problem in that Word keeps the EMF file open until Word quits
        ' which means we can't delete or overwrite the file.
        filePath = doc.Path & sep & "IgorGraph.emf"
        cmd = "SavePICT/O/E=-2 as " & quote & filePath & quote
    End If

    result = ExecuteIgorProCmd(cmd, 1, 1)
    If result <> 0 Then
        Exit Sub
    End If

    ' Replace current selection in the Word document with the graph
    With Selection
        .TypeText Text:="Igor graph:"
        .TypeParagraph
        If useClipboard Then
            .PasteSpecial dataType:=wdPasteEnhancedMetafile, placement:=wdInLine
            If .Type = wdSelectionShape Then
                 .ShapeRange.ConvertToInlineShape
            End If
            .MoveRight Unit:=wdCharacter, Count:=1
        Else
            .InlineShapes.AddPicture FileName:=filePath, LinkToFile:=False
        End If
        .TypeParagraph
    End With
End Sub

Loading and Saving Igor Experiments
The IApplication interface provides methods for loading (LoadExperiment) and saving (SaveExperiment) Igor experiment files.

Loading and Calling Igor Procedures
The IApplication interface provides methods for #including (InsertInclude) and deleting (DeleteInclude) Igor procedure files. After #including a file and compiling it (CompileProcedures), you can then use Execute or Execute2 to invoke your procedures.

Killing Data Folders
You can kill a data folder using the Remove method of the IDataFolders interface or by executing a KillDataFolder command via the Execute or Execute2 methods of the IApplication interface. But it is not that simple.

You can't kill a data folder if it or its contents are in use. When your program creates a reference to a data folder Igor considers it "in use" so you can't kill a data folder while you hold a reference to it. Therefore you must release your reference before killing the data folder.

In Microsoft's .NET languages, such as C# and VB.NET, releasing a reference to a data folder requires releasing the associated COM object that was created when you created the reference. This requires calling a special routine - ReleaseCOMObject.

Here is an example in VB.NET:

Imports System.Runtime.InteropServices.Marshal

Sub Test()
	Dim dfcol As IgorPro.DataFolders	' A collection of data folders
	dfcol = IgorApp.DataFolders(":")

	Dim df As IgorPro.DataFolder	' Create reference to a particular data folder
	df = dfcol.Item("Junk")

	ReleaseComObject(df)	' Release COM reference to data folder. df is no longer valid.

	dfcol.Remove("Junk")	' Kill data folder

	dfcol = Nothing
	df = Nothing
End Sub

Note that the statement
	df = Nothing
does not release the COM object. You must call ReleaseComObject.

You can kill the data folder by executing
	IgorApp.Execute("KillDataFolder Junk")
instead of
	dfcol.Remove("Junk")
but you still must release the reference using ReleaseComObject in order for KillDataFolder to succeed.

In order to kill a data folder, all  references to it must be released, so if you have created multiple references to the data folder in your program you need multiple calls to ReleaseComObject.

In VB the Imports statement must be appear before any code.

In C# you would use:
	using System.Runtime.InteropServices;
instead of
	Imports System.Runtime.InteropServices.Marshal

You also can't kill a data folder if you are holding a reference to a wave, wave collection, variable or variable collection in the data folder. This example illustrates the issue and the solution:

Sub Test()
	Dim dfcol As IgorPro.DataFolders	' A collection of data folders
	dfcol = IgorApp.DataFolders(":")

	Dim df As IgorPro.DataFolder	' Create reference to a particular data folder
	df = dfcol.Item("Junk")

	Dim vcol as IgorPro.Variables	' Create a reference to a variable collection
	vcol = df.Variables

	Dim v as IgorPro.Variable	' Create a reference to a variable
	v = vcol.Add("temp", IgorPro.IgorProDataType.ipDataTypeDouble, 1)

	ReleaseComObject(v)	' Release COM reference to variable. v is no longer valid.

	vcol.Remove("temp")	' Kill the variable

	ReleaseComObject(vcol)	' Release COM reference to variable collection. vcol is no longer valid.

	ReleaseComObject(df)	' Release COM reference to data folder. df is no longer valid.

	dfcol.Remove("Junk")	' Kill data folder

	v = Nothing
	vcol = Nothing
	dfcol = Nothing
	df = Nothing
End Sub

Finally here is something to keep in mind. If your program creates a reference to an Igor data folder, wave, wave collection, variable or variable collection and crashes before you call ReleaseComObject, the COM object is never released and Igor is not notified that your program is no longer referencing the data folder. You will not be able to kill it until you do New Experiment or Open Experiment in Igor.

Killing Waves
The consideration for killing waves are the same as for killing data folders. See Killing Data Folders for details.

Killing Variables
Unlike waves and data folders, Igor allows you to kill a variable even if you hold a reference to it. However it is best to release the reference anyway as shown but the example above.

#	Visual Basic Example
Here are some instructions and tips for creating a Visual Basic program that uses Igor Pro as an Automation server. These instructions assume that you have some familiarity with Visual Basic.

See Running ActiveX Automation in Newer Windows Versions for issues with recent version of Windows.

You first need to create a project.

In Visual Basic For Application 2000 (e.g., from Word 2000 or Excel 2000):
Choose Tools->Macro->Visual Basic Editor.

In Visual Basic 6:
Choose File->New Project and make a "Standard EXE" project.

In Visual Basic 2003-2010:
Choose File->New->Project and create a new "Windows Application" project.

Next you must configure Visual Basic so that it knows about Igor's classes and methods by creating a reference to its type library.

In Visual Basic For Application 2000 (e.g., from Word 2000 or Excel 2000):
Choose Tools->References, find "Igor Pro 6.0 Type Library" and click its checkbox.

In Visual Basic 6:
Choose Project->References, find "Igor Pro 6.0 Type Library" and click its checkbox.

In Visual Basic 2003-2010:
Open the Solution Explorer. Right-click the References icon. Choose Add Reference. Click the COM tab in the resulting dialog. Select "Igor Pro 6.0 Type Library". Click the Select button. Click the OK button.

If you are using a later version of Igor Pro, the type library name might be slightly different.

The "Igor Pro 6.0 Type Library" entry appears in the References dialog as a result of the Igor installer creating a registry entry that points to the IgorProTypeLib.tlb file which resides in the Igor Pro folder.

Now you can create a simple program by adding a module to the project and entering the following code:

Option Explicit

Dim IgorApp As IgorPro.Application  ' Early binding

Function ConnectToIgorProServer(allowConnectingToAlreadyRunningInstance As Boolean) As IgorPro.Application
    Dim app As IgorPro.Application

    If allowConnectingToAlreadyRunningInstance = True Then
        ' Try to connect with running Igor.
        On Error Resume Next
        Set app = GetObject(, "IgorPro.Application")
        If Err.Number = 429 Then
            ' The following works in XP but not in recent versions of Windows.
            ' Could not connect so try to create new Igor application.
            Set app = CreateObject("IgorPro.Application")
        End If
    Else
       ' The following works in XP but not in recent versions of Windows.
         Set app = CreateObject("IgorPro.Application")
       ' Set IgorApp = New IgorPro.Application             ' This also works.
    End If

    Set ConnectToIgorProServer = app
End Function

Sub Test()
    Set IgorApp = ConnectToIgorProServer(1)

    Dim cmd As String
    Dim errorCode As Long
    Dim errorMsg As String
    Dim history As String
    Dim results As String

    cmd = "Make/O jack=x; Display jack"
    IgorApp.Execute cmd

    cmd = "WaveStats/Q jack; fprintf 0, ""%g"", V_avg"
    IgorApp.Execute2 0, 0, cmd, errorCode, errorMsg, history, results
    MsgBox(results)
End Sub

In Visual Basic 6, add this to the Form1 module code window:

Private Sub Form_Load()
    Test
End Sub

In Visual Basic 2003-2010, add this to the New subroutine in the Form1 module code window:
    Test

Now launch Igor.

Now press F5 to run the program. It will execute, create a wave and a graph in Igor, run the WaveStats operation, and finally display the average value of the wave in a message box.

A more comprehensive example Visual Basic program is included in the Windows Automation folder.

The ConnectToIgorProServer subroutine connects to an already-running instance of Igor if one exists. If one does not exist, it launches a new instance of Igor. In this case, the new instance will be invisible. You can make it visible by executing:
	IgorApp.Visible = True

If the instance of Igor is visible then it will remain running when the Visual Basic program quits. If the instance is invisible, it will automatically quit when the Visual Basic program quits (actually, when your program releases the IgorApp object). You can use the Windows Task Manager to see if there is an invisible instance of Igor Pro running.

If you were using Igor Pro as a server in a finished program, you might want it to be invisible and to quit automatically when you program quits. However, during development, you want to be able to see what is going on in Igor and you don't want to wait for it to launch each time you revise and re-run your program, so you will want to manually launch Igor and then connect to the already-running instance of it, as we did in the preceding example.

#	Visual C++ Example
Here are some instructions and tips for creating a Visual C++ program that uses Igor Pro as an Automation server. These instructions assume that you have some familiarity with Visual C++ and Automation.

See Running ActiveX Automation in Newer Windows Versions for issues with recent version of Windows.

You will probably want to start from the C++ example project provided by WaveMetrics. It is named IgorClient. Here is a description of the files in the IgorClient project:

IgorClient.dsw	Visual C++ 6 workspace file.
IgorClient.dsp	Visual C++ 6 project file.
IgorClient.sln	Visual C++ 2003-2010 solution file.
IgorClient.vcproj	Visual C++ 2003-2008 project file.
IgorClient.vcxproj	Visual C++ 2010 project file.
IgorProTypeLib.h	Contains an automatically-generated description of the interfaces that Igor exposes, the methods that each interface provides and the parameters and return types of each method. You need to include this in any file from which you can to call an Igor Pro method.
IgorClient.h	Contains prototypes and other declarations used by all of the files in the project.
IgorClient.cpp	Contains the following functions:
	WinMain: The main function called when the program starts.
	IgorClientWndProc: The window procedure for the window created by the program. This handles menu selection events and calls the appropriate test routine.
	DisplayIgorErrorInfo: Displays detailed error information when a call to Igor causes an error.
TestDataFolders.cpp	Contains methods that illustrate all of Igor's IDataFolder interface methods.
TestWaves.cpp	Contains methods that illustrate all of Igor's IWave interface methods.
TestVariables.cpp	Contains methods that illustrate all of Igor's IWave interface methods.
Utilities.cpp	Contains utility routines including ConnectToIgorProServer which establishes a connection to an already-running instance of Igor Pro or launches a new instance.
IgorClient.rc	Contains resource descriptions for the menu.

The IgorClient program is configured to be compiled as a Unicode program or as a regular program. The file Utilities.cpp contains comments about Unicode.

Open IgorClient.dsw in Visual C++ 6 or IgorClient.sln in Visual C++ 2003-2010. Compile and run the program.

The ConnectToIgorProServer function connects to an already-running instance of Igor if one exists. If one does not exist, it launches a new instance of Igor. In this case, the new instance will be invisible. You can make it visible by executing:
	pApplication->put_Visible(TRUE);
where pApplication is a pointer to Igor's application interface returned from ConnectToIgorProServer.

If the instance of Igor is visible then it will remain running when the Visual C++ program quits. If the instance is invisible, it will automatically quit when the Visual C++ program quits (actually, when your program releases the IgorApp object). You can use the Windows Task Manager to see if there is an invisible instance of Igor Pro running.

If you were using Igor Pro as a server in a finished program, you might want it to be invisible and to quit automatically when you program quits. However, during development, you want to be able to see what is going on in Igor and you don't want to wait for it to launch each time you revise and re-run your program, so you will want to manually launch Igor and then connect to the already-running instance of it.

#	Strings, Paths and Names
Automation uses Unicode to encode strings. Therefore the client program must use Unicode when calling Igor through Automation.

Igor Pro uses the "multibyte character system" (MBCS). Therefore Igor has to translate all strings passed between it and the client program.

Igor Pro methods used to pass string data include a "codePage" parameter which Igor uses to translate between Unicode and MBCS. For most purposes, you should pass zero for the codePage parameter. Zero means "use the default ANSI code page". In C++, you can also use the symbol CP_ACP.

If you deal with string data that is not in the "default ANSI code page" (the meaning of which is unclear - see Microsoft documentation), you must pass a different code. In this case, you will have to learn all about code pages and the MBCS, a murky area.

If you limit yourself to "low ASCII" characters (characters with codes below 0x7F), then you can always use zero for the codePage parameter. Low ASCII characters include the numerals 0 through 9, the standard letters without diacritical marks, and the common punctuation marks.

Methods used to pass Igor object names (e.g., wave names, data folder names) and Igor data paths do not have a codePage parameter. Internally Igor uses the value 0 (CP_ACP for the default ANSI code page). This means that you will run into trouble if you create waves or data folders whose names do not fall in the low ASCII range and are not in the default ANSI code page.

#	Interface Hierarchy
To give you an overview, this section lists the interfaces, methods and properties that Igor Pro exposes through Automation. The interfaces, methods and properties are described in greater detail below.

Methods with names like get_<property> and put_<property> are "properties". This means that Visual Basic hides the actual names of the functions and instead allows you to write:
	object.property = <value>
	<variable> = object.property

For example:
	IgorApp.visible = 1
	MyVariable = IgorApp.visible

However, from C++, you would use the actual names, get_Visible and put_Visible.

Some properties are read-only so there is a get_<property> function but no put_<property> function.

IApplication Methods
get_Application	Returns the application interface.
get_FullName	Returns the full path to the Igor.exe file.
get_Name	Returns "Igor Pro".
get_Parent	Returns the application interface.
get_Visible	Returns truth that Igor Pro is visible.
put_Visible	Sets Igor Pro to visible or invisible.
get_Status1	Returns assorted bits of Igor status information.
Quit	Causes Igor Pro to quit.
SendToHistory	Sends a text message to Igor's history area.
Execute	Executes a command on Igor's command line.
Execute2	Executes a command on Igor's command line. Provides more error and output options than Execute.
DataFolderExists	Checks if a data folder exists.
DataFolder	Gets data folder interface from Igor.
get_DataFolders	Returns a collection of all the data folders in a specified parent data folder.
NewExperiment	Creates a new, empty experiment.
LoadExperiment	Loads an experiment file into memory.
SaveExperiment	Saves the current experiment to a file.
OpenFile	Opens a procedure, notebook or help file.
InsertInclude	Inserts a #include statement in the Procedure window.
DeleteInclude	Deletes a #include statement from the Procedure window.
CompileProcedures	Causes Igor procedure files to be compiled.
See IApplication Methods for details.

IDataFolder Methods
get_Name	Gets data folder name.
put_Name	Sets data folder name.
get_InUse	Returns the truth that the data folder is in use and can't be removed.
get_IsRoot	Returns the truth that the data folder is the root data folder.
get_Path	Gets full or relative path to data folder.
get_Next	Gets next data folder.
get_ParentDataFolder	Returns the interface for the data folder's parent data folder.
get_SubDataFolders	Returns a collection of the data folders within the data folder.
WaveExists	Checks if a wave exists.
Wave	Gets wave interface from Igor.
get_Waves	Returns a collection of all the waves in the data folder.
VariableExists	Checks if a global variable exists.
Variable	Gets global variable interface from Igor.
get_Variables	Returns a collection of all the global variables in the data folder.
See IDataFolder Methods for details.

IDataFolders Methods
Add	Adds a new data folder to the collection.
get_Count	Returns the number of data folders in the collection.
get_Item	Given an integer index or a data folder name, returns one of the data folders in the collection.
get__NewEnum	Used internally by Visual Basic to implement For Each loop.
Remove	Given an integer index or a data folder name, kills the data folder.
DataFolderExists	Checks if a data folder exists in the collection.
See IDataFolders Methods for details.

IWave Methods
get_Name	Gets wave name.
put_Name	Sets wave name.
get_InUse	Returns the truth that the wave is in use and can't be removed.
get_Path	Gets full or relative path to wave.
get_Next	Gets next wave in linked list or NULL.
get_ParentDataFolder	Returns the interface for the wave's parent data folder.
GetDimensions	Gets the wave's data type and size of each dimension.
SetDimensions	Sets the wave's data type and size of each dimension.
GetScaling	Gets the wave's dimension scaling.
SetScaling	Sets the wave's dimension scaling.
get_Units	Gets wave units.
put_Units	Sets wave units.
get_DimensionLabel	Gets the dimension label for a particular element.
put_DimensionLabel	Sets the dimension label for a particular element
GetMiscInfo	Gets the wave's creation and modification dates.
get_Note	Gets wave name.
put_Note	Sets wave name.
get_Lock	Gets wave lock status.
put_Lock	Sets wave lock status.
GetNumericWaveData	Gets numeric wave data from Igor using any numeric data type.
SetNumericWaveData	Sets numeric wave data in Igor using any numeric data type.
GetNumericWaveDataAsDouble	Gets numeric wave data from Igor using double-precision floating point.
SetNumericWaveDataAsDouble	Sets numeric wave data from Igor using double-precision floating point.
GetNumericWavePointValue	Gets value of a single point of a 1D, real numeric wave.
SetNumericWavePointValue	Sets value of a single point of a 1D, real numeric wave.
MDGetNumericWavePointValue	Gets value of a single point of a real or complex numeric wave of any dimension.
MDSetNumericWavePointValue	Sets value of a single point of a real or complex numeric wave of any dimension.
GetTextWavePointValue	Gets value of a single point of a 1D text wave as Unicode.
SetTextWavePointValue	Sets value of a single point of a 1D text wave as Unicode.
MDGetTextWavePointValue	Gets value of a single point of a text wave of any dimension as Unicode.
MDSetTextWavePointValue	Sets value of a single point of a text wave of any dimension as Unicode.
GetRawTextWaveData	Gets all of the text of a text wave of any dimension as multi-byte characters.
SetRawTextWaveData	Sets all of the text of a text wave of any dimension as multi-byte characters.
See IWave Methods for details.

IWaves Methods
Add	Adds a new wave to collection's data folder.
get_Count	Returns the number of waves in the collection.
get_Item	Given an integer index or a wave name, returns one of the waves in the collection.
get__NewEnum	Used internally by Visual Basic to implement For Each loop.
Remove	Given an integer index or a wave name, kills the wave.
WaveExists	Checks if a wave exists in the collection.
See IWaves Methods for details.

IVariable Methods
get_Name	Gets global variable name.
get_Path	Gets full or relative path to global variable.
get_Next	Gets next global variable in the variable's parent data folder.
get_ParentDataFolder	Returns the interface for the global variable's parent data folder.
get_DataType	Returns the global variable's data type.
GetNumericValue	Gets the value of a global numeric variable.
SetNumericValue	Sets the value of a global numeric variable.
GetStringValue	Gets the value of a global string variable.
SetStringValue	Sets the value of a global string variable.
See IVariable Methods for details.

IVariables Methods
Add	Adds a new global variable to collection's data folder.
get_Count	Returns the number of global variables in the collection.
get_Item	Given an integer index or a global variable name, returns one of the global variables in the collection.
get__NewEnum	Used internally by Visual Basic to implement For Each loop.
Remove	Given an integer index or a global variable name, kills the global variable.
VariableExists	Checks if a global variable exists in the collection.
See IVariables Methods for details.

#	IApplication Methods

get_FullName(BSTR* fullName)
Returns the full path to the Igor.exe file.

get_Name(BSTR* name)
Returns the name of the application ("IgorPro").

get_Visible(VARIANT_BOOL* visible)
Returns the truth that the Igor Pro frame window is visible.

put_Visible(VARIANT_BOOL visible)
Makes Igor Pro visible or invisible.

get_Status1(IgorProStatus1Code what, double* pStatus)
Returns assorted bits of information about Igor.
Parameters
what [input] indicates what kind of information you want.
pStatus [output] receives the status value.
Details
The what parameter can take the following values:
ipStatusIgorVersion 	Returns Igor's version number.
ipStatusRunningProcedure	Returns non-zero if an Igor macro or function is running, zero otherwise.
ipStatusOperationQueueIsEmpty
	Returns non-zero if Igor's operation queue is empty. The Operation Queue contains the commands that you have submitted to Igor through Igor's Execute/P operation (not Igor's Automation Execute method). This status bit allows you to submit commands through Execute/P and then wait until they are finished executing.
ipStatusPauseForUser	Returns non-zero if Igor's PauseForUser operation is executing.
ipStatusExperimentModified	Returns non-zero if the current experiment has been modified since it was open. There is no way to tell if specific parts of the experiment have been modified.
ipStatusExperimentNeverSaved	Returns non-zero if the current experiment is not associated with a file on disk (i.e., it is "Untitled").
ipStatusProceduresCompiled	Returns non-zero if procedures are in a compiled state in Igor, zero if they need to be compiled.
For details on using enumerations such as ipStatusProceduresCompiled in Visual Basic, see Enumerations in Visual Basic.
Visual Basic Example
    ' Check Igor's version
    Dim result as Double
    result = IgorApp.Status1(IgorPro.IgorProStatus1Code.ipStatusIgorVersion)
    Debug.WriteLine "Igor Pro Version Is: " & result  ' Use Debug.Print in VB 6


SendToHistory(int codePage, BSTR* message)
Writes the message to Igor's history area.
Parameters
codePage [input] is a value that Igor uses to convert Igor text data into Unicode. In most cases you should pass zero, which represents the system default ANSI code page (known as CP_ACP in Microsoft's Windows API).
message [input] is a Basic string containing one or more lines of text to be sent to Igor's history area. Each line must be terminated by a CR (carriage-return). Do not use LF (linefeed) or CRLF pair (carriage-return/linefeed).

Execute(BSTR cmds)
Executes the commands in Igor's command line.
If you need specific error output, history output or fprintf output, you must use the more complex Execute2 method.
Parameters
cmds [input] is a Basic string containing one or more command lines to be executed. If more than command line is present, each command line should be terminated by a CR (carriage-return) or CRLF pair (carriage-return/linefeed).
Details
The command or commands in cmds are executed just as if you typed them into Igor's command line. This means that no local variables are accessible from the command, liberal names must be quoted, and all other syntactic rules apply.
Normally commands are logged to Igor's history area. You can suppress history logging by executing the command, "Silent 2". This persists until you turn logging back on by executing "Silent 3".
When you call the Execute method, the user's Igor preferences are off by default. If you want preferences to apply, you must include the command "Preferences 1" in the cmds parameter. The "Preferences 1" command applies only to one invocation of the Execute method so you must include it every time you want preferences to apply.
The Execute method return an Automation error (E_FAIL) if the execution of the commands generates an error in Igor.
For more error and output control, see the Execute2 method.

Execute2(int flags, int codePage, BSTR cmds, int* pIgorErrorCode, BSTR* errorMsg, BSTR* history, BSTR* results)
Executes the commands in Igor's command line.
If you do not need specific error information, history output or fprintf output, you can use the simpler the Execute method.
Parameters
flags [input] is a bitwise parameter that controls aspects of command execution. Currently only one bit is defined:
Bit 0:	If set, the command will not be logged in the history area. If cleared, normally logging will occur unless Silent 2 is in effect. Use the symbol ipExecute2Silent to set this bit.
For details on using the enumerations such as ipExecute2Silent in Visual Basic, see Enumerations in Visual Basic.
codePage [input] is a value that Igor uses to convert Igor text data into Unicode. In most cases you should pass zero, which represents the system default ANSI code page (known as CP_ACP in Microsoft's Windows API).
cmds [input] is a Basic string containing one or more command lines to be executed. If more than command line is present, each command line should be terminated by a CR (carriage-return) or CRLF pair (carriage-return/linefeed).
pIgorErrorCode [output] receives an Igor error code that indicates if the command execution succeeded (*pIgorErrorCode is set to zero) or failed (*pIgorErrorCode is set to a non-zero value).
errorMsg  [output] is a Basic string. On output it contains any error message from Igor or "" if the commands executed without error.
history  [output] is a Basic string. On output it contains any text sent to Igor's history area by the commands.
results  [output] is a Basic string. On output it contains any result text created by fprintf commands that use a file reference of zero within the cmds.
Details
The command or commands in cmds are executed just as if you typed them into Igor's command line. This means that no local variables are accessible from the command, liberal names must be quoted, and all other syntactic rules apply.
Normally commands are logged to Igor's history area. You can suppress history logging by setting bit 0 of the flags parameter which affects just the current invocation of Execute2. You can also suppress logging by executing the command, "Silent 2" which persists until you turn logging back on by executing "Silent 3".
When you call the Execute method, the user's Igor preferences are off by default. If you want preferences to apply, you must include the command "Preferences 1" in the cmds parameter. The "Preferences 1" command applies only to one invocation of the Execute method so you must include it every time you want preferences to apply.
Unlike Execute, the Execute2 method does not return an Automation error if the execution of the commands generates an error in Igor. Instead it returns the Automation NOERROR code. To determine if an error occurred, check the pIgorErrorCode output. If it is non-zero then an error occurred.

DataFolder(BSTR nameOrPath, IDataFolder** retval )
Returns a reference to the named data folder.
Parameters
nameOrPath [input] is a Basic string containing the name of a data folder in the current data folder or a data folder path (absolute or relative to the current data folder).
If nameOrPath is "", a reference to the first data folder in the current data folder is returned. If there are no data folders in the current data folder, NULL is returned but no error is generated.
If nameOrPath is not "" and does not reference an existing data folder, NULL is returned and an error is generated.
Unlike the behavior of Igor's command line, nameOrPath does not need to be quoted even if it contains liberal names. Both quoted and unquoted paths are accepted.
retval [output] receives an IDataFolder interface pointer if the function succeeds. See IDataFolder Methods.

get_DataFolders(BSTR nameOrPath, IDataFolders** retval )
Returns a reference to the collection of data folders in the named parent data folder.
Parameters
nameOrPath [input] is a Basic string that designates the parent data folder - that is, the data folder containing the collection of data folders of interest.
nameOrPath can be a simple name of a data folder in the current data folder or a data folder path (absolute or relative to the current data folder).
If nameOrPath does not reference an existing data folder, NULL is returned and an error is generated.
Unlike the behavior of Igor's command line, nameOrPath does not need to be quoted even if it contains liberal names. Both quoted and unquoted paths are accepted.
retval [output] receives an IDataFolders interface pointer if the function succeeds. See IDataFolders Methods.

NewExperiment(int flags)
Creates a new empty experiment in memory. The previous current experiment, if any is disposed.
Parameters
flags [input] is a bitwise parameter. Currently no bits are defined and you must pass zero.
Details
NOTE: NewExperiment deletes the data from the old experiment without asking if you want to save changes. If you do want to save changes, call the SaveExperiment method before calling the NewExperiment method.

LoadExperiment(int flags, int loadType, BSTR symbolicPathName, BSTR filePath)
Loads an experiment from a file into memory, replacing the previous experiment data.
Parameters
flags [input] is a bitwise parameter. Currently no bits are defined and you must pass zero.
loadType [input] is one of the following:
ipLoadTypeOpen:	Does a normal experiment open, like Igor's File->Open Experiment menu command.
ipLoadTypeStationery:	Like ipLoadTypeOpen except that after the load, the current experiment is disassociated from the experiment file and acts like an unsaved experiment.
ipLoadTypeMerge:	Merges an experiment file into the current experiment. See Merging Experiments.
For details on using enumerations such as ipLoadTypeOpen in Visual Basic, see Enumerations in Visual Basic.
symbolicPathName [input] is the name of an Igor symbolic path or "".  See Symbolic Paths.
filePath [input] is a simple file name, a partial path to an experiment file or a full path to an experiment file.
Details
NOTE: LoadExperiment does not ask if you want to save changes to the previous current experiment. If you do want to save changes, call the SaveExperiment method before calling the LoadExperiment method.
The experiment file to be loaded can be specified as a full Windows path or by reference to an Igor symbolic path in combination with a partial path or a simple file name.
C++ programmers should remember that, to embed a backslash in a literal string, you need to use a double-backslash. This is not the case in Visual Basic.
Visual Basic Examples
Dim filePath as String

' Open experiment using full path.
filePath = "C:\Program Files\WaveMetrics\Igor Pro Folder\Examples\Sample Graphs\Demo Experiment #1.pxp"
IgorApp.LoadExperiment 0, IgorPro.IgorProLoadType.ipLoadTypeOpen, "", filePath

' Open experiment using a symbolic path and a partial path.
filePath = "\Examples\Sample Graphs\Demo Experiment #1.pxp"
IgorApp.LoadExperiment 0, IgorPro.IgorProLoadType.ipLoadTypeOpen, "Igor", filePath

' Open experiment using a symbolic path and a simple file name. (Assumes a symbolic path named SampleGraphs exists.)
filePath = "Demo Experiment #1.pxp"
IgorApp.LoadExperiment 0, IgorPro.IgorProLoadType.ipLoadTypeOpen, "SampleGraphs", filePath

SaveExperiment(int flags, IgorProSaveType saveType, IgorProExpFileType expFileType, BSTR symbolicPathName, BSTR filePath)
Saves the current experiment to a file.
Parameters
flags [input] is a bitwise parameter. Currently no bits are defined and you must pass zero.
saveType [input] is one of the following:
ipSaveTypeSave:	Does a normal experiment save, like Igor's File->Save Experiment menu command. Use this to save the current experiment to the same file from which it was loaded.
	If the experiment was never saved, ipSaveTypeSave acts just like ipSaveTypeSaveAs. If the experiment was previously saved then the values in symbolicPathName and filePath are not used.
ipSaveTypeSaveAs:	Does a Save As, like Igor's File->Save Experiment As menu command. Use this to save the current experiment to a different file and associate it with the new file.
ipSaveTypeSaveCopy:	Does a Save A Copy, like Igor's File->Save Experiment Copy menu command. Use this to save the current experiment to a different file without associating it with the new file.
For details on using enumerations such as ipSaveTypeSave in Visual Basic, see Enumerations in Visual Basic.
expFileType [input] controls whether the experiment is saved as a packed file or as an unpacked file. It is ignored when doing a normal experiment save (saveType=ipSaveTypeSave). expFileType is one of the following:
ipExpFileTypeDefault:	Saves the experiment in the format in which it was loaded. This is the value that should be used for most purposes.
ipExpFileTypePacked	Saves the experiment as a packed experiment file even if it was loaded from an unpacked file.
ipExpFileTypeUnpacked	Saves the experiment as a unpacked experiment file even if it was loaded from a packed file.
For details on using enumerations such as ipExpFileTypeDefault in Visual Basic, see Enumerations in Visual Basic.
symbolicPathName [input] is the name of an Igor symbolic path or "".  See Symbolic Paths.
filePath [input] is a simple file name, a partial path to an experiment file or a full path to an experiment file.
Details
NOTE: SaveExperiment does not ask if you want want to overwrite any pre-existing file - it just overwrites it.
NOTE: If you use ipExpFileTypeUnpacked, SaveExperiment does not ask if you want want to overwrite any pre-existing folder - it just overwrites it. The folder name used is the same as the experiment file name specified by the filePath parameter with the extension removed and " Folder" added.
The experiment file to be written can be specified as a full Windows path or by reference to an Igor symbolic path in combination with a partial path or a simple file name.
C++ programmers should remember that, to embed a backslash in a literal string, you need to use a double-backslash. This is not the case in Visual Basic.
Visual Basic Examples
Dim filePath as String

' Save experiment using full path.
filePath = "C:\Program Files\WaveMetrics\Igor Pro Folder\Examples\Sample Graphs\Demo Experiment #1.pxp"
IgorApp.SaveExperiment 0, IgorPro.IgorProSaveType.ipSaveTypeSave, IgorPro.IgorProExpFileType.ipExpFileTypeDefault, "", filePath

' Save experiment using a symbolic path and a partial path.
filePath = "\Examples\Sample Graphs\Demo Experiment #1.pxp"
IgorApp.SaveExperiment 0, IgorPro.IgorProSaveType.ipSaveTypeSave, IgorPro.IgorProExpFileType.ipExpFileTypeDefault, "Igor", filePath

' Save experiment using a symbolic path and a simple file name. (Assumes a symbolic path named SampleGraphs exists.)
filePath = "Demo Experiment #1.pxp"
IgorApp.SaveExperiment(0, IgorPro.IgorProSaveType.ipSaveTypeSave, IgorPro.IgorProExpFileType.ipExpFileTypeDefault, "SampleGraphs", filePath

OpenFile(int flags, int fileKind, BSTR symbolicPathName, BSTR filePath)
Opens a procedure file, notebook file or help file.
Parameters
flags [input] is a bitwise parameter and can be zero or any combination of the following:
ipOpenFileReadOnly
ipOpenFileInvisible
For details on using enumerations such as ipOpenFileReadOnly in Visual Basic, see Enumerations in Visual Basic.
fileKind [input] is one of the following:
ipFileKindNotebook
ipFileKindProcedure
ipFileKindHelp
For details on using enumerations such as ipFileKindNotebook in Visual Basic, see Enumerations in Visual Basic.
symbolicPathName [input] is the name of an Igor symbolic path or "".  See Symbolic Paths.
filePath [input] is a simple file name, a partial path to an experiment file or a full path to an experiment file.
Details
If the file is already open, the OpenFile method does nothing except to bring the file to the top of the front if ipOpenFileInvisible was not used.
The experiment file to be loaded can be specified as a full Windows path or by reference to an Igor symbolic path in combination with a partial path or a simple file name.
C++ programmers should remember that, to embed a backslash in a literal string, you need to use a double-backslash. This is not the case in Visual Basic.
For must purposes, you should use the InsertInclude method to open a procedure file rather than the OpenFile method because you can use DeleteInclude to remove the procedure file and there is no way to remove a procedure file that you have opened with OpenFile.
After calling OpenFile to open a procedure file, call the CompileProcedures method to make Igor compiles procedures.
Visual Basic Examples
Dim filePath as String

' Open help file using full path.
filePath = "C:\Program Files\WaveMetrics\Igor Pro Folder\More Help Files\Advanced Topics.ihf"
IgorApp.OpenFile 0, IgorPro.IgorProFileKind.ipFileKindHelp, "", filePath

' Open procedure file using a symbolic path and a partial path.
filePath = "\WaveMetrics Procedures\Waves\Compare Waves.ipf"
IgorApp.OpenFile IgorPro.IgorProOpenFileFlags.ipOpenFileInvisible, "Igor", filePath

InsertInclude(int flags, BSTR procedureSpec)
Inserts a #include statement in the Procedure window.
Parameters
flags [input] is a bitwise parameter. Currently no bits are defined and you must pass zero.
procedureSpec [input] is the part of the #include statement that you would type after "#include " if you were entering the statement manually in the procedure window. See Including a Procedure File.
Details
The InsertInclude method does nothing if the include statement is already in the procedure window.
After calling InsertInclude, call the CompileProcedures method to make Igor compiles procedures.
Visual Basic Example
Dim procedureSpec as String
procedureSpec = "<Compare Waves>"
IgorApp.InsertInclude 0, procedureSpec

DeleteInclude(int flags, BSTR procedureSpec)
Deletes a #include statement from the Procedure window.
Parameters
flags [input] is a bitwise parameter. Currently no bits are defined and you must pass zero.
procedureSpec [input] is the part of the #include statement that you would type after "#include " if you were entering the statement manually in the procedure window. See Including a Procedure File.
Details
The DeleteInclude method does nothing if the include statement does not appear in the procedure window.
After calling DeleteInclude, call the CompileProcedures method to make Igor compiles procedures.
Visual Basic Example
Dim procedureSpec as String
procedureSpec = "<Compare Waves>"
IgorApp.DeleteInclude 0, procedureSpec

CompileProcedures(int flags)
Causes Igor procedure files to be compiled. You must call this after calling InsertInclude, DeleteInclude or OpenFile to open a procedure file.
Parameters
flags [input] is a bitwise parameter. Currently only one bit is defined. You can pass zero or:
ipCompileProceduresNoErrorDialog
	Tells Igor to not display an error dialog if procedures can not be compiled because of an error.
For details on using enumerations such as ipCompileProceduresNoErrorDialog in Visual Basic, see Enumerations in Visual Basic.
Details
The CompileProcedures method does nothing if the procedures are already compiled.
CompileProcedures returns an error if procedures can not be compiled because of a programming error in a procedure file.
Visual Basic Example
IgorApp.CompileProcedures 0

Quit()
Makes Igor Pro quit.
Details
NOTE: Quit does not ask if you want to save changes to the current experiment. If you do want to save changes, call the SaveExperiment method before calling the Quit method.

#	IDataFolder Methods
The IDataFolder interface allows you to navigate Igor's data folder hierarchy and to access the contents of data folders (waves, numeric variables, string variables and other data folders). Use the IApplication::DataFolder method to obtain an IDataFolder interface pointer.

get_Name(BSTR* name)
Retrieves the name of the data folder.
Parameters
name [output] receives the name of the data folder.

put_Name(BSTR name)
Sets the name of the data folder.
Parameters
name [input] is the new name for the data folder.
An error is generated if the name is not a legal data folder name or is in conflict with the name of another data folder at the same level of the data folder hierarchy.

get_InUse(VARIANT_BOOL* inUse)
Sets inUse to -1 if the data folder is in use or to zero otherwise.
Parameters
inUse [output] is set to -1 if the data folder is in use or to zero otherwise.
Details
A data folder that is in use can not be removed (killed). A data folder is in use if it contains one or more waves that are in use (displayed in a graph or table, used in a dependency, used by an XOP) or locked or if it contains a sub-data folder that is in use.

get_IsRoot(VARIANT_BOOL* isRoot)
Sets isRoot to a -1 if the data folder is the root data folder or to zero if not.
Parameters
isRoot [output] is set to -1 if the data folder is the root data folder or to zero if not.
Despite its name, VARIANT_BOOL (a Microsoft data type) is a short, not a VARIANT.

get_Path(int getRelativePath, int getQuotedPath, BSTR* path)
Retrieves the path to the data folder.
Parameters
getRelativePath [input]: If zero, the output path will be a full path starting from root. If non-zero, the output path will be relative to the current data folder.
getQuotedPath [input]: If zero, the output path will be unquoted. If non-zero, liberal names in the output path will be quoted. When calling the IApplication.Execute method you must use quoted paths.
path [output] receives the path to the data folder.
Details
Relative paths always start with a colon.
The output path will always end with a colon.

get_Next
Returns a reference to the next data folder at the same level of the data folder hierarchy or NULL if there are no more data folders.
This method returns the data folders in order of creation.
Visual Basic Example
	Private Sub ListAllDataFoldersInCurrentDataFolder(IgorApp as  As IgorPro.Application)
		Dim df As DataFolder
		Set df = IgorApp.DataFolder("")	' Get first data folder in current data folder
		Do Until df Is Nothing
			Debug.WriteLine df.Name        ' Use Debug.Print in VB 6
			Set df = df.Next	' Get next data folder in current data folder
		Loop
	End Sub

get_ParentDataFolder(IDataFolder** retval )
Returns a reference to the data folder's parent data folder.
The root data folder has no parent and Igor will generate an error if you try to get its parent.
Parameters
retval [output] receives an IDataFolder interface pointer if the function succeeds. See IDataFolder Methods.

get_SubDataFolders( IDataFolders** retval )
Returns a reference to the collection of data folders in the parent data folder.
Parameters
retval [output] receives an IDataFolders interface pointer if the function succeeds. See IDataFolders Methods.

WaveExists(BSTR waveNameOrPath, VARIANT_BOOL* exists)
Sets exists to a -1 if the wave exists or to zero if it does not exist.
Parameters
waveNameOrPath [input] is a Basic string containing the name of a wave in the associated data folder or a data folder path (absolute or relative to the associated data folder) to an Igor wave.
Unlike the behavior of Igor's command line, waveNameOrPath does not need to be quoted even if it contains liberal names. Both quoted and unquoted paths are accepted.
exists [output] is set to -1 if the wave exists, zero if not.
Despite its name, VARIANT_BOOL (a Microsoft data type) is a short, not a VARIANT.

Wave(BSTR waveNameOrPath, IWave** retval )
Returns a reference to the named wave.
Parameters
waveNameOrPath [input] is a Basic string containing the name of a wave in the associated data folder or a data folder path (absolute or relative to the associated data folder) to an Igor wave.
If waveNameOrPath is "", a reference to the first wave in the associated data folder is returned. If there are no waves in the associated data folder, NULL is returned.
Unlike the behavior of Igor's command line, waveNameOrPath does not need to be quoted even if it contains liberal names. Both quoted and unquoted paths are accepted.
retval [output] receives an IWave interface pointer if the function succeeds. See IWave Methods.

get_Waves(IWaves** retval )
Returns a reference to a collection of all of the waves in the associated data folder.
Parameters
retval [output] receives an IWaves interface pointer if the function succeeds. See IWaves Methods.

VariableExists(BSTR varNameOrPath, VARIANT_BOOL* exists)
Sets exists to a -1 if the global variable exists or to zero if it does not exist.
Parameters
varNameOrPath [input] is a Basic string containing the name of a global variable in the associated data folder or a data folder path (absolute or relative to the associated data folder) to an Igor global variable.
Unlike the behavior of Igor's command line, varNameOrPath does not need to be quoted even if it contains liberal names. Both quoted and unquoted paths are accepted.
exists [output] is set to -1 if the global variable exists, zero if not.
Despite its name, VARIANT_BOOL (a Microsoft data type) is a short, not a VARIANT.

Variable(BSTR varNameOrPath, IVariable** retval )
Returns a reference to the named global variable.
Parameters
varNameOrPath [input] is a Basic string containing the name of a global variable in the associated data folder or a data folder path (absolute or relative to the associated data folder) to an Igor global variable.
If varNameOrPath is "", a reference to the first global variable in the associated data folder is returned. If there are no global variables in the associated data folder, NULL is returned.
Unlike the behavior of Igor's command line, varNameOrPath does not need to be quoted even if it contains liberal names. Both quoted and unquoted paths are accepted.
retval [output] receives an IVariable interface pointer if the function succeeds. See IVariable Methods.

get_Variables(IVariables** retval )
Returns a reference to a collection of all of the global variables in the associated data folder.
Parameters
retval [output] receives an IVariaiables interface pointer if the function succeeds. See IVariables Methods.

#	IDataFolders Methods
The IDataFolders class is a collection class used to navigate through Igor's data folder hierarchy. An instance of the IDataFolders class represents all of the data folders in a particular "parent" data folder. You specify the parent data folder when you create an instance of IDataFolders by calling IApplication.DataFolders.

get_Count(long* retval)
Returns via retval the number of data folders in the parent data folder.
Parameters
retval  [output] receives the number of data folders in the parent data folder.

get_Item(VARIANT vIndex, IDataFolder** retval)
Given the zero-based index of a data folder or the name of a data folder in the parent data folder, returns a reference to that data folder.
Parameters
vIndex [input] is a variant that specifies the data folder of interest.
retval [output] receives a reference to the data folder.
Details
In Visual Basic, Item is the default property of the IDataFolders interface.
If vIndex contains a numeric value then that value is taken to be a zero-based index of the data folder of interest and Item returns a reference to the corresponding data folder.
If vIndex contains a BSTR then Item returns a reference to the data folder with that name. In this case, the BSTR must contain just a simple data folder name, not a path to a data folder.
If the data folder does not exist, Item returns an error. You can use the DataFolders.Count method or the DataFolders.DataFolderExists method to determine if a data folder exists before you call the Item method.

DataFolderExists(BSTR name, VARIANT_BOOL* exists)
Returns the truth that the named data folder exists via the exists parameter.
Parameters
name [input] is the simple name of the data folder (not a path to the data folder) in the parent data folder.
exists [output] receives the truth that the data folder exists (0=no, -1=yes).

Add(BSTR bstrName, int overwrite, IDataFolder** ppDataFolder)
Creates a new data folder in the parent data folder and returns a reference to it.
It is an error if a data folder with the specified name exists in the data folder.
Parameters
bstrName [input] contains the simple name (not a path) of the data folder to create in the parent data folder.
overwrite [input]. If overwrite is zero and the data folder exists, an error is returned. If overwrite is non-zero and the data folder exists, no error is returned and the data folder is left intact.
ppDataFolder [output] receives the reference to the newly-created data folder.
Details
If overwrite is non-zero and the data folder exists, it is not actually overwritten. Instead it is left intact, as it was before the Add method was called.

Remove(VARIANT vIndex)
Given the zero-based index of a data folder or the name of a data folder in the parent data folder, removes (kills) the data folder.
Parameters
vIndex [input] is a variant that specifies the data folder to be killed.
Details
If vIndex contains a numeric value then that value is taken to be a zero-based index of the data folder to be killed.
If vIndex contains a BSTR then it is the name of the data folder to be killed. In this case, the BSTR must contain just a simple data folder name, not a path to a data folder.
If the data folder does not exist, Remove returns an error. You can use the DataFolders.Count method or the DataFolders.DataFolderExists method to determine if a data folder exists before you call the Remove method.
A data folder can not be killed if in use (see IDataFolder.InUse for details). If you try to kill a data folder that is in use, Remove will return an error.
Remove will also return an error if your program holds an interface pointer to an IDataFolder object representing the data folder. See Killing Data Folders for a discussion of how to deal with this issue.
The root data folder is special. When you "remove it", Igor really removes just its contents while the root data folder itself is not removed.

get__NewEnum(IUnknown** retval)
This method is used by Visual Basic to implement the For-Each loop. You can not use it directly.
C++ clients should use the DataFolders.Next method or the DataFolder.Next method to iterate through all data folders.
Visual Basic Example
	Dim wcol As DataFolders ' Collection of data folders in the current data folder
	Set wcol = IgorApp.DataFolders
	For Each w In wcol
		Debug.WriteLine w.Name    ' Use Debug.Print in VB 6
	Next

#	IWave Methods

get_Name(BSTR* name)
Retrieves the name of the wave.
Parameters
name [output] receives the name of the wave.

put_Name(BSTR name)
Sets the name of the wave.
Parameters
name [input] is the new name for the wave.
An error is generated if the name is not a legal wave name or is in conflict with the name of another object at the same level of the data folder hierarchy.

get_InUse(VARIANT_BOOL* inUse)
Sets inUse to -1 if the wave is in use or to zero otherwise.
Parameters
inUse [output] is set to -1 if the wave is in use or to zero otherwise.
Details
A wave that is in use can not be removed (killed). A wave is in use if it  is displayed in a graph or table, is used in a dependency, or is used by an XOP, or is locked.
A wave is also in use if an Automation client (e.g., your program) holds an interface pointer to an IWave object representing the wave.

get_Path(int getRelativePath, int getQuotedPath, BSTR* path)
Retrieves the path to the wave.
Parameters
getRelativePath [input]: If zero, the output path will be a full path starting from root. If non-zero, the output path will be relative to the current data folder.
getQuotedPath [input]: If zero, the output path will be unquoted. If non-zero, liberal names in the output path will be quoted. When calling the IApplication.Execute method you must use quoted paths.
path [output] receives the path to the wave.
Details
Relative paths always start with a colon.

get_ParentDataFolder(IDataFolder** retval )
Returns a reference to the wave's parent data folder.
Parameters
retval [output] receives an IDataFolder interface pointer if the function succeeds. See IDataFolder Methods.

GetDimensions(IgorProDataType* pDataType, long* pNumRows, long* pNumColumns, long* pNumLayers, long* pNumChunks)
Returns the wave's data type and the size of each dimension.
Parameters
pDataType [output] is the data type of the wave. See IgorProDataType Enumeration for allowable values.
pNumRows [output] is the number of rows in the wave.
pNumColumns [output] is the number of layers in the wave. For a wave of dimension less than two, this will be zero.
pNumLayers [output] is the number of layers in the wave. For a wave of dimension less than three, this will be zero.
pNumChunks [output] is the number of chunks in the wave. For a wave of dimension less than four, this will be zero.

SetDimensions(IgorProDataType dataType, long numRows, long numColumns, long numLayers, long numChunks)
Sets the data type of the wave and the size of each dimension.
Parameters
dataType [input] specifies the data type of the wave. See IgorProDataType Enumeration for allowable values.
numRows [input] specifies the number of rows that the wave should have.
numColumns [input] specifies the number of layers that the wave should have. For a wave of dimension less than two, this must be zero.
numLayers [input] specifies the number of layers that the wave should have. For a wave of dimension less than three, this must be zero.
numChunks [input] specifies the number of chunks that the wave should have. For a wave of dimension less than four, this must be zero.

GetScaling(int dimension, double* psfA, double* pSFB)
Returns the wave's dimension scaling.
Parameters
dimension [input] identifies the dimension of interest. 0 is the rows dimension, 1 is the columns dimension, 2 is the layers dimension and 3 is the chunks dimension. If dimension is -1, the wave's data full scale value is returned.
psfA [output] is the intercept of the dimension scaling.
pSFB [output] is the slope of the dimension scaling.
Details
The scaled dimension index for element e is equal to:
	sfA + sfB * e

SetScaling(int dimension, double sfA, double sfB)
Sets the wave's dimension scaling.
Parameters
dimension [input] identifies the dimension of interest. 0 is the rows dimension, 1 is the columns dimension, 2 is the layers dimension and 3 is the chunks dimension. If dimension is -1, the wave's data full scale value is set.
sfA [input] is the intercept of the dimension scaling.
sfB [input] is the slope of the dimension scaling.
Details
The scaled dimension index for element e is equal to:
	sfA + sfB * e

get_Units(int dimension, int codePage, BSTR* pUnits)
Retrieves the wave's units for the specified dimension.
Parameters
dimension [input] identifies the dimension of interest. 0 is the rows dimension, 1 is the columns dimension, 2 is the layers dimension and 3 is the chunks dimension. If dimension is -1, the wave's data units are returned.
codePage [input] is a value that Igor uses to convert Igor text data into Unicode. In most cases you should pass zero, which represents the system default ANSI code page (known as CP_ACP in Microsoft's Windows API).
pUnits [output] receives the wave's units for the specified dimension.

put_Units(int dimension, int codePage, BSTR units)
Sets the wave's units for the specified dimension.
Parameters
dimension [input] identifies the dimension of interest. 0 is the rows dimension, 1 is the columns dimension, 2 is the layers dimension and 3 is the chunks dimension. If dimension is -1, the wave's data units are set.
codePage [input] is a value that Igor uses to convert Igor text data into Unicode. In most cases you should pass zero, which represents the system default ANSI code page (known as CP_ACP in Microsoft's Windows API).
units [input] specifies the new units.

get_DimensionLabels(int dimension, long element, int codePage, BSTR* pDimensionLabel)
Retrieves the dimension label for the specified element of the specified dimension.
Parameters
dimension [input] identifies the dimension of interest. 0 is the rows dimension, 1 is the columns dimension, 2 is the layers dimension and 3 is the chunks dimension.
element [input] is the zero-based element number (row, column, layer or chunk). If element is -1, the overall label for the entire dimension is returned.
codePage [input] is a value that Igor uses to convert Igor text data into Unicode. In most cases you should pass zero, which represents the system default ANSI code page (known as CP_ACP in Microsoft's Windows API).
pDimensionLabel [output] receives the dimension label for the specified element of the specified dimension.

put_DimensionLabel(int dimension, long element,, int codePage, BSTR dimensionLabel)
Sets the dimension label for the specified element of the specified dimension.
Parameters
dimension [input] identifies the dimension of interest. 0 is the rows dimension, 1 is the columns dimension, 2 is the layers dimension and 3 is the chunks dimension. If dimension is -1, the wave's data units are set.
element [input] is the zero-based element number (row, column, layer or chunk). If element is -1, the overall label for the entire dimension is returned.
codePage [input] is a value that Igor uses to convert Igor text data into Unicode. In most cases you should pass zero, which represents the system default ANSI code page (known as CP_ACP in Microsoft's Windows API).
dimensionLabel [input] specifies the dimension label for the specified element of the specified dimension.

GetMiscInfo(double* pCreateDate, double* pModDate, double* pModCount)
Returns the wave's data type and the size of each dimension.
Parameters
pCreateDate [output] is the wave's creation date/time expressed as seconds since midnight, January 1, 1904. This has a resolution of one second.
pModDate [output] is the wave's modifcation date/time expressed as seconds since midnight, January 1, 1904. This has a resolution of one second.
pModCount [output] is a number that Igor changes each time the wave is modified. You can use it to see if a wave has changed since the last time you checked. You should not depend on the exact value of this output but merely check to see if it has changed since the last time you checked.

get_Note(int codePage, BSTR* pNote)
Retrieves the wave's note.
Parameters
codePage [input] is a value that Igor uses to convert Igor text data into Unicode. In most cases you should pass zero, which represents the system default ANSI code page (known as CP_ACP in Microsoft's Windows API).
pNote [output] receives the contents of the wave's note.

put_Note(int codePage, BSTR note)
Sets the wave's note.
Parameters
codePage [input] is a value that Igor uses to convert Igor text data into Unicode. In most cases you should pass zero, which represents the system default ANSI code page (known as CP_ACP in Microsoft's Windows API).
note [input] is the new wave note text.

get_Lock(int* pLock)
Retrieves the wave's lock status.
Parameters
pLock [output] receives the wave's lock status. 0 means unlocked, 1 means locked. All other bits in this value are reserved for future use.

put_Lock(int lock)
Sets the wave's lock status.
Parameters
lock [input] is the new wave lock status. 0 means unlocked, 1 means locked. All other bits in this value are reserved for future use.

GetNumericWaveData(IgorProDataType ipDataType, VARIANT* vDataPtr)
Retrieves a wave's data from Igor.
The data is returned via a variant that references a SafeArray created by Igor. On return, the client owns the variant and the SafeArray.
Parameters
ipDataType  [input] specifies the numeric data type of the returned data. ipDataType does not need to be the same as the wave's data type, but if it is not Igor will have to do a data type conversion.
For details on using the IgoProDataType enumeration in Visual Basic, see Enumerations in Visual Basic.
vDataPtr [output] points to a variant that is set by GetNumericWaveData. The variant will contain a SafeArray.
Details
A wave with zero points is treated as a 1D wave with zero elements.
Complex waves have the real and imaginary part of each data point interleaved in the rows dimension.
In a 2D wave, the data for each row of the first column is contiguous, followed by the data for each row of the next column, and so on. In a 3D wave, the data for each layer is contigous, is organized like a 2D wave, and is followed by the data for the next layer.

SetNumericWaveData(IgorProDataType ipDataType, VARIANT vData)
Sets a wave's data.
The data is passed to Igor via a variant that references a SafeArray created by the client. On return, the client owns the variant and the SafeArray.
Parameters
ipDataType  [input] specifies the numeric data type of the data in the SafeArray referenced by vData. ipDataType does not need to be the same as the wave's data type, but if it is not Igor will have to do a data type conversion.
For details on using the IgoProDataType enumeration in Visual Basic, see Enumerations in Visual Basic.
vData [input] is a variant that is created by the client. The variant must contain a SafeArray of a data type compatible with ipDataType. On return, the client still owns the variant and the SafeArray.
Details
A wave with zero points is treated as a 1D wave with zero elements.
Complex waves have the real and imaginary part of each data point interleaved in the rows dimension.
In a 2D wave, the data for each row of the first column is contiguous, followed by the data for each row of the next column, and so on. In a 3D wave, the data for each layer is contigous, is organized like a 2D wave, and is followed by the data for the next layer.

GetNumericWaveDataAsDouble(SAFEARRAY** ppsa)
Retrieves a wave's data as double-precision floating point.
The data is returned via a SafeArray created by Igor. On return, the client owns the SafeArray.
Parameters
ppsa [output]. GetNumericWaveDataAsDouble sets *ppsa to point to a newly created SafeArray containing double-precision data.
Details
A wave with zero points is treated as a 1D wave with zero elements.
Complex waves have the real and imaginary part of each data point interleaved in the rows dimension.
In a 2D wave, the data for each row of the first column is contiguous, followed by the data for each row of the next column, and so on. In a 3D wave, the data for each layer is contigous, is organized like a 2D wave, and is followed by the data for the next layer.

SetNumericWaveDataAsDouble(SAFEARRAY** ppsa)
Sets a wave's data from double-precision floating data.
The data is passed to Igor via a SafeArray created by the client. On return, the client owns the SafeArray.
Parameters
ppsa [input]. *ppsa points to a SafeArray created by the client containing double-precision data.
Details
Logic says that the parameter should be a SAFEARRAY*, not a SAFEARRAY**. We had to make it a SAFEARRAY** because SAFEARRAY* did not work with Visual Basic 6 using early binding.
A wave with zero points is treated as a 1D wave with zero elements.
Complex waves have the real and imaginary part of each data point interleaved in the rows dimension.
In a 2D wave, the data for each row of the first column is contiguous, followed by the data for each row of the next column, and so on. In a 3D wave, the data for each layer is contigous, is organized like a 2D wave, and is followed by the data for the next layer.

GetNumericWavePointValue(long index, double* pValue)
Returns 1D numeric wave data for a single data point. Supports real data only.
Parameters
index [input] is a point number index into a 1D wave.
pValue [output] receives the value of the wave at that point.
Details
This method is for use with 1D, real waves only. Use MDGetNumericWavePointValue for multi-dimensional or complex waves.

SetNumericWavePointValue(long index, double value)
Sets 1D numeric wave data for a single data point. Supports real data only.
Parameters
index [input] is a point number index into a 1D wave.
value [input] is the value to be stored in the wave.
Details
This method is for use with 1D, real waves only. Use MDSetNumericWavePointValue for multi-dimensional or complex waves.

MDGetNumericWavePointValue(long row, long column, long layer, long chunk, double* pRealValue, double* pImagValue)
Returns numeric wave data for a single data point.
Parameters
row [input] is the row index into the wave.
column [input] is the column index into the wave. Pass zero for column if the wave's dimensionality is less than 2.
layer [input] is the layer index into the wave. Pass zero for layer if the wave's dimensionality is less than 3.
chunk [input] is the chunk index into the wave. Pass zero for chunk if the wave's dimensionality is less than 4.
pRealValue [output] receives the value of the real part of the specified element of the wave.
pImagValue [output] receives the value of the imaginary part of the specified element of the wave. If the wave is not complex, the value returned via pImagValue is undefined.
Details
This method is for use with real or complex numeric waves of any dimension. For 1D, real waves, GetNumericWavePointValue will be slightly more efficient.

MDSetNumericWavePointValue(long row, long column, long layer, long chunk, double realValue, double imagValue)
Sets numeric wave data for a single data point.
Parameters
row [input] is the row index into the wave.
column [input] is the column index into the wave. Pass zero for column if the wave's dimensionality is less than 2.
layer [input] is the layer index into the wave. Pass zero for layer if the wave's dimensionality is less than 3.
chunk [input] is the chunk index into the wave. Pass zero for chunk if the wave's dimensionality is less than 4.
realValue [input] is the value to be stored in the real part of the specified element of the wave.
imagValue [input] is the value to be stored in the imaginary part of the specified element of the wave. If the wave is not complex, the value passed as imagValue is immaterial.
Details
This method is for use with real or complex numeric waves of any dimension. For 1D, real waves, SetNumericWavePointValue will be slightly more efficient.

GetTextWavePointValue(long index, int codePage, BSTR* pValue)
Returns 1D text wave data for a single data point.
The value is returned from Igor as a Basic string which uses Unicode encoding.
Parameters
index [input] is a point number index into a 1D wave.
codePage [input] is a value that Igor uses to convert Igor text data into Unicode. In most cases you should pass zero, which represents the system default ANSI code page (known as CP_ACP in Microsoft's Windows API).
pValue [output] receives the value of the wave at that point.
Details
This method is for use with 1D text waves only. Use MDGetTextWavePointValue for multi-dimensional text waves.

SetTextWavePointValue(long index, int codePage, BSTR value)
Sets 1D text wave data for a single data point.
The value is passed to Igor as a Basic string which uses Unicode encoding.
Parameters
index [input] is a point number index into a 1D wave.
codePage [input] is a value that Igor uses to convert the Unicode parameter into Igor text data. In most cases you should pass zero, which represents the system default ANSI code page (known as CP_ACP in Microsoft's Windows API).
value [input] is the value to be stored in the wave.
Details
This method is for use with 1D text waves only. Use MDSetTextWavePointValue for multi-dimensional text waves.

MDGetTextWavePointValue(long row, long column, long layer, long chunk, int codePage, BSTR* pValue)
Returns text wave data for a single data point.
The value is returned from Igor as a Basic string which uses Unicode encoding.
Parameters
row [input] is the row index into the wave.
column [input] is the column index into the wave. Pass zero for column if the wave's dimensionality is less than 2.
layer [input] is the layer index into the wave. Pass zero for layer if the wave's dimensionality is less than 3.
chunk [input] is the chunk index into the wave. Pass zero for chunk if the wave's dimensionality is less than 4.
codePage [input] is a value that Igor uses to convert Igor text data into Unicode. In most cases you should pass zero, which represents the system default ANSI code page (known as CP_ACP in Microsoft's Windows API).
pValue [output] receives the value of the wave at that point.
Details
This method is for use with text waves of any dimension. For 1D text waves, GetTextWavePointValue will be slightly more efficient.

MDSetTextWavePointValue(long row, long column, long layer, long chunk, int codePage, BSTR value)
Sets text wave data for a single data point.
The value is passed to Igor as a Basic string which uses Unicode encoding.
Parameters
row [input] is the row index into the wave.
column [input] is the column index into the wave. Pass zero for column if the wave's dimensionality is less than 2.
layer [input] is the layer index into the wave. Pass zero for layer if the wave's dimensionality is less than 3.
chunk [input] is the chunk index into the wave. Pass zero for chunk if the wave's dimensionality is less than 4.
codePage [input] is a value that Igor uses to convert the Unicode parameter into Igor text data. In most cases you should pass zero, which represents the system default ANSI code page (known as CP_ACP in Microsoft's Windows API).
value [input] is the value to be stored in the wave.
Details
This method is for use with text waves of any dimension. For 1D text waves, SetTextWavePointValue will be slightly more efficient.

GetRawTextWaveData(SAFEARRAY** ppsa)
Retrieves a text wave's data.
NOTE: This method is provided for advanced programmers only and is needed only in situations where you have to deal with large text waves. For most uses you should use GetTextWavePointValue instead of this method.
The data is returned via a SafeArray created by Igor. On return, the client owns the SafeArray.
The data is returned as Igor stores it. It is not Unicode. It is a simple array of bytes which can be thought of as "multibyte" characters. See Microsoft's documentation for the "Multibyte character system." For many purposes you can think of it as "ASCII text".
Parameters
ppsa [output]. GetRawTextWaveData sets *ppsa to point to a newly created SafeArray containing unsigned char data.
Details
The returned SafeArray contains unsigned char data which you must interpret as three contiguous sections:
	A 256 byte header
	A variable length array of indices
	A variable length array of raw text
The header is defined as follows:
	#pragma pack(2)
	struct RawTextHeader {
		unsigned long numBytesTotal;	// Number of bytes in header plus following data.
		unsigned long numElements;	// Number of elements in the wave.
		unsigned long numBytesOfIndices;	// Size of indices in bytes.
		unsigned long numBytesOfData;	// Size of raw text data in bytes.
		unsigned char reserved[240];	// Reserved. Must be set to zero.
	};
	typedef struct RawTextHeader RawTextHeader;
	typedef struct RawTextHeader *RawTextHeaderPtr;
	#pragma pack()

The indices section is an array of longs which are offsets into the array of raw text. For each element in the text wave, indices[i] is the offset from the start of the raw text array to the start of the text for the next element. For example, if the text wave contains "Red" in point 0, "Green" in point 1, and "blue" in point 2, the indices array would contain 3, 8, and 12. The raw text for the first element runs from byte 0 through byte indices[0]-1. The second element runs from byte indices[0] through byte indices[1]-1. The third element runs from byte indices[1] through byte indices[2]-1.
The SafeArray is a 1D array regardless of the dimensionality of the text wave. For a 2D text wave, the raw text for each row in column 0 appears first, followed by the raw text for each row in column 2, and so on.
C++ Example
See the IgorClient sample program for an illustration of the use of GetRawTextWaveData.

SetRawTextWaveData(SAFEARRAY** ppsa)
Sets a text wave's data.
NOTE: This method is provided for advanced programmers only and is needed only in situations where you have to deal with large text waves. For most uses you should use SetTextWavePointValue instead of this method.
The data is passed to Igor via a SafeArray created by the client. On return, the client owns the SafeArray.
The data is passed to Igor as Igor stores it. It is not Unicode. It is a simple array of bytes which can be thought of as "multibyte" characters. See Microsoft's documentation for the "Multibyte character system." For many purposes you can think of it as "ASCII text".
Parameters
ppsa [input]. *ppsa points to a SafeArray created by the client containing unsigned char data.
Details
Logic says that the parameter should be a SAFEARRAY*, not a SAFEARRAY**. We had to make it a SAFEARRAY** because SAFEARRAY* did not work with Visual Basic 6 using early binding.
See the Details section of the documentation for GetRawTextWaveData for a description of how the data in the SafeArray must be formatted.
C++ Example
See the IgorClient sample program for an illustration of the use of SetRawTextWaveData.

get_Next(IWave** retval )
Returns the a reference to the next wave in Igor's linked list of waves or NULL if there are no more waves.
All waves in a given data folder are in a linked list in the order of the waves' creation.
Visual Basic Example
	Private Sub ListAllWavesInCurrentDataFolder(IgorApp as  As IgorPro.Application)
	    Dim df As DataFolder
	    Set df = IgorApp.DataFolder(":") ' Get reference to current data folder
	    Dim w As Wave
	    Set w = df.Wave("")         ' Get first wave in current data folder
	    Do Until w Is Nothing
	        Debug.WriteLine w.Name  ' Use Debug.Print in VB 6
	        Set w = w.Next          ' Get next wave in current data folder
	    Loop
	End Sub

#	IWaves Methods
The IWaves class is a collection class. An instance of the IWaves class represents all of the waves in a particular parent data folder. You specify the parent data folder when you create the IWaves interface by calling IDataFolder.Waves.

get_Count(long* retval)
Returns via retval the number of waves in the parent data folder.
Parameters
retval  [output] receives the number of waves in the parent data folder.

get_Item(VARIANT vIndex, IWave** retval)
Given the zero-based index of a wave or the name of a wave in the parent data folder, returns a reference to that wave.
Parameters
vIndex [input] is a variant that specifies the wave of interest.
retval [output] receives a reference to the wave.
Details
In Visual Basic, Item is the default property of the IWaves interface.
If vIndex contains a numeric value then that value is taken to be a zero-based index of the wave of interest and Item returns a reference to the corresponding wave.
If vIndex contains a BSTR then Item returns a reference to the wave with that name. In this case, the BSTR must contain just a simple wave name, not a path to a wave.
If the wave does not exist, Item returns an error. You can use the Waves.Count method or the Waves.WaveExists method to determine if a wave exists before you call the Item method.

WaveExists(BSTR name, VARIANT_BOOL* exists)
Returns the truth that the named wave exists via the exists parameter.
Parameters
name [input] is the simple name of the wave (not a path to the wave) in the data folder.
exists [output] receives the truth that the wave exists (0=no, -1=yes).

Add(BSTR bstrName, IgorProDataType dataType, long numRows, long numColumns, long numLayers, long numChunks, int overwrite, IWave** ppWave)
Creates a new wave in the data folder and returns a reference to it.
It is an error if a wave with the specified name exists in the data folder.
Parameters
bstrName [input] contains the simple name (not a path) of the wave to create in the data folder.
dataType [input] specifies the data type of the wave. See IgorProDataType Enumeration for allowable values.
numRows [input] specifies the number of rows that the wave should have.
numColumns [input] specifies the number of layers that the wave should have. For a wave of dimension less than two, this must be zero.
numLayers [input] specifies the number of layers that the wave should have. For a wave of dimension less than three, this must be zero.
numChunks [input] specifies the number of chunks that the wave should have. For a wave of dimension less than four, this must be zero.
overwrite [input]. If overwrite is zero and the wave exists, an error is returned. If overwrite is non-zero and the wave exists, the wave is overwritten (like a Make/O operation).
ppWave [output] receives the reference to the newly-created wave.

Remove(VARIANT vIndex)
Given the zero-based index of a wave or the name of a wave in the parent data folder, removes (kills) the wave.
Parameters
vIndex [input] is a variant that specifies the wave to be killed.
Details
If vIndex contains a numeric value then that value is taken to be a zero-based index of the wave to be killed.
If vIndex contains a BSTR then it is the name of the wave to be killed. In this case, the BSTR must contain just a simple wave name, not a path to a wave.
If the wave does not exist, Remove returns an error. You can use the Waves.Count method or the Waves.WaveExists method to determine if a wave exists before you call the Remove method.
A wave can not be killed if in use (see IWave.InUse for details). If you try to kill a wave that is in use, Remove will return an error.
Remove will also return an error if your program holds an interface pointer to an IWave object representing the wave. See Killing Waves for a discussion of how to deal with this issue.

get__NewEnum(IUnknown** retval)
This method is used by Visual Basic to implement the For-Each loop. You can not use it directly.
C++ clients should use the Waves.Next method or the Wave.Next method to iterate through all waves.
Visual Basic Example
	Dim wcol As Waves	' Collection of waves in the current data folder
	Set wcol = IgorApp.Waves
	For Each w In wcol
		Debug.WriteLine w.Name            ' Use Debug.Print in VB 6
	Next


#	IVariable Methods

get_Name(BSTR* name)
Retrieves the name of the global variable.
Parameters
name [output] receives the name of the global variable.

get_Path(int getRelativePath, int getQuotedPath, BSTR* path)
Retrieves the path to the global variable.
Parameters
getRelativePath [input]: If zero, the output path will be a full path starting from root. If non-zero, the output path will be relative to the current data folder.
getQuotedPath [input]: If zero, the output path will be unquoted. If non-zero, liberal names in the output path will be quoted. When calling the IApplication.Execute method you must use quoted paths.
path [output] receives the path to the global variable.
Details
Relative paths always start with a colon.

get_ParentDataFolder(IDataFolder** retval )
Returns a reference to the global variable's parent data folder.
Parameters
retval [output] receives an IDataFolder interface pointer if the function succeeds. See IDataFolder Methods.

get_DataType(IgorProDataType* pDataType)
Returns the global variable's data type.
Parameters
pDataType [output] is the data type of the global variable. The data type will be one of the following:
	ipDataTypeDouble
	ipDataTypeDouble | ipDataTypeComplex
	ipDataTypeText
For details on using the IgoProDataType enumeration in Visual Basic, see Enumerations in Visual Basic.
GetNumericValue(double* pRealValue, double* pImagValue)
Returns the value of a numeric global variable.
Parameters
pRealValue [output] receives the real part of the numeric value.
pImagValue [output] receives the imaginary part of the numeric value or zero if the variable is not complex.

SetNumericValue(double realValue, double imagValue)
Sets the value of a numeric global variable.
Parameters
realValue [input] is the real part of the numeric value.
imagValue [input] is the imaginary part of the numeric value or zero if the variable is not complex.

GetStringValue(int codePage, BSTR* pValue)
Retrieves the value of a global string variable.
Parameters
codePage [input] is a value that Igor uses to convert Igor text data into Unicode. In most cases you should pass zero, which represents the system default ANSI code page (known as CP_ACP in Microsoft's Windows API).
pValue [output] receives the global string variable's value.

SetStringValue(int codePage, BSTR value)
Sets the value of a global string variable.
Parameters
codePage [input] is a value that Igor uses to convert Unicode into Igor text data. In most cases you should pass zero, which represents the system default ANSI code page (known as CP_ACP in Microsoft's Windows API).
value [input] is the global string variable's new value.

get_Next(IVariable** retval )
Returns the a reference to the next global variable in the parent data folder or NULL if there are no more global variables.
The order of global variables in a data folder is not predictable.
Visual Basic Example
	Private Sub ListAllVariablesInCurrentDataFolder(IgorApp as  As IgorPro.Application)
	    Dim df As DataFolder
	    Set df = IgorApp.DataFolder(":") ' Get reference to current data folder
	    Dim v As Variable
	    Set v = df.Variable("")       ' Get first variable in current data folder
	    Do Until v Is Nothing
	        Debug.WriteLine v.Name    ' Use Debug.Print in VB 6
	        Set v = v.Next            ' Get next variable in current data folder
	    Loop
	End Sub

#	IVariables Methods
The IVariables class is a collection class. An instance of the IVariables class represents all of the global variables in a particular parent data folder. You specify the parent data folder when you create the IWaves interface by calling IDataFolder.Waves.

get_Count(long* retval)
Returns via retval the number of global variables in the parent data folder.
Parameters
retval  [output] receives the number of global variables in the parent data folder.

get_Item(VARIANT vIndex, IWave** retval)
Given the zero-based index of a global variable or the name of a global variable in the parent data folder, returns a reference to that global variable.
Parameters
vIndex [input] is a variant that specifies the global variable of interest.
retval [output] receives a reference to the global variable.
Details
In Visual Basic, Item is the default property of the IWaves interface.
If vIndex contains a numeric value then that value is taken to be a zero-based index of the global variable of interest and Item returns a reference to the corresponding global variable.
If vIndex contains a BSTR then Item returns a reference to the global variable with that name. In this case, the BSTR must contain just a simple global variable name, not a path to a global variable.
If the global variable does not exist, Item returns an error. You can use the Waves.Count method or the Waves.WaveExists method to determine if a global variable exists before you call the Item method.

VariableExists(BSTR name, VARIANT_BOOL* exists)
Returns the truth that the named global variable exists via the exists parameter.
Parameters
name [input] is the simple name of the global variable (not a path to the global variable) in the data folder.
exists [output] receives the truth that the global variable exists (0=no, -1=yes).

Add(BSTR bstrName, IgorProDataType dataType, int overwrite, IVariable** ppWave)
Creates a new global variable in the data folder and returns a reference to it.
It is an error if a global variable with the specified name exists in the data folder.
Parameters
bstrName [input] contains the simple name (not a path) of the global variable to create in the data folder.
dataType [input] specifies the data type of the global variable and must be one of the following:
	ipDataTypeDouble
	ipDataTypeDouble | ipDataTypeComplex
	ipDataTypeText
See IgorProDataType Enumeration for further description of IgorProDataType.
overwrite [input]. If overwrite is zero and the variable exists, an error is returned. If overwrite is non-zero and the variable exists, the variable left intact (like a Variable/G operation).
ppWave [output] receives the reference to the newly-created global variable.

Remove(VARIANT vIndex)
Given the zero-based index of a global variable or the name of a global variable in the parent data folder, removes (kills) the global variable.
Parameters
vIndex [input] is a variant that specifies the global variable to be killed.
Details
If vIndex contains a numeric value then that value is taken to be a zero-based index of the global variable to be killed.
If vIndex contains a BSTR then it is the name of the global variable to be killed. In this case, the BSTR must contain just a simple global variable name, not a path to a global variable.
If the global variable does not exist, Remove returns an error. You can use the Waves.Count method or the Waves.WaveExists method to determine if a global variable exists before you call the Remove method.
A global variable can not be killed if in use (see IWave.InUse for details). If you try to kill a global variable that is in use, Remove will return an error.
Remove will also return an error if your program holds an interface pointer to an IVariable object representing the global variable.

get__NewEnum(IUnknown** retval)
This method is used by Visual Basic to implement the For-Each loop. You can not use it directly.
C++ clients should use the Variables.Next method or the Variable.Next method to iterate through all global variables.
Visual Basic Example
	Dim wcol As Waves	' Collection of waves in the current data folder
	Set wcol = IgorApp.Waves
	For Each w In wcol
		Debug.WriteLine w.Name           ' Use Debug.Print in VB 6
	Next


#	Enumerations
The enumerations shown in this section are defined in the IgorProTypeLib.h file.

Enumerations in Visual Basic
In Visual Basic versions after VB6, you need to fully-qualify enums that you reference. For example:

Dim filePath as String
filePath = "C:\Test.pxp"	// Note: In Windows 7, writing to C: requires administrator privileges.

// Error
IgorApp.SaveExperiment(0, ipSaveTypeSave, ipExpFileTypeDefault, "", filePath)

// OK
IgorApp.SaveExperiment(0, IgorPro.IgorProSaveType.ipSaveTypeSave, IgorPro.IgorProExpFileType.ipExpFileTypeDefault, "", filePath)

Here IgorPro is a namespace created by Visual Basic when it imports the Igor Pro type library.

If you declare that you want to use names in the IgorPro namespace, like this:
	Imports IgorPro
or if you check IgorPro in the Imported Namespaces section of the References tab of the project properties window, then you do not need the "IgorPro." part:

// OK if you have imported the IgorPro namespace
IgorApp.SaveExperiment(0, IgorProSaveType.ipSaveTypeSave, IgorProExpFileType.ipExpFileTypeDefault, "", filePath)

The Imports statement affects just the file in which it appears while the Imported Names checkboxes affect all files in the Visual Basic project.

Enumerations in C++
C++ does not require qualification of enums like Visual Basic. In C++ the SaveExperiment example would be written like this:

BSTR bstrFilePath = SysAllocString(L"C:\\Test.pxp");		// Note: In Windows 7, writing to C: requires administrator privileges.
BSTR bstrSymbolicPathName = SysAllocString(L"");		// Empty string
hr = pApplication->SaveExperiment(0, ipSaveTypeSave, ipExpFileTypeDefault, bstrSymbolicPathName, bstrFilePath);
SysFreeString(bstrFilePath);
SysFreeString(bstrSymbolicPathName);

IgorProDataType Enumeration
Defines the values that can be used to specify the data type of an Igor wave or variable.

typedef enum {
	ipDataTypeFloat = 0x02,
	ipDataTypeDouble = 0x04,
	ipDataTypeSignedByte = 0x08,
	ipDataTypeSignedShort = 0x10,
	ipDataTypeSignedLong = 0x20,
	ipDataTypeUnsignedByte = 0x48,
	ipDataTypeUnsignedShort = 0x50,
	ipDataTypeUnsignedLong = 0x60,
	ipDataTypeComplex = 0x01,	// Use this in combination with another value.
	ipDataTypeText = 0
} IgorProDataType;

ipDataTypeComplex is used only in conjunction with another numeric value. For example:

// Visual Basic
Dim w as Wave
w = IgorApp.Waves.Add("wave0", IgorPro.IgorProDataType.ipDataTypeFloat + IgorPro.IgorProDataType.ipDataTypeComplex, 10, 0, 0, 0, 0)

It is not legal to combine ipDataTypeComplex with ipDataTypeText.

For variables only the following types are allowed:
ipDataTypeDouble
ipDataTypeDouble | ipDataTypeComplex
ipDataTypeText

IgorProStatus1Code Enumeration
Defines the values that can be used with the get_Status1 method.

typedef enum {
	ipStatusIgorVersion = 1,
	ipStatusRunningProcedure = 2,
	ipStatusOperationQueueIsEmpty = 3,
	ipStatusPauseForUser = 4,
	ipStatusExperimentModified = 5,
	ipStatusExperimentNeverSaved = 6,
	ipStatusProceduresCompiled = 7
}	IgorProStatus1Code;

IgorProExecute2Flag Enumeration
Defines the values that can be used with the Execute2 method.

typedef enum {
	ipExecute2Silent = 1
}	IgorProExecute2Flag;

IgorProLoadType Enumeration
Defines the values that can be used with the LoadExperiment method.

typedef enum {
	ipLoadTypeOpen = 2,
	ipLoadTypeStationery = 4,
	ipLoadTypeMerge = 5
} IgorProLoadType;

IgorProSaveType Enumeration
Defines the values that can be used to with the SaveExperiment method.

typedef enum {
	ipSaveTypeSave = 1,
	ipSaveTypeSaveAs = 2,
	ipSaveTypeSaveCopy = 3
} IgorProSaveType;

IgorProExpFileType Enumeration
Defines the values that can be used to with the SaveExperiment.

typedef enum {
	ipExpFileTypeDefault = -1,
	ipExpFileTypeUnpacked = 0,
	ipExpFileTypePacked = 1
} IgorProExpFileType;

IgorProFileKind Enumeration
Defines the values that can be used to with the OpenFile method.

typedef enum {
	ipFileKindNotebook = 11,
	ipFileKindProcedure = 12,
	ipFileKindHelp = 13
} IgorProFileKind;

IgorProOpenFileFlags Enumeration
Defines the values that can be used with the OpenFile method.

typedef enum {
	ipOpenFileReadOnly = 1,
	ipOpenFileInvisible = 2
} IgorProOpenFileFlags;

IgorProCompileProcedureFlag Enumeration
Defines the values that can be used with the CompileProcedures method.

typedef enum {
	ipCompileProceduresNoErrorDialog = 1
} IgorProCompileProceduresFlag;

#	Killing Waves and Data Folders
Igor normally does not permit waves or data folders to be killed if they are in use. The only exception is when the current experiment is closed, in which case all waves and data folders are killed, except for the root data folder.

For waves "in use" means:
	Displayed in a graph or table
	Used in a dependency formula
	Locked
	In use by an XOP
	In use by an Automation client

Since you are the programmer of an automation client, you need to know about this last item.

When you create reference to a wave, that wave can not be killed until you release the reference. For example, in Visual Basic:
	Dim df as DataFolder
	Set df = IgorApp.DataFolder(":")	' Obtain reference to the current data folder.
	Dim w as Wave
	Set w = df.Waves("wave0")	' Obtain reference to wave0 in that data folder

Now Igor will not allow wave0 to be killed. Igor will also not allow the DataFolder referenced by df to be killed.

Because of this, it is a good practice to release references as soon as possible. In Visual Basic a reference is released when the variable containing the reference goes out of scope, when you set the variable to "Nothing", or when you assign it to refer to another object. For example:

Subroutine Test
	Dim df as DataFolder
	Set df = IgorApp.DataFolder(":")	' Obtain reference to the current data folder.
	Dim w as Wave
	Set w = df.Waves("wave0")	' Obtain reference to wave0 in that data folder
	Debug.WriteLine w.Name              ' Use Debug.Print in VB 6
	Set w = Nothing	' w is released
	Set df = Nothing	' df is released
End Sub

In this case, it is not necessary to set df and w to Nothing because this would happen automatically when they go out of scope, that is, when the subroutine ends.

In C++ this would be:

void Test(IApplication* pApplication)
{
	IDataFolder* pDataFolder = NULL;
	IWave* pWave = NULL;
	BSTR bstrDataFolderName = NULL;
	BSTR bstrWaveName = NULL;

	// Obtain reference to the current data folder.
	bstrDataFolderName = SysAllocString(L":");
	pApplication->DataFolder(bstrDataFolderName, &pDataFolder);
	SysFreeString(bstrDataFolderName);

	// Obtain reference to wave0 in that data folder.
	bstrWaveName = SysAllocString(L"wave0");
	pDataFolder->Wave(bstrWaveName, &pWave);
	SysFreeString(bstrWaveName);

	pWave->get_Name(&bstrWaveName);
	SysFreeString(bstrWaveName);

	pWave->Release();		// pWave is released.
	pDataFolder->Release();		// pDataFolder is released.
}

Unlike Visual Basic, in C++ you must always explicitly release wave and data folder references.

When the current Igor experiment is close, which occurs when a New Experiment is done or when another experiment is open, all waves and data folders are killed, even if you hold a reference to them. If you subsequently try to use one of these references, Igor will return an error.

Normally the current experiment will never be killed unless you (the client) does something to cause it to be killed. However, if Igor Pro is visible, which it normally will be when you are developing your client, the current experiment can be closed manually.

#	Known Problems
When the client program causes Igor to be launched, the client program becomes deactivated (loses focus) even though Igor is invisible at that point. This occurs when the operating systems sends a WM_USER message to a hidden window created in Igor by COM named "OleMainThreadWinName". We have been unable to find a way to prevent this from happening.
