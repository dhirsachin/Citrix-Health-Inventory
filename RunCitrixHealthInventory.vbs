Option Explicit

' Starts the companion CMD launcher without a visible console window.
' The WPF Citrix Health & Inventory window remains visible to the operator.
Dim fileSystem, shell, launcherPath, launcherLogPath, commandLine, exitCode

Set fileSystem = CreateObject("Scripting.FileSystemObject")
launcherPath = fileSystem.BuildPath( _
    fileSystem.GetParentFolderName(WScript.ScriptFullName), _
    "RunCitrixHealthInventory.cmd")

If Not fileSystem.FileExists(launcherPath) Then
    MsgBox "The launcher file was not found:" & vbCrLf & launcherPath, _
        vbCritical, "Citrix Health & Inventory"
    WScript.Quit 2
End If

Set shell = CreateObject("WScript.Shell")
' Start cmd.exe from a local Windows directory. The launcher and script still
' use absolute paths, so a release stored on UNC does not trigger cmd.exe's
' unsupported-UNC-current-directory warning.
shell.CurrentDirectory = shell.ExpandEnvironmentStrings("%SystemRoot%")
launcherLogPath = fileSystem.BuildPath( _
    shell.ExpandEnvironmentStrings("%LOCALAPPDATA%\CitrixDataPull\Logs"), _
    "CitrixHealthInventory_Launcher.log")
commandLine = Chr(34) & launcherPath & Chr(34) & " /quiet"

' Keep the console hidden but wait for the WPF process. On startup failure this
' lets the launcher return a reliable exit code and present the bootstrap log.
exitCode = shell.Run(commandLine, 0, True)
If exitCode <> 0 Then
    MsgBox "Citrix Health and Inventory could not start." & vbCrLf & vbCrLf & _
        "Exit code: " & CStr(exitCode) & vbCrLf & _
        "Launcher log:" & vbCrLf & launcherLogPath, _
        vbCritical, "Citrix Health & Inventory"
End If
WScript.Quit exitCode
