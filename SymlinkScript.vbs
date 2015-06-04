Set shell = CreateObject("Wscript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

title="Symlink Script"

path = Wscript.ScriptFullName
Set this = FSO.GetFile(path)
parent = FSO.GetParentFolderName(this)
Set dir = FSO.GetFolder(parent)
Set colFiles = dir.Files
Set colFolders = dir.SubFolders

newPath=InputBox("Target Directory Don't forget \ at the end",title)

For Each Folder in colFolders
	shell.run "cmd.exe /C mklink " + newPath + Folder.Name + " /D " + Folder.Path & Chr(34), 0
Next

For Each File in colFiles
	shell.run "cmd.exe /C mklink " + newPath + File.Name + " " + File.Path & Chr(34), 0
Next

WScript.Echo "Completed!"
