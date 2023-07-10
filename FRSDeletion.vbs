Set WShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

NameDelete = InputBox("Please enter the string you want to delete.","File Rename Script Deletion (FRS Deletion)","")
CurrentDirectory = WShell.CurrentDirectory : CurrentDirectory = UCase(CurrentDirectory)
RenameFile FSO.GetFolder(CurrentDirectory)

Sub RenameFile(Folder)
Dim File
For Each File In Folder.Files
If File.Name <> "FRSDeletion.vbs" Then
FileName = FSO.GetBaseName(File.Name) : FileExtension = LCase(FSO.GetExtensionName(File.Name))
If InStr(FileName,NameDelete) <> 0 Then
  FileName = Replace(FileName,NameDelete,"")
  File.Name = FileName & "." & FileExtension
End If
End If
Next
End Sub
MsgBox "Done.",vbOKOnly,"File Rename Script Deletion (FRS Deletion)"
