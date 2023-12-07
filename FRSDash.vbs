Set WShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
CurrentDirectory = WShell.CurrentDirectory : CurrentDirectory = UCase(CurrentDirectory)
RenameFile FSO.GetFolder(CurrentDirectory)

Sub RenameFile(Folder)
  Dim File, SeriesShortName, RES
  Set RES = New RegExp
  RES.Global = True
  RES.Pattern = "S[0-9][0-9]"
  For Each File In Folder.Files
    If File.Name <> "FRSDash.vbs" Then
      FileName = FSO.GetBaseName(File.Name) : FileExtension = LCase(FSO.GetExtensionName(File.Name))
      Set SeriesMatches = RES.Execute(FileName)
      If SeriesMatches.Count > 0 Then
        Set SeriesMatche = SeriesMatches(0)
      Else
        WScript.Echo "Cannot find correct pattern. Please rename manually."
      End If
      If InStr(FileName,SeriesMatche) <> 0 Then
        FileName = Replace(FileName,SeriesMatche,SeriesMatche & "-")
        File.Name = FileName & "." & FileExtension
      End If
    End If
  Next
End Sub

Set WShell = Nothing
Set FSO = Nothing
Set RES = Nothing
Set SeriesMatches = Nothing
Set SeriesMatche = Nothing

MsgBox "Done.",vbOKOnly,"File Rename Script Dash (FRS Dash)"
