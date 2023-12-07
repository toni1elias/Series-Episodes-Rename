Set WShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
SeriesName = InputBox("Please enter the series name.","File Rename Script (FRS)","")
CurrentDirectory = WShell.CurrentDirectory : CurrentDirectory = UCase(CurrentDirectory)
RenameFile FSO.GetFolder(CurrentDirectory)

Sub RenameFile(Folder)
    Dim File, SeriesShortName, EpisodeShortName, SeriesEpisode, RES, REE
    Set RES = New RegExp
    Set REE = New RegExp
    RES.Global = True
    REE.Global = True
    RES.Pattern = "S[0-9][0-9]"
    REE.Pattern = "E[0-9][0-9]"
    For Each File In Folder.Files
        If File.Name <> "FRS.vbs" Then
            FileName = UCase(FSO.GetBaseName(File.Name)) : FileExtension = LCase(FSO.GetExtensionName(File.Name))
            Set SeriesMatches = RES.Execute(FileName)
            Set EpisodeMatches = REE.Execute(FileName)
            If SeriesMatches.Count > 0 And EpisodeMatches.Count > 0 Then
                Set SeriesMatche = SeriesMatches(0)
                Set EpisodeMatche = EpisodeMatches(0)
            Else
                WScript.Echo "Cannot find correct pattern. Please rename manually."
            End If
            SeriesSearchPosition = InStr(FileName , SeriesMatche)
            EpisodeSearchPosition = InStr(FileName , EpisodeMatche)
            SeriesShortName = Mid(FileName,SeriesSearchPosition,3)
            EpisodeShortName = Mid(FileName,EpisodeSearchPosition,3)
            SeriesEpisode = SeriesShortName & "-" & EpisodeShortName
            File.Name = SeriesEpisode & " " & SeriesName & "." & FileExtension
        End If
    Next
End Sub

Set WShell = Nothing
Set FSO = Nothing
Set RES = Nothing
Set REE = Nothing
Set SeriesMatches = Nothing
Set EpisodeMatches = Nothing
Set SeriesMatche = Nothing
Set EpisodeMatche = Nothing

MsgBox "Done.",vbOKOnly,"File Rename Script (FRS)"
