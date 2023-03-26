Sub CombineTextFiles()
Dim lFile As Long
Dim sFile As String
Dim vNewFile As Variant
Dim sPath As String
Dim sTxt As String
Dim sLine As String
With Application.FileDialog(msoFileDialogFolderPicker)
.AllowMultiSelect = False
If .Show Then
sPath = .SelectedItems(1)
If Right(sPath, 1) <> Application.PathSeparator Then
sPath = sPath & Application.PathSeparator
End If
Else
'Path cancelled, exit
Exit Sub
End If
End With
vNewFile = Application.GetSaveAsFilename("CombinedFile.txt", "Text files (*.txt), *.txt", , "Please enter the combined filename.")
If TypeName(vNewFile) = "Boolean" Then Exit Sub
sFile = Dir(sPath & "*.txt")
Do While Len(sFile) > 0
lFile = FreeFile
Open CStr(sFile) For Input As #lFile
Do Until EOF(lFile)
Line Input #1, sLine
sTxt = sTxt & vbNewLine & sLine
Loop
Close lFile
sFile = Dir()
Loop
lFile = FreeFile
Open CStr(vNewFile) For Output As #lFile
Print #lFile, sTxt
Close lFile
End Sub