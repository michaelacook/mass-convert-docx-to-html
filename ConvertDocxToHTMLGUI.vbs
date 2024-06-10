' Visual Basic script to convert docx files to filtered HTML files
' Author: Mike Cook <mcook@hrdownloads.com>

Const wdFormatFilteredHTML = 10
Const ForReading = 1
Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

' Display file picker dialog
Set objFolder = objShell.BrowseForFolder(0, "Select the folder containing .docx files to convert:", 0, "ITS")

If objFolder Is Nothing Then
    MsgBox "No folder selected. Conversion aborted.", 0, "ITS"
    WScript.Quit
End If

MsgBox "Starting conversion. This could take a couple minutes. You will be notified when complete.", 0, "ITS"

folderPath = objFolder.Self.Path
Set objFolder = objFSO.GetFolder(folderPath)

For Each objFile In objFolder.Files

    If LCase(objFSO.GetExtensionName(objFile.Path)) = "docx" Then

        Set objWord = CreateObject("Word.Application")
        
        Set objDoc = objWord.Documents.Open(objFile.Path)
		
		Do While objDoc.Comments.Count > 0
            objDoc.Comments(1).Delete
        Loop
		
		objDoc.AcceptAllRevisions
		
		objDoc.Save

        fileNameWithoutExt = objFSO.GetBaseName(objFile.Path)
        htmlFilePath = objFSO.BuildPath(objFolder.Path, fileNameWithoutExt & ".htm")
        objDoc.SaveAs2 htmlFilePath, wdFormatFilteredHTML

        objDoc.Close False

        objWord.Quit

        Set objWord = Nothing
        Set objDoc = Nothing

    End If
Next

Set objFolder = Nothing
Set objFSO = Nothing
Set objShell = Nothing

MsgBox "All files converted", 0, "ITS"
