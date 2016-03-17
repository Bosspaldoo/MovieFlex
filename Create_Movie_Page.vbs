' By Daniel Ferguson March 2016
' Purpose: Create a Movie webpage to lis and play local movies.

Option Explicit

Const ForReading = 1

Dim objShell, objFSO, objFile, objStream
Dim strResult, strFile, strDir, strExt

Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Prompt for a directory containing movies
strDir = selectFolder("Select your movie folder:", "")
If strDir = vbNull Then
    MsgBox("Movie page creation FAILED: Movies folder not selected.")
    Wscript.Quit
 End If

' Begin movie page html
strResult = "<html>" & vbCRLF
strResult = strResult & "<head></head>"  & vbCRLF
strResult = strResult & "<body>Hello World</body>" & vbCRLF
strResult = strResult & "<ul>" & vbCRLF

' Loop files in directory, create links to movie files.
For Each objFile in objFSO.GetFolder(strDir).Files
    strExt = LCase(Right(objFile.Name,4))
    If strExt = ".mp4" OR strExt = ".m4v"  Then
        strResult = strResult & "<li><a href=" & chr(34) & "file:///" & _
            strDir & "/" & objFile.Name & chr(34) & ">" & objFile.Name & _
            "</a></li>"
    End If
Next

' Close html
strResult = strResult & "</ul>" & vbCRLF
strResult = strResult & "</html>" & vbCRLF

' Save html string to file
strFile = "movies.html"
Set objStream = objFSO.CreateTextFile(strFile,True)
objStream.Write strResult
objStream.Close

' Open webpage in browser
objShell.Run strFile

MsgBox("Movie page created SUCCESSFULLY!")

Set objStream = Nothing
Set objFSO = Nothing
Set objShell = Nothing

Function selectFolder(strPrompt, strPath)
    Dim objFolder, objShell
    selectFolder = vbNull
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, strPrompt, 0, strPath)
    If IsObject(objfolder) Then selectFolder = objFolder.Self.Path
    Set objFolder = Nothing
    Set objShell = Nothing
End Function
