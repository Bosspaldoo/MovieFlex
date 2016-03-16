' By Daniel Ferguson March 2016
' Purpose: Create a Movie webpage to lis and play local movies.

Option Explicit

Const ForReading = 1

Dim objShell, objFSO, objFile, objStream
Dim strResult, strFile

Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Create html string
strResult = "<html>" & vbCRLF
strResult = strResult & "<head></head>"  & vbCRLF
strResult = strResult & "<body>Hello World</body>" & vbCRLF
strResult = strResult & "</html>" & vbCRLF

'Save html string to file
strFile = "movies.html"
Set objStream = objFSO.CreateTextFile(strFile,True)
objStream.Write strResult
objStream.Close

'Open webpage in browser
objShell.Run strFile

Set objStream = Nothing
Set objFSO = Nothing
Set objShell = Nothing
