' By Daniel Ferguson March 2016
' Purpose: Create a Movie webpage to lis and play local movies.

Option Explicit

Const ForReading = 1

Dim objShell, objFSO, objFile, objStream
Dim s, strFile, strDir, strExt

Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Prompt for a directory containing movies
strDir = selectFolder("Select your movie folder:", "")
If strDir = "" Then
    MsgBox("Movie page creation FAILED: Movies folder not selected.")
    Wscript.Quit
 End If

 
' Begin movie page html
s = "<!DOCTYPE html>" & vbCRLF
s = s & "<html>" & vbCRLF

' Create header with css and javascript
s = s & "<head>" & vbCRLF
s = s & "	<meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>" & vbCRLF
s = s & "	<style>" & vbCRLF
s = s & "		body { background: #333; }" & vbCRLF
s = s & "		h1 { color: white; }" & vbCRLF
s = s & "		.thumb { height: 180px;" & vbCRLF
s = s & "			width: 225px;" & vbCRLF
s = s & "			float: left;" & vbCRLF
s = s & "			margin: 10px;" & vbCRLF
s = s & "			vertical-align: middle;" & vbCRLF
s = s & "			display:inline-block; }" & vbCRLF
s = s & "		.thumb a { vertical-align: middle; }" & vbCRLF
s = s & "		.thumb img { border: 2px solid #444; width: 100%; vertical-align: middle; }" & vbCRLF
s = s & "	</style>" & vbCRLF
s = s & "	<script>" & vbCRLF

' Javascript creates links for each movie.
s = s & "		function addItem(strFile) {" & vbCRLF
s = s & "			try {" & vbCRLF
s = s & "				var strRoot = 'file:///" & Replace(strDir, "\", "/") & "/';" & vbCRLF
s = s & "				var strNoExt = strFile.substring(0,strFile.lastIndexOf('.'));" & vbCRLF
s = s & "				var oDiv = document.createElement('div');" & vbCRLF
s = s & "				var oMov = document.getElementById('movies');" & vbCRLF
s = s & "				oDiv.className = 'thumb';" & vbCRLF
s = s & "				oMov.appendChild(oDiv);" & vbCRLF
s = s & "				var oA = document.createElement('a');" & vbCRLF
s = s & "				oA.href = strRoot + strFile;" & vbCRLF
s = s & "				oDiv.appendChild(oA);" & vbCRLF
s = s & "				var oImg = document.createElement('img');" & vbCRLF
s = s & "				oA.appendChild(oImg);" & vbCRLF
s = s & "				oImg.src = 'file:///" & Replace(strDir, "\", "/") & "/thumbnails/' + strNoExt + '.png';" & vbCRLF
s = s & "				oImg.title = strNoExt;" & vbCRLF
s = s & "			} catch(err) {" & vbCRLF
s = s & "				alert('Error trying to load ' + strFile);" & vbCRLF
s = s & "			}" & vbCRLF
s = s & "		}" & vbCRLF
s = s & "	</script>" & vbCRLF
s = s & "	<title id='title'>Movies</title>" & vbCRLF
s = s & "</head>" & vbCRLF
s = s & "" & vbCRLF

' Begin html body with inline script which calls addItem for each movie.
s = s & "<body>" & vbCRLF
s = s & "" & vbCRLF
s = s & "<h1>Movies</h1>" & vbCRLF
s = s & "" & vbCRLF
s = s & "<div id='movies'>" & vbCRLF
s = s & "" & vbCRLF
s = s & "" & vbCRLF
s = s & "<script>" & vbCRLF

' Loop files in directory, create call for each movie file.
For Each objFile in objFSO.GetFolder(strDir).Files
    strExt = LCase(Right(objFile.Name,4))
    If strExt = ".mp4" OR strExt = ".m4v" OR strExt = ".mpg" _
        OR strExt = "mpeg" Then
        s = s & "	addItem(""" & objFile.Name & """);" & vbCRLF
    End If
Next

' Close html
s = s & "</script>" & vbCRLF
s = s & "</div>" & vbCRLF
s = s & "</body>" & vbCRLF
s = s & "</html>" & vbCRLF

' Save html string to file
strFile = "movies.html"
Set objStream = objFSO.CreateTextFile(strFile,True)
objStream.Write s
objStream.Close

' Open webpage in browser
objShell.Run strFile

' Display status message
MsgBox("Movie page created SUCCESSFULLY!")

' Cleanup objects
Set objStream = Nothing
Set objFSO = Nothing
Set objShell = Nothing

' Function prompts user to select a folder
Function selectFolder(strPrompt, strPath)
    Dim objFolder, objShell
    selectFolder = ""
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, strPrompt, 0, strPath)
    If Not objFolder Is Nothing Then selectFolder = objFolder.Self.Path
    Set objFolder = Nothing
    Set objShell = Nothing
End Function
