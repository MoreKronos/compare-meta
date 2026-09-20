Set objArgs = WScript.Arguments
If objArgs.Count < 2 Then WScript.Quit

strMusicFolder = objArgs(0)
strOutputFile = objArgs(1)

Set objFSO = CreateObject("Scripting.FileSystemObject")
If Not objFSO.FolderExists(strMusicFolder) Then
    WScript.Echo "[ERROR] music_files folder not found."
    WScript.Quit
End If

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.NameSpace(strMusicFolder)
Set objOut = objFSO.CreateTextFile(strOutputFile, True, True)

' Find the Title column index dynamically
intTitleIdx = 21
For i = 0 To 299
    strHeader = objFolder.GetDetailsOf(Null, i)
    If strHeader = "Title" Or strHeader = "Titre" Or strHeader = "Título" Then
        intTitleIdx = i
        Exit For
    End If
Next

intCount = 0
For Each objItem In objFolder.Items
    strExt = LCase(objFSO.GetExtensionName(objItem.Path))
    If strExt = "mp3" Or strExt = "flac" Or strExt = "m4a" Or strExt = "wav" Or strExt = "ogg" Then
        ' Read ID3 Tag Title
        strTitle = objFolder.GetDetailsOf(objItem, intTitleIdx)
        
        ' Fallback to file name if no ID3 Title tag exists
        If Trim(strTitle) = "" Then
            strTitle = objFSO.GetBaseName(objItem.Path)
        End If
        
        objOut.WriteLine strTitle
        WScript.Echo " [+] " & strTitle
        intCount = intCount + 1
    End If
Next

objOut.Close
WScript.Echo ""
WScript.Echo "Successfully exported " & intCount & " titles."