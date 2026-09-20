Option Explicit

Dim fso, reg, seenDict, scriptFolder, targetFolder, startTime
startTime = Timer

Set fso = CreateObject("Scripting.FileSystemObject")
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName) & "\"

' 1. Automatically resolve folder paths
If fso.FileExists(scriptFolder & "user_playlist.txt") Then
    targetFolder = scriptFolder
Else
    Dim toolboxFolder
    toolboxFolder = fso.GetParentFolderName(fso.GetParentFolderName(scriptFolder)) & "\"
    If fso.FolderExists(toolboxFolder & "playlist_check\") Then
        targetFolder = toolboxFolder & "playlist_check\"
    Else
        targetFolder = scriptFolder
    End If
End If

Set reg = New RegExp
reg.IgnoreCase = True
reg.Global = True

Set seenDict = CreateObject("Scripting.Dictionary")

' 2. Check input files
Dim userPath, ytPath
userPath = targetFolder & "user_playlist.txt"
ytPath = targetFolder & "youtube_playlist.txt"

If Not fso.FileExists(userPath) Or Not fso.FileExists(ytPath) Then
    WScript.Echo "ERROR: Files not found!"
    WScript.Echo "Expected location: " & targetFolder
    WScript.Echo "Ensure 'user_playlist.txt' and 'youtube_playlist.txt' are in that folder."
    WScript.Quit
End If

' 3. Read input files
Dim userContent, ytContent
userContent = ReadFile(userPath)
ytContent = ReadFile(ytPath)

Dim userLines, ytLines
userLines = Split(Replace(userContent, vbCrLf, vbLf), vbLf)
ytLines = Split(Replace(ytContent, vbCrLf, vbLf), vbLf)

' Calculate totals and initial estimates
Dim totalUserTracks, totalYTTracks, totalComparisons, estSeconds
totalUserTracks = UBound(userLines) + 1
totalYTTracks = UBound(ytLines) + 1
totalComparisons = totalUserTracks * totalYTTracks

estSeconds = Round(totalComparisons / 15000)
If estSeconds < 3 Then estSeconds = 3

WScript.Echo "Loaded: " & totalUserTracks & " local tracks | " & totalYTTracks & " YouTube tracks."
WScript.Echo "Initial estimate: ~" & estSeconds & " seconds"
WScript.Echo "--------------------------------------------------"
WScript.Echo "Processing playlist. This may take some time depending on its size..."
' 4. Parse User Playlist
Dim userCleanFull(), userCleanTitles(), userCount, i, line
userCount = 0
ReDim userCleanFull(UBound(userLines))
ReDim userCleanTitles(UBound(userLines))

For i = 0 To UBound(userLines)
    line = Trim(userLines(i))
    If line <> "" Then
        userCleanFull(userCount) = CleanText(line)
        userCleanTitles(userCount) = CleanText(GetTitlePart(line))
        userCount = userCount + 1
    End If
Next

' 5. Process YouTube Playlist entries with Live Countdown
Dim missingTracks, ytTrack, cleanYTFull, cleanYTTitle, countMissing
missingTracks = ""
countMissing = 0

Dim percentDone, elapsedNow, avgTimePerTrack, estRemainingSecs

For i = 0 To UBound(ytLines)
    ytTrack = Trim(ytLines(i))
    If ytTrack <> "" Then
        cleanYTFull = CleanText(ytTrack)
        cleanYTTitle = CleanText(GetTitlePart(ytTrack))
        
        ' Skip duplicate YouTube entries
        If Not seenDict.Exists(cleanYTFull) Then
            seenDict.Add cleanYTFull, True
            
            ' Check against user library
            If Not IsInUserPlaylist(cleanYTFull, cleanYTTitle, userCleanFull, userCleanTitles, userCount) Then
                missingTracks = missingTracks & ytTrack & vbCrLf
                countMissing = countMissing + 1
            End If
        End If
    End If

    ' Update live dynamic countdown every 3 tracks or on last track
    If (i Mod 3 = 0) Or (i = UBound(ytLines)) Then
        percentDone = Round(((i + 1) / totalYTTracks) * 100)
        elapsedNow = Timer - startTime
        If elapsedNow < 0 Then elapsedNow = elapsedNow + 86400

        If i > 0 Then
            avgTimePerTrack = elapsedNow / (i + 1)
            estRemainingSecs = Round(avgTimePerTrack * (totalYTTracks - (i + 1)))
        Else
            estRemainingSecs = estSeconds
        End If

        ' Overwrite current terminal line live
        WScript.StdOut.Write Chr(13) & "Progress: " & percentDone & "% | Timer: ~" & estRemainingSecs & "s remaining...   "
    End If
Next

' Move cursor to new line after loop finishes
WScript.StdOut.WriteLine ""

' 6. Calculate total actual time elapsed
Dim elapsedTime
elapsedTime = Round(Timer - startTime, 1)
If elapsedTime < 0 Then elapsedTime = elapsedTime + 86400

' 7. Output report
Dim outputPath
outputPath = targetFolder & "missing_tracks.txt"
WriteFile outputPath, missingTracks

WScript.Echo "--------------------------------------------------"
WScript.Echo "Done in " & elapsedTime & " seconds!"
WScript.Echo "Found " & countMissing & " truly missing tracks."
WScript.Echo "Report saved to: " & outputPath

' --- HELPER FUNCTIONS ---

Function GetTitlePart(str)
    If InStr(str, "-") > 0 Then
        GetTitlePart = Mid(str, InStr(str, "-") + 1)
    Else
        GetTitlePart = str
    End If
End Function

Function CleanText(str)
    Dim s
    s = LCase(str)
    reg.Pattern = "\([^\)]*\)"
    s = reg.Replace(s, "")
    reg.Pattern = "\[[^\]]*\]"
    s = reg.Replace(s, "")
    reg.Pattern = "\b(feat|ft|produced by|prod|official video|official audio|lyric video|lyrics)\b"
    s = reg.Replace(s, "")
    reg.Pattern = "[^a-z0-9\s]"
    s = reg.Replace(s, "")
    reg.Pattern = "\s+"
    s = Trim(reg.Replace(s, " "))
    CleanText = s
End Function

Function IsInUserPlaylist(ytFull, ytTitle, userFullList, userTitleList, userCount)
    IsInUserPlaylist = False
    If userCount = 0 Then Exit Function

    Dim uFull, uTitle, j
    For j = 0 To userCount - 1
        uFull = userFullList(j)
        uTitle = userTitleList(j)
        
        ' Direct Song Title Match
        If Len(ytTitle) >= 3 And Len(uTitle) >= 3 Then
            If ytTitle = uTitle Or InStr(ytTitle, uTitle) > 0 Or InStr(uTitle, ytTitle) > 0 Then
                IsInUserPlaylist = True
                Exit Function
            End If
        End If

        ' Full Line Substring Match
        If uFull <> "" Then
            If InStr(ytFull, uFull) > 0 Or InStr(uFull, ytFull) > 0 Then
                IsInUserPlaylist = True
                Exit Function
            End If
            
            ' Fuzzy match fallback
            If SimilarityRatio(ytFull, uFull) >= 0.72 Then
                IsInUserPlaylist = True
                Exit Function
            End If
        End If
    Next
End Function

Function SimilarityRatio(s1, s2)
    Dim len1, len2, maxLen, dist
    len1 = Len(s1)
    len2 = Len(s2)
    If len1 = 0 Or len2 = 0 Then
        SimilarityRatio = 0
        Exit Function
    End If
    If len1 > len2 Then maxLen = len1 Else maxLen = len2
    dist = LevenshteinDistance(s1, s2)
    SimilarityRatio = 1 - (dist / maxLen)
End Function

Function LevenshteinDistance(s1, s2)
    Dim m, n, i, j, cost
    m = Len(s1)
    n = Len(s2)
    Dim d()
    ReDim d(m, n)

    For i = 0 To m : d(i, 0) = i : Next
    For j = 0 To n : d(0, j) = j : Next

    For i = 1 To m
        For j = 1 To n
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            d(i, j) = Min3(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + cost)
        Next
    Next
    LevenshteinDistance = d(m, n)
End Function

Function Min3(a, b, c)
    Dim m
    m = a
    If b < m Then m = b
    If c < m Then m = c
    Min3 = m
End Function

Function ReadFile(path)
    On Error Resume Next
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    ReadFile = stream.ReadText
    stream.Close
End Function

Function WriteFile(path, text)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile path, 2
    stream.Close
End Function