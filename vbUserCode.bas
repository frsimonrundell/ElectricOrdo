Option Compare Database



'============== PROCESS =================
' Every year you forget how to do this, so I am going to document my steps (and mis-steps)

' 1) Download Universalis for Month of Dec <current year> and Year <next year> as ePub
' 2) Rename ePub to Zip and unzip
' 3) clear("tblUniversalis") to Clesar tables but in tblUniuversalis create record with revised_date of "01/01/1970"
' 4) Run ProcessUniversalis to import to tblUniversalis
' 5) Download ics file from http://almanac.oremus.org/ and import to CW table
' 6) Run  modifyUniversalis to strip out unwanted characters
' 7) AppendUniversalis to insert into tblMass
' 8) AppendCW to insert CW
' 9) create_iCal to create the file

Global Const ForReading = 1, ForWriting = 2, ForAppending = 3
'===========================================================================
' DECLARATIONS
'===========================================================================
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)


'===========================================================================
' DATA GATHER FUNCTIONS
'===========================================================================

Sub processUniversalis()
' export the universalis to epub, change suffix to zip, unzip to desktop for processing

'Advent to Pentecost
processList = GetFileList("D:\Dropbox\Development\churchdates\2023\data0", "u*.xhtml")

For Each FileName In processList

    ExtractUniversalis (FileName)
    
Next

'Pentecost to Advent
processList = GetFileList("D:\Dropbox\Development\churchdates\2023\data", "u*.xhtml")

For Each FileName In processList

    ExtractUniversalis (FileName)
    
Next

Debug.Print "Process Universalis Completed"
Call modifyUniversalis

End Sub


'===========================================================================
' XML MANIPULATION FUNCTIONS
'===========================================================================



Sub ExtractUniversalis(sSourceFile As String)

Dim ret As Boolean, CurrentDate As String, sLectYear As String, sPsalmWeek As String
Dim sfestival As String, sColour As String
Dim sInvitatory As String, sMorningPrayer As String, sEveningPrayer As String
Dim sLiturgy As String, sReadings As String, sentrance As String
Dim scollect As String, sfirstreading As String, sResppsalm As String, ssecondreading As String
Dim sgospelacclamation As String, sgospel As String, sprayeroff As String
Dim scommant As String, spostcomm As String
Dim tmpCalDate As String, revised_date As String
Dim fileString As String

If sSourceFile = "" Then
    Debug.Print "No Source File"
    Exit Sub
End If

'Opens the file for input
Open sSourceFile For Input As #1

                Dim dbs As DAO.Database
                Dim rst As DAO.Recordset
    
                'Get the database and Recordset
                Set dbs = CurrentDb
                Set rst = dbs.OpenRecordset("tblUniversalis")
                
                fileString = ""

While Not EOF(1)
   Line Input #1, LineString
   fileString = fileString & LineString
Wend

'Closes the text file
Close #1

sSourceFileName = PullText(sSourceFile, "OEBPS\", ".xhtml")

Debug.Print "=== Processing " & sSourceFileName & " ==="
   
   
   If sContenttype = "" Then sContenttype = PullText(fileString, "<title>", "</title>")
   
      'Debug.Print "File: " & sSourceFile

   
   Select Case sContenttype
   
   
   Case "About Today"
        CurrentDate = PullText(fileString, "Table of Contents</a><h2>", "</h2>")
        If sLectYear = "" Then sLectYear = PullText(fileString, "Year: ", "Psalm week: ")
        If sPsalmWeek = "" Then sPsalmWeek = PullText(fileString, "Psalm week: ", "Liturgical Colour:")
        If sfestival = "" Then sfestival = PullText(fileString, "<strong>", "About Today</h2>")
        If sColour = "" Then sColour = PullText(fileString, "Liturgical Colour: ", "</i></p>")
        tmpCalDate = Mid(CurrentDate, InStr(1, CurrentDate, "day") + 4)
        revised_date = CDate(tmpCalDate)
        
        If sPsalmWeek = "" Then sPsalmWeek = 0
        
       
       Debug.Print "About Today: " & CurrentDate & " Year " & sLectYear & " Psalm Week " & sPsalmWeek & " Festival: " & sfestival
        

                found = False
                
                'Search for the first matching record
                rst.MoveFirst
                    
                Do Until rst.EOF

                If CDate(rst.Fields("revised_date")) = CDate(revised_date) Then
                    
                    Debug.Print "About Today (Update): " & revised_date
                    ret = appendFieldtoRst(rst, "fileref", sSourceFileName)
                    ret = updateFieldtoRst(rst, "lectionaryyear", sLectYear)
                    ret = updateFieldtoRst(rst, "colour", sColour)
                    ret = updateFieldtoRst(rst, "festival", sfestival)
                    
                    found = True
                    
                End If

                rst.MoveNext
                Loop
                
                If found = False Then
                    Debug.Print "About Today (New): " & revised_date
                    rst.AddNew
                    ret = addFieldtoRst(rst, "revised_date", revised_date)
                    ret = addFieldtoRst(rst, "lectionaryyear", sLectYear)
                    ret = addFieldtoRst(rst, "office", sPsalmWeek)
                    ret = addFieldtoRst(rst, "colour", sColour)
                    ret = addFieldtoRst(rst, "festival", sfestival)
                    rst.Update
                    
                    found = True
                End If
                
    Case "Invitatory Psalm"
        If CurrentDate = "" Then CurrentDate = PullText(fileString, "Table of Contents</a><h2>", "</h2>")
        tmpCalDate = Mid(CurrentDate, InStr(1, CurrentDate, "day") + 4)
        revised_date = CDate(tmpCalDate)
        If sfestival = "" Then sfestival = PullText(fileString, "<strong>", "About Today</h2>")
        If sInvitatory = "" Then sInvitatory = Trim(PullText(fileString, "Invitatory Psalm</h2>", "</body></html>"))
        
        'Search for the first matching record
                rst.MoveFirst
                    
                Do Until rst.EOF

                If CDate(rst.Fields("revised_date")) = CDate(revised_date) Then
                    
                    Debug.Print "Invitatory (Update): " & revised_date
                    ret = updateFieldtoRst(rst, "invitatory", sInvitatory)
                    
                    found = True
                    
                End If

                rst.MoveNext
                Loop
                
                If found = False Then
                    Debug.Print "Invitatory (Add): " & revised_date
                    rst.AddNew
                    ret = addFieldtoRst(rst, "revised_date", revised_date)
                    ret = addFieldtoRst(rst, "invitatory", sInvitatory)
                    rst.Update
                    
                    found = True
                End If
    
        Case "Morning Prayer (Lauds)"
        
            If CurrentDate = "" Then CurrentDate = PullText(fileString, "Table of Contents</a><h2>", "</h2>")
            tmpCalDate = Mid(CurrentDate, InStr(1, CurrentDate, "day") + 4)
            revised_date = CDate(tmpCalDate)
            If sfestival = "" Then sfestival = PullText(fileString, "<strong>", "About Today</h2>")
            If sMorningPrayer = "" Then sMorningPrayer = regexOfficeReadings(PullText(fileString, "Morning Prayer (Lauds)", "</body></html>"))
            
            'Search for the first matching record
                    rst.MoveFirst
                        
                    Do Until rst.EOF
    
                    If CDate(rst.Fields("revised_date")) = CDate(revised_date) Then
                        
                        Debug.Print "MP (Update): " & revised_date
                        ret = appendFieldtoRst(rst, "fileref", sSourceFileName)
                        ret = updateFieldtoRst(rst, "morningprayer", sMorningPrayer)
                        ret = updateFieldtoRst(rst, "festival", sfestival)
                        
                        found = True
                        
                    End If
    
                    rst.MoveNext
                    Loop
                    
                    If found = False Then
                    Debug.Print "MP (Add): " & revised_date
                        rst.AddNew
                        ret = addFieldtoRst(rst, "revised_date", revised_date)
                        ret = addFieldtoRst(rst, "morningprayer", sMorningPrayer)
                        ret = addFieldtoRst(rst, "festival", sfestival)
                        rst.Update
                       
                        found = True
                    End If
            
        Case "Vespers (Evening Prayer)"
            If CurrentDate = "" Then CurrentDate = PullText(fileString, "Table of Contents</a><h2>", "</h2>")
            tmpCalDate = Mid(CurrentDate, InStr(1, CurrentDate, "day") + 4)
            revised_date = CDate(tmpCalDate)
            If sfestival = "" Then sfestival = PullText(fileString, "<strong>", "About Today</h2>")
            If sEveningPrayer = "" Then sEveningPrayer = regexOfficeReadings(PullText(fileString, "Vespers (Evening Prayer)</h2>", "</body></html>"))
            
            'Search for the first matching record
                    rst.MoveFirst
                        
                    Do Until rst.EOF
    
                    If CDate(rst.Fields("revised_date")) = CDate(revised_date) Then
                    
                    Debug.Print "EP (Update): " & revised_date
                        
                        ret = appendFieldtoRst(rst, "fileref", sSourceFileName)
                        ret = updateFieldtoRst(rst, "eveningprayer", sEveningPrayer)
                        ret = updateFieldtoRst(rst, "festival", sfestival)
                        
                        found = True
                        
                    End If
    
                    rst.MoveNext
                    Loop
                    
                    If found = False Then
                    Debug.Print "EP (Add): " & revised_date
                        rst.AddNew
                        ret = addFieldtoRst(rst, "revised_date", revised_date)
                        ret = addFieldtoRst(rst, "eveningprayer", sEveningPrayer)
                        ret = addFieldtoRst(rst, "festival", sfestival)
                        rst.Update
                        
                        found = True
                    End If


' Mass Today seems no longer to be a page, using readings at mass instead

    Case "Readings at Mass"
        If CurrentDate = "" Then CurrentDate = PullText(fileString, "Table of Contents</a><h2>", "</h2>")
            If CurrentDate <> "" Then
                tmpCalDate = Mid(CurrentDate, InStr(1, CurrentDate, "day") + 4)
                revised_date = CDate(tmpCalDate)
                sDay = twoDigit(Day(revised_date))
                sMonth = twoDigit(Month(revised_date))
                sYear = Year(revised_date)
                cal_date = sYear & sMonth & sDay
            End If
            
        If sfestival = "" Then sfestival = PullText(fileString, "<strong>", "About Today")
        If sReadings = "" Then sReadings = PullText(fileString, "<h2>Readings at Mass</h2>", "</body></html>")
        If sColour = "" Then sColour = PullText(fileString, "Liturgical Colour: ", "</i></p>")
        If sentrance = "" Then sentrance = separateReadings(PullText(fileString, "Entrance Antiphon</th>", "<hr class="))
        If scollect = "" Then scollect = PullText(fileString, "Collect</th>", "<hr class=")
        
        Debug.Print scollect
        
        If sfirstreading = "" Then sfirstreading = regexReadings(PullText(fileString, "First reading</th>", "<hr class="))
                
                If InStr(1, fileString, "Second Reading") = 0 Then
                    If sResppsalm = "" Then sResppsalm = regexReadings(PullText(fileString, "Responsorial Psalm</th>", "<hr class="))
                Else
                    If sResppsalm = "" Then sResppsalm = regexReadings(PullText(fileString, "Responsorial Psalm</th>", "<hr class="))
                    If ssecondreading = "" Then ssecondreading = regexReadings(PullText(fileString, "Second reading</th>", "Gospel Acclamation")) ' all vigil readings to go in here
                End If
                
         If sgospelacclamation = "" Then sgospelacclamation = separateReadings(PullText(fileString, "Gospel Acclamation</th>", "<hr class="))
         If sgospel = "" Then sgospel = regexReadings(PullText(fileString, "Gospel</th>", "<hr class="))
         If sprayeroff = "" Then sprayeroff = PullText(fileString, "Prayer over the Offerings</th>", "<hr class=")
         If scommant = "" Then scommant = separateReadings(PullText(fileString, "Communion Antiphon</th>", "<hr class="))
         If spostcomm = "" Then spostcomm = PullText(fileString, "Prayer after Communion</th>", "<hr class=")
                

                 
                Debug.Print "========================================================================="
                
                Debug.Print "Festival:" & sfestival
                Debug.Print "Entrance Antiphon: " & sentrance
                Debug.Print "Collect: " & scollect
                Debug.Print "1st Reading: " & sfirstreading
                Debug.Print "Psalm: " & sResppsalm
                Debug.Print "2nd Reading: " & ssecondreading
                Debug.Print "Gospel Acclamation: " & sgospelacclamation
                Debug.Print "Gospel: " & sgospel
                Debug.Print "Offertory: " & sprayeroff
                Debug.Print "Communion Antiphon: " & scommant
                Debug.Print "Post Communion: " & spostcomm
                
                
                'Search for the first matching record
                rst.MoveFirst
                found = False
                    
                Do Until rst.EOF

                If CDate(rst.Fields("revised_date")) = CDate(revised_date) Then
                Debug.Print "Readings at Mass (Update): " & CurrentDate & ": Converts to: " & sDay & "/" & sMonth & "/" & sYear & " -> " & cal_date
                    ret = appendFieldtoRst(rst, "fileref", sSourceFileName)
                    ret = updateFieldtoRst(rst, "festival", sfestival)
                    ret = updateFieldtoRst(rst, "full_date", CurrentDate)
                    ret = updateFieldtoRst(rst, "revised_date", revised_date)
                    ret = updateFieldtoRst(rst, "cal_date", cal_date)
                    ret = updateFieldtoRst(rst, "day", sDay)
                    ret = updateFieldtoRst(rst, "month", sMonth)
                    ret = updateFieldtoRst(rst, "year", sYear)
                    ret = updateFieldtoRst(rst, "colour", sColour)
                    ret = updateFieldtoRst(rst, "collect", scollect)
                    ret = updateFieldtoRst(rst, "entrance", BibleNames(sentrance))
                    ret = updateFieldtoRst(rst, "firstreading", sfirstreading)
                    ret = updateFieldtoRst(rst, "psalm", sResppsalm)
                    ret = updateFieldtoRst(rst, "secondreading", ssecondreading)
                    ret = updateFieldtoRst(rst, "gospelacclamation", BibleNames(sgospelacclamation))
                    ret = updateFieldtoRst(rst, "gospel", sgospel)
                    ret = updateFieldtoRst(rst, "potg", sprayeroff)
                    ret = updateFieldtoRst(rst, "communionantiphon", BibleNames(scommant))
                    ret = updateFieldtoRst(rst, "postcommunion", spostcomm)
                    found = True
                    
                End If

                rst.MoveNext
                Loop
                
                If found = False Then
                Debug.Print "Readings at Mass (Add): " & CurrentDate & ": Converts to: " & sDay & "/" & sMonth & "/" & sYear & " -> " & cal_date
                rst.AddNew
                    ret = addFieldtoRst(rst, "fileref", sSourceFileName)
                    ret = addFieldtoRst(rst, "festival", sfestival)
                    ret = addFieldtoRst(rst, "full_date", CurrentDate)
                    ret = addFieldtoRst(rst, "revised_date", revised_date)
                    ret = addFieldtoRst(rst, "cal_date", cal_date)
                    ret = addFieldtoRst(rst, "day", sDay)
                    ret = addFieldtoRst(rst, "month", sMonth)
                    ret = addFieldtoRst(rst, "year", sYear)
                    ret = addFieldtoRst(rst, "colour", sColour)
                    ret = addFieldtoRst(rst, "entrance", sentrance)
                    ret = addFieldtoRst(rst, "firstreading", sfirstreading)
                    ret = addFieldtoRst(rst, "psalm", sResppsalm)
                    ret = addFieldtoRst(rst, "secondreading", ssecondreading)
                    ret = addFieldtoRst(rst, "gospelacclamation", sgospelacclamation)
                    ret = addFieldtoRst(rst, "gospel", sgospel)
                    ret = addFieldtoRst(rst, "potg", sprayeroff)
                    ret = addFieldtoRst(rst, "communionantiphon", scommant)
                    ret = addFieldtoRst(rst, "postcommunion", spostcomm)
                rst.Update
                    found = True
                    
                End If
                
                
        
        
        
    End Select
    
    ' pause so we can see what is going on...
    ' WaitSeconds (1)
    DoEvents

        rst.Close
        Set rst = Nothing
        Set dbs = Nothing
                

End Sub


'===========================================================================
' OUTPUT FUNCTIONS
'===========================================================================

Sub create_iCal()
' create iCalendar (Version 2)
' create according to RFC 2445 Specification
' with some mods as discovered by trial and error
' by Fr. Simon

Dim db As Database, rs As DAO.Recordset, strSQL As String
Set db = CurrentDb
Dim strFilename As String

' For monthcount = 1 To 12

    ' strFilename = "churchdates" & (Year(Now()) + 1) & "_" & monthcount & ".ics"
    
    strFilename = "churchdates" & (Year(Now()) + 1) & ".ics"
    
    Debug.Print "Writing: " & strFilename
    
    
    strSQL = "SELECT * FROM tblMass"
    Set rs = db.OpenRecordset(strSQL)
      
     
    strOutput = "BEGIN:VCALENDAR" & vbCrLf
    strOutput = strOutput & "PRODID:-//Google Inc//Google Calendar 70.9054//EN" & vbCrLf
    strOutput = strOutput & "VERSION:2.0" & vbCrLf
    strOutput = strOutput & "CALSCALE:GREGORIAN" & vbCrLf
    strOutput = strOutput & "METHOD:PUBLISH" & vbCrLf
    
    Do While Not rs.EOF
    
    DoEvents ' prevent freezes
    
    Debug.Print "Writing " & rs.Fields("cal_date")
   
    strOutput = strOutput & "BEGIN:VEVENT" & vbCrLf
    strOutput = strOutput & "DTSTAMP:" & return_today() & "T000001Z" & vbCrLf
    strOutput = strOutput & "DTSTART;VALUE=DATE:" & rs.Fields("cal_date") & vbCrLf
    strOutput = strOutput & "DTEND;VALUE=DATE:" & OneMore(rs.Fields("cal_date")) & vbCrLf
    strOutput = strOutput & "ORGANIZER;CN=www.bsp.church:MAILTO:null@null.com" & vbCrLf
    strOutput = strOutput & "SUMMARY:" & rs.Fields("festival") & vbCrLf
    
    strOutput = strOutput & "DESCRIPTION:" & cr_to_slashn(createdescription(rs)) & vbCrLf
                            
    'strOutput = strOutput & "UID:" & "08101967@rundell.org.uk" & vbCrLf
    strOutput = strOutput & "STATUS:CONFIRMED" & vbCrLf
    strOutput = strOutput & "TRANSP:TRANSPARENT" & vbCrLf
    strOutput = strOutput & "X-MICROSOFT-CDO-BUSYSTATUS:FREE" & vbCrLf
    strOutput = strOutput & "X-MICROSOFT-CDO-INSTTYPE:0" & vbCrLf
    strOutput = strOutput & "X-MICROSOFT-CDO-INTENDEDSTATUS:FREE" & vbCrLf
    strOutput = strOutput & "X-MICROSOFT-CDO-ALLDAYEVENT:TRUE" & vbCrLf
    strOutput = strOutput & "X-MICROSOFT-CDO-IMPORTANCE:1" & vbCrLf
    strOutput = strOutput & "END:VEVENT" & vbCrLf
    
    rs.MoveNext
    Loop
    
    strOutput = strOutput & "END:VCALENDAR"
    
    rs.Close
    Set rs = Nothing
    
   
    'now write this string object to file
       
        If Right$(CurDir$(), 1) <> "\" Then
            sFile = CurDir$() & "\" & strFilename
        Else
             sFile = CurDir$() & strFilename
        End If
        
    
    Open sFile For Output As #1
    Print #1, strOutput
    Close #1

Debug.Print "File Complete and Ready at " & sFile

'Next monthcount

End Sub

Function createdescription(rs As Recordset)

strOutput = ""

strOutput = strOutput & rs.Fields("CWFestival") & "\n\nToday is a " & rs.Fields("festivaltype") & " Colour: " & Trim(rs.Fields("colour")) & " " & Trim(rs.Fields("gloria")) & " " & Trim(rs.Fields("creed")) & " " & _
                        "Lectionary Year: " & Trim(rs.Fields("lectionary_year")) & " " & _
                        "Divine Office Week:" & Trim(rs.Fields("office")) & "\n\n"
                        
                        
strOutput = strOutput & "Readings:\n" & _
                       "---------\n\n" & Trim(rs.Fields("MassR")) & " \n\n" & _
                       "Anglican Lectionary\n================\n\n" & Trim(rs.Fields("CW")) & "\n\n" & _
                       Trim(rs.Fields("comment"))

strOutput = strOutput & "\n\n\nCompiled from a variety of Liturgical Sources for his own use by Fr. Simon Rundell SCP\n" & _
                        "(simon@rundell.org.uk) and made available to others as PRAYERWARE: Available without cost or charge except that your prayers are asked for Fr. Simon and the Parishes of Bickleigh and Shaugh Prior, Plymouth\n" & _
                        "This edition: Advent 2022 to Advent 2023. See http://www.bsp.church\n\n"


createdescription = strOutput

End Function


'===========================================================================
' STRING FUNCTIONS
'===========================================================================

Function getDOW(intDOW)

Select Case intDOW
    Case 1:
        getDOW = "Sunday"
    Case 2:
        getDOW = "Monday"
    Case 3:
        getDOW = "Tuesday"
    Case 4:
        getDOW = "Wednesday"
    Case 5:
        getDOW = "Thursday"
    Case 6:
        getDOW = "Friday"
    Case 7:
        getDOW = "Saturday"
End Select
End Function

Function StripHTMLTags(ByVal HTML As String) As String
Dim a() As String
Dim v As Variant
a() = Split(HTML, "<")
For Each v In a
    StripHTMLTags = StripHTMLTags & Mid$(v, InStr(v, ">") + 1)
Next v
End Function

Function breakdown(s As String)
' debugging function

ans = ""
For n = 1 To Len(s)
    MyChar = Mid(s, n, 1)
    ans = ans & Asc(MyChar) & "-> " & MyChar & " || "
Next

'MsgBox (ans)
Debug.Print ans

breakdown = ans

End Function

Function cr_to_slashn(strReading As Variant)
strTemp = Replace(strReading, Chr(13), "\n")
strTemp = Replace(strTemp, Chr(10), "")
cr_to_slashn = strTemp
End Function

Function cr_to_nothing(strReading As Variant)
strTemp = Replace(strReading, Chr(13), "")
strTemp = Replace(strTemp, Chr(10), "")
cr_to_nothing = strTemp
End Function


Function OneMore(dtDate As String)
' add date to yyyymmdd

d = CInt(Mid(dtDate, 7, 2))
m = CInt(Mid(dtDate, 5, 2))
y = CInt(Mid(dtDate, 1, 4))

strOutput = return_date(DateAdd("d", 1, CDate(d & "/" & m & "/" & y)))

OneMore = strOutput

End Function

Function return_today()

d = CStr(DatePart("d", Now))
m = CStr(DatePart("m", Now))
y = DatePart("yyyy", Now)

If Len(d) = 1 Then d = "0" & d
If Len(m) = 1 Then m = "0" & m

return_today = y & m & d
End Function

Function return_date(dtDate As Date)

d = CStr(DatePart("d", dtDate))
m = CStr(DatePart("m", dtDate))
y = DatePart("yyyy", dtDate)

If Len(d) = 1 Then d = "0" & d
If Len(m) = 1 Then m = "0" & m

return_date = y & m & d

End Function

Function splitDate(sDate As String) As String
'take yyyymmdd and return dd/mm/yyyy

y = Left(sDate, 4)
m = Mid(sDate, 5, 2)
d = Mid(sDate, 7)

splitDate = d & "/" & m & "/" & y

End Function

Private Function LeadingZeros(ExpectedLen As Integer, ActualLen As Integer) As String
' Pad out with leading zeros
      LeadingZeros = String$(ExpectedLen - ActualLen, "0")
End Function
      
Function twoDigit(MyInt As String)

    If Len(MyInt) < 2 Then
        twoDigit = "0" & MyInt
    Else
        twoDigit = MyInt
    End If

End Function

Public Function sDate(UnixDateStamp As Double)
    sDate = DateAdd("s", UnixDateStamp, #1/1/1970#)
End Function

Public Function uDate(NormalDate As Date)
   uDate = DateDiff("s", #1/1/1970#, NormalDate)
End Function



Function replaceCRs(strFieldName)
'work through a string and convert vbcrlf to \n

strNewField = ""
If Not IsNull(strFieldName) Then
    For n = 1 To Len(strFieldName)
        strChar = Mid(strFieldName, n, 1)
        
        If Asc(strChar) = 13 Then strChar = "\n"  'together this represents the \n token for a carriage return
        If Asc(strChar) = 10 Then strChar = ""

        'If Asc(strChar) = 34 Then strChar = " " ' convert double quotes
        'If Asc(strChar) = 39 Then strChar = " " ' convert single quotes
        strNewField = strNewField & strChar
    Next n
End If

replaceCRs = strNewField

End Function



Function StripHTML(strString As String) As String
 Dim regex As Object
 Set regex = CreateObject("vbscript.regexp")

 Dim sInput As String
 Dim sOut As String
 sInput = Replace(strString, "<\\", "\\")

 With regex
    .Global = True
    .IgnoreCase = True
    .MultiLine = True
    .Pattern = "<[^>]+>" 'Regular Expression for HTML Tags.
 End With

 sOut = regex.Replace(sInput, "")
 StripHTML = Replace(Replace(Replace(sOut, "&nbsp;", vbCrLf, 1, -1), "&quot;", "'", 1, -1), "\\", "<\\", 1, -1)
 Set regex = Nothing
End Function



Function PullText(sLine, sStartTag, sEndTag) As String

Dim p1 As Integer
Dim p2 As Integer

If IsNull(sLine) Then
    PullText = ""

Else
    p1 = InStr(1, sLine, sStartTag, vbTextCompare) + Len(startTag)
    If p1 <> 0 And sEndTag <> "#EOF" Then p2 = InStr(p1, sLine, sEndTag, vbBinaryCompare) ' case sensitive from the NEXT occcurance
    If sEndTag = "#EOF" Then p2 = Len(sLine)
    
    If (p1 = 0) Or (p2 = 0) Then
            PullText = ""
    Else
            p1 = p1 + Len(sStartTag)
            sLine = Mid(sLine, p1, p2 - p1)
            sLine = Replace(sLine, vbLf, "")
            sLine = Replace(sLine, vbCr, "")
            sLine = Replace(sLine, vbTab, "")
            PullText = sLine
            
    End If
End If

End Function



Function replaceVClass(sString) As String

sString = Replace(sString, Chr(34), "|")

sSearch = "<div class=|v|>"

    p1 = InStr(1, sString, sSearch)
    
    If p1 = 0 Then
        replaceVClass = Replace(sString, "|", Chr(34))
    Else
        
        'Debug.Print "*** FOUND VCLASS ***"
        sString = Replace(sString, sSearch, "")
        sString = Replace(sString, "</div>", "\n")
        replaceVClass = Replace(sString, "|", Chr(34))

    End If

End Function

Function PullText2(sLine As String, sStartTag As String, iStart As Integer, sEndTag As String)
' select the istart occurance of sStartTag


'count occurence of sStartTag in sLine
Dim p1, p2 As Integer
Dim iCount As Integer
Dim sParts() As String

iCount = CountChrInString(sLine, sStartTag)
sParts = Split(sLine, sStartTag)

If iStart > iCount Then iStart = iCount

If iCount > 0 Then sLine = sParts(iStart)   ' otherwise just return string

p1 = 0

If sEndTag = "#EOF" Then
    p2 = Len(sLine)
Else
    p2 = InStr(1, sLine, sEndTag, vbTextCompare)
End If

If p2 = 0 Then
        PullText2 = ""
Else
        p1 = p1 + Len(sStartTag)
        raw = StripHTML(Mid(sLine, p1 - 7, p2))
        raw = Replace(raw, "EITHER:", "\n" & "EITHER: " & "\n\n")
        raw = Replace(raw, "OR:", "\n" & "OR: " & "\n\n")
        'Debug.Print "-----------> Extracted " & raw & " -> " & Len(raw) & " <------------"
        PullText2 = raw
End If

End Function

Function PullText3(sLine As String, sStartTag As String, sEndTag As String)
' select the LAST occurance of sStartTag


'count occurence of sStartTag in sLine
Dim p1, p2 As Integer
Dim iCount As Integer
Dim sParts() As String

iCount = CountChrInString(sLine, sStartTag)
sParts = Split(sLine, sStartTag)

sLine = sParts(iCount)

p1 = 0
If p1 <> 0 And sEndTag <> "#EOF" Then p2 = InStr(0, sLine, sEndTag, vbTextCompare)
If sEndTag = "#EOF" Then p2 = Len(sLine)

If p2 = 0 Then
        PullText3 = ""
Else
        p1 = p1 + Len(sStartTag)
        raw = StripHTML(Mid(sLine, p1, p2))
        raw = Replace(raw, "EITHER:", "\n" & "EITHER: " & "\n\n")
        raw = Replace(raw, "OR:", "\n" & "OR: " & "\n\n")
        'Debug.Print "-----------> Extracted " & raw & " -> " & Len(raw) & " <------------"
        PullText3 = raw
End If

End Function

Public Function CountChrInString(Expression As String, Character As String) As Long
' credit to https://stackoverflow.com/questions/9260982/how-to-find-number-of-occurences-of-slash-from-a-strings

    Dim iResult As Long
    Dim sParts() As String

    sParts = Split(Expression, Character)

    iResult = UBound(sParts, 1)

    If (iResult = -1) Then
    iResult = 0
    End If

    CountChrInString = iResult

End Function


Public Function GetFileList(ByVal ParentDir As String, Optional ByVal sSearch As String) As Variant

' build a list of files in a folder to process

If sSearch = "" Then sSearch = "*.*"

Dim Counter As Long

'Create a dynamic array variable, and then declare its initial size
Dim DirectoryListArray() As String
ReDim DirectoryListArray(10000)

'Loop through all the files in the directory by using Dir$ function
MyFile = Dir$(ParentDir & "\" & sSearch)
Do While MyFile <> ""
    DirectoryListArray(Counter) = ParentDir & "\" & MyFile
    MyFile = Dir$
    Counter = Counter + 1
Loop

If Counter = 0 Then Counter = 1

ReDim Preserve DirectoryListArray(Counter - 1)

GetFileList = DirectoryListArray

End Function


Function removeText(sText, sStart, sEnd) As String
'removge from a string everything between sStart and sEnd (including)

p1 = InStr(sText, sStart)
p2 = InStr(sText, sEnd)

    If p2 <> 0 Then p2 = p2 + Len(sEnd)

    s1 = Left(sText, p1 - 1)
    s2 = Mid(sText, p2)  ' If p2 is not found returns the whole string again.
        
removeText = s1 & s2

End Function


Sub removeOddChars(strField, strDB)

Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
        
    'Get the database and Recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset(strDB, dbOpenTable, dbDenyRead, dbPessimistic)
    
    rst.MoveFirst

            Do Until rst.EOF
            If Not IsNull(rst.Fields(strField)) Then
                rst.Edit
                    rst.Fields(strField) = Replace(rst.Fields(strField), "&#8216;", "'") ' use single instead of double quotes
                    rst.Fields(strField) = Replace(rst.Fields(strField), "&#8217;", "'")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "&#8220;", "'")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "&#8221;", "'")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "&#160;", " ")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "&#225;", "a")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "&#281;", "a")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "&#7841;", "a")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "&#233;", "e")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "&#232;", "e")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "&#361;", "u")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "&#x2013;", " ")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "orAlleluia", "\nor\n\nAlleluia.\n\n")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Alleluia!Or", "\n\nAlleluia!\n\nOr\n\n")
                    rst.Fields(strField) = Replace(rst.Fields(strField), ".", "\n\n")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "INTRODUCTION", "", 1, -1, vbBinaryCompare)
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Hymn", "\nHymn\n\n")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "antiphon*)", "antiphon*)\n\n")
                    rst.Fields(strField) = Replace(rst.Fields(strField), Chr(34), "") ' extraneous double quotes
                    rst.Fields(strField) = Replace(rst.Fields(strField), "/>", "")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Contents", "")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "\n the 1\n", "")
                    
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Lentor", "Lent or")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Timeor", "Lent or")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Easteror", "Easter or")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Christmasor", "Christmas or")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Adventor", "Advent or")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Religiousor", "Religious or")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Adventor", "Advent or")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Eastertideor", "Easter or")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Lentor", "Lent or")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Christmasor", "Christmas or")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Timeor", "Time or")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Eastertide", "Easter")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "&#x2013;&#160;", " - ")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Popeor", "Pope or")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Doctoror", "Doctor or")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Martyror", "Martyr or")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Religiousor", "Religious or")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Priestor", "Priest or")
                    
                                    
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Canticle", "")
                    rst.Fields(strField) = Replace(rst.Fields(strField), " Reading:", "****")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Reading", "")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "****", " Reading:")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Apocalypse", "Revelation")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Gospel", "Gospel: ")
                    rst.Fields(strField) = Replace(rst.Fields(strField), "Gospel: :", "Gospel: ")
                    
                    'replace vbcrlf to \n
                    rst.Fields(strField) = replaceCRs(rst.Fields(strField))
              
                             
                rst.Update
            End If
                rst.MoveNext
            Loop
        
        rst.Close
        Set rst = Nothing
        Set dbs = Nothing
        
        Debug.Print "Finished Removing odd characters in the " & strField & " field"


End Sub

Sub modifyUniversalis()

    Call removeOddChars("festival", "tblUniversalis")
    Call removeOddChars("psalm", "tblUniversalis")
    Call removeOddChars("collect", "tblUniversalis")
    Call removeOddChars("gospelacclamation", "tblUniversalis")
    Call removeOddChars("potg", "tblUniversalis")
    Call removeOddChars("communionantiphon", "tblUniversalis")
    Call removeOddChars("postcommunion", "tblUniversalis")
    Call removeOddChars("invitatory", "tblUniversalis")
    Call removeOddChars("morningprayer", "tblUniversalis")
    Call removeOddChars("eveningprayer", "tblUniversalis")
    
    
    Debug.Print "Finished Modifying Universalis"
    
End Sub

'=====================================================================================
' DATABASE FUNCTIONS
'=====================================================================================



Function recordExists(sDate)

Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("tblMass")

    'Search for the first matching record
    rst.MoveFirst
    tmpFound = False
    
    Do Until rst.EOF
        If rst.Fields("cal_date") = sDate Then
            tmpFound = True
        End If
        rst.MoveNext
    Loop
    
    recordExists = tmpFound
       
Cleanup:
        rst.Close
        Set rst = Nothing
        Set dbs = Nothing



End Function

Sub InsertCW()
' add contents of CW table into tblMass

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim CWData As DAO.Recordset
    Dim MassDate As String
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("tblMass", dbOpenTable, dbDenyRead, dbPessimistic)
    Set CWData = dbs.OpenRecordset("CW", dbOpenTable, dbDenyRead, dbPessimistic)

    rst.MoveFirst
    
        Do Until rst.EOF
        
            MassDate = rst.Fields("cal_date")
                    
                    CWData.MoveFirst
                    Do Until CWData.Fields("dDate") = MassDate
                        CWData.MoveNext
                    Loop

            
            rst.Edit
            rst.Fields("CWFestival") = CWData.Fields("Festival")
            
            Debug.Print splitDate(MassDate)
            
            If CWData.Fields("CWMass") <> "" Then
                rst.Fields("CW") = "Holy Communion:\n"
                rst.Fields("CW") = rst.Fields("CW") & CWData.Fields("CWMass") & "\n\n"
            End If
            If CWData.Fields("CWMP") <> "" Then
                rst.Fields("CW") = rst.Fields("CW") & "Mattins:\n"
                rst.Fields("CW") = rst.Fields("CW") & CWData.Fields("CWMP") & "\n\n"
            End If
            If CWData.Fields("CWEP") <> "" Then
                rst.Fields("CW") = rst.Fields("CW") & "Evensong:\n"
                rst.Fields("CW") = rst.Fields("CW") & CWData.Fields("CWEP") & "\n\n"
            End If
            If CWData.Fields("CWAdditionalReadings") <> "" Then
                rst.Fields("CW") = rst.Fields("CW") & "Additional Readings:\n"
                rst.Fields("CW") = rst.Fields("CW") & CWData.Fields("CWAdditionalReadings") & "\n\n"
            End If
            If CWData.Fields("CWCollect") <> "" Then
                rst.Fields("CW") = rst.Fields("CW") & "Collect:\n"
                rst.Fields("CW") = rst.Fields("CW") & CWData.Fields("CWCollect") & "\n\n"
            End If
            
            rst.Update
            
            CWData.MoveFirst
            rst.MoveNext
        Loop

        Debug.Print "Finished Adding CW Data "
        
End Sub


Function getCW(sDate)

' get data from CW tabled mased on sDate yyyymmdd

Dim dbs As DAO.Database
Dim rst As DAO.Recordset
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("CW", dbOpenTable, dbDenyRead, dbPessimistic)
    
    If IsNull(sDate) Then
    
    ' dpo nothing
    
    Else
    
        rst.MoveFirst
    
                Do Until rst.EOF
    
                    If rst.Fields("dDate") = sDate Then
                       getCW = rst
                    End If
                    rst.MoveNext
                    
                Loop
            
            rst.Close
            Set rst = Nothing
            Set dbs = Nothing
        
     End If
     
End Function

Sub AppendUniversalis()
' RUN THIS SUB - the function is called by this.
'Universalis is now the base data for tblMass, to which we add Common Worship



Dim dbs As DAO.Database
    Dim rstUniversalis As DAO.Recordset
    Dim strBuild As String
        
    'Get the database and Recordset
    Set dbs = CurrentDb
    Set rstUniversalis = dbs.OpenRecordset("tblUniversalis", dbOpenTable, dbDenyRead, dbPessimistic)
    
    rstUniversalis.MoveFirst

            Do Until rstUniversalis.EOF
                
                ans = InsertUniversalis(rstUniversalis)
                DoEvents
 
                rstUniversalis.MoveNext
            Loop
        
        rstUniversalis.Close
        
        Set rstMass = Nothing
        Set rstUniversalis = Nothing
        Set dbs = Nothing
        
        Debug.Print "Finished Building Universalis"

End Sub

Function InsertUniversalis(rsu As Recordset)

'Universalis is now the base data for tblMass, to which we add Common Worship

Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("tblMass", dbOpenTable, dbDenyRead, dbPessimistic)
    
    rst.AddNew
    
        rst.Fields("day") = rsu.Fields("day")
        rst.Fields("month") = rsu.Fields("month")
        rst.Fields("year") = rsu.Fields("year")
        rst.Fields("dow") = getDOW(Weekday(rsu.Fields("revised_date")))
        rst.Fields("revised_date") = rsu.Fields("revised_date")
        rst.Fields("cal_date") = rsu.Fields("cal_date")
        rst.Fields("festival") = rsu.Fields("festival")
        rst.Fields("colour") = rsu.Fields("colour")
        rst.Fields("lectionary_year") = rsu.Fields("lectionaryyear")
        rst.Fields("office") = rsu.Fields("office")
        
        
                   tmpRoman = "Roman Lectionary" & "\n"
        tmpRoman = tmpRoman & "================\n\n"
        tmpRoman = tmpRoman & "Morning Prayer\n"
        tmpRoman = tmpRoman & rsu.Fields("morningprayer") & "\n\n"
        tmpRoman = tmpRoman & "EveningPrayer\n"
        tmpRoman = tmpRoman & rsu.Fields("eveningprayer") & "\n\n"
        tmpRoman = tmpRoman & "Readings at Mass\n"
        tmpRoman = tmpRoman & "----------------\n"
        tmpRoman = tmpRoman & "Entrance: " & rsu.Fields("entrance") & "\n"
        tmpRoman = tmpRoman & "Collect: " & rsu.Fields("collect") & "\n"
        tmpRoman = tmpRoman & "First Reading: " & rsu.Fields("firstreading") & "\n"
        tmpRoman = tmpRoman & "Second Reading: " & rsu.Fields("secondreading") & "\n"
        tmpRoman = tmpRoman & "Psalm: " & rsu.Fields("psalm") & "\n"
        tmpRoman = tmpRoman & "Gospel: " & rsu.Fields("gospel") & "\n"
        tmpRoman = tmpRoman & "Prayer over the Gifts: " & rsu.Fields("potg") & "\n"
        tmpRoman = tmpRoman & "Post Communion: " & rsu.Fields("postcommunion") & "\n"
        
        rst.Fields("MassR") = tmpRoman
        
    
    rst.Update
        
        rst.Close
        Set rst = Nothing
        Set dbs = Nothing
        
InsertUniversalis = True
End Function


Function addFieldtoRst(rst As Recordset, sFieldname As String, sFieldValue As Variant) As Boolean
    
    rst.Fields(sFieldname) = Trim(sFieldValue)
    addFieldtoRst = True
    
End Function

Function updateFieldtoRst(rst As Recordset, sFieldname As String, sFieldValue As Variant) As Boolean
    rst.Edit
        If sFieldValue <> "" Then rst.Fields(sFieldname) = Trim(sFieldValue)
    rst.Update
    updateFieldtoRst = True
End Function

Function appendFieldtoRst(rst As Recordset, sFieldname As String, sFieldValue As Variant) As Boolean
    rst.Edit
        tmp = rst.Fields(sFieldname)
        If sFieldValue <> "" Then rst.Fields(sFieldname) = Trim(sFieldValue) & " " & tmp
    rst.Update
    appendFieldtoRst = True
End Function


Public Sub WaitSeconds(intSeconds As Integer)
  ' Comments: Waits for a specified number of seconds
  ' Params  : intSeconds      Number of seconds to wait
  ' Source  : Total Visual SourceBook

  On Error GoTo PROC_ERR

  Dim datTime As Date

  datTime = DateAdd("s", intSeconds, Now)

  Do
   ' Yield to other programs (better than using DoEvents which eats up all the CPU cycles)
    Sleep 100
    DoEvents
  Loop Until Now >= datTime

PROC_EXIT:
  Exit Sub

PROC_ERR:
  Debug.Print "WaitSeconds Error: " & Err.Number & ". " & Err.Description
  Resume PROC_EXIT
End Sub


Function regexReadings(sReading As String) As String
' need to enable Microsoft VBScipt Regular Expressions 5.5 in Tools/References
' Regex to extract Bible Verse https://stackoverflow.com/questions/22254746/bible-verse-regex
' (\d*)\s*([a-z]+)\s*(\d+)(?::(\d+))?(\s*-\s*(\d+)(?:\s*([a-z]+)\s*(\d+))?(?::(\d+))?)?


Dim rExp As Object, rMatch As Object, r_item As Object, result As String

If IsNull(sReading) Then sReading = ""

If Len(sReading) = 0 Then
    regexReadings = ""
Else

'first clean up sloppy bible refs
sReading = Replace(sReading, ": ", ":")
sReading = Replace(sReading, "; ", ";")
sReading = Replace(sReading, ", ", ",")

' now process regex
    Set rExp = CreateObject("vbscript.regexp")
    With rExp
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        ' .Pattern = "(\d*)\s*([a-z]+)\s*(\d+)(?::(\d+))?(\s*-\s*(\d+)(?:\s*([a-z]+)\s*(\d+))?(?::(\d+))?)?"
        'Improved Regex 23/11/21
        .Pattern = "(?:(?:[123]|I{1,3})\s*)?(?:[A-Z][a-zA-Z]+|Song of Songs|Song of Solomon).?\s*(?:1?[0-9]?[0-9]):\s*\d{1,3}(?:[,-]\s*\d{1,3})*(?:;\s*(?:(?:[123]|I{1,3})\s*)?(?:[A-Z][a-zA-Z]+|Song of Songs|Song of Solomon)?.?\s*(?:1?[0-9]?[0-9]):\s*\d{1,3}(?:[,-]\s*\d{1,3})*)*"
    End With
    
    Set rMatch = rExp.Execute(sReading)
    If rMatch.Count > 0 Then
        For Each r_item In rMatch
            result = result & r_item.Value & "\n"
        Next r_item
    End If
    
    regexReadings = result

End If

End Function

Function regexOfficeReadings(sReading As String) As String
' need to enable Microsoft VBScipt Regular Expressions 5.5 in Tools/References
' Regex to extract Bible Verse https://stackoverflow.com/questions/22254746/bible-verse-regex
' (\d*)\s*([a-z]+)\s*(\d+)(?::(\d+))?(\s*-\s*(\d+)(?:\s*([a-z]+)\s*(\d+))?(?::(\d+))?)?


Dim rExp As Object, rMatch As Object, r_item As Object, result As String

If IsNull(sReading) Then sReading = ""

If Len(sReading) = 0 Then
    regexOfficeReadings = ""
Else

'first clean up sloppy bible refs
sReading = Replace(sReading, ": ", ":")
sReading = Replace(sReading, "; ", ";")
sReading = Replace(sReading, ", ", ",")

' now process regex
    Set rExp = CreateObject("vbscript.regexp")
    With rExp
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = "(\d*)\s*([a-z]+)\s*(\d+)(?::(\d+))?(\s*-\s*(\d+)(?:\s*([a-z]+)\s*(\d+))?(?::(\d+))?)?"
    End With
    
    Set rMatch = rExp.Execute(sReading)
    If rMatch.Count > 0 Then
        For Each r_item In rMatch
            result = result & r_item.Value & "\n"
        Next r_item
    End If
    
    regexOfficeReadings = result

End If

End Function

Sub tidyMassR()

Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("tblMass", dbOpenTable, dbDenyRead, dbPessimistic)
    
    rst.MoveFirst

            Do Until rst.EOF
                
                If IsNull(rst.Fields("MassR")) Then
                    ' do nothing
                Else
                rst.Edit
                    rst.Fields("MassR") = Replace(rst.Fields("MassR"), "Canticle", "")
                    rst.Fields("MassR") = Replace(rst.Fields("MassR"), " Reading:", "****")
                    rst.Fields("MassR") = Replace(rst.Fields("MassR"), "Reading", "")
                    rst.Fields("MassR") = Replace(rst.Fields("MassR"), "****", " Reading:")
                    rst.Fields("MassR") = Replace(rst.Fields("MassR"), "Apocalypse", "Revelation")
                    rst.Fields("MassR") = Replace(rst.Fields("MassR"), "Gospel", "Gospel: ")
                    rst.Fields("MassR") = Replace(rst.Fields("MassR"), "Gospel: :", "Gospel: ")
                rst.Update
                End If
                             
                rst.MoveNext
            Loop
        
        rst.Close
        Set rst = Nothing
        Set dbs = Nothing
        
        Debug.Print "Finished Sorting out MassR field"

End Sub



Sub spaceReadings(strField, tblData)

Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim p1 As Integer
    Dim sSource As String, strTemp As String, strNew As String, strReading As String
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset(tblData, dbOpenTable, dbDenyRead, dbPessimistic)
    
    rst.MoveFirst
    
            Do Until rst.EOF
                rst.Edit
                
                If IsNull(rst.Fields(strField)) Then
                    sSource = ""
                Else
                    sSource = rst.Fields(strField)
                
                    
                    'first clean up sloppy bible refs
                    sSouce = Replace(sSouce, ": ", ":")
                    sSouce = Replace(sSouce, "; ", ";")
                    sSouce = Replace(sSouce, ", ", ",")

                    strReading = regexReadings(sSource)
                    Debug.Print "Extracted Reading= " & strReading & " Length: " & Len(strReading)
                    If Len(strReading) > 0 Then
                        p1 = InStr(sSource, strReading)
                        Debug.Print "p1: " & p1
                        strTemp = Mid(sSource, p1 + Len(strReading))
                        
                        Debug.Print strTemp
                        strNew = strReading & " - " & strTemp
                        Debug.Print "--->" & strNew
                        rst.Fields(strField) = strNew
                    End If
                                   
                rst.Update
                rst.MoveNext
            Loop
        
        rst.Close
        Set rst = Nothing
        Set dbs = Nothing
        
        Debug.Print "Finished Spacing Bible Refs in " & strField

End Sub


Option Compare Database

'additional functions

Sub DeleteEnd(strFind, strField, tblData)
' delete everything after a string

Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset(tblData, dbOpenTable, dbDenyRead, dbPessimistic)
    
    rst.MoveFirst

            Do Until rst.EOF


                rst.Edit
                
                p1 = InStr(rst.Fields(strField), strFind)
                
                Debug.Print "Position: " & p1 & ")" & Left(rst.Fields(strField), p1)
                
                If p1 > 0 Then
                
                    rst.Fields(strField) = Left(rst.Fields(strField), p1 - 1)
                
                End If
               
                rst.Update
                rst.MoveNext
            Loop
        
        rst.Close
        Set rst = Nothing
        Set dbs = Nothing
        
        Debug.Print "Finished removing" & strFind

End Sub

Sub DeleteBegin(strFind, strField, tblData)
' Find a string and only keep everything after it

Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset(tblData, dbOpenTable, dbDenyRead, dbPessimistic)
    
    rst.MoveFirst

            Do Until rst.EOF


                rst.Edit
                
                p1 = InStr(rst.Fields(strField), strFind)
                
                
                
                If p1 > 0 Then
                
                    Debug.Print rst.Fields(strField)
                    Debug.Print "Position: " & p1 & ")" & Mid(rst.Fields(strField), p1 + Len(strFind) + 1)
                    Debug.Print "================================================"
                    rst.Fields(strField) = Mid(rst.Fields(strField), p1 + Len(strFind) + 1)
                rst.Update
                End If
               
                
                rst.MoveNext
            Loop
        
        rst.Close
        Set rst = Nothing
        Set dbs = Nothing
        
        Debug.Print "Finished removing" & strFind

End Sub


Sub clear(sTable As String)
' clear out a table for a restart (and add a single record)

DoCmd.RunSQL "DELETE * FROM " & sTable

Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset(sTable, dbOpenTable, dbDenyRead, dbPessimistic)
    
                rst.AddNew
                
                rst.Fields("revised_date") = "01/01/1970"
               
                rst.Update
        
        rst.Close
        Set rst = Nothing
        Set dbs = Nothing
        
        Debug.Print "Finished Clearing " & sTable

End Sub




Function BibleNames(sString As String) As String
' expand bible names to their full titles.

    'OT
    sString = Replace(sString, "Gen", "Genesis ", , , vbBinaryCompare)
    sString = Replace(sString, "Ex", "Exodus ", , , vbBinaryCompare)
    sString = Replace(sString, "Lev", "Leviticus ", , , vbBinaryCompare)
    sString = Replace(sString, "Num", "Numbers ", , , vbBinaryCompare)
    sString = Replace(sString, "Deut", "Deuteronomy ", , , vbBinaryCompare)
    sString = Replace(sString, "Josh", "Joshua ", , , vbBinaryCompare)
    sString = Replace(sString, "Judg", "Judges ", , , vbBinaryCompare)
    ' is Ruth ever shortened?
    sString = Replace(sString, "1Sam", "1 Samuel ", , , vbBinaryCompare)
    sString = Replace(sString, "2Sam", "2 Samuel ", , , vbBinaryCompare)
    sString = Replace(sString, "1Kings", "1 Kings ", , , vbBinaryCompare)
    sString = Replace(sString, "2Kings", "2 Kings ", , , vbBinaryCompare)
    sString = Replace(sString, "1Chron", "1 Chronicles ", , , vbBinaryCompare)
    sString = Replace(sString, "2Chron", "2 Chronicles ", , , vbBinaryCompare)
    ' Ezra
    sString = Replace(sString, "Neh", "Nehemiah ", , , vbBinaryCompare)
    sString = Replace(sString, "Est", "Esther ", , , vbBinaryCompare)
    'Job
    sString = Replace(sString, "Ps", "Psalms ", , , vbBinaryCompare)
    sString = Replace(sString, "Prov", "Proverbs ", , , vbBinaryCompare)
    sString = Replace(sString, "Eccles", "Ecclesiastes ", , , vbBinaryCompare)
    sString = Replace(sString, "Song", "Song of Solomon ", , , vbBinaryCompare)
    sString = Replace(sString, "Is", "Isaiah ", , , vbBinaryCompare)
    sString = Replace(sString, "Jer", "Jeremiah ", , , vbBinaryCompare)
    sString = Replace(sString, "Lam", "Lamentations ", , , vbBinaryCompare)
    sString = Replace(sString, "Ez", "Exekiel ", , , vbBinaryCompare)
    sString = Replace(sString, "Dan", "Daniel ", , , vbBinaryCompare)
    sString = Replace(sString, "Hos", "Hosea ", , , vbBinaryCompare)
    'Joel
    'Amos
    sString = Replace(sString, "Obad", "Obadiah ", , , vbBinaryCompare)
    'Jonah
    sString = Replace(sString, "Mic", "Micah ", , , vbBinaryCompare)
    sString = Replace(sString, "Nah", "Nahum ", , , vbBinaryCompare)
    sString = Replace(sString, "Hab", "Habakkuk ", , , vbBinaryCompare)
    sString = Replace(sString, "Zeph", "Zephaniah ", , , vbBinaryCompare)
    sString = Replace(sString, "Hag", "Haggai ", , , vbBinaryCompare)
    sString = Replace(sString, "Zech", "Zechariah ", , , vbBinaryCompare)
    sString = Replace(sString, "Mal", "Malachi ", , , vbBinaryCompare)
    
    'Apocrypha
    sString = Replace(sString, "Tob", "Tobit ", , , vbBinaryCompare)
    sString = Replace(sString, "Jth", "Judith ", , , vbBinaryCompare)
    sString = Replace(sString, "Wis", "Wisdom ", , , vbBinaryCompare)
    sString = Replace(sString, "Sir", "Sirach ", , , vbBinaryCompare)
    sString = Replace(sString, "Ecclus", "Ecclesiaticus ", , , vbBinaryCompare)
    sString = Replace(sString, "Bar", "Baruch ", , , vbBinaryCompare)
    sString = Replace(sString, "Sus", "Susanna ", , , vbBinaryCompare)
    sString = Replace(sString, "1Macc", "1 Maccabees ", , , vbBinaryCompare)
    sString = Replace(sString, "2Macc", "2 Maccabees ", , , vbBinaryCompare)
    sString = Replace(sString, "3Macc", "3 Maccabees ", , , vbBinaryCompare)
    sString = Replace(sString, "4Macc", "4 Maccabees ", , , vbBinaryCompare)
    sString = Replace(sString, "1Esd", "1 Esdras ", , , vbBinaryCompare)
    sString = Replace(sString, "2Esd", "2 Esdras ", , , vbBinaryCompare)
    sString = Replace(sString, "Man", "Prayer of Manasseh ", , , vbBinaryCompare)
    
    'NT
    sString = Replace(sString, "Matt", "Matthew ", , , vbBinaryCompare)
    sString = Replace(sString, "Mt", "Matthew ", , , vbBinaryCompare)
    sString = Replace(sString, "Mk", "Mark ", , , vbBinaryCompare)
    sString = Replace(sString, "Lk", "Luke ", , , vbBinaryCompare)
    sString = Replace(sString, "Jn", "John ", , , vbBinaryCompare)
    'Acts
    sString = Replace(sString, "Rom", "Romans ", , , vbBinaryCompare)
    sString = Replace(sString, "Rm", "Romans ", , , vbBinaryCompare)
    sString = Replace(sString, "1Cor", "1 Corinthians ", , , vbBinaryCompare)
    sString = Replace(sString, "2Cor", "2 Corinthians ", , , vbBinaryCompare)
    sString = Replace(sString, "Gal", "Galatians ", , , vbBinaryCompare)
    sString = Replace(sString, "Eph", "Ephesians ", , , vbBinaryCompare)
    sString = Replace(sString, "Phil", "Philippians ", , , vbBinaryCompare)
    sString = Replace(sString, "Col", "Colossians ", , , vbBinaryCompare)
    sString = Replace(sString, "1Thess", "1 Thessalonians ", , , vbBinaryCompare)
    sString = Replace(sString, "2Thess", "2 Thessalonians ", , , vbBinaryCompare)
    sString = Replace(sString, "1Tim", "1 Timothy ", , , vbBinaryCompare)
    sString = Replace(sString, "2Tim", "2 Timothy ", , , vbBinaryCompare)
    'Titus
    sString = Replace(sString, "Philem", "Philemon ", , , vbBinaryCompare)
    sString = Replace(sString, "Heb", "Hebrews ", , , vbBinaryCompare)
    ' James
    sString = Replace(sString, "1Pet", "1 Peter ", , , vbBinaryCompare)
    sString = Replace(sString, "2Pet", "2 Peter ", , , vbBinaryCompare)
    sString = Replace(sString, "1Jn", "1 John ", , , vbBinaryCompare)
    sString = Replace(sString, "2Jn", "2 John ", , , vbBinaryCompare)
    sString = Replace(sString, "3Jn", "3 John ", , , vbBinaryCompare)
    'Jude
    sString = Replace(sString, "Rev", "Revelation ", , , vbBinaryCompare)
    sString = Replace(sString, "Rv", "Revelation ", , , vbBinaryCompare)

    BibleNames = sString

End Function

Function separateReadings(sReading As String) As String

    sRef = regexReadings(sReading)
    sLen = Len(sRef)
    If sLen > 0 Then
        sSearch = Left(Right(sRef, 3), 1) ' minus the \n
        sTemp = Mid(sReading, sSearch + 1)
        sPos = InStr(sTemp, sSearch)
    
        separateReadings = sRef & Mid(sTemp, sPos + 1)
    Else
        separateReadings = sReading
    End If

End Function

Sub SortFestivals()
' Work out festival type, Creed and Gloria from the data

Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("tblMass", dbOpenTable, dbDenyRead, dbPessimistic)
    
    rst.MoveFirst

            Do Until rst.EOF
            
            
            rst.Edit
            
            rst.Fields("festivaltype") = "Feria" ' baseline is Feria
            
                If InStr(1, rst.Fields("festival"), "Saint") <> 0 Then
                       rst.Fields("festivaltype") = "Memorial"
                End If
                
                If InStr(1, rst.Fields("festival"), "Dedication") <> 0 Then
                       rst.Fields("festivaltype") = "Memorial"
                End If
                
                If InStr(1, rst.Fields("festival"), "Memorial") <> 0 Then
                       rst.Fields("festivaltype") = "Memorial"
                End If
                
                If InStr(1, rst.Fields("festival"), "Presentation") <> 0 Then
                       rst.Fields("festivaltype") = "Memorial"
                End If
                
                If InStr(1, rst.Fields("festival"), "Priest") <> 0 Then
                       rst.Fields("festivaltype") = "Memorial"
                End If
                
                If InStr(1, rst.Fields("festival"), "Doctor") <> 0 Then
                       rst.Fields("festivaltype") = "Memorial"
                End If
                
                If InStr(1, rst.Fields("festival"), "Bishop") <> 0 Then
                       rst.Fields("festivaltype") = "Memorial"
                End If
                
                If InStr(1, rst.Fields("festival"), "Martyr") <> 0 Then
                       rst.Fields("festivaltype") = "Memorial"
                End If
                
                If InStr(1, rst.Fields("festival"), "Pope") <> 0 Then
                       rst.Fields("festivaltype") = "Memorial"
                End If
                
                If InStr(1, rst.Fields("festival"), "Virgin") <> 0 Then
                       rst.Fields("festivaltype") = "Memorial"
                End If
                
                If InStr(1, rst.Fields("festival"), "Deacon") <> 0 Then
                       rst.Fields("festivaltype") = "Memorial"
                End If
                
                If InStr(1, rst.Fields("festival"), "Apostle") <> 0 Then
                       rst.Fields("festivaltype") = "Memorial"
                       rst.Fields("gloria") = "Gloria"
                End If
                
                If InStr(1, rst.Fields("festival"), "Missionary") <> 0 Then
                       rst.Fields("festivaltype") = "Memorial"
                End If
                
                'end with these
                
                If (rst.Fields("dow") = "Sunday") Then
                       rst.Fields("festivaltype") = "Solemnity"
                       rst.Fields("gloria") = "Gloria"
                       rst.Fields("creed") = "Creed"
                End If
               
                If (rst.Fields("dow") = "Sunday") And (InStr(1, rst.Fields("festival"), "Advent") <> 0) Then
                       rst.Fields("festivaltype") = "Solemnity"
                       rst.Fields("creed") = "Creed"
                End If
                
                If (rst.Fields("dow") = "Sunday") And (InStr(1, rst.Fields("festival"), "Lent") <> 0) Then
                       rst.Fields("festivaltype") = "Solemnity"
                       rst.Fields("creed") = "Creed"
                End If
                
                If InStr(1, rst.Fields("festival"), "Solemnity") <> 0 Then
                       rst.Fields("festivaltype") = "Solemnity"
                       rst.Fields("gloria") = "Gloria"
                       rst.Fields("creed") = "Creed"
                End If
                
                If InStr(1, rst.Fields("festival"), "Immaculate") <> 0 Then
                       rst.Fields("festivaltype") = "Solemnity"
                       rst.Fields("gloria") = "Gloria"
                       rst.Fields("creed") = "Creed"
                End If
                               
                If InStr(1, rst.Fields("festival"), "Feast") <> 0 Then
                       rst.Fields("festivaltype") = "Feast"
                       rst.Fields("gloria") = "Gloria"
                End If
                
                Debug.Print rst.Fields("revised_date") & " " & rst.Fields("festival") & " is a " & rst.Fields("festivaltype")
                
                rst.Update
                rst.MoveNext
            Loop
        
        rst.Close
        Set rst = Nothing
        Set dbs = Nothing
        
        Debug.Print "Finished Sorting Out the Festivals"

End Sub

Function CountStringOccurances(strStringToCheck As String, strValueToCheck As String) As Integer
'Purpose: Counts the number of times a string appears in another string.

    Dim intStringPosition As Long
    Dim intCursorPosition As Long
    Dim i As Double
    CountStringOccurances = 0
    intCursorPosition = 1
    For i = 0 To Len(strStringToCheck)
        intStringPosition = InStr(intCursorPosition, strStringToCheck, strValueToCheck)
        If intStringPosition = 0 Then
            Exit Function
        Else
            CountStringOccurances = CountStringOccurances + 1
            intCursorPosition = intStringPosition + Len(strValueToCheck)
        End If
    Next i
    Exit Function

End Function

Function IsFound(strData As String, strSearch As String) As Boolean

If InStr(1, strData, strSearch) > 0 Then
    IsFound = True
Else
    IsFound = False
End If

End Function

Function GetText(ByVal sLine As String, sStartTag As String, sEndTag As String) As String

Dim p1 As Integer
Dim p2 As Integer

If IsNull(sLine) Then
    GetText = ""

Else
    p1 = InStr(1, sLine, sStartTag)
    If p1 <> 0 And sEndTag <> "#EOF" Then p2 = InStr(p1, sLine, sEndTag) ' case sensitive from the NEXT occcurance
    If sEndTag = "#EOF" Then p2 = Len(sLine)
    
    If (p1 = 0) Or (p2 = 0) Then
            GetText = ""
    Else
            p1 = p1 + Len(sStartTag)
            
            If p2 > p1 Then
                GetText = Mid(sLine, p1, p2 - p1)
                GetText = Replace(GetText, "\,", ",")
            Else
                GetText = ""
            End If
            
    End If
End If

End Function
Function SplitDays(sLine As String, sStartTag As String)
' return array split by sStartTag

    Dim sParts() As String
    sParts = Split(sLine, sStartTag)

SplitDays = sParts

End Function

Sub readICS()

    Dim sSourceFile As String
    Dim sPosition As Long
    Dim fileString As String
    Dim LineString As String
    Dim EventDate As String
    Dim strSummary As String
    Dim strFestival As String
    Dim strHC As String
    Dim strMP As String
    Dim strEP As String
    Dim strSdrLect As String
    Dim strCollect As String
    Dim strPostCommunion As String
    Dim strAdditionalCollect As String
    Dim sBCPTitle As String
    Dim sBCPMass As String
    Dim sBCPCollect As String
    Dim TotalDay As Integer
    Dim iCount As Integer
    Dim DaysList() As String
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("CW", dbOpenTable, dbDenyRead, dbPessimistic)


sSourceFile = "C:\Users\simon\Dropbox\Development\churchdates\2023\CW.ics"
Debug.Print "importing " & sSourceFile

Open sSourceFile For Input As #1
               
While Not EOF(1)
   Line Input #1, LineString
   fileString = fileString & Trim(LineString)
Wend

'Closes the text file
Close #1

'Now process filestring

TotalDay = CountStringOccurances(fileString, "BEGIN:VEVENT")
  
Debug.Print "There are " & TotalDay & " Events in this Calendar"

For iCount = 1 To TotalDay

    DaysList = SplitDays(fileString, "BEGIN:VEVENT")

    EventDate = GetText(DaysList(iCount), "VALUE=DATE:", "DTEND")
    
    strSummary = GetText(DaysList(iCount), "SUMMARY:", "DESCRIPTION:")
    
     If IsFound(DaysList(iCount), "DESCRIPTION:") Then
        Debug.Print "FOUND DESCRIPTION " & GetText(DaysList(iCount), "DESCRIPTION:", "https")
        strFestival = GetText(DaysList(iCount), "DESCRIPTION:", "https")
        If strFestival = "" Then strFestival = strSummary
        If Len(strFestival) < 4 Then strFestival = "Feria"
    End If
    
    If IsFound(DaysList(iCount), "Holy Communion\n") Then
        Debug.Print "FOUND HC" & GetText(DaysList(iCount), "Principal Service\n", "\n\n")
        strHC = GetText(DaysList(iCount), "Holy Communion\n", "\n\n")
            If strHC = "" Then strHC = GetText(DaysList(iCount), "Principal Service\n", "\n\n")
    End If
    
    If IsFound(DaysList(iCount), "Morning Prayer\n") Then
        strMP = GetText(DaysList(iCount), "Morning Prayer\n", "\n\n")
            If strMP = "" Then strMP = GetText(DaysList(iCount), "Third Service\n", "\n\n")
    End If
    
    If IsFound(DaysList(iCount), "Evening Prayer\n") Then
        strEP = GetText(DaysList(iCount), "Evening Prayer\n", "\n\n")
            If strEP = "" Then strEP = GetText(DaysList(iCount), "Second Service\n", "\n\n")
    End If
        
    If IsFound(DaysList(iCount), "Additional Weekday Lectionary\n") Then
        strSdrLect = GetText(DaysList(iCount), "Additional Weekday Lectionary\n", "\n\n")
    End If
    
    If IsFound(DaysList(iCount), "Collect\n\n") Then
        strCollect = GetText(DaysList(iCount), "Collect\n\n", "Post Communion")
    End If
    
    If IsFound(DaysList(iCount), "Post Communion\n\n") Then
        strPostCommunion = GetText(DaysList(iCount), "Post Communion\n\n", "\n\n")
    End If
    
    If IsFound(DaysList(iCount), "Additional Collect\n\n") Then
        strAdditionalCollect = GetText(DaysList(iCount), "Additional Collect\n\n", "\n\n")
    End If
    
    If IsFound(DaysList(iCount), "BCP:") Then
        sBCPTitle = GetText(DaysList(iCount), "BCP:", "\n\n")
    End If
    
    If IsFound(DaysList(iCount), "BCP Holy Communion\n\n") Then
        sBCPMass = GetText(DaysList(iCount), "BCP Holy Communion\n\n", "\n\n")
    End If
    
    If IsFound(DaysList(iCount), "BCP Collect\n\n") Then
        sBCPCollect = GetText(DaysList(iCount), "BCP Collect\n\n", "\n\n")
    End If

   
    Debug.Print "Extracted:"
    Debug.Print "Date: " & EventDate & "-> " & splitDate(EventDate)
    Debug.Print "Festival: " & strFestival
    Debug.Print "CW Mass: " & strHC
    Debug.Print "MP: " & strMP
    Debug.Print "EP: " & strEP
    Debug.Print "Additional: " & strSdrLect
    Debug.Print "Collect: " & strCollect
    Debug.Print "Post Communion: " & strPostCommunion
    Debug.Print "Additional Collect: " & strAdditionalCollect
    Debug.Print "BCP: " & sBCPTitle
    Debug.Print "Mass: " & sBCPMass
    Debug.Print "Collect: " & sBCPCollect
    Debug.Print "========================================"
    
    DoEvents
    
    ' Add to CW Table
    
    rst.AddNew
    
        rst.Fields("dDate") = EventDate
        rst.Fields("Festival") = strFestival
        rst.Fields("CWMass") = strHC
        rst.Fields("CWMP") = strMP
        rst.Fields("CWEP") = strEP
        rst.Fields("CWAdditionalReadings") = strSdrLect
        rst.Fields("CWCollect") = strCollect
        rst.Fields("CWPostCommunion") = strPostCommunion
        rst.Fields("CWAdditionalCollect") = strAdditionalCollect
        rst.Fields("BCPTitle") = sBCPTitle
        rst.Fields("BCPMass") = sBCPMass
        rst.Fields("BCPCollect") = sBCPCollect
    rst.Update
    
        
Next iCount

Debug.Print "Finished Processing ICS"


End Sub

Sub FestivalType()

Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("tblMass", dbOpenTable, dbDenyRead, dbPessimistic)
    
    rst.MoveFirst

            Do Until rst.EOF


                rst.Edit
                
                rst.Fields("FestivalType") = "Feria" ' default
                
                    If IsFound(rst.Fields("festival"), "Ordinary") Then rst.Fields("festivaltype") = "Feria"
                    If IsFound(rst.Fields("festival"), " or ") Then rst.Fields("festivaltype") = "Memorial"
                    If IsFound(rst.Fields("festival"), "Bishop") Then rst.Fields("festivaltype") = "Memorial"
                    If IsFound(rst.Fields("festival"), "Martyr") Then rst.Fields("festivaltype") = "Memorial"
                    If IsFound(rst.Fields("festival"), "Doctor") Then rst.Fields("festivaltype") = "Memorial"
                    If IsFound(rst.Fields("festival"), "Virgin") Then rst.Fields("festivaltype") = "Memorial"
                    If IsFound(rst.Fields("festival"), "Priest") Then rst.Fields("festivaltype") = "Memorial"
                    If IsFound(rst.Fields("festival"), "Religious") Then rst.Fields("festivaltype") = "Memorial"
                    If IsFound(rst.Fields("festival"), "Missionary") Then rst.Fields("festivaltype") = "Memorial"
                    If IsFound(rst.Fields("festival"), "memorial") Then rst.Fields("festivaltype") = "Memorial"
                    If IsFound(rst.Fields("festival"), "Saint") Then rst.Fields("festivaltype") = "Memorial"
                    If IsFound(rst.Fields("festival"), "Lady") Then rst.Fields("festivaltype") = "Memorial"
                    If IsFound(rst.Fields("festival"), "Solemnity") Then rst.Fields("festivaltype") = "Solemnity"
                    If IsFound(rst.Fields("festival"), "Feast") Then rst.Fields("festivaltype") = "Feast"
                    
                    If IsFound(rst.Fields("dow"), "Sunday") Then rst.Fields("festivaltype") = "Solemnity"
                    
                    If IsFound(rst.Fields("festivaltype"), "Solemnity") Then
                        rst.Fields("gloria") = "Gloria"
                        rst.Fields("creed") = "Creed"
                    End If
                    
                     If IsFound(rst.Fields("festivaltype"), "Feast") Then
                        rst.Fields("gloria") = "Gloria"
                    End If
                    
                    If IsFound(rst.Fields("festival"), "Lent") Then
                        rst.Fields("gloria") = ""
                    End If
                    
                    If IsFound(rst.Fields("festival"), "Advent") Then
                        rst.Fields("gloria") = ""
                    End If
                    
                    Debug.Print rst.Fields("revised_date") & " is a "; rst.Fields("festivaltype")
                    
               
                rst.Update
                rst.MoveNext
            Loop
        
        rst.Close
        Set rst = Nothing
        Set dbs = Nothing
        
        Debug.Print "Finished Database Manipulation"

End Sub


Sub Z_DBStandard(strField, tblData)
' Template for Data Manipulation, keep at bottom of code

Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset(tblData, dbOpenTable, dbDenyRead, dbPessimistic)
    
    rst.MoveFirst

            Do Until rst.EOF


                rst.Edit
               
                rst.Update
                rst.MoveNext
            Loop
        
        rst.Close
        Set rst = Nothing
        Set dbs = Nothing
        
        Debug.Print "Finished Database Manipulation"

End Sub
