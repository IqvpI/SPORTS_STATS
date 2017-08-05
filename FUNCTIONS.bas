Attribute VB_Name = "FUNCTIONS"
'CHECK IF FILE OPEN
Function FILEOPEN(WB, path1)

        For Each Workbook In Application.Workbooks
            If InStr(1, Workbook.Name, WB, 1) <> 0 Then
                Exit Function 'FILE OPEN
            End If
        Next
        
        'OPEN FILE
        Workbooks.Open Filename:=path1 & WB

End Function
'ATTRIBUTE EXTRACT
Function attextract(objie, ByRef tagname1 As Variant, ByRef classname1 As Variant, _
ByRef PNclassname As Variant, ByRef dataid As Variant, ByRef attributename As Variant, _
attributenameVAL)
    
    Dim attributte As Object, attributes As Object
    
    With objie.document
        Set links1 = .getElementsByTagName(tagname1)
        For Each link1 In links1
            
            'VERIFY CORRECT LINK BY PARENT NODES AND DATA-ID
            If PNclassname <> "" Then
                If link1.ParentNode.ParentNode.className = PNclassname _
                And InStr(1, link1.ParentNode.ParentNode.innerHTML, dataid, 1) > 0 Then
                    GoTo ESA:
                End If
                GoTo Nextlink:
            End If
            
            If classname1 = "" Then 'URL LISTED ON PAGE
                If link1.getAttribute(attributename) = attributenameVAL Then
                    attributeextract = link1.innerHTML
                    Exit For
                End If
            Else
ESA:
                If link1.className = classname1 Then 'EXTRACT SPECIFIC ATTRIBUTE
                    attributeextract = link1.getAttribute(attributename)
                    Exit For
                End If
            End If
Nextlink:
        Next
    End With
    
    'RELEASE VARS
    attextract = attributeextract
    tagname1 = ""
    classname1 = ""
    PNclassname = ""
    dataid = ""
    attributename = ""
    
End Function
'CLICK BUTTON
Sub buttonclick(objie, ByRef buttontag As Variant, _
ByRef buttonclass As Variant, ByRef buttontext As Variant)

    Dim button As Object, buttons As Object
    
        With objie.document
            Set buttons = .getElementsByTagName(buttontag)
            For Each button In buttons
                If button.className = buttonclass Then
                    button.Click
                    Do While objie.Busy = True Or objie.readyState <> 4: DoEvents: Loop
                    Exit For
                End If
            Next
        End With

    'RELEASE VARS
    buttontag = ""
    buttonclass = ""
    buttontext = ""

End Sub
'CHECK TEXT
Function checktext(objie, tofind)

    With objie.document
        For i = 1 To 2
        
            'SET VARS
            If i = 1 Then
                findText = InStr(1, .body.innerText, tofind, 1) 'CHECK TEXT
            ElseIf i = 2 Then
                findText = InStr(6000, .body.innerHTML, tofind, 1) 'CHECK HYPERLINK
            End If
            
            'SEARCH WEBPAGE
            If findText > 0 Then
                checktext = True
                Exit Function
            End If
        
        Next
    End With
    
End Function
'CHECK EXACT TEXT
Function exact_check(WB, ws, text)

With Workbooks(WB).Worksheets(ws)
    If Not .Cells.Find(text, lookat:=xlWhole) Is Nothing Then
        exact_check = .Cells.Find(text, lookat:=xlWhole).Column
'    ElseIf Not .Cells.Find(text, lookat:=xlPart) Is Nothing Then
'        exact_check = .Cells.Find(text, lookat:=xlPart).Column
    Else: exact_check = False
    End If
End With

End Function
'EXTRACT WEB TEXT
Function extractText(objie, ByRef tofind As Variant)

    With objie.document
        fp = InStr(1, .body.innerText, tofind, 1)
        lp = InStr(1, .body.innerText, ".com", 1) + 4
        extractText = Mid(.body.innerText, fp, lp - fp)
    End With
    
    'RELEASE VARS
    tofind = ""
    
End Function
'EXTRACT SPECIFIC TEXT
Function extract_specific_text(text, fp, lp)

    fp = InStr(1, text, fp, 1)
'    lp = InStr(fp, text, lp, 1)
    extract_specific_text = Left(text, fp - 2)
    
End Function
'DATA AVOIDANCE
Function data_avoidance(dataset, SSWB, text)

    'EXTRACT SPECIFIC TEXT
    If dataset = "TS" And InStr(1, text, "[", 1) > 0 Then
        fp = "["
        text = extract_specific_text(text, fp, lp)
    End If

    ws = "Cols"
    
    'CHECK IF DATA NEEDS TO BE AVOIDED
    With Workbooks(SSWB).Worksheets(ws)
        If Not .Cells.Find("Avoid", lookat:=xlWhole).CurrentRegion.Find(text, lookat:=xlWhole) Is Nothing Then
            data_avoidance = True
        End If
    End With

End Function
'CHECK COLUMN EXCEPTIONS
Function column_exceptions(WB, text, dataset)

    ws = "Cols"

    With Workbooks(WB).Worksheets(ws)
        If Not .Cells.Find(dataset, lookat:=xlWhole).CurrentRegion.Find(text, lookat:=xlWhole) Is Nothing Then
            column_exceptions = .Cells.Find(dataset, lookat:=xlWhole).CurrentRegion.Find(text, lookat:=xlWhole) _
            .Offset(0, -1)
        End If
    End With

End Function
'DIVISION_PARSE
Sub Division_Parse(SSWB, SSws, nxrow, DIVcol, stats)
        
    With Workbooks(SSWB).Worksheets(SSws)
        If stats = "East" Or stats = "Central" Or stats = "West" Then
            .Cells(nxrow, DIVcol) = stats
        Else:
            If .Cells(nxrow, DIVcol) = "" Then
                .Cells(nxrow, DIVcol) = divvar
            End If
        End If
    End With

End Sub
'LEAGUE_PARSE AND LINE DELETIONS
Sub League_Parse(SSWB, SSws)

    Dim team As String
    AL = "American"
    NL = "National"
    ws = "Cols"
redo:
    'DECLARE COLUMN/ROWS
    For i = 1 To 2
        With Workbooks(SSWB).Worksheets(SSws)
            teamcol = .Cells.Find("Team", lookat:=xlWhole).Column
            fr = .Cells.Find("Team", lookat:=xlWhole).Offset(1).Row
            lr = .Cells.Find("Team", lookat:=xlWhole).End(xlDown).Row
            
            If i = 1 Then 'DELETE ERRONEOUS ROWS
                
                For ii = fr To lr
                    If .Cells(ii, teamcol) = "East" _
                    Or .Cells(ii, teamcol) = "Central" _
                    Or .Cells(ii, teamcol) = "West" Then
                        .Rows(ii).Delete
                    End If
                Next
            
            ElseIf i = 2 Then 'ADD LEAGUE
                
                For iii = fr To lr
                    team = .Cells(iii, teamcol)
                    If Workbooks(SSWB).Worksheets(ws).Cells.Find(AL, lookat:=xlWhole) _
                    .CurrentRegion.Find(team, lookat:=xlPart) Is Nothing Then
                        .Cells(iii, teamcol).Offset(0, -2) = NL
                    Else: .Cells(iii, teamcol).Offset(0, -2) = AL
                    End If
                Next
                
            End If
        End With
    Next

End Sub
'WIPE DB DATA
Sub wipe_data(SSWB, SSws)

    With Workbooks(SSWB).Worksheets(SSws)
        lr = .Range("A10000").End(xlUp).Row
        If lr <> 2 Then
            .Rows("3:" & lr).Delete
        End If
    End With

End Sub
'VERIFY JOB ID NOT PRESENT IN DATABASE
Function DBcheck(tocheck) As Boolean

JLWB = "jobs log.xlsb"
Jws = "Jobs"
ESws = "External Sites"

    For i = 1 To 2
        If i = 1 Then
            ws = Jws 'JOBS tab'
        ElseIf i = 2 Then
            ws = ESws 'External sites tab
        End If
        
        'SEARCH SHEET
        With Workbooks(JLWB).Worksheets(ws)
            dataIDcol = .Cells.Find("data-id", lookat:=xlWhole).Column
            If Not .Columns(dataIDcol).Find(tocheck, lookat:=xlWhole) Is Nothing Then
                DBcheck = True 'ID FOUND
                Exit Function
            End If
        End With
    Next
    
    DBcheck = False

End Function
'TEAM STANDINGS DIVISION DETERMINATION
 Function ts_Division_Determination(elemcol, num, num2)
    
    'VARS
    Dim AL As String: AL = "American League"
    Dim NL As String: NL = "National League"
    DIV = ""
    divvar = ""
    league = ""
    leaguevar = ""
                    
    'DIVISION
    For i = 1 To 3
        If i = 1 Then
            divvar = "West"
        ElseIf i = 2 Then: divvar = "Central"
        ElseIf i = 3 Then: divvar = "West"
        End If
'            If InStr(1, elemcol(t).Rows(r).ParentNode.innerText, DIVvar, 1) > 0 Then

            MsgBox elemcol(t).ParentNode.innerText
'
'            MsgBox elemcol(t).Rows(1).Cells(0).innerText
            
            If InStr(1, elemcol(t).ParentNode.innerText, divvar, 1) > 0 Then
                DIV = divvar
                Exit For
            End If
    Next
        
    'LEAGUE
    For ii = 1 To 2
        If ii = 1 Then
            leaguevar = AL
        ElseIf ii = 2 Then: leaguevar = NL
        End If
            If InStr(1, elemcol(t).ParentNode.ParentNode.innerText, leaguevar, 1) > 0 Then
                league = leaguevar
                Exit For
            End If
    Next
                
    'COMBINE L+D
    If league <> "" And DIV <> "" Then
        ts_Division_Determination = "[" & league & "-" & DIV & "]"
    Else: Application.Speech.Speak "Error detected in league or division determination!"
        Stop
    End If

End Function
'EMAIL OUTREACH
Sub send_email(email_subject, email_recipient, email_body)

'ERROR HADNLING
On Error GoTo errcatching:

Dim mymail As CDO.Message
Dim attachmentsarray(1) As Variant

'VARS
schemaconfig = "http://schemas.microsoft.com/cdo/configuration/"
emailADRS = "zacharylenat@gmail.com"
PWORD = "P11nkfloyd!"
fldrpath = "C:\Users\qp\Desktop\RESUME\"
RESpath = fldrpath & "ZLRESUME.pdf"
CLetterpath = fldrpath & "Cover Letter.pdf"

attachmentsarray(0) = RESpath
attachmentsarray(1) = CLetterpath

Set mymail = New CDO.Message

'SSL
mymail.Configuration.Fields.Item(schemaconfig & "smtpusessl") = True
'AUTHENTICATE
mymail.Configuration.Fields.Item(schemaconfig & "smtpauthenticate") = 1
'SERVER
mymail.Configuration.Fields.Item(schemaconfig & "smtpserver") = "smtp.gmail.com"
'SERVER PORT
mymail.Configuration.Fields.Item(schemaconfig & "smtpserverport") = 465
'SEND USING
mymail.Configuration.Fields.Item(schemaconfig & "sendusing") = 2
'USER NAME
mymail.Configuration.Fields.Item(schemaconfig & "sendusername") = emailADRS
'PASS WORD
mymail.Configuration.Fields.Item(schemaconfig & "sendpassword") = PWORD
'UPDATE
mymail.Configuration.Fields.Update

'EMAIL DATA
With mymail
    .Subject = email_subject
    .From = emailADRS
    .To = email_recipient

    .HTMLBody = "Hello there," & "<br>" & "<br>" & email_body _
        & "<br>" & "<br>" & "Thanks," & "<br>" & "<br>" & "Zach"
    
    'ADD ATTACHMENTS
    For i = LBound(attachmentsarray) To UBound(attachmentsarray)
        .AddAttachment attachmentsarray(i)
    Next
    
    .send
End With

'RELEASE MEMORY
Set mymail = Nothing

'FIN
'Application.Speech.Speak "Gmail sent!"

'ERROR HANDLING
If Errors = 100 Then
errcatching:
    Application.Speech.Speak "Gmail issue detected!"
    
    'HIGHLIGHT PROBLEMATIC EMAIL
    Call email_highlight(email_recipient)
    
    Exit Sub
End If

End Sub
'HIGHLIGHT PROBLEMATIC EMAIL
Sub email_highlight(email_recipient)

'LOOP THROUGH WORKSHEETS
For Each Worksheet In Application.Workbooks
    If Not Cells.Find(email_recipient, lookat:=xlWhole) Is Nothing Then
        Cells.Find(email_recipient, lookat:=xlWhole).EntireRow.Interior.Color = vbRed
        Exit For
    End If
Next

End Sub
'EMAIL BODY TEXT
Function email_body_text()

    email_body_text = _
        "My name is Zach and I recently completed a contract focused on financial planning and analysis for Wells Fargo and I am an excellent fit for this role. " & _
        "My overall technical skills are very strong, and my knowledge of MS Office is advanced (particularly Excel and VB - I can code an algorithm to automate any task). " & _
        "I can handle big data, I possess strong systems and office administration experience, as well as processing invoices and GL accounting, not to mention supply chain/vendor relationship management, and logistics. " _
        & "I have a keen analytical mind and can make substantive process-improvement recommendations for your organization. " _
        & "I have significant automation experience (one of the many processes that I automated was invoice processing, which reduced the processing time required by ~90%). " & _
        "This role is very much aligned with my qualifications and career aspirations and I would love to set up a time to discuss this position and the contributions that I could make to your organization in greater detail."

End Function
'PROCESS TIMEFRAME TRACKING
Function PTT(xxtime, PTTmsg)
    
    Application.Speech.Speak PTTmsg & xxtime
    End

End Function
