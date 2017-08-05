Attribute VB_Name = "Craigslist"
'CRAIGSLIST APPLICATION AUTOMATION
Sub craigapply()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

'DIMENSIONING
Dim objie As InternetExplorer
Dim objShellWindows As New SHDocVw.ShellWindows
Dim link As Object, link1 As Object, links As Object, links1 As Object, _
link2 As Object, links2 As Object, link0 As Object, links0 As Object, _
Clink As Object, Clinks As Object, clink0 As Object, clinks0 As Object

Dim JOBID As String
Dim postingsnum As Integer

'Dim ar1 As Variant
'ReDim ar1(1, 1)

Dim DATARRAY(6) As Variant
Dim starttime As Long, endtime As Long, xtime As Long
Dim catchafound As Boolean

'VARIABLES
accountingnav = "https://sfbay.craigslist.org/search/sfc/acc"
JLWB = "jobs log.xlsb"
Jws = "Jobs"
ESws = "External Sites"
JLpath = "C:\Users\qp\Desktop\EXCELCIOR\JOB SEARCH\"
job_listings = 0
job_applied = 0
email_issues = 0
external_site = 0
postreviewed = 0
current_listing = 0
starttime = Timer()

'# OF POSTINGS TO REVIEW
postingsnum1 = InputBox("Please enter the number of job posts to review")
If postingsnum1 = vbNullString Then
    Application.Speech.Speak "Procedure terminated!"
    End
End If
postingsnum = postingsnum1

'CHECK IF DB OPEN>OPEN IF NOT
WB = JLWB
path1 = JLpath
Call FILEOPEN(WB, path1)

'INITIALIZE COLUMN VARS
With Workbooks(JLWB).Worksheets(Jws)
    IDcol = .Cells.Find("id", lookat:=xlWhole).Column
    dataIDcol = .Cells.Find("data-id", lookat:=xlWhole).Column
    PSTDcol = .Cells.Find("date posted", lookat:=xlWhole).Column
    APPLYcol = .Cells.Find("date applied", lookat:=xlWhole).Column
    SRCcol = .Cells.Find("source", lookat:=xlWhole).Column
    CONTACTcol = .Cells.Find("contact", lookat:=xlWhole).Column
    URLcol = .Cells.Find("posting url", lookat:=xlWhole).Column
    TITLEcol = .Cells.Find("title", lookat:=xlWhole).Column
End With

'WEB MANIPULATION
Set objie = New InternetExplorer
objie.Visible = True
objie.Width = 1200
objie.Height = 1800
objie.navigate accountingnav

Do While objie.Busy = True Or objie.readyState <> 4: DoEvents: Loop
Application.StatusBar = "Navigating..."
    With objie.document
        .getElementById("subArea").Value = "sfc"
        
        'COUNT JOB LISTINGS
        Set Clinks = .getElementsByTagName("li")
        For Each Clink In Clinks
            If Clink.className = "result-row" Then
                job_listings = job_listings + 1
            End If
        Next
parselistings:
        'POSTING REVIEW CONTROL
        If postreviewed = postingsnum Then
            GoTo fin:
        End If
        
        'PARSE JOBS LISTINGS
        Set links = .getElementsByTagName("a")
        For Each link In links
            If link.className = "result-title hdrlnk" _
            And link.ParentNode.className = "result-info" _
            And link.ParentNode.ParentNode.className <> "result-row banished" Then
                    
                'CURRENT LISTING COUNTER
                current_listing = current_listing + 1
                         
                'JOB TITLE
                JOBTITLE = link.text
                    
                'JOB DATA-ID
                JOBID = link.getAttribute("data-id")
                    'VERIFY JOB ID NOT PRESENT IN DATABASE
                    tocheck = JOBID
                    DBchecked1 = DBcheck(tocheck)
                    
                    'REPOST???
                    If link.ParentNode.ParentNode.getAttribute("data-repost-of") <> 0 Then
                        repostID = link.ParentNode.ParentNode.getAttribute("data-repost-of")
                        tocheck = repostID
                        DBchecked2 = DBcheck(tocheck)
                    End If
                    
                    'ID FOUND>NEXT POST
                    If DBchecked1 = True Or DBchecked2 = True Then
                        GoTo nextpost:
                    End If
                    
                    'TIMESTAMP
                    dataid = JOBID
                    tagname1 = "time"
                    classname1 = "result-date"
                    attributename = "datetime"
                    PNclassname = "result-row"
                    POSTED = attextract(objie, tagname1, classname1, _
                        PNclassname, dataid, attributename, attributenameVAL)
                    
                    'ACCESS/EVALUATE LISTING
                    link.Click
                    Do While objie.Busy = True Or objie.readyState <> 4: DoEvents: Loop
                    
                    'JOB POSTING URL
                    JOBURL = objie.LocationURL
                    
                    'CHECK IF EXTERNAL JOB SITE
                    textfound = ""
                    tofind = "reply below"
                    textfound = checktext(objie, tofind)
                    
                    'CANT EMAIL DIRECTLY: MOVE ON TO NEXT POSTING
                    If textfound = True Then
                    
                        'CHECK IF EMAIL LISTED ON PAGE
                        textfound0 = ""
                        tofind = "@"
                        
                        findText = InStr(1, .body.innerText, tofind, 1)
                        If findText > 0 Then
                            textfound0 = True
                        End If
                        
'                        textfound0 = checktext(objie, tofind)

                        'EXTRACT EMAIL LISTED ON PAGE
                        If textfound0 = True Then
                            fp = InStr(1, .body.innerText, tofind, 1)
                            lp = InStr(fp, .body.innerText, " ", 1)
                            fp = InStrRev(.body.innerText, " ", fp, 1) + 1
                            JOBEMAIL = Mid(.body.innerText, fp, lp - fp)
                            GoTo addtoarray:
                        End If
                        
                        'CHECK IF EMAIL LISTED ON PAGE
                        textfound1 = ""
                        tofind = "http://"
                        textfound1 = checktext(objie, tofind)
                        
                        'EXTRACT URL IF LISTED ON PAGE
                        If textfound1 = True Then
                            emailstring = ""
                            tagname1 = "a"
                            attributename = "rel"
                            attributenameVAL = "nofollow"
                            emailstring = attextract(objie, tagname1, classname1, _
                                PNclassname, dataid, attributename, attributenameVAL)
                                
'                            'EXTRACT HREF
'                            If InStr(1, emailstring, " ", 1) > 0 Then
'                                tagname1 = "a"
'                                attributename = "ref"
'                                emailstring = attextract(objie, tagname1, classname1, _
'                                PNclassname, dataid, attributename, attributenameVAL)
'                            End If
'
                            JOBEMAIL = emailstring
                        End If
                        
                    Else: 'DIRECT EMAIL AVAILABLE
                    
                        'EXTRACT REPLY EMAIL
                        buttontag = "button"
                        buttonclass = "reply_button js-only"
                        buttontext = "reply"
                        Call buttonclick(objie, buttontag, buttonclass, buttontext)
                        
                        'CAPTCHA DETECTED?
'                        catchafound = ""
                        tofind = "I'm not a robot"
                        catchafound = checktext(objie, tofind)
                        If catchafound = True Then
                            Application.Speech.Speak "Captcha detected! Procedure terminated!"
                            GoTo fin:
                        End If
                        
'                        .getElementById("replylink").Click
                        Do While objie.Busy = True Or objie.readyState <> 4: DoEvents: Loop
                        Application.Wait (Now + TimeValue("00:00:07"))
                        
'                        .getElementById("recaptcha-anchor").Click
'
'                        Set clinks0 = .getElementsByTagName("input")
'                        For Each Item In clinks0
'                            If clink0.role = "checkbox" Then
'                                fdsfdfds = 1
'                            End If
'                        Next
'
'                        Set Clinks = .getElementsByTagName("a")
'                        For Each Clink In Clinks
'                            If Clink.className = "mailapp" Then
'                                JOBEMAIL = Clink.innerHTML
'                                Exit For
'                            End If
'                        Next
'
'                        objie.GoBack
'                        Do While objie.Busy = True Or objie.readyState <> 4: DoEvents: Loop
                        
                        'PREVIOUS EMAIL EXTRACT METHOD
                        Set links2 = .getElementsByTagName("p")
                        For Each link2 In links2
                            If link2.className = "anonemail" Then
                                JOBEMAIL = link2.innerText
                                Exit For
                            End If
                        Next
                    End If
                    
'                'EXTERNAL SITE
'                If textfound0 = False And textfound1 = False _
'                And JOBEMAIL = "" Then
'                    Application.Speech.Speak "External site identified!"
'                    external_site = external_site + 1
'                    Jws = "external sites"
'                End If
addtoarray:
                'EXTERNAL SITE
                If JOBEMAIL = "" _
                Or InStr(1, JOBEMAIL, "@", 1) = 0 Then
                    external_site = external_site + 1
                    Jws = "external sites"
                End If

                'ADD DATA TO ARRAY
                DATARRAY(0) = JOBTITLE
                DATARRAY(1) = JOBID
                DATARRAY(2) = POSTED
                DATARRAY(3) = JOBURL
                DATARRAY(4) = JOBEMAIL
                  
                'ADD DATA TO DATABASE
                With Workbooks(JLWB).Worksheets(Jws)
                
                    nxrow = .Range("A1000").End(xlUp).Offset(1).Row 'FIND NEXT ROW

                    'ADD ENTRY ID
                    If Not IsNumeric(.Cells(nxrow, IDcol).Offset(-1)) Then
                        .Cells(nxrow, IDcol) = 1
                    Else: .Cells(nxrow, IDcol) = .Cells(nxrow, IDcol).Offset(-1) + 1
                    End If

                    .Cells(nxrow, dataIDcol) = DATARRAY(1) 'DATA ID
                    .Cells(nxrow, PSTDcol) = DATARRAY(2) 'DATE POSTED
                    .Cells(nxrow, APPLYcol) = Date 'DATE APPLIED
                    .Cells(nxrow, SRCcol) = "Craigslist" 'SOURCE
                    .Cells(nxrow, CONTACTcol) = DATARRAY(4) 'JOB EMAIL
                    .Cells(nxrow, URLcol) = DATARRAY(3) 'URL
                    .Cells(nxrow, TITLEcol) = DATARRAY(0) 'TITLE
                    
                    'RETURN SHEET REFERENCE
                    If Jws = "external sites" Then
                        Jws = "Jobs"
                    End If
                End With
                
                'SEND EMAIL APPLICATION
                If InStr(1, JOBEMAIL, "@", 1) > 0 Then
                    email_subject = JOBTITLE
                    email_recipient = JOBEMAIL
                    email_body = email_body_text
                    Call send_email(email_subject, email_recipient, email_body)
                    
                    'RECORD SUCCESSFUL/FAILED EMAIL OUTREACH
                    If Workbooks(JLWB).Worksheets(Jws).Rows(nxrow).Interior.Color <> vbRed Then
                        job_applied = job_applied + 1
                    Else: email_issues = email_issues + 1
                    End If
                    
                    JOBEMAIL = ""
                End If
            
                'BANISH (HIDE) POSTING
                buttontag = "span"
                buttonclass = "banish"
                Call buttonclick(objie, buttontag, buttonclass, buttontext)
                
                
                'RETURN TO LISTINGS SIMPLE
                objie.GoBack
                Do While objie.Busy = True Or objie.readyState <> 4: DoEvents: Loop
                postreviewed = postreviewed + 1
                Application.Wait (Now + TimeValue("00:00:06"))
                
'                'RETURN TO LISTINGS
'                Set links0 = .getElementsByTagName("a")
'                For Each link0 In links0
'                    If link0.innerText = "accounting/finance" Then
'                        link0.Click
'                        Do While objie.Busy = True Or objie.readyState <> 4: DoEvents: Loop
'                        postreviewed = postreviewed + 1
'                        Application.Wait (Now + TimeValue("00:00:06"))
'                        Exit For
'                    End If
'                Next

                'ALL POSTS REVIEWED>MOVE TO NEXT PAGE
                If current_listing = job_listings Then
                    buttontag = "a"
                    buttonclass = "button next"
                    Call buttonclick(objie, buttontag, buttonclass, buttontext)
                    GoTo parselistings:
                Else: GoTo parselistings:
                End If

            End If
nextpost:
        Next
        
    End With
fin:
    'NAV TO BANISHED POSTS
    buttontag = "span"
    buttonclass = "icon icon-trash red"
    buttontext = "hidden"
    Call buttonclick(objie, buttontag, buttonclass, buttontext)
    'RESTORE BANISHED POSTS
    buttontag = "a"
    buttonclass = "clear-all-banished"
    buttontext = "unhide all"
    Call buttonclick(objie, buttontag, buttonclass, buttontext)
    Application.Wait (Now + TimeValue("00:00:01"))
    
    'FIN
    objie.Quit
    Workbooks(JLWB).Worksheets(Jws).Columns.AutoFit
    Workbooks(JLWB).Worksheets("Jobs").Activate
    Workbooks(JLWB).Save
'    Workbooks(JLWB).Close
    
    endtime = Timer()
    xtime = endtime - starttime
    Application.Speech.Speak "Process completed! Time elapsed: " & xtime & " seconds!" _
        & job_applied & " applications sent!" & external_site & " external sites identified!" _
        & email_issues & " invalied email addresses identified!"
    
End Sub
