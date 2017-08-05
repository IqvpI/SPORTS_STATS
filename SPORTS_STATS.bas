Attribute VB_Name = "SPORTS_STATS"
'SPORTS DATA EXTRACT SELECTION
Sub show_sports_form()

    Unload SPORTS_FORM
    SPORTS_FORM.Show vbModeless

End Sub
'MLB SPORTS STATISTICAL ANALYSIS
Sub MLB_STATS(dataset)

Application.ScreenUpdating = False
Application.DisplayAlerts = False

'DIMENSIONING
Dim objie As InternetExplorer
Dim link As Object, link1 As Object, links As Object, links1 As Object, _
link2 As Object, links2 As Object, elemcol As Object

Dim t As Integer, r As Integer, c As Integer
Dim webpage_loaded As Boolean
Dim starttime As Long, endtime As Long, xtime As Long
Dim rng As Range

Dim AL As String, NL As String, dscr As String
Dim DATARRAY(33, 1) As Variant, stats As Variant

'STATIC VARS
WBpath = "C:\Users\qp\Desktop\EXCELCIOR\sports stats\"
SSWB = "Sports stats.xlsb"
RowCount = 0
starttime = Timer()

'CHECK IF DB OPEN>OPEN IF NOT
WB = SSWB
path1 = WBpath
Call FILEOPEN(WB, path1)

'DEPENDENT VARS
If dataset = "PS" Then 'PLAYER STAS
    SFR = "RK"
    tofind = "rbi"
    SSws = "MLB"
    MLB_SITE = "MLB.com/stats"
    
    'COL VARS
    With Workbooks(SSWB).Worksheets(SSws)
        RANKcol = .Cells.Find("rank", lookat:=xlWhole).Column
        PLAYERcol = .Cells.Find("player", lookat:=xlWhole).Column
        PIDcol = .Cells.Find("pid", lookat:=xlWhole).Column
        teamcol = .Cells.Find("team", lookat:=xlWhole).Column
        POScol = .Cells.Find("pos", lookat:=xlWhole).Column
        Gcol = .Cells.Find("g", lookat:=xlWhole).Column
        ABcol = .Cells.Find("ab", lookat:=xlWhole).Column
        Rcol = .Cells.Find("r", lookat:=xlWhole).Column
        Hcol = .Cells.Find("h", lookat:=xlWhole).Column
        IIBcol = .Cells.Find("2b", lookat:=xlWhole).Column
        IIIBcol = .Cells.Find("3b", lookat:=xlWhole).Column
        HRcol = .Cells.Find("hr", lookat:=xlWhole).Column
        RBIcol = .Cells.Find("rbi", lookat:=xlWhole).Column
        BBcol = .Cells.Find("bb", lookat:=xlWhole).Column
        SOcol = .Cells.Find("so", lookat:=xlWhole).Column
        SBcol = .Cells.Find("sb", lookat:=xlWhole).Column
        CScol = .Cells.Find("cs", lookat:=xlWhole).Column
        AVGcol = .Cells.Find("avg", lookat:=xlWhole).Column
        OBPcol = .Cells.Find("obp", lookat:=xlWhole).Column
        SLGcol = .Cells.Find("slg", lookat:=xlWhole).Column
        OPScol = .Cells.Find("ops", lookat:=xlWhole).Column
        IBBcol = .Cells.Find("ibb", lookat:=xlWhole).Column
        HBPcol = .Cells.Find("hbp", lookat:=xlWhole).Column
        SACcol = .Cells.Find("sac", lookat:=xlWhole).Column
        SFcol = .Cells.Find("sf", lookat:=xlWhole).Column
        TBcol = .Cells.Find("tb", lookat:=xlWhole).Column
        XBHcol = .Cells.Find("xbh", lookat:=xlWhole).Column
        GDPcol = .Cells.Find("gdp", lookat:=xlWhole).Column
        GOcol = .Cells.Find("go", lookat:=xlWhole).Column
        AOcol = .Cells.Find("ao", lookat:=xlWhole).Column
        GO_AOcol = .Cells.Find("go_ao", lookat:=xlWhole).Column
        NPcol = .Cells.Find("np", lookat:=xlWhole).Column
        PAcol = .Cells.Find("pa", lookat:=xlWhole).Column
    End With
    
ElseIf dataset = "TS" Then 'TEAM STANDINGS
    SSws = "MLB Standings"
    MLB_SITE = "mlb.com/mlb/standings"
    AL = "American League"
    NL = "National League"
    
    'WIPE DATA
    Call wipe_data(SSWB, SSws)
    
    'COL VARS
    With Workbooks(SSWB).Worksheets(SSws)
        LEAGUEcol = .Cells.Find("league", lookat:=xlWhole).Column
        DIVcol = .Cells.Find("division", lookat:=xlWhole).Column
        teamcol = .Cells.Find("team", lookat:=xlWhole).Column
        WINcol = .Cells.Find("w", lookat:=xlWhole).Column
        LOSScol = .Cells.Find("L", lookat:=xlWhole).Column
        PCTcol = .Cells.Find("pct", lookat:=xlWhole).Column
        GBcol = .Cells.Find("gb", lookat:=xlWhole).Column
        ENUMcol = .Cells.Find("e#", lookat:=xlWhole).Column
        WCGBcol = .Cells.Find("wcgb", lookat:=xlWhole).Column
        LTENcol = .Cells.Find("L10", lookat:=xlWhole).Column
        STREAKcol = .Cells.Find("streak", lookat:=xlWhole).Column
        HOMEcol = .Cells.Find("home", lookat:=xlWhole).Column
        AWAYcol = .Cells.Find("away", lookat:=xlWhole).Column
        LGAMEcol = .Cells.Find("last game", lookat:=xlWhole).Column
        NGAMEcol = .Cells.Find("next game", lookat:=xlWhole).Column
        vsEcol = .Cells.Find("vs E", lookat:=xlWhole).Column
        vsCcol = .Cells.Find("vs C", lookat:=xlWhole).Column
        vsWcol = .Cells.Find("vs W", lookat:=xlWhole).Column
        vsAL_NLcol = .Cells.Find("vs AL/NL", lookat:=xlWhole).Column
        vsRcol = .Cells.Find("vs R", lookat:=xlWhole).Column
        vsLcol = .Cells.Find("vs L", lookat:=xlWhole).Column
        XTRAcol = .Cells.Find("xtra", lookat:=xlWhole).Column
        ONERUNcol = .Cells.Find("1-run", lookat:=xlWhole).Column
        RScol = .Cells.Find("rs", lookat:=xlWhole).Column
        RAcol = .Cells.Find("ra", lookat:=xlWhole).Column
        X_WLcol = .Cells.Find("x_wl", lookat:=xlWhole).Column
    End With
    
End If

'WIPE DB DATA
Call wipe_data(SSWB, SSws)

webmanip:
'WEB MANIPULATION
Set objie = New InternetExplorer
objie.Visible = True
objie.Width = 1200
objie.Height = 1800
objie.navigate MLB_SITE
Do While objie.Busy = True Or objie.readyState <> 4: DoEvents: Loop
rewait:
Application.Wait (Now + TimeValue("00:00:02"))

With objie.document
    
    'CONTROL FOR LOAD FAILURE
    tofind = "Can’t reach this page"
    webpage_loaded = checktext(objie, tofind)
    If webpage_loaded = True Then
        objie.Quit
        GoTo webmanip:
    End If
    
    'CONTROL FOR LONG LOAD TIMES
    If dataset = "PS" Then
        tofind = "RK"
    Else: tofind = tofind = "next game"
    End If
        webpage_loaded = checktext(objie, tofind)
        If webpage_loaded <> True Then
            GoTo rewait:
        End If
    
    'LOOP THROUGH TABLE>ROWS>COLUMNS>ADD DATA TO ARRAY
    Set elemcol = .getElementsByTagName("table")
    For t = 0 To (elemcol.Length - 1)
        For r = 0 To (elemcol(t).Rows.Length - 1)
        
            'CLEAR ARRAY
            Erase DATARRAY
        
            For c = 0 To (elemcol(t).Rows(r).Cells.Length - 1)
            
                If elemcol(t).Rows(r).Cells(c).innerText = "RK" Then
                    GoTo nextablerow:
                End If

                DATARRAY(c, 0) = elemcol(t).Rows(r).Cells(c).innerText
                DATARRAY(c, 1) = elemcol(t).Rows(1).Cells(c).className
next_cell:
            Next c 'CELL
            
            'ADD DATA TO DB
            With Workbooks(SSWB).Worksheets(SSws)
                nxrow = .Range("C10000").End(xlUp).Offset(1).Row 'FIND NEXT ROW
                
                'EXTRACT DATA FROM ARRAY
                For i = LBound(DATARRAY) To UBound(DATARRAY)
                    
                    'SKIP BLANK EMPTY ITEMS
                    If DATARRAY(i, 0) = "" Or DATARRAY(i, 1) = "" Then
                        GoTo nextarrayitem:
                    End If
                    
                    stats = DATARRAY(i, 0)

                    'END OF TABLE DATA CONTROL
                    If stats = "Su" Then
                    
                        'ADD LEAGUES
                        If dataset = "TS" Then
                            Call League_Parse(SSWB, SSws)
                        End If
                        
                        GoTo fin:
                    End If
                    
                    dscr = DATARRAY(i, 1)
                    
                    'DATA AVOIDANCE
                    avoid = ""
                    text = dscr
                    avoid = data_avoidance(dataset, SSWB, text)
                    If avoid = True Then
                        GoTo nextarrayitem:
                    End If
                    
                    'TRIM DESCRIPTION
                    dscr = Right(dscr, (Len(dscr) - InStr(1, dscr, "-", 1)))
                    
                    'CHECK FOR EXACT COLUMN MATCH
                    findcol = Null
                    WB = SSWB
                    ws = SSws
                    text = dscr
                    findcol = exact_check(WB, ws, text)
                    
                    'CHECK COLUMN EXCEPTIONS
                    If findcol = False Then
                    
                        WB = SSWB
                        text = dscr
                        findcol = column_exceptions(WB, text, dataset)
                        
                        'ID CORRECT COLUMN>ADD DATA
                        findcol = .Cells.Find(findcol, lookat:=xlWhole).Column
                        
                    End If
                  
                    'ADD DATA TO CELL
                    .Cells(nxrow, findcol) = "=" & """" & stats & """"
                    
                    'DIVISION PARSE
                    If dataset = "TS" Then
                        Call Division_Parse(SSWB, SSws, nxrow, DIVcol, stats)
                    End If
            
nextarrayitem:
                Next
            End With
nextablerow:
            RowCount = RowCount + 1
nextrow:
        Next r 'ROW
    Next t 'TABLE
    
    If dataset = "PS" Then
        'NEXT PAGE
        If RowCount = 51 Or RowCount = 102 Or RowCount = 153 Or RowCount = 204 Then
            buttontag = "button"
            buttonclass = "paginationWidget-next"
            Call buttonclick(objie, buttontag, buttonclass, buttontext)
            GoTo rewait:
        End If
    End If
    
End With
fin:
'FIN
objie.Quit

'ADD LINES/TIMESTAMP
With Workbooks(SSWB).Worksheets(SSws)
    .Cells.Find("last updated:", lookat:=xlWhole).Offset(0, 1) = Now
    .Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
End With

Workbooks(SSWB).Save

'PROCESS TIME-FRAME TRACKING
endtime = Timer()
xtime = endtime - starttime
xxtime = "Time elapse: " & xtime & " seconds!"

'APPLICATION SPEECH MSG CONTENT
If dataset = "PS" Then
    RowCount = RowCount - 4
    PTTmsg = "Player statistics data extract completed!" & RowCount & " rows processed!"
ElseIf dataset = "TS" Then
    PTTmsg = "American & National league team standings data extract completed!"
End If

Call PTT(xxtime, PTTmsg)
    
End Sub

'MLB STANDINGS
Sub MLB_STANDINGS()



End Sub
