Attribute VB_Name = "MHGA"
'SIGN UP FOR RAID
Sub RAIDSIGNUP()


'ERROR WARNING MESSAGE
On Error GoTo errorhandling:
If errorhandling = notlikely Then
errorhandling:
For Errv = 1 To 3
    Application.Speech.Speak "Error Detected!"
Next
End If

Dim objie As InternetExplorer
Dim objShellWindows As New SHDocVw.ShellWindows

Dim link As Object, links As Object, links1 As Object, optn As Object
Dim starttime As Long, endtime As Long, xtime As Long
Dim reps1 As Variant
Dim reps As Integer

'VARIABLES
MHGAsite = "http://makehordegreatagain.shivtr.com/events"
instanceswitch = "event_instance_id"
tofind = "You attended"
tofind1 = "You are attending"
repcount = 0
signedup = 0
starttime = Timer()

numreps:
'NUMBER OF REPETITIONS
reps1 = InputBox("Please enter the number of repetitions")
If reps1 = vbNullString Then
    MsgBox "Procedure terminated!", vbCritical + vbOKOnly
    End
End If
reps = reps1
mnthslct:
'MONTH('S) SELECTION
mnthspec = MsgBox("Start with current month?", vbQuestion + vbYesNoCancel)
If mnthspec = vbCancel Then
    GoTo numreps:
ElseIf mnthspec = vbNo Then
    months = InputBox("Please indicate how many months to proceed from current" _
    & vbCrLf & vbCrLf & "Ex: 2 months from current")
    If months = vbNullString Then
        GoTo mnthslct:
    End If
End If

'IE MANIPULATION
Set objie = New InternetExplorerMedium
objie.Visible = True
objie.Width = 1800
objie.Height = 1900
objie.navigate MHGAsite

'Do While objie.Busy = True Or objie.readyState <> 4: DoEvents: Loop

Application.Wait (Now + TimeValue("00:00:03"))

For Window = objShellWindows.Count To 1 Step -1

    Set objie = objShellWindows.Item(Window)
        If Not objie Is Nothing Then
            If InStr(1, objie.LocationURL, MHGAsite, 1) > 0 Then
                With objie.document
                
                    'MOVE TO ANOTHER MONTH?
                    If months <> "" Then
                        For monthcount = 1 To months
                            Set links = .getElementsByTagName("a")
                            For Each link In links
                                If InStr(1, link.Title, "Next Month", 1) > 0 Then
                                    link.Click
                                    Application.Wait (Now + TimeValue("00:00:01"))
                                    Exit For
                                End If
                            Next
                        Next
                    End If
                
                    'SEARCH FOR EVENTS
                    Set links = .getElementsByTagName("a")
                        For Each link In links
                            If InStr(1, link.innerText, "9p Raid Night", 1) > 0 Then
                                link.Click
                                Application.Wait (Now + TimeValue("00:00:02"))
                                Exit For
                            End If
                        Next
SIGNUPCHECK:
                        'SIGN UP FOR EVENT???
                        findText = InStr(1, .body.innerHTML, tofind, 1)
                        findText1 = InStr(1, .body.innerHTML, tofind1, 1)
                        If findText > 0 Or findText1 > 0 Then 'ALREADY SIGNED UP/ATTENDED
                            GoTo repcounting:
morereps:
                            Set links = .getElementsByTagName("a")
                            For Each link In links
                                If InStr(1, link.innerText, "switch instance", 1) > 0 Then
                                    link.Click
                                    Application.Wait (Now + TimeValue("00:00:01"))
                                    
                                    'ID SELECTED CURRENT DATE FROM DROPDOWN LIST
                                    Set links = .getElementsByTagName("option")
                                    For Each optn In links
                                        If optn.Selected = "True" Then
                                            optnhold = optn.Value
                                            Exit For
                                        End If
                                    Next
                                    
                                    'SELECT NEXT RAID DAY FROM LIST
                                    Application.Wait (Now + TimeValue("00:00:02"))
                                    .getElementById("event_instance_id").Value = optnhold + 1
                                    .getElementById("event_instance_id").FireEvent ("onchange")
                                    Application.Wait (Now + TimeValue("00:00:03"))
                                    GoTo SIGNUPCHECK:
                                End If
                            Next
                        Else 'NEED TO SIGN UP
                            Set links = .getElementsByTagName("a")
                            For Each button In links
                                If button.text = "Yes" Then
                                    button.Click 'YES BUTTON
                                    Application.Wait (Now + TimeValue("00:00:02"))
                                    .getElementById("event_participant_role").Value = "" 'CLASS ROLE
                                    Set links = .getElementsByTagName("input")
                                    For Each buttons In links
                                        If buttons.Value = "Signup" Then
                                            buttons.Click 'SIGNUP BUTTON
                                            signedup = signedup + 1
                                            GoTo repcounting:
                                        End If
                                    Next
                                End If
                            Next
                        End If
repcounting:
                        'REPETITIONS TRACKER
                        repcount = repcount + 1
                        If repcount <> reps Then
                            GoTo morereps:
                        Else: GoTo PTT:
                        End If
                        
                End With
            End If
        End If
                
Next
PTT:
'PROCESSING TIME
endtime = Timer()
xtime = endtime - starttime
If signedup <> 0 Then
    msg1 = "You have signed up for " & signedup & " more raids!"
Else
    msg1 = "You have not signed up for any additional raids!"
End If

Application.Speech.Speak msg1

'MsgBox msg1 & vbCrLf & vbCrLf & "Time elapsed: " & xtime & " seconds!", vbInformation + vbOKOnly

objie.Quit
End Sub

