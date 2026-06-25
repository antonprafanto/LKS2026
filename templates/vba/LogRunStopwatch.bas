Attribute VB_Name = "LogRunStopwatch"
Option Explicit

Private stopwatchRunning As Boolean
Private stopwatchStart As Date

Public Sub CatatSIAP()
    With Worksheets("Log Run Otonom").Range("B8")
        .Value = Time
        .NumberFormat = "hh:mm:ss"
    End With
End Sub

Public Sub CatatSTART()
    Dim ws As Worksheet
    Set ws = Worksheets("Log Run Otonom")
    With ws.Range("E8")
        .Value = Time
        .NumberFormat = "hh:mm:ss"
    End With
    stopwatchStart = Now
    stopwatchRunning = True
    ws.Range("K7").Value = 0
    ws.Range("K7").NumberFormat = "[h]:mm:ss"
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:01"), "LogRun_TickStopwatch", , False
    On Error GoTo 0
    Application.OnTime Now + TimeValue("00:00:01"), "LogRun_TickStopwatch"
End Sub

Public Sub CatatSELESAI()
    stopwatchRunning = False
    Dim ws As Worksheet
    Set ws = Worksheets("Log Run Otonom")
    With ws.Range("H8")
        .Value = Time
        .NumberFormat = "hh:mm:ss"
    End With
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:01"), "LogRun_TickStopwatch", , False
    On Error GoTo 0
End Sub

Public Sub LogRun_TickStopwatch()
    If Not stopwatchRunning Then Exit Sub
    Dim ws As Worksheet
    Set ws = Worksheets("Log Run Otonom")
    If ws.Range("H8").Value <> "" Then
        stopwatchRunning = False
        Exit Sub
    End If
    If ws.Range("E8").Value = "" Then Exit Sub
    ws.Range("K7").Value = Now - stopwatchStart
    ws.Range("K7").NumberFormat = "[h]:mm:ss"
    Application.OnTime Now + TimeValue("00:00:01"), "LogRun_TickStopwatch"
End Sub

Public Sub ResetWaktuRun()
    stopwatchRunning = False
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:01"), "LogRun_TickStopwatch", , False
    On Error GoTo 0
    Dim ws As Worksheet
    Set ws = Worksheets("Log Run Otonom")
    ws.Range("B8").ClearContents
    ws.Range("E8").ClearContents
    ws.Range("H8").ClearContents
    ws.Range("K7").ClearContents
End Sub
