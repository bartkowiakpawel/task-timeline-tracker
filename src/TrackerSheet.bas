'=============================================================
' Task Timeline Tracker
'
' Developed by: Pawel Bartkowiak
' AI assistance: ChatGPT by OpenAI
'
' This code belongs in the TRACKER worksheet module,
' not in a standard VBA module.
'=============================================================

Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    Dim wsUpdates As Worksheet
    Dim tblUpdates As ListObject
    Dim tblTracker As ListObject
    Dim taskRange As Range

    Dim taskName As String
    Dim taskColumn As Long

    On Error GoTo SafeExit

    Set wsUpdates = ThisWorkbook.Worksheets("UPDATES")
    Set tblUpdates = wsUpdates.ListObjects("tblUpdates")
    Set tblTracker = Me.ListObjects("tblTracker")

    ' More than one cell selected -> clear filter
    If Target.Cells.CountLarge > 1 Then
        ClearUpdatesFilter tblUpdates
        Exit Sub
    End If

    ' No rows in Tracker table
    If tblTracker.ListColumns("Task_name").DataBodyRange Is Nothing Then
        ClearUpdatesFilter tblUpdates
        Exit Sub
    End If

    Set taskRange = tblTracker.ListColumns("Task_name").DataBodyRange

    ' Clicked anywhere outside actual Task_name cells
    If Intersect(Target, taskRange) Is Nothing Then
        ClearUpdatesFilter tblUpdates
        Exit Sub
    End If

    ' Clicked an empty Task cell
    taskName = Trim(CStr(Target.Value))

    If taskName = "" Then
        ClearUpdatesFilter tblUpdates
        Exit Sub
    End If

    ' Remove previous filter
    ClearUpdatesFilter tblUpdates

    ' Find Task_name column in UPDATES
    taskColumn = tblUpdates.ListColumns("Task_name").Index

    ' Apply new filter
    tblUpdates.Range.AutoFilter _
        Field:=taskColumn, _
        Criteria1:=taskName

SafeExit:

End Sub


Private Sub ClearUpdatesFilter(ByVal tblUpdates As ListObject)

    On Error Resume Next

    tblUpdates.AutoFilter.ShowAllData

    On Error GoTo 0

End Sub
