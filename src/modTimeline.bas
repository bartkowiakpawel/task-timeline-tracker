'=============================================================
' Task Timeline Tracker
'
' Developed by: Pawel Bartkowiak
' AI assistance: ChatGPT by OpenAI
'
' Lightweight Excel/VBA tracker for task lifecycle monitoring,
' update history, weekly timeline visualization and validation.
'=============================================================

Option Explicit

Private Const HEADER_ROW As Long = 5
Private Const FIRST_DATA_ROW As Long = 6
Private Const FIRST_TIMELINE_COL As Long = 4
Private Const UNASSIGNED_GROUP As String = "Unassigned"


Public Sub GenerateTimeline()

    Dim wsTracker As Worksheet
    Dim wsUpdates As Worksheet
    Dim wsTimeline As Worksheet

    Dim lastTrackerRow As Long
    Dim lastUpdateRow As Long

    Dim trackerData As Variant
    Dim updatesData As Variant
    Dim order() As Long

    Dim dictTasks As Object
    Dim dictTimelineRows As Object
    Dim dictTaskEndDates As Object

    Dim validationErrors As String

    Dim r As Long
    Dim u As Long
    Dim i As Long
    Dim c As Long

    Dim sourceIndex As Long
    Dim outRow As Long

    Dim taskName As String
    Dim groupName As String
    Dim currentGroup As String

    Dim creationDate As Variant
    Dim startDate As Variant
    Dim endDate As Variant

    Dim updateTask As String
    Dim updateDate As Variant
    Dim updateType As String
    Dim updateReporter As String
    Dim updateComment As String

    Dim minDate As Date
    Dim maxDate As Date
    Dim firstWeek As Date
    Dim weekStart As Date
    Dim weekEnd As Date

    Dim lastTimelineCol As Long
    Dim lastTimelineRow As Long
    Dim firstGroupTaskRow As Long

    Dim targetRow As Long
    Dim targetCell As Range

    Dim noteText As String
    Dim key As String
    Dim trackerEndDate As Variant

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    Set wsTracker = ThisWorkbook.Worksheets("TRACKER")
    Set wsUpdates = ThisWorkbook.Worksheets("UPDATES")
    Set wsTimeline = ThisWorkbook.Worksheets("TIMELINE")

    Set dictTasks = CreateObject("Scripting.Dictionary")
    Set dictTimelineRows = CreateObject("Scripting.Dictionary")
    Set dictTaskEndDates = CreateObject("Scripting.Dictionary")

    dictTasks.CompareMode = vbTextCompare
    dictTimelineRows.CompareMode = vbTextCompare
    dictTaskEndDates.CompareMode = vbTextCompare

    lastTrackerRow = wsTracker.Cells(wsTracker.Rows.Count, "A").End(xlUp).Row
    lastUpdateRow = wsUpdates.Cells(wsUpdates.Rows.Count, "A").End(xlUp).Row

    If lastTrackerRow < 2 Then
        MsgBox "TRACKER does not contain any tasks.", vbExclamation, "Nothing to generate"
        GoTo CleanExit
    End If

    ' Read source data into memory. TRACKER and UPDATES remain read-only.
    trackerData = wsTracker.Range("A2:G" & lastTrackerRow).Value

    If lastUpdateRow >= 2 Then
        updatesData = wsUpdates.Range("A2:E" & lastUpdateRow).Value
    End If

    '=========================================================
    ' VALIDATION - TRACKER
    '=========================================================

    validationErrors = ""

    For r = 1 To UBound(trackerData, 1)

        taskName = Trim(CStr(trackerData(r, 1)))

        If taskName = "" Then

            If RowContainsData(trackerData, r, 2, 7) Then
                validationErrors = validationErrors & vbCrLf & _
                    "- TRACKER row " & (r + 1) & ": Task name is missing."
            End If

        Else

            key = NormalizeKey(taskName)

            If dictTasks.Exists(key) Then

                validationErrors = validationErrors & vbCrLf & _
                    "- Duplicate task in TRACKER: " & taskName

            Else

                dictTasks.Add key, True
                dictTaskEndDates.Add key, trackerData(r, 4)

            End If

        End If

    Next r

    '=========================================================
    ' VALIDATION - UPDATES
    '=========================================================

    If lastUpdateRow >= 2 Then

        For u = 1 To UBound(updatesData, 1)

            updateTask = Trim(CStr(updatesData(u, 1)))

            If updateTask = "" Then

                If RowContainsData(updatesData, u, 2, 5) Then
                    validationErrors = validationErrors & vbCrLf & _
                        "- UPDATES row " & (u + 1) & ": Task name is missing."
                End If

            Else

                key = NormalizeKey(updateTask)

                If Not dictTasks.Exists(key) Then

                    validationErrors = validationErrors & vbCrLf & _
                        "- Task from UPDATES not found in TRACKER: " & updateTask

                Else

                    If Not IsDate(updatesData(u, 2)) Then

                        validationErrors = validationErrors & vbCrLf & _
                            "- UPDATES row " & (u + 1) & ": Update Date is missing or invalid."

                    Else

                        trackerEndDate = dictTaskEndDates(key)

                        If IsDate(trackerEndDate) Then

                            If CDate(updatesData(u, 2)) > CDate(trackerEndDate) Then
                                validationErrors = validationErrors & vbCrLf & _
                                    "- Update after task closure: " & updateTask & _
                                    " | Update Date: " & Format(CDate(updatesData(u, 2)), "dd-mmm-yyyy") & _
                                    " | End Date: " & Format(CDate(trackerEndDate), "dd-mmm-yyyy")
                            End If

                        End If

                    End If

                End If

            End If

        Next u

    End If

    If validationErrors <> "" Then
        MsgBox "Timeline cannot be generated." & vbCrLf & vbCrLf & _
            "Please fix the following issues:" & vbCrLf & validationErrors, _
            vbCritical, "Validation failed"
        GoTo CleanExit
    End If

    '=========================================================
    ' CREATE SORT ORDER IN MEMORY
    '=========================================================

    ReDim order(1 To UBound(trackerData, 1))

    For i = 1 To UBound(order)
        order(i) = i
    Next i

    QuickSortTracker trackerData, order, LBound(order), UBound(order)

    '=========================================================
    ' DATE RANGE
    '=========================================================

    minDate = Date
    maxDate = Date

    For r = 1 To UBound(trackerData, 1)
        UpdateDateRange trackerData(r, 2), minDate, maxDate
        UpdateDateRange trackerData(r, 3), minDate, maxDate
        UpdateDateRange trackerData(r, 4), minDate, maxDate
    Next r

    If lastUpdateRow >= 2 Then
        For u = 1 To UBound(updatesData, 1)
            UpdateDateRange updatesData(u, 2), minDate, maxDate
        Next u
    End If

    firstWeek = minDate - Weekday(minDate, vbMonday) + 1

    '=========================================================
    ' CLEAR ONLY TIMELINE AREA
    ' Rows 1-4 remain available for buttons / labels.
    '=========================================================

    On Error Resume Next
    wsTimeline.Cells.ClearOutline
    wsTimeline.Rows(HEADER_ROW & ":" & wsTimeline.Rows.Count).ClearComments
    wsTimeline.Rows(HEADER_ROW & ":" & wsTimeline.Rows.Count).Clear
    On Error GoTo ErrorHandler

    '=========================================================
    ' STATIC HEADERS
    '=========================================================

    wsTimeline.Cells(HEADER_ROW, "A").Value = "Task"
    wsTimeline.Cells(HEADER_ROW, "B").Value = "Status"
    wsTimeline.Cells(HEADER_ROW, "C").Value = "Owner"

    With wsTimeline.Range(wsTimeline.Cells(HEADER_ROW, "A"), wsTimeline.Cells(HEADER_ROW, "C"))
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With

    '=========================================================
    ' WEEK HEADERS
    '=========================================================

    c = FIRST_TIMELINE_COL
    weekStart = firstWeek

    Do While weekStart <= maxDate

        wsTimeline.Cells(HEADER_ROW, c).Value = weekStart
        wsTimeline.Cells(HEADER_ROW, c).NumberFormat = "dd-mmm"

        With wsTimeline.Cells(HEADER_ROW, c)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        weekStart = weekStart + 7
        c = c + 1

    Loop

    lastTimelineCol = c - 1

    ' Remove stale cell fills from the control area.
    With wsTimeline.Range(wsTimeline.Cells(1, 1), wsTimeline.Cells(HEADER_ROW - 1, lastTimelineCol))
        .Interior.Pattern = xlNone
    End With

    '=========================================================
    ' BUILD TIMELINE
    '=========================================================

    outRow = FIRST_DATA_ROW
    currentGroup = ""
    firstGroupTaskRow = 0

    For i = 1 To UBound(order)

        sourceIndex = order(i)

        taskName = Trim(CStr(trackerData(sourceIndex, 1)))
        groupName = NormalizeGroupName(CStr(trackerData(sourceIndex, 7)))

        If taskName <> "" Then

            If StrComp(groupName, currentGroup, vbTextCompare) <> 0 Then

                ' Close previous group.
                If currentGroup <> "" And firstGroupTaskRow > 0 Then
                    If firstGroupTaskRow <= outRow - 1 Then
                        wsTimeline.Rows(firstGroupTaskRow & ":" & outRow - 1).Rows.Group
                    End If
                End If

                If outRow > FIRST_DATA_ROW Then
                    outRow = outRow + 1
                End If

                currentGroup = groupName

                ' Group header.
                wsTimeline.Cells(outRow, "A").Value = currentGroup

                With wsTimeline.Range(wsTimeline.Cells(outRow, 1), wsTimeline.Cells(outRow, lastTimelineCol))
                    .Font.Bold = True
                    .Interior.Color = RGB(220, 230, 241)
                End With

                outRow = outRow + 1
                firstGroupTaskRow = outRow

            End If

            ' Task row.
            wsTimeline.Cells(outRow, "A").Value = taskName
            wsTimeline.Cells(outRow, "B").Value = trackerData(sourceIndex, 5)
            wsTimeline.Cells(outRow, "C").Value = trackerData(sourceIndex, 6)

            dictTimelineRows.Add NormalizeKey(taskName), outRow

            creationDate = trackerData(sourceIndex, 2)
            startDate = trackerData(sourceIndex, 3)
            endDate = trackerData(sourceIndex, 4)

            For c = FIRST_TIMELINE_COL To lastTimelineCol

                weekStart = CDate(wsTimeline.Cells(HEADER_ROW, c).Value)
                weekEnd = weekStart + 6

                ' Waiting: Creation Date -> Start Date.
                If IsDate(creationDate) Then

                    If IsDate(startDate) Then

                        If CDate(creationDate) <= weekEnd And CDate(startDate) > weekStart Then
                            wsTimeline.Cells(outRow, c).Interior.Color = RGB(217, 217, 217)
                        End If

                    Else

                        If CDate(creationDate) <= weekEnd And weekStart <= Date Then
                            wsTimeline.Cells(outRow, c).Interior.Color = RGB(217, 217, 217)
                        End If

                    End If

                End If

                ' Active/completed: Start Date -> End Date or Today.
                If IsDate(startDate) Then

                    If IsDate(endDate) Then

                        If CDate(startDate) <= weekEnd And CDate(endDate) >= weekStart Then
                            wsTimeline.Cells(outRow, c).Interior.Color = RGB(91, 155, 213)
                        End If

                    Else

                        If CDate(startDate) <= weekEnd And Date >= weekStart Then
                            wsTimeline.Cells(outRow, c).Interior.Color = RGB(91, 155, 213)
                        End If

                    End If

                End If

            Next c

            outRow = outRow + 1

        End If

    Next i

    ' Close final group.
    If currentGroup <> "" And firstGroupTaskRow > 0 Then
        If firstGroupTaskRow <= outRow - 1 Then
            wsTimeline.Rows(firstGroupTaskRow & ":" & outRow - 1).Rows.Group
        End If
    End If

    lastTimelineRow = outRow - 1

    '=========================================================
    ' ADD UPDATES
    '=========================================================

    If lastUpdateRow >= 2 Then

        For u = 1 To UBound(updatesData, 1)

            updateTask = Trim(CStr(updatesData(u, 1)))
            updateDate = updatesData(u, 2)
            updateType = Trim(CStr(updatesData(u, 3)))
            updateReporter = Trim(CStr(updatesData(u, 4)))
            updateComment = Trim(CStr(updatesData(u, 5)))

            If updateTask <> "" And IsDate(updateDate) Then

                key = NormalizeKey(updateTask)

                If dictTimelineRows.Exists(key) Then

                    targetRow = dictTimelineRows(key)

                    For c = FIRST_TIMELINE_COL To lastTimelineCol

                        weekStart = CDate(wsTimeline.Cells(HEADER_ROW, c).Value)
                        weekEnd = weekStart + 6

                        If CDate(updateDate) >= weekStart And CDate(updateDate) <= weekEnd Then

                            Set targetCell = wsTimeline.Cells(targetRow, c)

                            If targetCell.Value = "" Then
                                targetCell.Value = "!"
                            Else
                                targetCell.Value = CStr(targetCell.Value) & "!"
                            End If

                            With targetCell
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter
                                .Font.Bold = True
                            End With

                            noteText = Format(CDate(updateDate), "dd-mmm-yyyy")

                            If updateType <> "" Then
                                noteText = noteText & " | " & updateType
                            End If

                            If updateReporter <> "" Then
                                noteText = noteText & " | " & updateReporter
                            End If

                            If updateComment <> "" Then
                                noteText = noteText & vbCrLf & updateComment
                            End If

                            If targetCell.Comment Is Nothing Then
                                targetCell.AddComment noteText
                            Else
                                targetCell.Comment.Text targetCell.Comment.Text & vbCrLf & vbCrLf & noteText
                            End If

                            With targetCell.Comment.Shape
                                .Width = 220
                                .Height = 140
                            End With

                            Exit For

                        End If

                    Next c

                End If

            End If

        Next u

    End If

    '=========================================================
    ' HIGHLIGHT CURRENT WEEK
    '=========================================================

    For c = FIRST_TIMELINE_COL To lastTimelineCol

        weekStart = CDate(wsTimeline.Cells(HEADER_ROW, c).Value)
        weekEnd = weekStart + 6

        If Date >= weekStart And Date <= weekEnd Then
            With wsTimeline.Cells(HEADER_ROW, c)
                .Interior.Color = RGB(255, 235, 156)
                .Font.Bold = True
            End With
        End If

    Next c

    '=========================================================
    ' FORMATTING
    '=========================================================

    wsTimeline.Outline.SummaryRow = xlAbove
    wsTimeline.Columns("A:C").AutoFit

    If wsTimeline.Columns("A").ColumnWidth > 45 Then
        wsTimeline.Columns("A").ColumnWidth = 45
    End If

    For c = FIRST_TIMELINE_COL To lastTimelineCol
        wsTimeline.Columns(c).ColumnWidth = 9
    Next c

    wsTimeline.Columns("A:C").VerticalAlignment = xlCenter

    If lastTimelineRow >= FIRST_DATA_ROW And lastTimelineCol >= FIRST_TIMELINE_COL Then
        wsTimeline.Range( _
            wsTimeline.Cells(HEADER_ROW, FIRST_TIMELINE_COL), _
            wsTimeline.Cells(lastTimelineRow, lastTimelineCol) _
        ).HorizontalAlignment = xlCenter
    End If

    ' Freeze rows 1-5 and columns A-C.
    wsTimeline.Activate
    ActiveWindow.FreezePanes = False
    wsTimeline.Cells(FIRST_DATA_ROW, FIRST_TIMELINE_COL).Select
    ActiveWindow.FreezePanes = True

    MsgBox "Timeline generated successfully." & vbCrLf & vbCrLf & _
        "TRACKER and UPDATES were not modified.", vbInformation, "Done"

CleanExit:

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Error " & Err.Number & ":" & vbCrLf & Err.Description, _
        vbCritical, "Generate Timeline"

End Sub


Private Function NormalizeKey(ByVal value As String) As String
    NormalizeKey = UCase$(Trim$(value))
End Function


Private Function NormalizeGroupName(ByVal value As String) As String

    Dim result As String

    result = Trim$(value)

    If result = "" Then
        result = UNASSIGNED_GROUP
    End If

    NormalizeGroupName = result

End Function


Private Function RowContainsData( _
    ByRef data As Variant, _
    ByVal rowNumber As Long, _
    ByVal firstCol As Long, _
    ByVal lastCol As Long _
) As Boolean

    Dim c As Long

    For c = firstCol To lastCol
        If Trim(CStr(data(rowNumber, c))) <> "" Then
            RowContainsData = True
            Exit Function
        End If
    Next c

    RowContainsData = False

End Function


Private Sub UpdateDateRange( _
    ByVal value As Variant, _
    ByRef minDate As Date, _
    ByRef maxDate As Date _
)

    If IsDate(value) Then

        If CDate(value) < minDate Then
            minDate = CDate(value)
        End If

        If CDate(value) > maxDate Then
            maxDate = CDate(value)
        End If

    End If

End Sub


Private Sub QuickSortTracker( _
    ByRef data As Variant, _
    ByRef order() As Long, _
    ByVal first As Long, _
    ByVal last As Long _
)

    Dim low As Long
    Dim high As Long
    Dim pivot As Long
    Dim temp As Long

    low = first
    high = last
    pivot = order((first + last) \ 2)

    Do While low <= high

        Do While low <= last
            If CompareTrackerRows(data, order(low), pivot) >= 0 Then Exit Do
            low = low + 1
        Loop

        Do While high >= first
            If CompareTrackerRows(data, order(high), pivot) <= 0 Then Exit Do
            high = high - 1
        Loop

        If low <= high Then
            temp = order(low)
            order(low) = order(high)
            order(high) = temp
            low = low + 1
            high = high - 1
        End If

    Loop

    If first < high Then QuickSortTracker data, order, first, high
    If low < last Then QuickSortTracker data, order, low, last

End Sub


Private Function CompareTrackerRows( _
    ByRef data As Variant, _
    ByVal rowA As Long, _
    ByVal rowB As Long _
) As Long

    Dim groupA As String
    Dim groupB As String
    Dim taskA As String
    Dim taskB As String
    Dim dateA As Variant
    Dim dateB As Variant

    groupA = NormalizeGroupName(CStr(data(rowA, 7)))
    groupB = NormalizeGroupName(CStr(data(rowB, 7)))

    taskA = Trim(CStr(data(rowA, 1)))
    taskB = Trim(CStr(data(rowB, 1)))

    dateA = data(rowA, 2)
    dateB = data(rowB, 2)

    ' Unassigned always last.
    If StrComp(groupA, UNASSIGNED_GROUP, vbTextCompare) = 0 And _
       StrComp(groupB, UNASSIGNED_GROUP, vbTextCompare) <> 0 Then
        CompareTrackerRows = 1
        Exit Function
    End If

    If StrComp(groupA, UNASSIGNED_GROUP, vbTextCompare) <> 0 And _
       StrComp(groupB, UNASSIGNED_GROUP, vbTextCompare) = 0 Then
        CompareTrackerRows = -1
        Exit Function
    End If

    ' Group.
    If StrComp(groupA, groupB, vbTextCompare) < 0 Then
        CompareTrackerRows = -1
        Exit Function
    End If

    If StrComp(groupA, groupB, vbTextCompare) > 0 Then
        CompareTrackerRows = 1
        Exit Function
    End If

    ' Creation Date.
    If IsDate(dateA) And IsDate(dateB) Then

        If CDate(dateA) < CDate(dateB) Then
            CompareTrackerRows = -1
            Exit Function
        End If

        If CDate(dateA) > CDate(dateB) Then
            CompareTrackerRows = 1
            Exit Function
        End If

    ElseIf IsDate(dateA) And Not IsDate(dateB) Then

        CompareTrackerRows = -1
        Exit Function

    ElseIf Not IsDate(dateA) And IsDate(dateB) Then

        CompareTrackerRows = 1
        Exit Function

    End If

    ' Task name.
    CompareTrackerRows = StrComp(taskA, taskB, vbTextCompare)

End Function
