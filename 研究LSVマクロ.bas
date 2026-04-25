Attribute VB_Name = "LSV_Import"
' ============================================================
' LSV Data Import & Chart Macro  v2
' - Multiple files selectable (overlaid on one chart)
' - No cycle detection (LSV = single sweep)
' - j in mA/cm2
' - Y axis: 0-centered nice intervals
' - X axis at TOP, tick labels above axis, tick marks inside
' - Compatible with CHI604E / Linear Sweep Voltammetry text files
' ============================================================

' NOTE: CalcNiceAxis is defined in CV_Import module.
'       If using LSV standalone (without CV module), paste CalcNiceAxis here too.

Sub ImportLSVData()

    ' ----------------------------------------------------------------
    ' 0. File selection (multiple files allowed)
    ' ----------------------------------------------------------------
    Dim filePaths As Variant
    filePaths = Application.GetOpenFilename( _
        "Text Files (*.txt),*.txt", , _
        "Select LSV Data Files (Ctrl+Click for multiple)", , True)
    If VarType(filePaths) = vbBoolean Then Exit Sub

    ' ----------------------------------------------------------------
    ' 1. Sheet setup
    ' ----------------------------------------------------------------
    Dim ws As Worksheet
    Dim wsName As String
    wsName = "LSV Data"

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(wsName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        If MsgBox("Sheet [" & wsName & "] already exists. Overwrite?", _
                  vbYesNo + vbQuestion, "Confirm") = vbNo Then Exit Sub
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = wsName

    ' ----------------------------------------------------------------
    ' 2. Electrode area input cell (B2)
    ' ----------------------------------------------------------------
    With ws
        .Range("A1").Value = "Electrode Area (cm2)"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Name = "Times New Roman"

        .Range("B2").Value = 1.0
        .Range("B2").Font.Color = RGB(0, 0, 255)
        .Range("B2").Font.Bold = True
        .Range("B2").Font.Name = "Times New Roman"
        .Range("B2").Interior.Color = RGB(255, 255, 200)

        .Range("A2").Value = "<-- Enter electrode area here (shared for all files)"
        .Range("A2").Font.Italic = True
        .Range("A2").Font.Color = RGB(100, 100, 100)
        .Range("A2").Font.Name = "Times New Roman"

        Dim hdr As Variant
        hdr = Array("File", "Potential / V", "Current / A", "j / mA cm-2")
        Dim col As Integer
        For col = 0 To 3
            With .Cells(5, col + 1)
                .Value = hdr(col)
                .Font.Bold = True
                .Font.Name = "Times New Roman"
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(46, 64, 87)
                .HorizontalAlignment = xlCenter
            End With
        Next col
    End With

    ' ----------------------------------------------------------------
    ' 3. Process each file
    ' ----------------------------------------------------------------
    Const MAX_PTS As Long = 30000

    Dim dataRow As Long
    dataRow = 6

    Dim fileCount As Integer
    fileCount = UBound(filePaths) - LBound(filePaths) + 1

    Dim fileStartRow() As Long
    Dim fileEndRow()   As Long
    Dim fileLabel()    As String
    ReDim fileStartRow(1 To fileCount)
    ReDim fileEndRow(1 To fileCount)
    ReDim fileLabel(1 To fileCount)

    Dim fi As Integer
    For fi = 1 To fileCount
        Dim filePath As String
        filePath = filePaths(LBound(filePaths) + fi - 1)

        Dim fname As String
        fname = Mid(filePath, InStrRev(filePath, "\") + 1)
        fname = Left(fname, InStrRev(fname, ".") - 1)
        fileLabel(fi) = fname

        Dim fileNum As Integer
        Dim fileLine As String
        Dim parts() As String
        Dim inData As Boolean
        inData = False

        fileNum = FreeFile
        fileStartRow(fi) = dataRow

        Open filePath For Input As #fileNum
        Do While Not EOF(fileNum)
            Line Input #fileNum, fileLine
            fileLine = Trim(fileLine)
            If fileLine = "Potential/V, Current/A" Then
                inData = True
            ElseIf inData And fileLine <> "" Then
                parts = Split(fileLine, ",")
                If UBound(parts) = 1 Then
                    On Error Resume Next
                    Dim tmpV As Double, tmpC As Double
                    tmpV = CDbl(Trim(parts(0)))
                    tmpC = CDbl(Trim(parts(1)))
                    If Err.Number = 0 Then
                        ws.Cells(dataRow, 1).Value   = fi
                        ws.Cells(dataRow, 2).Value   = tmpV
                        ws.Cells(dataRow, 3).Value   = tmpC
                        ' j [mA/cm2] = Current[A] * 1000 / Area[cm2]
                        ws.Cells(dataRow, 4).Formula = "=C" & dataRow & "*1000/$B$2"
                        dataRow = dataRow + 1
                    End If
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
        Loop
        Close #fileNum

        fileEndRow(fi) = dataRow - 1
    Next fi

    Dim lastRow As Long
    lastRow = dataRow - 1

    If lastRow < 6 Then
        MsgBox "No data found. Check the file format.", vbExclamation, "Error"
        Exit Sub
    End If

    ' ----------------------------------------------------------------
    ' 4. Formatting
    ' ----------------------------------------------------------------
    With ws
        .Columns("A").ColumnWidth = 6
        .Columns("B").ColumnWidth = 16
        .Columns("C").ColumnWidth = 16
        .Columns("D").ColumnWidth = 20

        Dim r As Long
        For r = 6 To lastRow
            If (r Mod 2) = 0 Then
                .Rows(r).Interior.Color = RGB(242, 246, 252)
            End If
        Next r

        .Range("B6:B" & lastRow).NumberFormat = "0.000"
        .Range("C6:C" & lastRow).NumberFormat = "0.00E+00"
        .Range("D6:D" & lastRow).NumberFormat = "0.0000"
        .Range("A6:D" & lastRow).Font.Name = "Times New Roman"
        .Range("A6:D" & lastRow).Font.Size = 10

        With .Range("A5:D" & lastRow).Borders
            .LineStyle = xlContinuous
            .Color = RGB(180, 180, 180)
            .Weight = xlThin
        End With
    End With

    ' ----------------------------------------------------------------
    ' 5. Y axis scale: fixed at -7 to 0, step 1
    ' ----------------------------------------------------------------
    Dim niceMin As Double, niceMax As Double, niceStep As Double
    niceMin  = -7
    niceMax  =  0
    niceStep =  1

    ' ----------------------------------------------------------------
    ' 6. Chart
    ' ----------------------------------------------------------------
    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(Left:=340, Top:=10, Width:=520, Height:=340)

    Dim palette(1 To 10) As Long
    palette(1)  = RGB(0, 0, 0)
    palette(2)  = RGB(0, 114, 189)
    palette(3)  = RGB(217, 83, 25)
    palette(4)  = RGB(126, 47, 142)
    palette(5)  = RGB(119, 172, 48)
    palette(6)  = RGB(77, 190, 238)
    palette(7)  = RGB(162, 20, 47)
    palette(8)  = RGB(0, 128, 0)
    palette(9)  = RGB(255, 153, 0)
    palette(10) = RGB(100, 100, 100)

    Dim ser As Series

    With cht.Chart
        .ChartType = xlXYScatterLinesNoMarkers
        .HasTitle = False

        For fi = 1 To fileCount
            .SeriesCollection.NewSeries
            Set ser = .SeriesCollection(fi)
            ser.XValues = ws.Range(ws.Cells(fileStartRow(fi), 2), ws.Cells(fileEndRow(fi), 2))
            ser.Values  = ws.Range(ws.Cells(fileStartRow(fi), 4), ws.Cells(fileEndRow(fi), 4))
            ser.Name    = fileLabel(fi)
            Dim cIdx As Integer
            cIdx = ((fi - 1) Mod 10) + 1
            ser.Format.Line.ForeColor.RGB = palette(cIdx)
            ser.Format.Line.Weight = 1.5
        Next fi

        ' ---- Y axis: nice scale fixed ----
        With .Axes(xlValue)
            .HasTitle = True
            .MinimumScaleIsAuto = False
            .MaximumScaleIsAuto = False
            .MajorUnitIsAuto = False
            .MinimumScale = niceMin
            .MaximumScale = niceMax
            .MajorUnit    = niceStep

            .AxisTitle.Text = "j / mA cm" & ChrW(8315) & ChrW(178)
            .AxisTitle.Font.Name = "Times New Roman"
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Italic = False
            .AxisTitle.Characters(1, 1).Font.Italic = True   ' j italic

            .TickLabels.Font.Name = "Times New Roman"
            .TickLabels.Font.Size = 11

            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            .Format.Line.Weight = 1

            .MajorTickMark = xlTickMarkInside
            .MinorTickMark = xlTickMarkNone
        End With

        ' ---- X axis: at TOP ----
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "E / V vs. RHE"
            .AxisTitle.Font.Name = "Times New Roman"
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Italic = False
            .AxisTitle.Characters(1, 1).Font.Italic = True   ' E italic

            ' Labels above the axis line
            .TickLabelPosition = xlHigh
            .TickLabels.Font.Name = "Times New Roman"
            .TickLabels.Font.Size = 11
            .TickLabels.NumberFormat = "0.0"

            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            .Format.Line.Weight = 1

            .MajorTickMark = xlTickMarkInside
            .MinorTickMark = xlTickMarkNone
        End With

        .PlotArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.Visible = msoFalse
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Format.Line.Visible = msoFalse

        On Error Resume Next
        .Axes(xlCategory).HasMajorGridlines = False
        .Axes(xlValue).HasMajorGridlines = False
        On Error GoTo 0

        .HasLegend = True
        With .Legend
            .Font.Name = "Times New Roman"
            .Font.Size = 11
            .Format.Fill.Visible = msoFalse
            .Format.Line.Visible = msoFalse
        End With

    End With   ' cht.Chart

    ' X axis at TOP: CrossesAt = niceMax, set AFTER With block
    cht.Chart.Axes(xlCategory).Crosses = xlCustom
    cht.Chart.Axes(xlCategory).CrossesAt = niceMax

    ' ----------------------------------------------------------------
    ' 7. Done
    ' ----------------------------------------------------------------
    Dim msg As String
    msg = "Import complete!" & vbCrLf & vbCrLf & _
          "Files loaded: " & fileCount & vbCrLf & vbCrLf
    For fi = 1 To fileCount
        msg = msg & fileLabel(fi) & ": " & _
              (fileEndRow(fi) - fileStartRow(fi) + 1) & " pts" & vbCrLf
    Next fi
    msg = msg & vbCrLf & "Tip: Edit B2 (yellow cell) to update electrode area."
    MsgBox msg, vbInformation, "LSV Import"

End Sub
