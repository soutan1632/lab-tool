Attribute VB_Name = "CV_Import"
' ============================================================
' CV Data Import & Chart Macro  v10
' - Multiple files selectable (overlaid on one chart)
' - Plots Cycle 3 only from each file
' - Requires CHI604E metadata (Init E, Segment) in file header
' - Supports comma-delimited and tab-delimited CHI formats
' - Cycle 3 = last LOW turn (V < 0.1, sign reversal)
'   -> works for both start=0.05V (seg=6) and start=0.45V (seg=7)
' - Y axis: fixed -80 to 80 uA/cm2, step 20
' ============================================================

' ------------------------------------------------------------------
' Shared helper: compute nice axis scale including 0
'   in:  dataMin, dataMax
'   out: axMin, axMax, axStep  (all rounded, 0 always included)
' ------------------------------------------------------------------
Sub CalcNiceAxis(dataMin As Double, dataMax As Double, _
                 axMin As Double, axMax As Double, axStep As Double)
    Dim span As Double
    span = Application.WorksheetFunction.Max(Abs(dataMin), Abs(dataMax))
    If span = 0 Then span = 1

    Dim rawStep As Double
    rawStep = span / 5   ' target ~5 intervals per side

    ' Round step up to nearest 1/2/5 * 10^n
    Dim mag As Double
    mag = 10 ^ Int(Log(rawStep) / Log(10))
    Dim s As Double
    Dim found As Boolean
    found = False
    Dim factors(1 To 4) As Double
    factors(1) = 1: factors(2) = 2: factors(3) = 5: factors(4) = 10
    Dim k As Integer
    For k = 1 To 4
        s = factors(k) * mag
        If s >= rawStep Then
            axStep = s
            found = True
            Exit For
        End If
    Next k
    If Not found Then axStep = 10 * mag

    ' Snap min/max outward to multiples of step, ensuring 0 is included
    axMin = Application.WorksheetFunction.Floor_Math(dataMin, axStep)
    axMax = Application.WorksheetFunction.Ceiling_Math(dataMax, axStep)
    If axMin > 0 Then axMin = 0
    If axMax < 0 Then axMax = 0

    ' Round to avoid floating-point noise
    Dim decimals As Integer
    If axStep >= 1 Then
        decimals = 0
    Else
        decimals = CInt(-Int(Log(axStep) / Log(10))) + 1
    End If
    axMin  = Application.WorksheetFunction.Round(axMin,  decimals)
    axMax  = Application.WorksheetFunction.Round(axMax,  decimals)
    axStep = Application.WorksheetFunction.Round(axStep, decimals)
End Sub

Sub ImportCVData()

    ' ----------------------------------------------------------------
    ' 0. File selection (multiple files allowed)
    ' ----------------------------------------------------------------
    Dim filePaths As Variant
    filePaths = Application.GetOpenFilename( _
        "Text Files (*.txt),*.txt", , _
        "Select CV Data Files (Ctrl+Click for multiple)", , True)
    If VarType(filePaths) = vbBoolean Then Exit Sub

    ' ----------------------------------------------------------------
    ' 1. Sheet setup
    ' ----------------------------------------------------------------
    Dim ws As Worksheet
    Dim wsName As String
    wsName = "CV Data"

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
    ' 2. Input cells: area (B2), start potential (B3), segments (B4)
    ' ----------------------------------------------------------------
    With ws
        ' --- row 1: area ---
        .Range("A1").Value = "Electrode Area (cm2)"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Name = "Times New Roman"

        .Range("B2").Value = 0.048
        .Range("B2").Font.Color = RGB(0, 0, 255)
        .Range("B2").Font.Bold = True
        .Range("B2").Font.Name = "Times New Roman"
        .Range("B2").Interior.Color = RGB(255, 255, 200)

        .Range("A2").Value = "<-- Enter electrode area here (shared for all files)"
        .Range("A2").Font.Italic = True
        .Range("A2").Font.Color = RGB(100, 100, 100)
        .Range("A2").Font.Name = "Times New Roman"

        ' --- row 3: start potential ---
        .Range("C3").Value = "Start Potential (V)"
        .Range("C3").Font.Bold = True
        .Range("C3").Font.Name = "Times New Roman"

        .Range("B3").Value = 0.45
        .Range("B3").Font.Color = RGB(0, 0, 255)
        .Range("B3").Font.Bold = True
        .Range("B3").Font.Name = "Times New Roman"
        .Range("B3").Interior.Color = RGB(255, 255, 200)

        .Range("A3").Value = "<-- Enter start potential"
        .Range("A3").Font.Italic = True
        .Range("A3").Font.Color = RGB(100, 100, 100)
        .Range("A3").Font.Name = "Times New Roman"

        ' --- row 4: number of segments ---
        .Range("C4").Value = "Number of Segments"
        .Range("C4").Font.Bold = True
        .Range("C4").Font.Name = "Times New Roman"

        .Range("B4").Value = 7
        .Range("B4").Font.Color = RGB(0, 0, 255)
        .Range("B4").Font.Bold = True
        .Range("B4").Font.Name = "Times New Roman"
        .Range("B4").Interior.Color = RGB(255, 255, 200)

        .Range("A4").Value = "<-- Enter number of sweep segments"
        .Range("A4").Font.Italic = True
        .Range("A4").Font.Color = RGB(100, 100, 100)
        .Range("A4").Font.Name = "Times New Roman"

        ' --- row 7: column headers (moved down from row 5) ---
        Dim hdr As Variant
        hdr = Array("File", "Cycle", "Potential / V", "Current / A", "j / uA cm-2")
        Dim col As Integer
        For col = 0 To 4
            With .Cells(7, col + 1)
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
    ' 3. Get mu character safely via ChrW -> cell -> string
    ' ----------------------------------------------------------------
    ws.Range("G1").Value = ChrW(956)
    Dim muChar As String
    muChar = ws.Range("G1").Value
    ws.Range("G1").Clear

    ' ----------------------------------------------------------------
    ' 4. Process each file: extract Cycle 3
    ' ----------------------------------------------------------------
    Const MAX_PTS As Long = 30000

    ' Read user inputs
    Dim electrodeArea As Double
    Dim startPotential As Double
    Dim numSegments As Integer
    electrodeArea   = ws.Range("B2").Value
    startPotential  = ws.Range("B3").Value
    numSegments     = CInt(ws.Range("B4").Value)

    Dim dataRow As Long
    dataRow = 8   ' data starts at row 8 (header at row 7)

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

        Dim allV() As Double
        Dim allC() As Double
        ReDim allV(MAX_PTS)
        ReDim allC(MAX_PTS)
        Dim totalPts As Long
        totalPts = 0

        Dim fileNum As Integer
        Dim fileLine As String
        Dim parts() As String
        Dim inData As Boolean
        inData = False

        ' ---- Read metadata and data ----
        ' Supports two CHI604E formats:
        '   Format A: comma-delimited, header = "Potential/V, Current/A"
        '   Format B: tab-delimited,   header = "Potential/V<tab>Current/A"
        ' Both formats always have CHI metadata (Init E, Segment).
        Dim fileStartV As Double
        Dim fileNumSeg As Integer
        Dim delim      As String
        fileStartV = 0
        fileNumSeg = 0
        delim      = ","

        fileNum = FreeFile
        Open filePath For Input As #fileNum
        Do While Not EOF(fileNum)
            Line Input #fileNum, fileLine
            fileLine = Trim(fileLine)

            ' -- Metadata --
            If InStr(fileLine, "Init E (V) =") > 0 Then
                On Error Resume Next
                fileStartV = CDbl(Trim(Split(fileLine, "=")(1)))
                On Error GoTo 0
            End If
            If InStr(fileLine, "Segment =") > 0 Then
                On Error Resume Next
                fileNumSeg = CInt(Trim(Split(fileLine, "=")(1)))
                On Error GoTo 0
            End If

            ' -- Data header: flexible match for comma or tab format --
            If InStr(fileLine, "Potential/V") > 0 And InStr(fileLine, "Current/A") > 0 Then
                inData = True
                If InStr(fileLine, Chr(9)) > 0 Then
                    delim = Chr(9)   ' tab-delimited (Format B)
                Else
                    delim = ","      ' comma-delimited (Format A)
                End If

            ElseIf inData And fileLine <> "" Then
                parts = Split(fileLine, delim)
                If UBound(parts) = 1 Then
                    On Error Resume Next
                    Dim tmpV As Double, tmpC As Double
                    tmpV = CDbl(Trim(parts(0)))
                    tmpC = CDbl(Trim(parts(1)))
                    If Err.Number = 0 And totalPts < MAX_PTS Then
                        allV(totalPts) = tmpV
                        allC(totalPts) = tmpC
                        totalPts = totalPts + 1
                    End If
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
        Loop
        Close #fileNum

        ' Reflect metadata values into B3/B4
        If fileStartV > 0 Then ws.Range("B3").Value = fileStartV
        If fileNumSeg > 0 Then ws.Range("B4").Value = fileNumSeg

        ' ---- Find Cycle 3 start: last LOW turn ----
        ' LOW turn = V < 0.1, direction reverses from down to up.
        ' CHI604E output is device-controlled, so simple sign-reversal is stable.
        ' Cycle 3 always begins at the LAST LOW turn regardless of start potential:
        '   start=0.05V, seg=6 -> 2 LOW turns -> last = Cycle3 start
        '   start=0.45V, seg=7 -> 3 LOW turns -> last = Cycle3 start
        Dim lowIdx(20) As Long
        Dim numLow As Integer
        numLow = 0

        Dim j As Long
        For j = 1 To totalPts - 2
            If allV(j) < 0.1 Then
                If allV(j) - allV(j - 1) < 0 And allV(j + 1) - allV(j) > 0 Then
                    If numLow < 20 Then
                        lowIdx(numLow) = j
                        numLow = numLow + 1
                    End If
                End If
            End If
        Next j

        Dim c3Start As Long
        If numLow >= 1 Then
            c3Start = lowIdx(numLow - 1)   ' always the last LOW turn
        Else
            c3Start = totalPts \ 2         ' fallback
        End If

        fileStartRow(fi) = dataRow
        Dim pt As Long
        For pt = c3Start To totalPts - 1
            ws.Cells(dataRow, 1).Value   = fi
            ws.Cells(dataRow, 2).Value   = 3
            ws.Cells(dataRow, 3).Value   = allV(pt)
            ws.Cells(dataRow, 4).Value   = allC(pt)
            ws.Cells(dataRow, 5).Formula = "=D" & dataRow & "*1000000/$B$2"
            dataRow = dataRow + 1
        Next pt
        fileEndRow(fi) = dataRow - 1
    Next fi

    Dim lastRow As Long
    lastRow = dataRow - 1

    If lastRow < 8 Then
        MsgBox "No data found. Check the file format.", vbExclamation, "Error"
        Exit Sub
    End If

    ' ----------------------------------------------------------------
    ' 5. Formatting
    ' ----------------------------------------------------------------
    With ws
        .Columns("A").ColumnWidth = 6
        .Columns("B").ColumnWidth = 6
        .Columns("C").ColumnWidth = 16
        .Columns("D").ColumnWidth = 16
        .Columns("E").ColumnWidth = 20

        Dim r As Long
        For r = 8 To lastRow
            If (r Mod 2) = 0 Then
                .Rows(r).Interior.Color = RGB(242, 246, 252)
            End If
        Next r

        .Range("C8:C" & lastRow).NumberFormat = "0.000"
        .Range("D8:D" & lastRow).NumberFormat = "0.00E+00"
        .Range("E8:E" & lastRow).NumberFormat = "0.0000"
        .Range("A8:E" & lastRow).Font.Name = "Times New Roman"
        .Range("A8:E" & lastRow).Font.Size = 10

        With .Range("A7:E" & lastRow).Borders
            .LineStyle = xlContinuous
            .Color = RGB(180, 180, 180)
            .Weight = xlThin
        End With
    End With

    ' ----------------------------------------------------------------
    ' 6. Y axis scale: fixed at -80 to 80, step 20
    ' ----------------------------------------------------------------
    Dim niceMin As Double, niceMax As Double, niceStep As Double
    niceMin  = -80
    niceMax  =  80
    niceStep =  20

    ' ----------------------------------------------------------------
    ' 7. Chart
    ' ----------------------------------------------------------------
    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(Left:=400, Top:=10, Width:=520, Height:=340)

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
            ser.XValues = ws.Range(ws.Cells(fileStartRow(fi), 3), ws.Cells(fileEndRow(fi), 3))
            ser.Values  = ws.Range(ws.Cells(fileStartRow(fi), 5), ws.Cells(fileEndRow(fi), 5))
            ser.Name    = fileLabel(fi)
            Dim cIdx As Integer
            cIdx = ((fi - 1) Mod 10) + 1
            ser.Format.Line.ForeColor.RGB = palette(cIdx)
            ser.Format.Line.Weight = 1.5
        Next fi

        ' ---- X axis ----
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "E / V vs. RHE"
            .AxisTitle.Font.Name = "Times New Roman"
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Italic = False
            .AxisTitle.Characters(1, 1).Font.Italic = True

            .TickLabels.Font.Name = "Times New Roman"
            .TickLabels.Font.Size = 11
            .TickLabels.NumberFormat = "0.0"

            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            .Format.Line.Weight = 1

            .MajorTickMark = xlTickMarkInside
            .MinorTickMark = xlTickMarkNone
        End With

        ' ---- Y axis: nice scale fixed ----
        With .Axes(xlValue)
            .HasTitle = True
            .MinimumScaleIsAuto = False
            .MaximumScaleIsAuto = False
            .MajorUnitIsAuto = False
            .MinimumScale = niceMin
            .MaximumScale = niceMax
            .MajorUnit    = niceStep

            .AxisTitle.Text = "j / " & muChar & "A cm" & ChrW(8315) & ChrW(178)
            .AxisTitle.Font.Name = "Times New Roman"
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Italic = False
            .AxisTitle.Characters(1, 1).Font.Italic = True

            .TickLabels.Font.Name = "Times New Roman"
            .TickLabels.Font.Size = 11

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

    ' Set X axis CrossesAt AFTER With block (Y scale now fully committed)
    cht.Chart.Axes(xlCategory).Crosses = xlCustom
    cht.Chart.Axes(xlCategory).CrossesAt = niceMin

    ' ----------------------------------------------------------------
    ' 8. Done
    ' ----------------------------------------------------------------
    Dim msg As String
    msg = "Import complete!" & vbCrLf & vbCrLf
    For fi = 1 To fileCount
        msg = msg & fileLabel(fi) & vbCrLf & _
              "  " & (fileEndRow(fi) - fileStartRow(fi) + 1) & " pts" & vbCrLf
    Next fi
    msg = msg & vbCrLf & _
          "* Files with metadata: Init E and Segment were read automatically" & vbCrLf & _
          "  and written back to B3 / B4." & vbCrLf & _
          "* Files without metadata: B3 / B4 cell values were used." & vbCrLf & vbCrLf & _
          "Tip: Check yellow cells (B2-B4) before re-running."
    MsgBox msg, vbInformation, "CV Import"

End Sub
