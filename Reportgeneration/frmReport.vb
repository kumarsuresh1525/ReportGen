Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Imports System.Globalization
Imports System.Text
Imports Newtonsoft.Json
Public Class frmReport
    Dim xlApp As Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet
    Dim start_sheet As Integer
    Dim lines(3), p_lines(3) As Integer
    Dim sheetCount As Integer
    Dim sheetCountOGS As Integer = 1
    Dim start_pos, last_pos, endSheet As Integer
    Dim da As New OleDbDataAdapter
    Dim dt As New DataTable
    Dim showMessage As Boolean = False
    Dim msgNumber As Integer
    Dim treeNodeStart As Integer = 0
    Dim partsNode As New ArrayList
    Dim qmsNode As New ArrayList
    Dim qmpNode As New ArrayList
    Dim freqAtInitNode As New ArrayList
    Dim freqNode As New ArrayList
    Dim measuringToolNode As New ArrayList
    Dim checkedByNode As New ArrayList
    Dim recordingSheetNode As New ArrayList
    Dim machineJigToolNode As New ArrayList
    Dim processNoNode, processNameNode As String
    Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + fileSource + "';Extended Properties = ""Excel 12.0 Xml;HDR=YES"""
    Dim PFC_NAME, REV_NO, PartNo As String
    Private Sub saveAllSetting(ByVal val As Integer)
        SaveSetting("st_pos", "st_pos", "st_pos", val)
    End Sub
    Private Sub getAllSetting()
        start_pos = GetSetting("st_pos", "st_pos", "st_pos")
        last_pos = GetSetting("last_pos", "last_pos", "last_pos")
        endSheet = GetSetting("endSheet", "endSheet", "endSheet")
    End Sub
    Private Sub generatePFC()
        Dim desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        If (Not System.IO.Directory.Exists(desktop + "\Reports\PFC")) Then
            System.IO.Directory.CreateDirectory(desktop + "\Reports\PFC")
        Else
            If (System.IO.File.Exists(desktop + "\Reports\PFC\PFC.xlsx")) Then
                My.Computer.FileSystem.DeleteFile(desktop + "\Reports\PFC\PFC.xlsx")
                My.Computer.FileSystem.CopyFile(Application.StartupPath + "\reference\PFC.xlsx",
                                            desktop + "\Reports\PFC\PFC.xlsx")
                startProcess(desktop + "\Reports\PFC\PFC.xlsx", "PFC", 8, 47)
            Else
                My.Computer.FileSystem.CopyFile(Application.StartupPath + "\reference\PFC.xlsx",
                                            desktop + "\Reports\PFC\PFC.xlsx")
                startProcess(desktop + "\Reports\PFC\PFC.xlsx", "PFC", 8, 47)
            End If
        End If
    End Sub
    Private Sub generatePCC()
        Dim desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        If (Not System.IO.Directory.Exists(desktop + "\Reports\PCC")) Then
            System.IO.Directory.CreateDirectory(desktop + "\Reports\PCC")
        Else
            If (System.IO.File.Exists(desktop + "\Reports\PCC\PCC.xlsx")) Then
                My.Computer.FileSystem.DeleteFile(desktop + "\Reports\PCC\PCC.xlsx")
                My.Computer.FileSystem.CopyFile(Application.StartupPath + "\reference\PCC.xlsx",
                                            desktop + "\Reports\PCC\PCC.xlsx")
                makePCC(desktop + "\Reports\PCC\PCC.xlsx", "PCC")
            Else
                My.Computer.FileSystem.CopyFile(Application.StartupPath + "\reference\PCC.xlsx",
                                            desktop + "\Reports\PCC\PCC.xlsx")
                makePCC(desktop + "\Reports\PCC\PCC.xlsx", "PCC")
            End If
        End If
    End Sub
    Private Sub generatePQCS()
        Dim desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim model = "MRPL"
        Dim startP, EndP As Integer
        startP = 12
        EndP = 66
        If (Not System.IO.Directory.Exists(desktop + "\Reports\PQCS")) Then
            System.IO.Directory.CreateDirectory(desktop + "\Reports\PQCS")
        Else
            If (System.IO.File.Exists(desktop + "\Reports\PQCS\PQCS.xlsx")) Then
                My.Computer.FileSystem.DeleteFile(desktop + "\Reports\PQCS\PQCS.xlsx")
                My.Computer.FileSystem.CopyFile(Application.StartupPath + "\reference\" + "PQCS_" + model + ".xlsx",
                                            desktop + "\Reports\PQCS\PQCS.xlsx")

                SaveSetting("st_pos", "st_pos", "st_pos", startP)
                SaveSetting("last_pos", "last_pos", "last_pos", startP)
                SaveSetting("endSheet", "endSheet", "endSheet", EndP)
                createPQCS(Application.StartupPath + "\db.xlsx", "P_MRPL", startP, EndP)
            Else
                My.Computer.FileSystem.CopyFile(Application.StartupPath + "\reference\" + "PQCS_" + model + ".xlsx",
                                            desktop + "\Reports\PQCS\PQCS.xlsx")
                SaveSetting("st_pos", "st_pos", "st_pos", startP)
                SaveSetting("last_pos", "last_pos", "last_pos", startP)
                SaveSetting("endSheet", "endSheet", "endSheet", EndP)
                createPQCS(Application.StartupPath + "\db.xlsx", "P_MRPL", startP, EndP)
            End If
        End If
    End Sub
    Private Sub createPQCS(ByVal fileSrc As String, ByVal sheetName As String, ByVal startNo As Integer, ByVal lastNo As Integer)
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open(fileSrc)
        Dim newSheetName = fnCreateSheet(xlWorkBook, sheetName, 1)
        xlWorkSheet = xlWorkBook.Worksheets(newSheetName)

        Dim ldtPartDetails As New DataTable
        Dim ldtPartDetailsNew As New DataTable
        Dim dr, rs As DataRow
        Dim xsProcess As Excel.Worksheet
        ldtPartDetails = gfnSelectQueryDt("SELECT DISTINCT pName, pNo FROM partDetails WHERE pfcName='" + PFC_NAME + "' order by pNo")
        Dim i, j, k As Integer
        Dim pr_no, flow_sy_type As String
        For Each dr In ldtPartDetails.Rows
            xsProcess = xlWorkBook.Worksheets("Processes")
            i = 2
            Do While Len(xsProcess.Range("A" & i).Value) <> 0
                If xsProcess.Range("E" & i).Value = dr("pName") Then
                    pr_no = xsProcess.Range("A" & i).Value
                    flow_sy_type = xsProcess.Range("F" & i).Value
                    Exit Do
                End If
                i = i + 1
            Loop
            xsProcess = xlWorkBook.Worksheets("PQCS_DB")
            i = 5
            Do While Len(xsProcess.Range("K" & i).Value) <> 0
                j = i
                Do While checkDBNull(xsProcess, i, j) = True
                    j = j + 1
                Loop
                If xsProcess.Range("K" & i).Value = pr_no Then
                    lines(k) = 1
                    getAllSetting()
                    ldtPartDetailsNew = gfnSelectQueryDt("SELECT * FROM partDetails WHERE pfcName='" + PFC_NAME + "' AND pName='" + dr("pName") + "'")
                    For Each rs In ldtPartDetailsNew.Rows
                        makeReact(rs("partNo"), rs("partName"), rs("qty"), rs("xyz"), xlWorkSheet)
                    Next
                    createFlowChart(flow_sy_type, Convert.ToInt16(dr("pNo")), dr("pName"), xlWorkSheet)
                    fillValuesPQCS(xlWorkSheet, xsProcess, i, j, dr("pNo"))
                    Exit Do
                End If
                i = j + 2
            Loop
        Next

        xlWorkBook.Save()
        xlWorkBook.Close() : xlApp.Quit()

        ' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook) : xlWorkBook = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet) : xlWorkSheet = Nothing
    End Sub
    Private Function checkDBNull(ByVal xs As Excel.Worksheet, ByVal i As Integer, ByVal j As Integer)
        If IsDBNull(xs.Range("K" & i & ":" & "N" & j + 1).MergeCells) Then
            Return False
        End If
        Return True
    End Function
    Private Function fnCreateSheet(ByVal xw As Excel.Workbook, ByVal sheetName As String, ByVal no As Integer)
        xlWorkSheet = xw.Worksheets(sheetName)
        xlWorkSheet.Visible = True
        xlWorkSheet.Copy(, xlWorkSheet)
        xlWorkSheet.Visible = False
        Dim sheetCount As String = 2
        Dim newSheetName As String = sheetName + " (" + sheetCount.ToString() + ")"
        xlWorkSheet = xlWorkBook.Worksheets(newSheetName)
        xlWorkSheet.Name = sheetName + "_" + no.ToString()
        Return sheetName + "_" + no.ToString()
    End Function
    Private Sub fillValuesPQCS(ByVal xsPQCS As Excel.Worksheet, ByVal xsPQCSDB As Excel.Worksheet, ByVal i As Integer, ByVal j As Integer, ByVal pNo As Integer)
        getAllSetting()
        If last_pos + j - i + 1 >= endSheet Then

        End If

        xsPQCSDB.Range("O" & i & ":AC" & j).Copy(xsPQCS.Range("P12"))
    End Sub
    Private Sub createDirectory()
        Dim desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        If (Not System.IO.Directory.Exists(desktop + "\" + "Reports")) Then
            System.IO.Directory.CreateDirectory(desktop + "\" + "Reports")
        End If
    End Sub
    Private Sub createSheet(ByVal xw As Excel.Workbook, ByVal sheetName As String)
        xlWorkSheet = xw.Worksheets(sheetName)
        xlWorkSheet.Visible = True
        xlWorkSheet.Copy(, xlWorkSheet)
        xlWorkSheet.Visible = False
    End Sub
    Private Sub makePCC(ByVal fileSrc As String, ByVal sheetName As String)
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open(fileSrc)
        createSheet(xlWorkBook, sheetName)
        sheetCount = 2
        Dim newSheetName As String = sheetName + " (" + sheetCount.ToString() + ")"
        xlWorkSheet = xlWorkBook.Worksheets(newSheetName)
        xlWorkSheet.Range("A3").Value = product_Name
        xlWorkSheet.Range("A5").Value = productNo
        xlWorkSheet.Range("A5", "C5").Justify()
        xlWorkSheet.Range("E5").Value = REV_NO
        Dim ldtPartDetails As New DataTable
        Dim dr As DataRow
        ldtPartDetails = gfnSelectQueryDt("SELECT DISTINCT pName, pNo FROM partDetails WHERE pfcName='" + PFC_NAME + "' order by pNo")
        Dim lpno As New StringBuilder
        Dim lpName As New StringBuilder
        For Each dr In ldtPartDetails.Rows
            lpno.Append(dr("pNo") + ",")
            lpName.Append(dr("pName") + ",")
        Next
        xlWorkSheet.Range("C6").Value = lpName.ToString().Substring(0, lpName.ToString().Length - 1)
        Dim lrange As Excel.Range = xlWorkSheet.Range("C6")
        lrange.WrapText = True
        xlWorkSheet.Range("C7").Value = lpno.ToString().Substring(0, lpno.ToString().Length - 1)
        xlWorkSheet.Range("Q3").Value = "ASSY"
        xlWorkSheet.Range("P28").Value = "Issue Date: " + Date.Now.ToString("dd.MM.yy", CultureInfo.InvariantCulture)
        xlWorkBook.Save()
        xlWorkBook.Close() : xlApp.Quit()

        ' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook) : xlWorkBook = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet) : xlWorkSheet = Nothing

    End Sub
    Private Sub startProcess(ByVal fileSrc As String, ByVal sheetName As String, ByVal startPosition As Integer, ByVal endSheetNo As Integer)
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open(fileSrc)
        createSheet(xlWorkBook, sheetName)
        sheetCount = 2
        Dim newSheetName As String = sheetName + " (" + sheetCount.ToString() + ")"
        xlWorkSheet = xlWorkBook.Worksheets(newSheetName)

        Dim i, k As Integer
        Dim fc As Boolean
        Dim H_LINE(3) As Integer
        i = 5
        'start_pos = 12
        fc = False
        If Not fc Then
            Dim ldtPartDetails As New DataTable
            ldtPartDetails = gfnSelectQueryDt("SELECT * FROM partDetails WHERE pfcName='" + PFC_NAME + "'")

            Dim Rs As DataRow
            If ldtPartDetails.Rows.Count = 0 Then
                Return
            End If
            saveAllSetting(startPosition)
            For Each Rs In ldtPartDetails.Rows
                If start_pos >= endSheetNo Then
                    createSheet(xlWorkBook, sheetName)
                    sheetCount = sheetCount + 1
                    newSheetName = sheetName + " (" + sheetCount.ToString() + ")"
                    xlWorkSheet = xlWorkBook.Worksheets(newSheetName)
                    saveAllSetting(startPosition)
                    endSheetNo = start_pos
                End If
                getAllSetting()
                lines(k) = 1
                makeReact(Rs("partNo"), Rs("partName"), Rs("qty"), Rs("xyz"), xlWorkSheet)
                PartNo = Rs("partNo")
                customerNo = Rs("customerNo")
                customerName = Rs("customerName")
                productNo = Rs("productNo")
                product_Name = Rs("productName")
                model = Rs("model")
                'fillPFC(xlWorkSheet)
            Next

            '
            ldtPartDetails.Clear()
            ldtPartDetails = gfnSelectQueryDt("SELECT DISTINCT pName, pNo FROM partDetails WHERE pfcName='" + PFC_NAME + "' order by pNo")

            If ldtPartDetails.Rows.Count = 0 Then
                Return
            End If
            For Each Rs In ldtPartDetails.Rows
                If Not IsDBNull(Rs("pName")) AndAlso Not Rs("pName") = "" Then
                    If start_pos >= endSheetNo Then
                        createSheet(xlWorkBook, sheetName)
                        sheetCount = sheetCount + 1
                        newSheetName = sheetName + " (" + sheetCount.ToString() + ")"
                        xlWorkSheet = xlWorkBook.Worksheets(newSheetName)
                        saveAllSetting(startPosition)
                        endSheetNo = start_pos
                    End If
                    getAllSetting()
                    createFlowChart("OP", Convert.ToInt16(Rs("pNo")), Rs("pName"), xlWorkSheet)
                    'fillValues(Rs("pName"), xlWorkSheet)
                End If
            Next
        End If

        If sheetName = "PFC" Then
            fillPFC(xlWorkSheet)
        ElseIf sheetName = "PQCS" Then
            fillPQCS(xlWorkSheet)
        End If

        xlWorkBook.Save()
        xlWorkBook.Close() : xlApp.Quit()

        ' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook) : xlWorkBook = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet) : xlWorkSheet = Nothing
        'createKakoTora()
        'createOGS()


    End Sub
    Private Sub createOGS()
        Dim desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        If (Not System.IO.Directory.Exists(desktop + "\Reports\OGS")) Then
            System.IO.Directory.CreateDirectory(desktop + "\Reports\OGS")
        Else
            If (System.IO.File.Exists(desktop + "\Reports\OGS\OGS.xlsx")) Then
                My.Computer.FileSystem.DeleteFile(desktop + "\Reports\OGS\OGS.xlsx")
                My.Computer.FileSystem.CopyFile(Application.StartupPath + "\reference\OGS.xlsx",
                                            desktop + "\Reports\OGS\OGS.xlsx")

                makeOGS(desktop + "\Reports\OGS\OGS.xlsx", "OGS")
            Else
                My.Computer.FileSystem.CopyFile(Application.StartupPath + "\reference\OGS.xlsx",
                                            desktop + "\Reports\OGS\OGS.xlsx")
                makeOGS(desktop + "\Reports\OGS\OGS.xlsx", "OGS")
            End If
        End If
    End Sub
    Private Sub makeOGS(ByVal fileSrc As String, ByVal sheetName As String)
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open(fileSrc)
        createSheet(xlWorkBook, sheetName)
        sheetCountOGS = sheetCountOGS + 1
        Dim newSheetName As String = sheetName + " (" + sheetCount.ToString() + ")"
        xlWorkSheet = xlWorkBook.Worksheets(newSheetName)
        Dim ldtPartDetails As New DataTable
        ldtPartDetails = gfnSelectQueryDt("SELECT DISTINCT pName, pNo FROM partDetails WHERE pfcName='" + PFC_NAME + "' ORDER BY pNo")
        Dim pNoSb As New StringBuilder
        Dim pNameSb As New StringBuilder
        If ldtPartDetails.Rows.Count = 0 Then
            Return
        End If
        Dim ogsStartPos As Integer = 5
        Dim circleStart As Integer = 65
        Dim reactStart As Integer = 8
        For Each Rs In ldtPartDetails.Rows
            pNoSb.Append(Rs("pNo"))
            pNameSb.Append(Rs("pName"))
            Dim sc = xlWorkSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeFlowchartConnector, 0, 0, 36, 36)
            sc.Left = xlWorkSheet.Range(Chr(circleStart) & ogsStartPos).Left + 3
            sc.Top = xlWorkSheet.Cells(ogsStartPos, 7).Top
            sc.Fill.Visible = MsoTriState.msoTrue
            sc.Line.Visible = MsoTriState.msoTrue
            sc.Line.Style = MsoLineStyle.msoLineSingle
            sc.Fill.ForeColor.RGB = RGB(255, 255, 255)
            sc.Line.Weight = 1
            sc.TextFrame.Characters.Text = Convert.ToInt16(Rs("pNo"))
            sc.TextFrame.Characters.Font.Name = "MS PGothic"
            sc.TextFrame.Characters.Font.Size = 10
            sc.TextFrame.Characters.Font.Color = RGB(0, 0, 0)
            sc.TextFrame.Characters.Font.FontStyle = "Regular"
            sc.TextFrame.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            sc.TextFrame.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            sc.TextFrame.Orientation = MsoTextOrientation.msoTextOrientationHorizontal
            sc.TextFrame.Characters.Font.Strikethrough = False
            sc.TextFrame.Characters.Font.Superscript = False
            sc.TextFrame.Characters.Font.Subscript = False
            sc.TextFrame.Characters.Font.Underline = Microsoft.Office.Core.XlUnderlineStyle.xlUnderlineStyleNone
            sc.TextFrame.Characters.Font.ColorIndex = Excel.Constants.xlAutomatic
            Dim ldtQMS As New DataTable
            Dim dr As DataRow
            reactStart = 8
            ldtQMS = gfnSelectQueryDt("SELECT * FROM QMS WHERE pName='" + Rs("pName") + "' AND pfcName='" + PFC_NAME + "'")
            If ldtQMS.Rows.Count > 0 Then
                For Each dr In ldtQMS.Rows
                    Dim reactangle = xlWorkSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 200, 30)
                    reactangle.Left = xlWorkSheet.Range(Chr(circleStart) & reactStart).Left + 3
                    reactangle.Top = xlWorkSheet.Cells(reactStart, 7).Top
                    reactangle.Line.Visible = MsoTriState.msoTrue
                    reactangle.Line.Style = MsoLineStyle.msoLineSingle
                    reactangle.Fill.ForeColor.RGB = RGB(255, 255, 0)
                    reactangle.Line.Weight = 2
                    reactangle.TextFrame.Characters.Text = dr("QMS")
                    reactangle.TextFrame.Characters.Font.Name = "MS PGothic"
                    reactangle.TextFrame.Characters.Font.Size = 10
                    reactangle.TextFrame.Characters.Font.FontStyle = "Regular"
                    reactangle.TextFrame.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    reactangle.TextFrame.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    reactangle.TextFrame.Orientation = MsoTextOrientation.msoTextOrientationHorizontal
                    reactangle.TextFrame.Characters.Font.Strikethrough = False
                    reactangle.TextFrame.Characters.Font.Superscript = False
                    reactangle.TextFrame.Characters.Font.Subscript = False
                    reactangle.TextFrame.Characters.Font.Underline = Microsoft.Office.Core.XlUnderlineStyle.xlUnderlineStyleNone
                    reactangle.TextFrame.Characters.Font.ColorIndex = Excel.Constants.xlAutomatic
                    reactStart = reactStart + 2.5
                Next
            End If
            If ldtPartDetails.Rows.Count > 0 Then
                circleStart = 70
            End If
        Next
        For Each Rs In ldtPartDetails.Rows

        Next
        fillOGS(xlWorkSheet, pNoSb, pNameSb)
        xlWorkBook.Save()
        xlWorkBook.Close() : xlApp.Quit()

        ' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook) : xlWorkBook = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet) : xlWorkSheet = Nothing
    End Sub
    Private Sub fillOGS(ByVal xs As Excel.Worksheet, ByVal pNoSb As StringBuilder, ByVal pNameSb As StringBuilder)
        xs.Range("C2").Value = product_Name
        xs.Range("G2").Value = "SWITCH NO. : " + PartNo
        xs.Range("I17").Value = "REV. NO - " + REV_NO
        xs.Range("G18").Value = "Issue Date: " + Date.Now.ToString("dd.MM.yy", CultureInfo.InvariantCulture)
        xs.Range("G3").Value = "PROCESS NO. : " + pNoSb.ToString()
        xs.Range("C3").Value = pNameSb.ToString()
    End Sub
    Private Sub fillPFC(ByVal xs As Excel.Worksheet)
        xs.Range("B3").Value = product_Name
        xs.Range("B4").Value = productNo
        xs.Range("I3").Value = customerName
        xs.Range("I4").Value = customerNo
        xs.Range("L4").Value = model
        xs.Range("L1").Value = REV_NO
    End Sub
    Private Sub fillPQCS(ByVal xs As Excel.Worksheet)
        xs.Range("T1").Value = product_Name
        xs.Range("T2").Value = productNo
        xs.Range("L5").Value = customerName
        xs.Range("I5").Value = customerNo
        xs.Range("Q5").Value = model
        xs.Range("L6").Value = REV_NO
        xs.Range("S76").Value = "Issue Date: " + Date.Now.ToString("dd.MM.yy", CultureInfo.InvariantCulture)
    End Sub
    Private Sub createKakoTora()
        Dim desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        If (Not System.IO.Directory.Exists(desktop + "\Reports\Kakotora")) Then
            System.IO.Directory.CreateDirectory(desktop + "\Reports\Kakotora")
        Else
            If (System.IO.File.Exists(desktop + "\Reports\Kakotora\Kakotora.xlsx")) Then
                My.Computer.FileSystem.DeleteFile(desktop + "\Reports\Kakotora\Kakotora.xlsx")
                My.Computer.FileSystem.CopyFile(Application.StartupPath + "\reference\Kakotora.xlsx",
                                            desktop + "\Reports\Kakotora\Kakotora.xlsx")

                fillKakoTora(desktop + "\Reports\Kakotora\Kakotora.xlsx", "Kakotora")
            Else
                My.Computer.FileSystem.CopyFile(Application.StartupPath + "\reference\Kakotora.xlsx",
                                            desktop + "\Reports\Kakotora\Kakotora.xlsx")
                fillKakoTora(desktop + "\Reports\Kakotora\Kakotora.xlsx", "Kakotora")
            End If
        End If
    End Sub
    Private Sub fillKakoTora(ByVal fileSrc As String, ByVal sheetName As String)
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open(fileSrc)
        createSheet(xlWorkBook, sheetName)
        sheetCount = 2
        Dim newSheetName As String = sheetName + " (" + sheetCount.ToString() + ")"
        xlWorkSheet = xlWorkBook.Worksheets(newSheetName)
        If xlWorkSheet.Range("BG10").Value = "" Then
            Dim str As String = InputBox("Problem Observed")
            xlWorkSheet.Range("BG10").Value = str
        End If
        If xlWorkSheet.Range("BG17").Value = "" Then
            Dim str As String = InputBox("ContinueMeasure Taken")
            xlWorkSheet.Range("BG17").Value = str
        End If
        If xlWorkSheet.Range("AZ69").Value = "" Then
            Dim str As String = InputBox("Effect")
            xlWorkSheet.Range("AZ69").Value = str
        End If
        If xlWorkSheet.Range("AZ78").Value = "" Then
            Dim str As String = InputBox("How to Check first text")
            xlWorkSheet.Range("AZ78").Value = str
        End If
        If xlWorkSheet.Range("AZ84").Value <> "" Then
            Dim str As String = InputBox("How to check second text")
            xlWorkSheet.Range("AZ84").Value = str
        End If
        xlWorkBook.Save()
        xlWorkBook.Close() : xlApp.Quit()

        ' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook) : xlWorkBook = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet) : xlWorkSheet = Nothing
    End Sub
    Private Sub makeReact(ByVal part_no As String, ByVal part_name As String, ByVal qty As String, ByVal xyz As String, ByVal xs As Excel.Worksheet)
        Dim lRange = xs.Range("B" & start_pos, "E" & start_pos + 2)
        lRange.Select()
        With lRange
            .Font.Name = "MS PGothic"
            .Font.Size = 10
            .Font.Strikethrough = False
            .Font.FontStyle = "Bold"
            .Font.Subscript = False
            .Font.Superscript = False
            .Font.TintAndShade = 0
        End With
        makeBorder(lRange, True, True, True, True, True, True)
        lRange = xs.Range("B" & start_pos, "D" & start_pos)
        lRange.Select()
        With lRange
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.XlOrder.xlDownThenOver
            .MergeCells = True
        End With
        lRange = xs.Range("B" & start_pos + 1, "D" & start_pos + 2)
        lRange.Select()
        With lRange
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            .Font.FontStyle = "Regular"
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.XlOrder.xlDownThenOver
            .MergeCells = True
        End With
        lRange = xs.Range("E" & start_pos, "E" & start_pos)
        lRange.Select()
        With lRange
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.XlOrder.xlDownThenOver
            .MergeCells = False
        End With
        lRange = xs.Range("E" & start_pos + 1, "E" & start_pos + 2)
        lRange.Select()
        With lRange
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.XlOrder.xlDownThenOver
            .MergeCells = True
        End With

        xs.Cells(start_pos, 2) = part_no
        xs.Cells(start_pos + 1, 2) = part_name
        xs.Cells(start_pos, 5) = qty
        xs.Cells(start_pos + 1, 5) = xyz
        makeLine(start_pos, start_pos + 3, xs, 70)
        start_pos = start_pos + 4
        saveAllSetting(start_pos)
        'createFlowChart("OP", 10, "MANUAL ASS", xs)
        'fillValues(xs)

    End Sub
    Private Sub makeBorder(ByVal i As Excel.Range, Optional ByVal left As Boolean = False, Optional ByVal top As Boolean = False, Optional ByVal bottom As Boolean = False, Optional ByVal right As Boolean = False, Optional ByVal vertical As Boolean = False, Optional ByVal horizontal As Boolean = False)
        Dim bord(5) As Boolean
        Dim k As Integer
        bord(0) = left
        bord(1) = top
        bord(2) = bottom
        bord(3) = right
        bord(4) = vertical
        bord(5) = horizontal
        For k = 0 To 5
            If bord(k) = True Then
                With i.Borders(k + 7)
                    .LineStyle = Excel.XlLineStyle.xlLineStyleNone
                    .ThemeColor = 2
                    .TintAndShade = 0
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
            End If
        Next
    End Sub
    Public Sub makeLine(ByVal start_l As Integer, ByVal end_l As Integer, ByVal xs As Excel.Worksheet, Optional ByVal h_start As Integer = -1)
        Dim k As Integer
        For k = 0 To 3
            If lines(k) = 1 Then
                If p_lines(k) = 0 Then
                    makeBorder(xs.Range(Chr(73 - k) & start_l + 1 & ":" & Chr(73 - k) & end_l), False, False, False, True, False, False)
                Else
                    makeBorder(xs.Range(Chr(73 - k) & start_l & ":" & Chr(73 - k) & end_l), False, False, False, True, False, False)
                End If
            Else
                k = k - 1
                Exit For
            End If
        Next
        If h_start <> -1 Then
            makeBorder(xs.Range(Chr(h_start) & start_l + 1 & ":" & Chr(73 - k) & start_l + 1), False, True, False, False, False, False)
        End If

        For k = 0 To 3
            p_lines(k) = lines(k)
        Next
    End Sub
    Private Sub createFlowChart(ByVal ftype As String, ByVal lpNo As String, ByVal lpName As String, ByVal xs As Excel.Worksheet)
        Dim k As Integer
        For k = 0 To 3
            If lines(k) = 0 Then
                k = k - 1
                'If OptionButton0.Value = True Then Sheets("info").Cells(7 + k, 2).Value = pr_name
                Exit For
            Else
                If k = 3 Then
                    Exit For
                End If
            End If
        Next
        If ftype = "OP" Then
            Dim sc = xs.Shapes.AddShape(MsoAutoShapeType.msoShapeFlowchartConnector, 0, 0, 36, 36)
            sc.Left = xs.Cells(start_pos, 10).Left - 18
            sc.Top = xs.Cells(start_pos, 7).Top
            sc.Fill.Visible = MsoTriState.msoTrue
            sc.Line.Visible = MsoTriState.msoTrue
            sc.Line.Style = MsoLineStyle.msoLineSingle
            sc.Fill.ForeColor.RGB = RGB(255, 255, 255)
            sc.Line.Weight = 1
            sc.TextFrame.Characters.Text = lpNo
            sc.TextFrame.Characters.Font.Name = "MS PGothic"
            sc.TextFrame.Characters.Font.Size = 10
            sc.TextFrame.Characters.Font.Color = RGB(0, 0, 0)
            sc.TextFrame.Characters.Font.FontStyle = "Regular"
            sc.TextFrame.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            sc.TextFrame.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            sc.TextFrame.Orientation = MsoTextOrientation.msoTextOrientationHorizontal
            sc.TextFrame.Characters.Font.Strikethrough = False
            sc.TextFrame.Characters.Font.Superscript = False
            sc.TextFrame.Characters.Font.Subscript = False
            sc.TextFrame.Characters.Font.Underline = Microsoft.Office.Core.XlUnderlineStyle.xlUnderlineStyleNone
            sc.TextFrame.Characters.Font.ColorIndex = Excel.Constants.xlAutomatic

            Dim lRange = xs.Range("B" & start_pos, Chr(74 - k - 2) & start_pos + 2)
            lRange.Select()
            With lRange
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = Excel.XlOrder.xlDownThenOver
                .MergeCells = True
                With .Font
                    .Bold = True
                    .Name = "MS PGothic"
                    .Size = 13
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .Underline = Microsoft.Office.Core.XlUnderlineStyle.xlUnderlineStyleNone
                    .ColorIndex = Excel.Constants.xlAutomatic
                    .TintAndShade = 0
                    .ThemeFont = Excel.XlThemeFont.xlThemeFontNone
                End With
            End With
            xs.Cells(start_pos, 2) = lpName
            If True Then makeLine(start_pos, start_pos + 3, xs)
            start_pos = start_pos + 4
            saveAllSetting(start_pos)
        End If
    End Sub
    Private Sub toBeContinue(ByVal xs As Excel.Worksheet)
        Dim i, j, k As Integer
        start_pos = start_pos + 3
        i = 1
        For k = 3 To 0 Step -1
            If lines(k) = 1 Then
                For j = k To 0 Step -1
                    makeBorder(xs.Range(Chr(73 - j - 3) & start_pos & ":" & Chr(73 - j) & start_pos), False, False, False, True, False, False)
                Next
                Dim lRange = xs.Range(Chr(73 - k) & start_pos - 1, Chr(73 - k) & start_pos - 1)
                lRange.Select()
                With lRange
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = True
                    .ReadingOrder = Excel.XlOrder.xlDownThenOver
                    .MergeCells = True
                    With .Font
                        .Bold = True
                        .Name = "MS PGothic"
                        .Size = 8
                        .Strikethrough = False
                        .Superscript = False
                        .Subscript = False
                        .Underline = Microsoft.Office.Core.XlUnderlineStyle.xlUnderlineStyleNone
                        .ColorIndex = Excel.Constants.xlAutomatic
                        .TintAndShade = 0
                        .ThemeFont = Excel.XlThemeFont.xlThemeFontNone
                    End With
                End With
                makeBorder(xs.Range(Chr(73 - k - 3) & start_pos & ":" & Chr(73 - k) & start_pos), False, True, False, True, False, False)
                xs.Cells(start_pos - 1, 9 - k - 3) = "CONTD. FROM P.NO." + 10
                start_pos = start_pos - 1
                i = i + 1
            Else
                start_pos = start_pos - 1
            End If
        Next
        start_pos = start_pos + i
        saveAllSetting(start_pos)
    End Sub

    Private Sub fillValues(ByVal pFullName As String, ByVal xs As Excel.Worksheet)
        Dim i As Integer
        Dim lRange As Excel.Range
        Dim query As String
        Dim ldtPartDetails As New DataTable
        ldtPartDetails = gfnSelectQueryDt("SELECT DISTINCT pName FROM partDetails WHERE pfcName='" + PFC_NAME + "'")
        Dim dr As DataRow
        i = 12
        lRange = xs.Range("P" & i)
        'Dim pFullName As String

        'For Each dr In ldtPartDetails.Rows
        '    pFullName = dr("pFullName")
        'Next

        lRange.Value = 10 ' Process Number
        lRange = xs.Range("Q" & i)
        lRange.WrapText = False
        lRange.Value = pFullName
        lRange = xs.Range("Q" & i, "R" & i + 2)
        lRange.Justify()

        'QMP
        query = "select * from QMP where pName='" & pFullName & "' AND pfcName='" + PFC_NAME + "'"
        Dim ldtQMP As New DataTable
        ldtQMP = gfnSelectQueryDt(query)
        For Each dr In ldtQMP.Rows
            lRange = xs.Range("U" & i)
            lRange.Value = dr("qmp")
            lRange = xs.Range("U" & i, "U" & i + 1)
            lRange.Justify()
        Next

        'QMS
        query = "select * from QMS where pName='" & pFullName & "' AND pfcName='" + PFC_NAME + "'"
        Dim ldtQMS As New DataTable
        ldtQMS = gfnSelectQueryDt(query)

        For Each dr In ldtQMS.Rows
            lRange = xs.Range("V" & i)
            lRange.Value = dr("qms")
            lRange = xs.Range("V" & i, "V" & i + 4)
            lRange.Justify()
            i = i + 5
        Next

        'Frequency
        query = "select * from frequency where pName='" & pFullName & "' AND pfcName='" + PFC_NAME + "'"
        Dim ldtFrequency As New DataTable
        ldtFrequency = gfnSelectQueryDt(query)

        For Each dr In ldtFrequency.Rows
            lRange = xs.Range("Y" & i)
            lRange.Value = dr("freq")
            lRange = xs.Range("Y" & i, "Y" & i + 1)
            i = i + 5
        Next

        'measuring tool
        query = "select * from measuringTool where pName='" & pFullName & "' AND pfcName='" + PFC_NAME + "'"
        Dim ldtMeasuringTool As New DataTable
        ldtMeasuringTool = gfnSelectQueryDt(query)

        For Each dr In ldtMeasuringTool.Rows
            lRange = xs.Range("Z" & i)
            lRange.Value = dr("measTool")
            lRange = xs.Range("Z" & i, "Z" & i + 1)
            i = i + 5
        Next

        'checked by
        query = "select * from checkedBy where pName='" & pFullName & "' AND pfcName='" + PFC_NAME + "'"
        Dim ldtCheckedBy As New DataTable
        ldtCheckedBy = gfnSelectQueryDt(query)

        For Each dr In ldtCheckedBy.Rows
            lRange = xs.Range("AA" & i)
            lRange.Value = dr("checkedBy")
            lRange = xs.Range("AA" & i, "AA" & i + 1)
            i = i + 5
        Next

        'recording sheet
        query = "select * from recordingSheet where pName='" & pFullName & "' AND pfcName='" + PFC_NAME + "'"
        Dim ldtRecordingSheet As New DataTable
        ldtRecordingSheet = gfnSelectQueryDt(query)

        For Each dr In ldtRecordingSheet.Rows
            lRange = xs.Range("AB" & i)
            lRange.Value = dr("recordSheet")
            lRange = xs.Range("AB" & i, "AB" & i + 1)
            i = i + 5
        Next
    End Sub
    Private Sub EditExcel(ByVal sFile As String)
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open(sFile)           ' WORKBOOK TO OPEN THE EXCEL FILE.
        xlWorkSheet = xlWorkBook.Worksheets("PQCS")
        xlWorkSheet.Cells(1, 1) = 25

        ''        circle()
        Dim sc = xlWorkSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeFlowchartConnector, xlWorkSheet.Cells(2, 2).Left, xlWorkSheet.Cells(2, 2).Top, 38, 38)
        sc.Fill.Visible = MsoTriState.msoTrue
        sc.TextFrame.Characters.Text = 10
        sc.Fill.ForeColor.RGB = RGB(128, 0, 0)
        sc.TextFrame.Characters.Font.Name = "MS PGothic"
        sc.TextFrame.Characters.Font.Size = 10
        sc.TextFrame.Characters.Font.FontStyle = "Regular"
        sc.TextFrame.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        sc.TextFrame.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        sc.TextFrame.Orientation = MsoTextOrientation.msoTextOrientationHorizontal

        '' rectangle
        'Dim rectangle = xlWorkSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, xlWorkSheet.Cells(5, 5).left, xlWorkSheet.Cells(5, 5).Top, 200, 30)
        'rectangle.TextFrame.Characters.Text = "S13050"
        'rectangle.Fill.ForeColor.RGB = RGB(255, 0, 0)
        'rectangle.Line.Weight = 2
        'rectangle.Line.Visible = MsoTriState.msoTrue
        'rectangle.Line.Style = MsoLineStyle.msoLineSingle
        'rectangle.SetShapesDefaultProperties()
        'rectangle.TextFrame.Characters.Font.Name = "MS PGothic"
        'rectangle.TextFrame.Characters.Font.Size = 10
        'rectangle.TextFrame.Characters.Font.FontStyle = "Regular"
        'rectangle.TextFrame.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        'rectangle.TextFrame.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        'rectangle.TextFrame.Orientation = MsoTextOrientation.msoTextOrientationHorizontal

        'Dim decision = xlWorkSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeFlowchartDecision, xlWorkSheet.Cells(7, 7).left, xlWorkSheet.Cells(7, 7).Top, 50, 36)

        'xlWorkSheet.Range("B7:C7").MergeCells = True
        'xlWorkSheet.Cells("B7") = "S13080"
        'xlWorkSheet.Range("B8:D8").MergeCells = True
        ' save and close excel
        xlWorkBook.Save()
        xlWorkBook.Close() : xlApp.Quit()

        ' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook) : xlWorkBook = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet) : xlWorkSheet = Nothing
        MsgBox("Edited Succesfully")
    End Sub
    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        main.Show()
        Me.Close()
    End Sub

    Private Sub cmdReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReport.Click
        createDirectory()
        generatePFC()
        generatePQCS()
        generatePCC()
        MsgBox("Edited Succesfully")
    End Sub

    Private Sub frmReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        PFC_NAME = main.cboPFC.Text
        Dim ldtRevNo As New DataTable
        ldtRevNo = gfnSelectQueryDt("SELECT * FROM revisionNumber WHERE pfcName='" + PFC_NAME + "'")
        Dim dr As DataRow
        For Each dr In ldtRevNo.Rows
            REV_NO = dr("revNo")
        Next

        Dim ldtJoin As New DataTable
        ldtJoin = gfnSelectQueryDt("SELECT * from QMS")
        Dim _json As String = GetJson(ldtJoin)
        createPQCS(Application.StartupPath + "\db.xlsx", "P_MRPL", 12, 66)



    End Sub
    Private Function GetJson(ByVal dt As DataTable) As String
        'Dim Jserializer = New System.Web.Script.Serialization()
        Dim rowsList As New List(Of Dictionary(Of String, Object))()
        Dim row As Dictionary(Of String, Object)
        For Each dr As DataRow In dt.Rows
            row = New Dictionary(Of String, Object)()
            For Each col As DataColumn In dt.Columns
                row.Add(col.ColumnName, dr(col))
            Next
            rowsList.Add(row)
        Next
        Return JsonConvert.SerializeObject(rowsList)
    End Function
End Class