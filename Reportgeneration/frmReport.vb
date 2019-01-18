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
    Dim xsInfo As Excel.Worksheet
    Dim start_sheet As Integer
    Dim lines(3), p_lines(3) As Integer
    Dim sheetCount As Integer
    Dim sheetCountOGS As Integer = 1
    Dim start_pos, last_pos, endSheet, currentSheetNo, sheetNo As Integer
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
    Dim PFC_NAME, REV_NO, PartNo, RELEASE_TYPE, PRODUCT_RATING As String
    Dim PROCESS_NO As String = ""
    Dim LAST_PROCESS As String
    Dim PROCESS_NAME As String = ""
    Private Sub getAllSetting()
        start_pos = GetSetting("st_pos", "st_pos", "st_pos")
        last_pos = GetSetting("last_pos", "last_pos", "last_pos")
        endSheet = GetSetting("endSheet", "endSheet", "endSheet")
        currentSheetNo = GetSetting("currentSheetNo", "currentSheetNo", "currentSheetNo")
        sheetNo = GetSetting("sheetNo", "sheetNo", "sheetNo")
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
    Private Sub PCHART()
        Dim desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        If (Not System.IO.Directory.Exists(desktop + "\Reports\PCHART")) Then
            System.IO.Directory.CreateDirectory(desktop + "\Reports\PCHART")
        Else
            If (System.IO.File.Exists(desktop + "\Reports\PCHART\PCHART.xlsx")) Then
                My.Computer.FileSystem.DeleteFile(desktop + "\Reports\PCHART\PCHART.xlsx")
                My.Computer.FileSystem.CopyFile(Application.StartupPath + "\reference\PCHART.xlsx",
                                            desktop + "\Reports\PCHART\PCHART.xlsx")
                startPCHART(desktop + "\Reports\PCHART\PCHART.xlsx", "PCHART")
            Else
                My.Computer.FileSystem.CopyFile(Application.StartupPath + "\reference\PCHART.xlsx",
                                            desktop + "\Reports\PCHART\PCHART.xlsx")
                startPCHART(desktop + "\Reports\PCHART\PCHART.xlsx", "PCHART")
            End If
        End If
    End Sub
    Private Sub startPCHART(ByVal fileSrc As String, ByVal sheetName As String)

        Dim ldtPartDetails As New DataTable
        ldtPartDetails = gfnSelectQueryDt("Select * from partDetails where pfcName='" + PFC_NAME + "'")

        If ldtPartDetails.Rows.Count = 0 Then
            Exit Sub
        End If

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open(fileSrc)
        Dim newSheetName = fnCreateSheet(xlWorkBook, sheetName, currentSheetNo)
        xlWorkSheet = xlWorkBook.Worksheets(newSheetName)
        Dim text As String = "As per table"
        xlWorkSheet.Range("B16").Value = text
        xlWorkSheet.Range("B21").Value = text
        xlWorkSheet.Range("B25").Value = text
        xlWorkSheet.Range("C66").Value = "CUSTOMER - " & text
        xlWorkSheet.Range("M66").Value = "PRODUCT NAME - " & text
        xlWorkSheet.Range("X66").Value = "PRODUCT NO - " & text

        Dim today = DateTime.Today
        xlWorkSheet.Range("A7").Value = DateTime.Now.ToString("MMM")
        xlWorkSheet.Range("B7").Value = today.Year
        Dim startPartName As Integer = 68
        Dim dr As DataRow
        For Each dr In ldtPartDetails.Rows
            startPartName = startPartName + 2
            xlWorkSheet.Range("A" & startPartName).Value = dr("partName")
        Next
        xlWorkSheet = xlWorkBook.Worksheets("Sheet1")
        startPartName = 1
        For Each dr In ldtPartDetails.Rows
            startPartName = startPartName + 1
            xlWorkSheet.Range("A" & startPartName).Value = dr("partName")
            xlWorkSheet.Range("B" & startPartName).Value = dr("productName")
            xlWorkSheet.Range("C" & startPartName).Value = dr("productNo")
            xlWorkSheet.Range("D" & startPartName).Value = dr("customer")
        Next
        xlWorkSheet.Range("A1:" & "D" & startPartName).CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture)
        Dim xlRange As Excel.Range
        xlRange = xlWorkSheet.Range("A1:" & "D" & startPartName)
        If System.IO.File.Exists(Application.StartupPath & "\pic.jpg") Then
            My.Computer.FileSystem.DeleteFile(Application.StartupPath & "\pic.jpg")
        End If
        Dim oChtobj As Excel.ChartObject = xlWorkSheet.ChartObjects.add(xlRange.Left, xlRange.Top, xlRange.Width, xlRange.Height)
        Dim oCht As Excel.Chart
        oCht = oChtobj.Chart
        oCht.Paste()
        oCht.Export(Filename:=Application.StartupPath & "\pic.jpg")
        oChtobj.Delete()
        xlWorkSheet = xlWorkBook.Worksheets(newSheetName)
        xlWorkSheet.Shapes.AddPicture(Application.StartupPath & "\pic.jpg", False, True, 0, 0, 500, 500)
        xlWorkBook.Save()
        xlWorkBook.Close() : xlApp.Quit()

        ' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook) : xlWorkBook = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet) : xlWorkSheet = Nothing
    End Sub
    Private Sub generatePQCS()
        Dim desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim model = "MRPL"
        Dim startP, EndP As Integer
        startP = DEFAULTPQCSSTARTVALUE
        EndP = DEFAULTPQCSSTARTVALUE
        currentSheetNo = 1
        sheetNo = 1
        If (Not System.IO.Directory.Exists(desktop + "\Reports\PQCS")) Then
            System.IO.Directory.CreateDirectory(desktop + "\Reports\PQCS")
        Else
            If (System.IO.File.Exists(desktop + "\Reports\db.xlsx")) Then
                My.Computer.FileSystem.DeleteFile(desktop + "\Reports\db.xlsx")
                My.Computer.FileSystem.CopyFile(Application.StartupPath + "\db.xlsx",
                                            desktop + "\Reports\db.xlsx")
                saveSettings("st_pos", startP)
                saveSettings("last_pos", startP)
                saveSettings("endSheet", EndP)
                saveSettings("currentSheetNo", currentSheetNo)
                saveSettings("sheetNo", sheetNo)
                getAllSetting()
                createPQCS(desktop + "\Reports\db.xlsx", "P_MRPL", startP, EndP)
            Else
                My.Computer.FileSystem.CopyFile(Application.StartupPath + "\db.xlsx",
                                            desktop + "\Reports\db.xlsx")
                saveSettings("st_pos", startP)
                saveSettings("last_pos", startP)
                saveSettings("endSheet", EndP)
                saveSettings("currentSheetNo", currentSheetNo)
                getAllSetting()
                createPQCS(desktop + "\Reports\db.xlsx", "P_MRPL", startP, EndP)
            End If
        End If
    End Sub
    Private Sub saveSettings(ByVal name As String, ByVal val As Integer)
        SaveSetting(name, name, name, val)
    End Sub
    Private Sub createPQCS(ByVal fileSrc As String, ByVal sheetName As String, ByVal startNo As Integer, ByVal lastNo As Integer)
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open(fileSrc)
        Dim newSheetName = fnCreateSheet(xlWorkBook, sheetName, currentSheetNo)
        xlWorkSheet = xlWorkBook.Worksheets(newSheetName)
        Dim flowChart As Boolean = False
        Dim ldtPartDetails As New DataTable
        Dim ldtPartDetailsNew As New DataTable
        Dim dr, rs As DataRow
        Dim xsProcess As Excel.Worksheet
        xsInfo = xlWorkBook.Worksheets("INFO")
        ldtPartDetails = gfnSelectQueryDt("SELECT DISTINCT pName, pNo FROM partDetails WHERE pfcName='" + PFC_NAME + "' order by pNo")
        Dim i, j, k As Integer
        Dim pr_no, flow_sy_type As String
        For Each dr In ldtPartDetails.Rows
            xsProcess = xlWorkBook.Worksheets("Processes")
            i = 2
            Do While Len(xsProcess.Range("A" & i).Value) <> 0
                If Trim(xsProcess.Range("E" & i).Value) = Trim(dr("pName")) Then
                    pr_no = xsProcess.Range("A" & i).Value
                    flow_sy_type = xsProcess.Range("F" & i).Value
                    Exit Do
                End If
                i = i + 1
            Loop
            xsProcess = xlWorkBook.Worksheets("PQCS_DB")
            i = 5
            flowChart = False
            Do While Len(xsProcess.Range("K" & i).Value) <> 0
                j = i
                Do While checkDBNull(xsProcess, i, j) = True
                    j = j + 1
                Loop
                If xsProcess.Range("K" & i).Value = pr_no Then
                    If flowChart = False Then
                        'For k = 0 To 3
                        '    If lines(k) = 0 Then
                        '        lines(k) = 1
                        '        Exit For
                        '    End If
                        'Next
                        lines(k) = 1
                        getAllSetting()
                        If start_pos + j - i + 1 >= DEFAULTPQCSENDVALUE Or start_pos + 3 >= DEFAULTPQCSENDVALUE Then
                            If start_pos < DEFAULTPQCSENDVALUE Then makeLine(start_pos, DEFAULTPQCSENDVALUE - 1, xlWorkSheet)
                            newSheetName = fnCreateSheet(xlWorkBook, sheetName, currentSheetNo + 1)
                            xlWorkSheet = xlWorkBook.Worksheets(newSheetName)
                            saveSettings("currentSheetNo", currentSheetNo + 1)
                            toBeContinue(xlWorkSheet, dr("pNo"))
                        End If
                        last_pos = start_pos
                        saveSettings("last_pos", last_pos)
                        getAllSetting()
                        xlWorkSheet = xlWorkBook.Worksheets(sheetName & "_" & currentSheetNo)
                        ldtPartDetailsNew = gfnSelectQueryDt("SELECT * FROM partDetails WHERE pfcName='" + PFC_NAME + "' AND pName='" + dr("pName") + "'")
                        For Each rs In ldtPartDetailsNew.Rows
                            makeReact(rs("partNo"), rs("partName"), rs("qty"), rs("xyz"), xlWorkSheet, xlWorkBook, dr("pNo"))
                        Next
                        createFlowChart(flow_sy_type, Convert.ToInt16(dr("pNo")), dr("pName"), xlWorkSheet, xlWorkBook)
                        fillValuesPQCS(xlWorkSheet, xsProcess, i, j, dr("pNo"), xlWorkBook, dr("pName"))
                        flowChart = True
                    Else
                        'xlWorkSheet = xlWorkBook.Worksheets(sheetName & "_" & currentSheetNo)
                        fillValuesPQCS(xlWorkSheet, xsProcess, i, j, dr("pNo"), xlWorkBook, dr("pName"))
                    End If
                End If
                i = j + 2
            Loop
        Next
        xlWorkBook.Application.DisplayAlerts = False
        'xlWorkBook.Sheets("INFO").Delete
        xlWorkBook.Sheets("Processes").Delete
        xlWorkBook.Sheets("PQCS_DB").Delete
        xlWorkBook.Application.DisplayAlerts = True
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
    Private Sub fillValuesPQCS(ByVal xsPQCS As Excel.Worksheet, ByVal xsPQCSDB As Excel.Worksheet, ByVal i As Integer, ByVal j As Integer, ByVal pNo As Integer, ByVal xw As Excel.Workbook, ByVal pName As String)
        getAllSetting()
        Dim defaultSheetName As String = "P_MRPL"
        If last_pos + j - i + 1 >= DEFAULTPQCSENDVALUE Then
            If sheetNo = currentSheetNo Then
                If last_pos < DEFAULTPQCSENDVALUE Then makeLine(last_pos, DEFAULTPQCSENDVALUE - 1, xsPQCS)
                Dim sheetName = fnCreateSheet(xw, defaultSheetName, currentSheetNo + 1)
                xsPQCS = xw.Worksheets(sheetName)
                saveSettings("currentSheetNo", currentSheetNo + 1)
                saveSettings("sheetNo", currentSheetNo + 1)
                saveSettings("st_pos", DEFAULTPQCSSTARTVALUE)
                saveSettings("last_pos", DEFAULTPQCSSTARTVALUE)
                getAllSetting()
                toBeContinue(xsPQCS, pNo)
            Else
                saveSettings("currentSheetNo", currentSheetNo + 1)
                saveSettings("sheetNo", currentSheetNo + 1)
                saveSettings("last_pos", DEFAULTPQCSSTARTVALUE)
                getAllSetting()
                For k = 0 To 3
                    If lines(k) = 1 Then last_pos = last_pos + 1
                Next
                saveSettings("last_pos", last_pos)
                getAllSetting()
            End If
        End If
        xsPQCS = xw.Worksheets(defaultSheetName & "_" & currentSheetNo)
        If RELEASE_TYPE = "Prototype" Then
            xsPQCS.Shapes.Item("Check Box 1").OLEFormat.Object.Value = True
            xsPQCS.Shapes.Item("Check Box 2").OLEFormat.Object.Value = False
            xsPQCS.Shapes.Item("Check Box 3").OLEFormat.Object.Value = False
        ElseIf RELEASE_TYPE = "Prelaunch" Then
            xsPQCS.Shapes.Item("Check Box 1").OLEFormat.Object.Value = False
            xsPQCS.Shapes.Item("Check Box 2").OLEFormat.Object.Value = False
            xsPQCS.Shapes.Item("Check Box 3").OLEFormat.Object.Value = True
        ElseIf RELEASE_TYPE = "Production" Then
            xsPQCS.Shapes.Item("Check Box 1").OLEFormat.Object.Value = False
            xsPQCS.Shapes.Item("Check Box 2").OLEFormat.Object.Value = True
            xsPQCS.Shapes.Item("Check Box 3").OLEFormat.Object.Value = False
        End If
        xsPQCSDB.Range("O" & i & ":AC" & j).Copy(xsPQCS.Range("P" & last_pos))
        If xsPQCSDB.Range("AA" & last_pos).Value = "KAKOTORA " Then
            lblCurrentReport.Text = "KAKOTORA"
            createKakoTora()
        End If
        last_pos = last_pos + j - i + 2

        If sheetNo = currentSheetNo And last_pos > start_pos Then
            If True = False Then makeLine(start_pos, last_pos - 1, xsPQCS)
            start_pos = last_pos
            saveSettings("last_pos", start_pos)
        End If
        saveSettings("last_pos", last_pos)
        getAllSetting()
        If xsPQCSDB.Range("AF" & i).Value = 1 Then
            'PCC Here
            genPCC(xsPQCSDB.Range("AH" & i).Value, "", "", "", xsPQCSDB.Range("AI" & i).Value, xsPQCSDB.Range("AJ" & i).Value, xsPQCSDB.Range("AG" & i).Value, 1, xw, pNo, pName)
            lblCurrentReport.Text = "PCC"
        End If

    End Sub
    Private Sub genPCC(TXT1 As String, TXT2 As String, TXT3 As String, TXT4 As String, TXT5 As String, TXT6 As String, Optional PO As Integer = 1, Optional MER As Integer = 1, Optional xw As Excel.Workbook = Nothing, Optional lPno As Integer = 0, Optional pName As String = "")
        Dim pcc_file_name As String
        Dim current_sheet As Integer
        Dim BASE As Integer, offset_val As Integer
        Dim xsPCC As Excel.Worksheet
        If xsInfo.Cells(12, 7).Value > xsInfo.Cells(13, 7).Value Then
            current_sheet = xsInfo.Cells(12, 7).Value
        Else
            current_sheet = xsInfo.Cells(13, 7).Value
        End If
        If xsInfo.Cells(12 + PO, 6).Value > 4 Then
            If current_sheet = xsInfo.Cells(12 + PO, 7).Value Then
                If current_sheet > 0 Then
                    Dim result As String = String.Join(",", PROCESS_NO.Split(",").Distinct().ToArray())
                    xsPCC = xw.Worksheets("PCC_" & currentSheetNo)
                    xsPCC.Range("C7").Value = result
                    PROCESS_NO = ""
                    result = String.Join(",", PROCESS_NAME.Split(",").Distinct().ToArray())
                    xsPCC.Range("C6").Value = result
                End If
                Dim sheetName = fnCreateSheet(xw, "PCC", current_sheet + 1)
                xsPCC = xw.Worksheets(sheetName)
                xsPCC.Range("C7").Value = lPno
                xsPCC.Range("C6").Value = pName
                xsInfo.Cells(12 + PO, 7).Value = current_sheet + 1
                current_sheet = xsInfo.Cells(12 + PO, 7).Value
                xsInfo.Cells(9, 6).Value = xsInfo.Cells(12 + PO, 7).Value
                pcc_file_name = "PCC_" & current_sheet
                xsInfo.Cells(12 + PO, 6).Value = 0
            Else
                xsInfo.Cells(12 + PO, 7).Value = xsInfo.Cells(12 + PO, 7).Value + 1
                current_sheet = xsInfo.Cells(12 + PO, 7).Value
                pcc_file_name = "PCC_" & current_sheet
                xsInfo.Cells(12 + PO, 6).Value = 0
            End If
        Else
            current_sheet = xsInfo.Cells(12 + PO, 7).Value
            pcc_file_name = "PCC_" & current_sheet
        End If
        PROCESS_NO = PROCESS_NO & lPno & ","
        PROCESS_NAME = PROCESS_NAME & pName & ","
        BASE = xsInfo.Cells(12 + PO, 5).Value
        offset_val = xsInfo.Cells(12 + PO, 6).Value
        xsInfo.Cells(12 + PO, 6).Value = xsInfo.Cells(12 + PO, 6).Value + 1

        If MER = 0 Then
            xsPCC = xw.Worksheets(pcc_file_name)
            xsPCC.Cells(6, BASE + offset_val).Value = TXT1
            xsPCC.Cells(4, BASE + offset_val).Value = TXT2
            xsPCC.Cells(4, BASE + offset_val).Value = TXT3
            xsPCC.Cells(3, BASE + offset_val).Value = TXT4
        Else
            xsPCC = xw.Worksheets(pcc_file_name)
            With xsPCC.Range(Chr(64 + BASE + offset_val) & 3 & ":" & Chr(64 + BASE + offset_val) & 7)
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .WrapText = True
                .Orientation = 90
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
            End With
            xsPCC.Cells(3, BASE + offset_val).Value = TXT1 & "(" & lPno & ")"
        End If
        xsPCC = xw.Worksheets(pcc_file_name)
        xsPCC.Cells(8, BASE + offset_val).Value = TXT5 & "            " & TXT6
        xsPCC.Range("E5").Value = REV_NO
        xsPCC.Range("S1").Value = current_sheet
        xsPCC.Range("Q3").Value = "ASSY"
        xsPCC.Range("D1").Value = "MRPL"

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
            saveSettings("st_pos", startPosition)
            For Each Rs In ldtPartDetails.Rows
                If start_pos >= endSheetNo Then
                    createSheet(xlWorkBook, sheetName)
                    sheetCount = sheetCount + 1
                    newSheetName = sheetName + " (" + sheetCount.ToString() + ")"
                    xlWorkSheet = xlWorkBook.Worksheets(newSheetName)
                    saveSettings("st_pos", startPosition)
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
                        saveSettings("st_pos", startPosition)
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
    Private Sub makeReact(ByVal part_no As String, ByVal part_name As String, ByVal qty As String, ByVal xyz As String, ByVal xs As Excel.Worksheet, Optional ByVal xb As Excel.Workbook = Nothing, Optional ByVal lPno As String = "")
        If start_pos + 3 >= DEFAULTPQCSENDVALUE Then
            If start_pos < DEFAULTPQCSENDVALUE Then makeLine(start_pos, DEFAULTPQCSENDVALUE - 1, xs)
            Dim newSheetName = fnCreateSheet(xb, "P_MRPL", currentSheetNo + 1)
            xs = xb.Worksheets(newSheetName)
            saveSettings("currentSheetNo", currentSheetNo + 1)
            saveSettings("st_pos", DEFAULTPQCSSTARTVALUE)
            saveSettings("last_pos", DEFAULTPQCSSTARTVALUE)
            getAllSetting()
            toBeContinue(xs, lPno)
        End If

        Dim lRange = xs.Range("B" & start_pos, "E" & start_pos + 2)
        'lRange.Select()
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
        'lRange.Select()
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
        'lRange.Select()
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
        'lRange.Select()
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
        'lRange.Select()
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
        saveSettings("st_pos", start_pos)
        getAllSetting()
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
    Private Sub createFlowChart(ByVal ftype As String, ByVal lpNo As String, ByVal lpName As String, ByVal xs As Excel.Worksheet, Optional ByVal xb As Excel.Workbook = Nothing)
        If start_pos + 3 >= DEFAULTPQCSENDVALUE Then
            If start_pos < DEFAULTPQCSENDVALUE Then makeLine(start_pos, DEFAULTPQCSENDVALUE - 1, xs)
            Dim newSheetName = fnCreateSheet(xb, "P_MRPL", currentSheetNo + 1)
            xs = xb.Worksheets(newSheetName)
            saveSettings("currentSheetNo", currentSheetNo + 1)
            saveSettings("st_pos", DEFAULTPQCSSTARTVALUE)
            saveSettings("last_pos", DEFAULTPQCSSTARTVALUE)
            getAllSetting()
            toBeContinue(xs, lpNo)
        End If

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
            'lRange.Select()
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
            saveSettings("st_pos", start_pos)
        End If
    End Sub
    Private Function toBeContinue(ByVal xs As Excel.Worksheet, Optional ByVal pno As String = "")
        Dim i, j, k As Integer
        start_pos = start_pos + 3
        i = 1
        For k = 3 To 0 Step -1
            If lines(k) = 1 Then
                For j = k To 0 Step -1
                    makeBorder(xs.Range(Chr(73 - j - 3) & start_pos & ":" & Chr(73 - j) & start_pos), False, False, False, True, False, False)
                Next
                Dim lRange = xs.Range(Chr(73 - k) & start_pos - 1, Chr(73 - k) & start_pos - 1)
                'lRange.Select()
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
                xs.Cells(start_pos - 1, 9 - k - 3) = "CONTD. FROM P.NO." + pno
                start_pos = start_pos - 1
                i = i + 1
            Else
                start_pos = start_pos - 1
            End If
        Next
        start_pos = start_pos + i
        saveSettings("st_pos", start_pos)
        getAllSetting()
        toBeContinue = i
    End Function

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
        statusBar.Value = statusBar.Value + 10
        lblCurrentReport.Text = "PFC"
        generatePFC()
        statusBar.Value = statusBar.Value + 20
        lblCurrentReport.Text = "PQCS"
        generatePQCS()
        statusBar.Value = statusBar.Value + 50
        'lblCurrentReport.Text = "OGS"
        'createOGS()
        statusBar.Value = statusBar.Value + 70
        lblCurrentReport.Text = "PCHART"
        PCHART()
        statusBar.Value = statusBar.Value + 100
        MsgBox("Edited Succesfully")
    End Sub

    Private Sub frmReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        PFC_NAME = main.cboPFC.Text
        Dim ldtRevNo As New DataTable
        ldtRevNo = gfnSelectQueryDt("SELECT * FROM revisionNumber WHERE pfcName='" + PFC_NAME + "'")
        Dim dr As DataRow
        For Each dr In ldtRevNo.Rows
            REV_NO = dr("revNo")
            RELEASE_TYPE = dr("releasingPhase")
            PRODUCT_RATING = dr("productRating")
        Next

        'Dim ldtJoin As New DataTable
        'ldtJoin = gfnSelectQueryDt("SELECT * from QMS")
        'Dim _json As String = GetJson(ldtJoin)
        'generatePQCS()
        'PCHART()
        MsgBox("Completed")
    End Sub
    'Private Function GetJson(ByVal dt As DataTable) As String
    '    'Dim Jserializer = New System.Web.Script.Serialization()
    '    Dim rowsList As New List(Of Dictionary(Of String, Object))()
    '    Dim row As Dictionary(Of String, Object)
    '    For Each dr As DataRow In dt.Rows
    '        row = New Dictionary(Of String, Object)()
    '        For Each col As DataColumn In dt.Columns
    '            row.Add(col.ColumnName, dr(col))
    '        Next
    '        rowsList.Add(row)
    '    Next
    '    Return JsonConvert.SerializeObject(rowsList)
    'End Function
End Class