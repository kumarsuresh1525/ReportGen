Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Public Class frmMakePFC
    Dim xlApp As Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet
    Dim start_sheet As Integer
    Dim lines(3), p_lines(3) As Integer
    Dim sheetCount As Integer
    Dim sheetCountOGS As Integer = 1
    Dim start_pos As Integer
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
    Private Sub LoadProcessFromExcel()
        Con.Close()
        'Create a standard SELECT SQL statement
        Dim selectStatement As String = "SELECT DISTINCT [PROCESS_NAME] FROM [Processes$]"
        'Create a DataAdapter that will be used to populate a DataTable with data
        Dim adapter As New OleDbDataAdapter(selectStatement, connectionString)

        'Populate a DataTable
        Dim excelData As New DataTable
        adapter.Fill(excelData)
        Dim Rs As DataRow
        If excelData.Rows.Count = 0 Then
            Return
        End If
        cboProcessName.Items.Clear()
        For Each Rs In excelData.Rows
            cboProcessName.Items.Add(Rs("PROCESS_NAME"))
        Next

        ''Create a standard SELECT SQL statement
        'selectStatement = "SELECT * FROM [MSIL$] WHERE PROCESS_NAME='MANUAL ASSEMBLY OF STEEL BALL & SPRING'"
        ''Create a DataAdapter that will be used to populate a DataTable with data
        'adapter = New OleDbDataAdapter(selectStatement, connectionString)

        ''Populate a DataTable
        'excelData = New DataTable
        'adapter.Fill(excelData)
        'If excelData.Rows.Count = 0 Then
        '    Return
        'End If
        'Dim str, str1 As String

        'For Each Rs In excelData.Rows
        '    str = Rs("QMP")
        '    str1 = Rs("QMS")
        'Next

    End Sub
    Private Sub LoadModels()
        On Error Resume Next
        Dim i As Integer
        Dim da As New OleDbDataAdapter
        Dim dt As New DataTable
        da.SelectCommand = New OleDbCommand("select * from modelSet", Con)
        da.Fill(dt)

        Dim Rs As DataRow
        If dt.Rows.Count = 0 Then
            Return
        End If

        DGVinfo.Rows.Clear()
        i = 1
        For Each Rs In dt.Rows
            DGVinfo.Rows.Add()
            'DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column1").Value = i
            DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column2").Value = Rs("partNumber")
            DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column3").Value = Rs("partName")
            DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column4").Value = Rs("qty")
            DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column5").Value = Rs("x_y_z")
            i = i + 1
        Next

    End Sub
    Private Sub saveAllSetting(ByVal val As Integer)
        SaveSetting("st_pos", "st_pos", "st_pos", val)
    End Sub
    Private Sub getAllSetting()
        start_pos = GetSetting("st_pos", "st_pos", "st_pos")
    End Sub
    Private Sub frmReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Enabled = True
        dgvSetting(DGVinfo)
        dgvSetting(dgvSelectedProcess)
        LoadProcessFromExcel()
    End Sub

    Private Sub DGVinfo_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVinfo.CellContentClick
        Dim row As Integer
        row = DGVinfo.CurrentRow.Index
        DGVinfo(0, row).Value = Not DGVinfo(0, row).Value
    End Sub
    Private Sub LoadProcess()
        On Error Resume Next
        Dim da As New OleDbDataAdapter
        Dim dt As New DataTable
        da.SelectCommand = New OleDbCommand("SELECT DISTINCT pName from ptable", Con)
        da.Fill(dt)

        Dim Rs As DataRow
        If dt.Rows.Count = 0 Then
            Return
        End If

        cboProcessName.Items.Clear()
        For Each Rs In dt.Rows
            cboProcessName.Items.Add(Rs("pName"))
        Next

    End Sub
    Private Sub LoadSubProcess()
        On Error Resume Next
        Dim da As New OleDbDataAdapter
        Dim dt As New DataTable
        da.SelectCommand = New OleDbCommand("SELECT * from ptable where pName='" + cboProcessName.Text + "'", Con)
        da.Fill(dt)

        Dim Rs As DataRow
        If dt.Rows.Count = 0 Then
            Return
        End If

        cboSubProcessName.Items.Clear()
        For Each Rs In dt.Rows
            cboSubProcessName.Items.Add(Rs("psubName"))
        Next

    End Sub

    Private Sub cboProcessName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboProcessName.SelectedIndexChanged
        If Not cboProcessName.Text = "" Then
            Dim selectStatement As String = "SELECT * FROM [Processes$] WHERE PROCESS_NAME='" + cboProcessName.Text + "'"
            'Create a DataAdapter that will be used to populate a DataTable with data
            Dim adapter As New OleDbDataAdapter(selectStatement, connectionString)

            'Populate a DataTable
            Dim excelData As New DataTable
            adapter.Fill(excelData)
            Dim Rs As DataRow
            If excelData.Rows.Count = 0 Then
                Return
            End If
            cboSubProcessName.Items.Clear()
            cboSubProcessName.Text = ""
            For Each Rs In excelData.Rows
                cboSubProcessName.Items.Add(Rs("adverb") + Rs("SUB_PROCESS_NAME"))
            Next
        End If
    End Sub

    Private Sub cmdAddProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddProcess.Click
        Dim selectedCount As Integer
        Dim alreadySelectedProcess As Boolean = False
        Dim selectedParts As Boolean = False
        Dim currentRow As Integer = DGVinfo.CurrentRow.Index

        If cboProcessName.Text = "" Or cboSubProcessName.Text = "" Then
            MsgBox("Select Process/Sub-Process")
            Exit Sub
        End If
        '' check quantity not to be zero
        For i = 0 To DGVinfo.RowCount - 1
            If DGVinfo(0, i).Value = True And Convert.ToInt32(DGVinfo(3, i).Value) < 1 Then
                selectedParts = False
                Exit For
            Else
                selectedParts = True
            End If
        Next
        '' check at least one part is selected
        If selectedParts Then
            For i = 0 To DGVinfo.RowCount - 1
                If DGVinfo(0, i).Value Then
                    selectedParts = True
                    Exit For
                Else
                    selectedParts = False
                End If
            Next
        End If
        If selectedParts = False Then
            MsgBox("May be you are not selected any part or check quantity")
            Exit Sub
        End If

        selectedCount = dgvSelectedProcess.RowCount
        If selectedCount > 0 Then
            For i = 0 To dgvSelectedProcess.RowCount - 1
                If dgvSelectedProcess(2, i).Value = Trim(cboProcessName.Text + cboSubProcessName.Text) Then
                    alreadySelectedProcess = True
                    Exit For
                End If
            Next
        End If
        If alreadySelectedProcess Then
            MsgBox("This Process is already selected")
            Exit Sub
        End If
        '' update qty for selected parts
        For i = 0 To DGVinfo.RowCount - 1
            If DGVinfo(0, i).Value Then
                DGVinfo(3, i).Value = DGVinfo(3, i).Value - 1
            End If
        Next
        If Not cboProcessName.Text = "" And Not cboSubProcessName.Text = "" Then
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Open(fileSource)           ' WORKBOOK TO OPEN THE EXCEL FILE.
            xlWorkSheet = xlWorkBook.Worksheets("MSIL")
            Dim oneTimeCall As Boolean = True
            txtProcessNo.Text = txtProcessNo.Text + 10
            TreeView1.Nodes.Add(New TreeNode("Process No- " + txtProcessNo.Text))
            TreeView1.Nodes(treeNodeStart).Nodes.Add(New TreeNode("Process Name- " + Trim(cboProcessName.Text + cboSubProcessName.Text)))
            TreeView1.Nodes(treeNodeStart).Nodes.Add(New TreeNode("Part Details"))
            For i = 0 To DGVinfo.RowCount - 1
                If DGVinfo(0, i).Value Then
                    Dim partsDetails As String = "No- " + CStr(DGVinfo(1, i).Value) + ", Name- " + CStr(DGVinfo(2, i).Value) + ", Qty- " + CStr(DGVinfo(3, i).Value) + ", XYZ- " + CStr(DGVinfo(4, i).Value) + ", Product Number- " + CStr(DGVinfo(5, i).Value) + ", Product Name- " + CStr(DGVinfo(6, i).Value) _
                    + ", Customer Name- " + CStr(DGVinfo(7, i).Value) + ", Customer No- " + CStr(DGVinfo(8, i).Value) + ", Customer- " + CStr(DGVinfo(9, i).Value) + ", Model- " + CStr(DGVinfo(10, i).Value) + ", Part Assy- " + CStr(DGVinfo(11, i).Value)
                    TreeView1.Nodes(treeNodeStart).Nodes(1).Nodes.Add(New TreeNode(partsDetails))
                End If
            Next

            TreeView1.Nodes(treeNodeStart).Nodes.Add(New TreeNode("Quality Manufacturing Parameter"))
            TreeView1.Nodes(treeNodeStart).Nodes.Add(New TreeNode("Quality Manufacturing Specifications"))
            TreeView1.Nodes(treeNodeStart).Nodes.Add(New TreeNode("Frequency at initial"))
            TreeView1.Nodes(treeNodeStart).Nodes.Add(New TreeNode("Frequency"))
            TreeView1.Nodes(treeNodeStart).Nodes.Add(New TreeNode("Measuring Tool"))
            TreeView1.Nodes(treeNodeStart).Nodes.Add(New TreeNode("Checked By"))
            TreeView1.Nodes(treeNodeStart).Nodes.Add(New TreeNode("Recording Sheet"))
            TreeView1.Nodes(treeNodeStart).Nodes.Add(New TreeNode("Machine Jig Tool"))
            For x As Integer = 1 To xlWorkSheet.UsedRange.Rows.Count
                If xlWorkSheet.Range("A" & x + 1).Value = Trim(cboProcessName.Text + cboSubProcessName.Text) Then
                    If Not xlWorkSheet.Range("C" & x + 1).Value = "" Then
                        TreeView1.Nodes(treeNodeStart).Nodes(2).Nodes.Add(New TreeNode(xlWorkSheet.Range("C" & x + 1).Value))
                    End If
                    If Not xlWorkSheet.Range("D" & x + 1).Value = "" Then
                        TreeView1.Nodes(treeNodeStart).Nodes(3).Nodes.Add(New TreeNode(xlWorkSheet.Range("D" & x + 1).Value))
                    End If
                    If Not xlWorkSheet.Range("E" & x + 1).Value = "" Then
                        TreeView1.Nodes(treeNodeStart).Nodes(4).Nodes.Add(New TreeNode(xlWorkSheet.Range("E" & x + 1).Value))
                    End If
                    If Not xlWorkSheet.Range("F" & x + 1).Value = "" Then
                        TreeView1.Nodes(treeNodeStart).Nodes(5).Nodes.Add(New TreeNode(xlWorkSheet.Range("F" & x + 1).Value))
                    End If
                    If Not xlWorkSheet.Range("G" & x + 1).Value = "" Then
                        TreeView1.Nodes(treeNodeStart).Nodes(6).Nodes.Add(New TreeNode(xlWorkSheet.Range("G" & x + 1).Value))
                    End If
                    If Not xlWorkSheet.Range("H" & x + 1).Value = "" Then
                        TreeView1.Nodes(treeNodeStart).Nodes(7).Nodes.Add(New TreeNode(xlWorkSheet.Range("H" & x + 1).Value))
                    End If
                    If Not xlWorkSheet.Range("I" & x + 1).Value = "" Then
                        TreeView1.Nodes(treeNodeStart).Nodes(8).Nodes.Add(New TreeNode(xlWorkSheet.Range("I" & x + 1).Value))
                    End If
                    If xlWorkSheet.Range("B" & x + 1).Interior.Color = RGB(255, 255, 0) Then
                        Dim str As String = InputBox("Machine Jig Tool", "Machine Jig Tool")
                        If Not str = "" Then
                            TreeView1.Nodes(treeNodeStart).Nodes(9).Nodes.Add(New TreeNode(str))
                        Else
                            MsgBox("Please Enter Valid input")
                        End If
                    ElseIf Not xlWorkSheet.Range("B" & x + 1).Interior.Color = RGB(255, 255, 0) And xlWorkSheet.Range("B" & x + 1).Value = "" Then
                        TreeView1.Nodes(treeNodeStart).Nodes(9).Nodes.Add(New TreeNode(xlWorkSheet.Range("B" & x + 1).Value))
                    End If
                End If
            Next
            treeNodeStart = treeNodeStart + 1
            xlWorkBook.Close() : xlApp.Quit()

            ' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook) : xlWorkBook = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet) : xlWorkSheet = Nothing
        Else
            MsgBox("Please Select appropriate Process")
        End If
    End Sub
    Private Sub FillProcessDetails(ByVal dgv As DataGridView, ByVal ws As Excel.Worksheet, ByVal sheetRange As String, ByVal index As Integer, ByVal cellName As String)
        If Not ws.Range(sheetRange & index).Value = "" Then
            dgv.Rows(dgv.RowCount - 1).Cells(cellName).Value = ws.Range(sheetRange & index).Value
        End If
    End Sub
    Private Sub cmdAddBOM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddBOM.Click
        OpenFileDialog1.Title = "Please Select excel file"
        OpenFileDialog1.Filter = "Excel Worksheets|*.xlsx"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim fileName As String = OpenFileDialog1.FileName
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Open(fileName)           ' WORKBOOK TO OPEN THE EXCEL FILE.
            xlWorkSheet = xlWorkBook.Worksheets("BOM")
            Dim i As Integer = 12
            Do While Not xlWorkSheet.Range("N" & i).Value = ""
                DGVinfo.Rows.Add()
                DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column2").Value = xlWorkSheet.Range("X" & i).Value
                DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column3").Value = xlWorkSheet.Range("N" & i).Value
                DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column4").Value = xlWorkSheet.Range("AW" & i).Value
                DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column5").Value = xlWorkSheet.Range("BC" & i).Value
                DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column6").Value = xlWorkSheet.Range("AD7").Value
                DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column7").Value = xlWorkSheet.Range("AD8").Value
                DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column8").Value = xlWorkSheet.Range("BQ6").Value
                DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column9").Value = xlWorkSheet.Range("BQ7").Value
                DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column10").Value = xlWorkSheet.Range("BK" & i).Value
                DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column11").Value = xlWorkSheet.Range("O7").Value
                DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column12").Value = xlWorkSheet.Range("O8").Value
                i = i + 1
            Loop
            xlWorkBook.Close() : xlApp.Quit()

            ' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook) : xlWorkBook = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet) : xlWorkSheet = Nothing
        End If
    End Sub
    Private Sub RecurseNodes(ByVal col As TreeNodeCollection)
        For Each tn As TreeNode In col
            'MsgBox(tn.Text)
            If tn.Text.Contains("Process No-") Then
                processNoNode = tn.Text
            End If
            If tn.Text.Contains("Process Name- ") Then
                processNameNode = tn.Text
            End If
            If (tn.Text = "Part Details") Then
                partsNode.Clear()
                For Each childTn As TreeNode In tn.Nodes
                    partsNode.Add(childTn.Text)
                Next
            End If
            If tn.Text = "Quality Manufacturing Parameter" Then
                qmpNode.Clear()
                For Each childTn As TreeNode In tn.Nodes
                    qmpNode.Add(childTn.Text)
                Next
            End If
            If tn.Text = "Quality Manufacturing Specifications" Then
                qmsNode.Clear()
                For Each childTn As TreeNode In tn.Nodes
                    qmsNode.Add(childTn.Text)
                Next
            End If
            If tn.Text = "Frequency at initial" Then
                freqAtInitNode.Clear()
                For Each childTn As TreeNode In tn.Nodes
                    freqAtInitNode.Add(childTn.Text)
                Next
            End If
            If tn.Text = "Frequency" Then
                freqNode.Clear()
                For Each childTn As TreeNode In tn.Nodes
                    freqNode.Add(childTn.Text)
                Next
            End If
            If tn.Text = "Measuring Tool" Then
                measuringToolNode.Clear()
                For Each childTn As TreeNode In tn.Nodes
                    measuringToolNode.Add(childTn.Text)
                Next
            End If
            If tn.Text = "Checked By" Then
                checkedByNode.Clear()
                For Each childTn As TreeNode In tn.Nodes
                    checkedByNode.Add(childTn.Text)
                Next
            End If
            If tn.Text = "Recording Sheet" Then
                recordingSheetNode.Clear()
                For Each childTn As TreeNode In tn.Nodes
                    recordingSheetNode.Add(childTn.Text)
                Next
            End If
            If tn.Text = "Machine Jig Tool" Then
                machineJigToolNode.Clear()
                For Each childTn As TreeNode In tn.Nodes
                    If childTn.Text <> "" Then
                        machineJigToolNode.Add(childTn.Text)
                    End If
                Next
                saveDetails(processNoNode, processNameNode, partsNode, qmpNode, qmsNode, freqAtInitNode, freqNode, _
                            measuringToolNode, checkedByNode, recordingSheetNode, machineJigToolNode)
            End If
            If tn.Nodes.Count > 0 Then
                RecurseNodes(tn.Nodes)
            End If
        Next tn
    End Sub
    Private Sub saveDetails(ByVal pno As String, ByVal pName As String, ByVal partsNode As ArrayList, _
                            ByVal qmpNode As ArrayList, ByVal qmsNode As ArrayList, _
                            ByVal freqAtInitNode As ArrayList, ByVal freqNode As ArrayList, _
                            ByVal measuringToolNode As ArrayList, ByVal checkedByNode As ArrayList, _
                            ByVal recordingSheetNode As ArrayList, ByVal machineJigToolNode As ArrayList)
        Dim lPno = pno.Replace("Process No-", "")
        Dim lPname = pName.Replace("Process Name- ", "")
        For Each item As String In partsNode
            Dim parts As String() = item.Split(New Char() {","c})
            Dim partNo As String = parts(0).Replace("No- ", "")
            Dim partName As String = parts(1).Replace(" Name- ", "")
            Dim qty As String = parts(2).Replace(" Qty- ", "")
            Dim xyz As String = parts(3).Replace(" XYZ- ", "")
            Dim product_no As String = parts(4).Replace(" Product Number- ", "")
            Dim product_name As String = parts(5).Replace(" Product Name- ", "")
            Dim customer_name As String = parts(6).Replace(" Customer Name- ", "")
            Dim customer_no As String = parts(7).Replace(" Customer No- ", "")
            Dim customer As String = parts(8).Replace(" Customer- ", "")
            Dim model As String = parts(9).Replace(" Model- ", "")
            Dim partAssy As String = parts(10).Replace(" Part Assy- ", "")

            gfnDBInsertRecord("partDetails", "pfcName, model, pName, pNo, qty, xyz, productNo, productName, customerNo, customerName, customer, partAssy, partName, partNo", _
                              "'" + Me.Text + "', '" + model + "', '" + lPname + "', '" + lPno + "', '" + qty + "', '" + xyz + "', '" + product_no + "', '" + product_name + "', '" + customer_no + "', '" + customer_name + "', '" + customer + "', '" + partAssy + "', '" + partName + "', '" + partNo + "'")
        Next
        For Each freq As String In freqNode
            gfnDBInsertRecord("frequency", "pfcName, freq, pName, pNo", _
                              "'" + Me.Text + "', '" + freq + "', '" + lPname + "', '" + lPno + "'")
        Next
        For Each freq As String In freqAtInitNode
            gfnDBInsertRecord("frequencyInitial", "pfcName, freqAtInit, pName, pNo", _
                              "'" + Me.Text + "', '" + freq + "', '" + lPname + "', '" + lPno + "'")
        Next
        For Each qmp As String In qmpNode
            gfnDBInsertRecord("QMP", "pfcName, qmp, pName, pNo", _
                              "'" + Me.Text + "', '" + qmp + "', '" + lPname + "', '" + lPno + "'")
        Next
        For Each qms As String In qmsNode
            gfnDBInsertRecord("QMS", "pfcName, qms, pName, pNo", _
                              "'" + Me.Text + "', '" + qms + "', '" + lPname + "', '" + lPno + "'")
        Next
        For Each item As String In measuringToolNode
            gfnDBInsertRecord("measuringTool", "pfcName, measTool, pName, pNo", _
                              "'" + Me.Text + "', '" + item + "', '" + lPname + "', '" + lPno + "'")
        Next
        For Each item As String In checkedByNode
            gfnDBInsertRecord("checkedBy", "pfcName, checkedBy, pName, pNo", _
                              "'" + Me.Text + "', '" + item + "', '" + lPname + "', '" + lPno + "'")
        Next
        For Each item As String In recordingSheetNode
            gfnDBInsertRecord("recordingSheet", "pfcName, recordSheet, pName, pNo", _
                              "'" + Me.Text + "', '" + item + "', '" + lPname + "', '" + lPno + "'")
        Next
        For Each item As String In machineJigToolNode
            gfnDBInsertRecord("machineJigTool", "pfcName, machineJigTool, pName, pNo", _
                              "'" + Me.Text + "', '" + item + "', '" + lPname + "', '" + lPno + "'")
        Next
    End Sub
    Private Sub cmdPFC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPFC.Click
        partsNode.Clear()
        qmsNode.Clear()
        qmpNode.Clear()
        freqAtInitNode.Clear()
        freqNode.Clear()
        measuringToolNode.Clear()
        checkedByNode.Clear()
        recordingSheetNode.Clear()
        machineJigToolNode.Clear()
        RecurseNodes(TreeView1.Nodes)
        MsgBox("PFC created successfully")
    End Sub
    Private Sub dgvSelectedProcess_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSelectedProcess.DoubleClick
        Try
            Dim row As Integer
            row = dgvSelectedProcess.CurrentRow.Index
            searchByProcessName = dgvSelectedProcess(2, row).Value
            If searchByProcessName <> "" Then
                frmProcessDetails.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox("Please add Process first")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Interval = 500
        Dim row As Integer
        Dim itemSelected As Boolean = False
        row = DGVinfo.RowCount
        If row > 0 Then
            msgNumber = 2
            gbProcess.Enabled = True
            cmdAddProcess.Enabled = True
            Dim rows As Integer
            rows = DGVinfo.CurrentRow.Index
            For i = 0 To DGVinfo.RowCount - 1
                If DGVinfo(0, i).Value Then
                    itemSelected = False
                    Exit For
                Else
                    itemSelected = True
                End If
            Next
            If itemSelected Then
                Me.Text = ""
            End If
            If DGVinfo(0, rows).Value And Me.Text = "" Then
                Me.Text = "PFC-" + DGVinfo(5, rows).Value
            End If
        Else
            msgNumber = 1
            gbProcess.Enabled = False
            cmdAddProcess.Enabled = False
        End If
        showMessage = Not showMessage
        If showMessage Then
            Select Case msgNumber
                Case 1
                    lblMessage.Text = "Please add BOM first"
                    lblMessage.ForeColor = Color.Red
                Case 2
                    lblMessage.Text = "Now you can add process or can add more BOM"
                    lblMessage.ForeColor = Color.Red
                Case 3
                    lblMessage.Text = "Process is already selected choose different one"
                    lblMessage.ForeColor = Color.Red
                Case Else
                    lblMessage.Text = ""
            End Select

        Else
            lblMessage.Text = ""
        End If
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        main.Show()
        Me.Close()
    End Sub
End Class