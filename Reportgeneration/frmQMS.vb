Imports System.Data.OleDb
Public Class frmQMS
    Private Sub frmProcess_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadData()
    End Sub
    Private Sub LoadData()
        On Error Resume Next
        Dim i As Integer
        Dim da As New OleDbDataAdapter
        Dim dt As New DataTable
        da.SelectCommand = New OleDbCommand("select * from QMS", Con)
        da.Fill(dt)

        Dim Rs As DataRow
        If dt.Rows.Count = 0 Then
            Return
        End If

        DGVinfo.Rows.Clear()
        i = 1
        For Each Rs In dt.Rows
            DGVinfo.Rows.Add()
            DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column1").Value = i
            DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column2").Value = Rs("pFullName")
            DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column3").Value = Rs("QMS_Data")
            DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column4").Value = Rs("freq_at_init")
            DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column5").Value = Rs("freq")
            DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column6").Value = Rs("measuring_tool")
            DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column7").Value = Rs("checked_by")
            DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column8").Value = Rs("recording_sheet")
            i = i + 1
        Next
        LoadComboBox()

    End Sub
    Private Sub LoadComboBox()
        On Error Resume Next
        Dim da As New OleDbDataAdapter
        Dim dt As New DataTable
        da.SelectCommand = New OleDbCommand("SELECT DISTINCT pFullName from ptable", Con)
        da.Fill(dt)

        Dim Rs As DataRow
        If dt.Rows.Count = 0 Then
            Return
        End If
        cboQMS.Items.Clear()
        For Each Rs In dt.Rows
            cboQMS.Items.Add(Rs("pFullName"))
        Next
    End Sub

    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        resetForm()
        cboQMS.Enabled = True
    End Sub
    Private Sub resetForm()
        txtQMS.Text = ""
        txtPName.Text = ""
        txtCheckedBy.Text = ""
        txtFreq.Text = ""
        txtFreqAtInit.Text = ""
        txtMeasuringTool.Text = ""
        txtRecordSheet.Text = ""
    End Sub
    Private Sub showErrorField(ByVal txt As TextBox)
        If txt.Text = "" Then
            txt.BackColor = Color.Red
        Else
            txt.BackColor = Color.White
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        main.Show()
        Me.Hide()
    End Sub

    Private Sub DGVinfo_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVinfo.DoubleClick
        Dim row As Integer
        row = DGVinfo.CurrentRow.Index
        txtPName.Text = DGVinfo(1, row).Value
        txtQMS.Text = DGVinfo(2, row).Value
        If IsDBNull(DGVinfo(3, row).Value) Then
            txtFreqAtInit.Text = ""
        Else
            txtFreqAtInit.Text = DGVinfo(3, row).Value
        End If

        txtFreq.Text = DGVinfo(4, row).Value
        txtMeasuringTool.Text = DGVinfo(5, row).Value
        txtCheckedBy.Text = DGVinfo(6, row).Value
        txtRecordSheet.Text = DGVinfo(7, row).Value
        cboQMS.Enabled = False
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If txtQMS.Text = "" Then
            showErrorField(txtQMS)
            MsgBox("Field Can't be blank")
            Exit Sub
        End If
        Dim qry As String = "select * from QMS where pFullName='" & (Trim(txtPName.Text)) & "' AND QMS_Data='" & (Trim(txtQMS.Text)) & "'"
        MakeConnection()
        Con.Open()
        Dim cmd As OleDbCommand = New OleDbCommand(qry, Con)

        Dim dr As OleDbDataReader = cmd.ExecuteReader()
        Dim r As Integer
        If dr.HasRows() Then

            qry = "update QMS set pFullName=@pFullName,QMS_Data=@QMS_Data,freq_at_init=@freq_at_init,freq=@freq,measuring_tool=@measuring_tool,checked_by=@checked_by,recording_sheet=@recording_sheet where pFullName='" & (Trim(txtPName.Text)) & "' AND QMS_Data='" & (Trim(txtQMS.Text)) & "'"
            Dim cmd1 As OleDbCommand = New OleDbCommand(qry, Con)
            cmd1.Parameters.AddWithValue("@pFullName", txtPName.Text)
            cmd1.Parameters.AddWithValue("@QMS_Data", txtQMS.Text)
            cmd1.Parameters.AddWithValue("@freq_at_init", txtFreqAtInit.Text)
            cmd1.Parameters.AddWithValue("@freq", txtFreq.Text)
            cmd1.Parameters.AddWithValue("@measuring_tool", txtMeasuringTool.Text)
            cmd1.Parameters.AddWithValue("@checked_by", txtCheckedBy.Text)
            cmd1.Parameters.AddWithValue("@recording_sheet", txtRecordSheet.Text)
            r = cmd1.ExecuteNonQuery()
            Con.Close()
            dr.Close()
        Else
            qry = "INSERT INTO QMS ( pFullName, QMS_Data, freq_at_init, freq, measuring_tool, checked_by, recording_sheet)"
            qry = qry & "  VALUES(@pFullName, @QMS_Data, @freq_at_init, @freq, @measuring_tool, @checked_by, @recording_sheet)"
            Dim cmd1 As OleDbCommand = New OleDbCommand(qry, Con)
            cmd1.Parameters.AddWithValue("@pFullName", cboQMS.Text)
            cmd1.Parameters.AddWithValue("@QMS_Data", txtQMS.Text)
            cmd1.Parameters.AddWithValue("@freq_at_init", txtFreqAtInit.Text)
            cmd1.Parameters.AddWithValue("@freq", txtFreq.Text)
            cmd1.Parameters.AddWithValue("@measuring_tool", txtMeasuringTool.Text)
            cmd1.Parameters.AddWithValue("@checked_by", txtCheckedBy.Text)
            cmd1.Parameters.AddWithValue("@recording_sheet", txtRecordSheet.Text)
            r = cmd1.ExecuteNonQuery()
            Con.Close()
            dr.Close()
        End If
        If r > 0 Then
            MsgBox("QMS saved successfully")
            LoadData()
            resetForm()
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click

    End Sub
End Class