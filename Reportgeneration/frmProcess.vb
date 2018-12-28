Imports System.Data.OleDb
Public Class frmProcess
    Private Sub frmProcess_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadData()

    End Sub
    Private Sub LoadData()
        On Error Resume Next
        Dim i As Integer
        Dim da As New OleDbDataAdapter
        Dim dt As New DataTable
        da.SelectCommand = New OleDbCommand("select * from ptable", Con)
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
            DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column2").Value = Rs("pName")
            DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column3").Value = Rs("padverb")
            DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column4").Value = Rs("psubName")
            DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column5").Value = Rs("pFlowSymbol")
            i = i + 1
        Next

    End Sub

    Private Sub DGVinfo_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVinfo.DoubleClick
        Dim row As Integer
        row = DGVinfo.CurrentRow.Index
        txtPName.Text = DGVinfo(1, row).Value
        txtAdverb.Text = DGVinfo(2, row).Value
        txtSubProcess.Text = DGVinfo(3, row).Value
        txtFlowType.Text = DGVinfo(4, row).Value
    End Sub

    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNew.Click
        resetForm()
    End Sub
    Private Sub resetForm()
        txtPName.Text = ""
        txtAdverb.Text = ""
        txtSubProcess.Text = ""
        txtFlowType.Text = ""
    End Sub
    Private Sub showErrorField(ByVal txt As TextBox)
        If txt.Text = "" Then
            txt.BackColor = Color.Red
        Else
            txt.BackColor = Color.White
        End If
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If txtPName.Text = "" Or txtFlowType.Text = "" Or txtSubProcess.Text = "" Then
            showErrorField(txtPName)
            showErrorField(txtFlowType)
            showErrorField(txtSubProcess)
            MsgBox("Field Can't be blank")
            Exit Sub
        End If
        Dim qry As String = "select * from ptable where pName='" & (Trim(txtPName.Text)) & "' AND psubName='" & (Trim(txtSubProcess.Text)) & "'"
        MakeConnection()
        Con.Open()
        Dim cmd As OleDbCommand = New OleDbCommand(qry, Con)

        Dim dr As OleDbDataReader = cmd.ExecuteReader()
        Dim r As Integer
        Dim pFullName As String
        pFullName = txtPName.Text + " " + txtAdverb.Text + " " + txtSubProcess.Text
        If dr.HasRows() Then

            qry = "update ptable set pName=@pName,padverb=@padverb,psubName=@psubName,pFlowSymbol=@pFlowSymbol where pName='" & (Trim(txtPName.Text)) & "' AND psubName='" & (Trim(txtSubProcess.Text)) & "'"
            Dim cmd1 As OleDbCommand = New OleDbCommand(qry, Con)
            cmd1.Parameters.AddWithValue("@pName", txtPName.Text)
            cmd1.Parameters.AddWithValue("@padverb", txtAdverb.Text)
            cmd1.Parameters.AddWithValue("@psubName", txtSubProcess.Text)
            cmd1.Parameters.AddWithValue("@pFlowSymbol", txtFlowType.Text)
            cmd1.Parameters.AddWithValue("@pFullName", pFullName)
            r = cmd1.ExecuteNonQuery()
            Con.Close()
            dr.Close()
        Else
            qry = "INSERT INTO ptable ( pName, padverb, psubName, pFlowSymbol, pFullName)"
            qry = qry & "  VALUES(@pName, @padverb, @psubName, @pFlowSymbol, @pFullName)"
            Dim cmd1 As OleDbCommand = New OleDbCommand(qry, Con)
            cmd1.Parameters.AddWithValue("@pName", txtPName.Text)
            cmd1.Parameters.AddWithValue("@padverb", txtAdverb.Text)
            cmd1.Parameters.AddWithValue("@psubName", txtSubProcess.Text)
            cmd1.Parameters.AddWithValue("@pFlowSymbol", txtFlowType.Text)
            cmd1.Parameters.AddWithValue("@pFullName", pFullName)
            r = cmd1.ExecuteNonQuery()
            Con.Close()
            dr.Close()
        End If
        If r > 0 Then
            MsgBox("Process saved successfully")
            LoadData()
            resetForm()
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        main.Show()
        Me.Hide()
    End Sub

End Class