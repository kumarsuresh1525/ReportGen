Imports System.Data.OleDb
Public Class frmQMP
    Private Sub frmProcess_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadData()
    End Sub
    Private Sub LoadData()
        On Error Resume Next
        Dim i As Integer
        Dim da As New OleDbDataAdapter
        Dim dt As New DataTable
        da.SelectCommand = New OleDbCommand("select * from QMP", Con)
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
            DGVinfo.Rows(DGVinfo.RowCount - 1).Cells("Column3").Value = Rs("QMP_Data")
            i = i + 1
        Next
        LoadComboData()
    End Sub
    Private Sub LoadComboData()
        On Error Resume Next
        Dim da As New OleDbDataAdapter
        Dim dt As New DataTable
        da.SelectCommand = New OleDbCommand("SELECT DISTINCT pFullName FROM ptable", Con)
        da.Fill(dt)

        Dim Rs As DataRow
        If dt.Rows.Count = 0 Then
            Return
        End If
        cboQMP.Items.Clear()
        For Each Rs In dt.Rows
            cboQMP.Items.Add(Rs("pFullName"))
        Next
    End Sub
    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        resetForm()
    End Sub
    Private Sub resetForm()
        txtQMP.Text = ""
        txtPName.Text = ""
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
        txtQMP.Text = DGVinfo(2, row).Value
        cboQMP.Items.Clear()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If txtQMP.Text = "" Then
            showErrorField(txtQMP)
            MsgBox("Field Can't be blank")
            Exit Sub
        End If
        Dim qry As String = "select * from QMP where pFullName='" & (Trim(txtPName.Text)) & "'"
        MakeConnection()
        Con.Open()
        Dim cmd As OleDbCommand = New OleDbCommand(qry, Con)

        Dim dr As OleDbDataReader = cmd.ExecuteReader()
        Dim r As Integer
        If dr.HasRows() Then

            qry = "update QMP set pFullName=@pFullName,QMP_Data=@QMP_Data where pFullName='" & (Trim(txtPName.Text)) & "'"
            Dim cmd1 As OleDbCommand = New OleDbCommand(qry, Con)
            cmd1.Parameters.AddWithValue("@pFullName", txtPName.Text)
            cmd1.Parameters.AddWithValue("@QMP_Data", txtQMP.Text)
            r = cmd1.ExecuteNonQuery()
            Con.Close()
            dr.Close()
        Else
            qry = "INSERT INTO QMP ( pFullName, QMP_Data)"
            qry = qry & "  VALUES(@pFullName, @QMP_Data)"
            Dim cmd1 As OleDbCommand = New OleDbCommand(qry, Con)
            cmd1.Parameters.AddWithValue("@pFullName", cboQMP.Text)
            cmd1.Parameters.AddWithValue("@QMP_Data", txtQMP.Text)
            r = cmd1.ExecuteNonQuery()
            Con.Close()
            dr.Close()
        End If
        If r > 0 Then
            MsgBox("QMP saved successfully")
            LoadData()
            resetForm()
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click

    End Sub
End Class