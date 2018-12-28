Imports System.Data.OleDb
Public Class login
    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        Dim Cmd As New OleDbCommand
        Dim qry As String

        If txtUserName.Text = "" Then
            MsgBox(UCase("Kindly Enter the UserName"), vbCritical)
            txtUserName.Focus()
            txtUserName.SelectionStart = 0
            txtUserName.SelectionLength = Len(txtUserName.Text)
            Exit Sub
        End If

        If txtPassword.Text = "" Then
            MsgBox(UCase("Kindly Enter the PassWord"), vbCritical)
            txtPassword.Focus()
            txtPassword.SelectionStart = 0
            txtPassword.SelectionLength = Len(txtPassword.Text)
            Exit Sub
        End If

        qry = "SELECT * FROM admin where userName='" & (txtUserName.Text) & "' and password='" & (txtPassword.Text) & "'"
        If MakeConnection() = True Then
            Try
                Cmd = New OleDbCommand(qry, Con)
                Dim ldtParts As New DataTable
                ldtParts = gfnSelectQueryDt("SELECT * FROM partDetails")
                If ldtParts.Rows.Count = 0 Then
                    MessageBox.Show("No PFC available Please create at least 1 PFC")
                    main.cmdReport.Text = "Create PFC"
                    main.lblPFC.Visible = False
                    main.cboPFC.Visible = False
                    main.Show()
                    Me.Close()
                Else
                    Dim result As Integer = MessageBox.Show("Press Yes to create new PFC and No for generate report from PFC", "caption", MessageBoxButtons.YesNo)
                    If result = DialogResult.Yes Then
                        main.cmdReport.Text = "Create PFC"
                        main.lblPFC.Visible = False
                        main.cboPFC.Visible = False
                        main.Show()
                        Me.Close()
                    Else
                        main.cmdReport.Text = "Generate Report from PFC"
                        main.lblPFC.Visible = True
                        main.cboPFC.Visible = True
                        main.Show()
                        Me.Close()
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        If Not MakeConnectionExcel() Then
            Application.Exit()
        End If
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        End
    End Sub

    Private Sub login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class