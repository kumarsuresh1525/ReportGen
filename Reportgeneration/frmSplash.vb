Public Class frmSplash

    Private Sub frmSplash_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If MakeConnection() = True Then
            Try
                Timer1.Enabled = True
                Timer1.Interval = 1000
                For i = 0 To 99
                    ProgressBar1.Value = ProgressBar1.Value + 1
                Next i
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        login.Show()
        Me.Hide()
    End Sub
End Class