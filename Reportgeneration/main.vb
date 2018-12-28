Imports System.Data.OleDb

Public Class main

    Private Sub main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cboSelectCustomer.Items.Clear()
        cboReleasingPhase.Visible = False
        Label2.Visible = False
        cboProductRating.Visible = False
        Label3.Visible = False
        cboSelectCustomer.Items.Add("HONDA")
        cboSelectCustomer.Items.Add("MSIL/M&M")
        cboSelectCustomer.Items.Add("TRMN")
        cboSelectCustomer.Items.Add("NISSAN")
        'End If
    End Sub

    Private Sub pSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pSet.Click
        frmProcess.Show()
        Me.Hide()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        frmMakePFC.Show()
        Me.Hide()
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        End
    End Sub

    Private Sub pqcsSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pqcsSet.Click
        frmQMP.Show()
        Me.Hide()
    End Sub

    Private Sub btnQMS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQMS.Click
        frmQMS.Show()
        Me.Hide()
    End Sub

    Private Sub cboSelectCustomer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSelectCustomer.SelectedIndexChanged
        Dim honda_msil = {"Prototype", "Prelanch", "Production"}
        Dim trmn = {"Prototype", "Prelanch", "Mass Production"}
        Dim nissan = {"Prototype", "Pre Production/Prelaunch", "Mass Production"}
        Dim hondaRating = {"HS", "HA", "HB", "Others"}
        Dim msilRating = {"HS", "HA", "HB", "Others"}
        Dim trmnRating = {"HS", "HA", "HB", "Others"}
        Dim nissanRating = {"A", "B", "OBD"}
        Dim selectPhase
        Dim selectRating
        If cboSelectCustomer.Text = "" Then
            cboReleasingPhase.Items.Clear()
            cboReleasingPhase.Items.Add("")
        Else
            cboReleasingPhase.Visible = True
            Label2.Visible = True
            cboReleasingPhase.Text = ""
            cboReleasingPhase.Items.Clear()
            If cboSelectCustomer.Text = "HONDA" Or cboSelectCustomer.Text = "MSIL/M&M" Then
                selectPhase = honda_msil
            ElseIf cboSelectCustomer.Text = "TRMN" Then
                selectPhase = trmn
            Else
                selectPhase = nissan
            End If
            For Each itm In selectPhase
                cboReleasingPhase.Items.Add(itm)
            Next

            cboProductRating.Visible = True
            Label3.Visible = True
            cboProductRating.Text = ""
            cboProductRating.Items.Clear()
            If cboSelectCustomer.Text = "HONDA" Then
                selectRating = hondaRating
            ElseIf cboSelectCustomer.Text = "MSIL/M&M" Then
                selectRating = msilRating
            ElseIf cboSelectCustomer.Text = "TRMN" Then
                selectRating = trmnRating
            Else
                selectRating = nissanRating
            End If
            For Each itm In selectRating
                cboProductRating.Items.Add(itm)
            Next

            If cmdReport.Text = "Generate Report from PFC" And cboSelectCustomer.Text <> "" Then
                Dim ldtPFC As New DataTable
                ldtPFC = gfnSelectQueryDt("SELECT DISTINCT pfcName FROM partDetails WHERE customer='" + cboSelectCustomer.Text + "'")
                cboPFC.Items.Clear()
                If ldtPFC.Rows.Count > 0 Then
                    Dim dr As DataRow
                    For Each dr In ldtPFC.Rows
                        cboPFC.Items.Add(dr("pfcName"))
                    Next
                End If
            End If
        End If
    End Sub

    Private Sub cmdReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReport.Click
        If cboReleasingPhase.Text <> "" And cboProductRating.Text <> "" Then
            If cmdReport.Text = "Generate Report from PFC" Then
                Dim ldtRevisionNumber As New DataTable
                ldtRevisionNumber = gfnSelectQueryDt("SELECT * FROM revisionNumber WHERE pfcName='" + cboPFC.Text + "'")
                Dim revNo As Integer = 0
                If ldtRevisionNumber.Rows.Count > 0 Then
                    Dim dr As DataRow
                    For Each dr In ldtRevisionNumber.Rows
                        revNo = dr("revNo")
                    Next
                    If cboReleasingPhase.Text <> "Prototype" Then
                        revNo = 0
                    Else
                        revNo = revNo + 1
                    End If
                    gfnDBUpdateRecord("revisionNumber", "revNo='" + revNo.ToString + "', productRating='" + cboProductRating.Text + "', releasingPhase='" + cboReleasingPhase.Text + "'", "pfcName='" + cboPFC.Text + "'")
                Else
                    If cboPFC.Text <> "" Then
                        gfnDBInsertRecord("revisionNumber", "pfcName, productRating, releasingPhase", "'" + cboPFC.Text + "', '" + cboProductRating.Text + "', '" + cboReleasingPhase.Text + "'")
                    Else
                        MsgBox("Please Select PFC")
                    End If
                End If

                ' report generation form here
                Dim frm = New frmReport()
                frm.Show()
                Me.Close()
            Else
                Dim frmMakePFC = New frmMakePFC()
                frmMakePFC.Show()
                Me.Close()
            End If
        Else
            MsgBox("Product rating/Releasing Phase can't be blank")
        End If

    End Sub
End Class
