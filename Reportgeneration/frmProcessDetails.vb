Public Class frmProcessDetails

    Private Sub frmProcessDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadQMS()
        loadQMP()
        loadFrequency()
        loadMeasuringTool()
        loadCheckedBy()
        loadRecordSheet()
    End Sub

    Private Sub loadQMS()
        Dim ldtQMS As New DataTable
        Dim dr As DataRow
        ldtQMS = gfnSelectQueryDtExcel("SELECT [QMS] FROM [MSIL$] WHERE Process_Name='" + searchByProcessName + "'")
        lbQMS.Items.Clear()
        For Each dr In ldtQMS.Rows
            lbQMS.Items.Add(dr("QMS"))
        Next
    End Sub
    Private Sub loadQMP()
        Dim ldtQMP As New DataTable
        Dim dr As DataRow
        ldtQMP = gfnSelectQueryDtExcel("SELECT [QMP] FROM [MSIL$] WHERE Process_Name='" + searchByProcessName + "'")
        lbQMP.Items.Clear()
        For Each dr In ldtQMP.Rows
            lbQMP.Items.Add(dr("QMP"))
        Next
    End Sub
    Private Sub loadFrequency()
        Dim ldtFreq As New DataTable
        Dim dr As DataRow
        ldtFreq = gfnSelectQueryDtExcel("SELECT [Frequency] FROM [MSIL$] WHERE Process_Name='" + searchByProcessName + "'")
        lbFrequency.Items.Clear()
        For Each dr In ldtFreq.Rows
            lbFrequency.Items.Add(dr("Frequency"))
        Next
    End Sub
    Private Sub loadMeasuringTool()
        Dim ldtMeasuring As New DataTable
        Dim dr As DataRow
        ldtMeasuring = gfnSelectQueryDtExcel("SELECT [Meas_Tool] FROM [MSIL$] WHERE Process_Name='" + searchByProcessName + "'")
        lbMeasuringTool.Items.Clear()
        For Each dr In ldtMeasuring.Rows
            lbMeasuringTool.Items.Add(dr("Meas_Tool"))
        Next
    End Sub
    Private Sub loadCheckedBy()
        Dim ldtCheckedBy As New DataTable
        Dim dr As DataRow
        ldtCheckedBy = gfnSelectQueryDtExcel("SELECT [Checked_by] FROM [MSIL$] WHERE Process_Name='" + searchByProcessName + "'")
        lbCheckedBy.Items.Clear()
        For Each dr In ldtCheckedBy.Rows
            lbCheckedBy.Items.Add(dr("Checked_by"))
        Next
    End Sub
    Private Sub loadRecordSheet()
        Dim ldtRecordSheet As New DataTable
        Dim dr As DataRow
        ldtRecordSheet = gfnSelectQueryDtExcel("SELECT [Recording_Sheet] FROM [MSIL$] WHERE Process_Name='" + searchByProcessName + "'")
        lbRecordingSheet.Items.Clear()
        For Each dr In ldtRecordSheet.Rows
            lbRecordingSheet.Items.Add(dr("Recording_Sheet"))
        Next
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Hide()
    End Sub
End Class