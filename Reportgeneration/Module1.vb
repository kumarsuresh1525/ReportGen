Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core

Module Module1
    Public Con As OleDbConnection
    Public conExcel As OleDbConnection
    Public loledbCommand As OleDbCommand
    Public loledbDataAdapter As New OleDbDataAdapter
    Public fileSource = Application.StartupPath + "\reference\database.xlsx"
    Public customerNo, customerName, model, product_Name, productNo As String
    Public revesionNumber As String
    Public searchByProcessName As String
    Dim xlApp As Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet
    Public Function MakeConnection() As Boolean
        Con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\database.mdb")
        Try
            Con.Open()
            Con.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try

    End Function
    Public Function MakeConnectionExcel() As Boolean
        conExcel = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + Application.StartupPath + "\database.xlsx" + "';Extended Properties = ""Excel 12.0 Xml;HDR=YES""")
        Try
            conExcel.Open()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            conExcel.Close()
            Return False
        End Try

    End Function
    Public Function gfnDBInsertRecord(ByVal lTablename As String, ByVal lFiels As String, ByVal lValue As String) As Boolean
        Try
            Dim lStrQuery As String = "Insert Into " & lTablename & " (" & lFiels & ") values (" & lValue & ") "
            If Con.State = ConnectionState.Closed Then Con.Open()
            loledbCommand = New OleDbCommand(lStrQuery, Con)
            loledbCommand.ExecuteNonQuery()
            loledbCommand.Dispose()
            Return True
        Catch ex As Exception
            loledbCommand.Dispose()
            Return False
        End Try
    End Function
    Public Function gfnDBUpdateRecord(ByVal lTablename As String, ByVal lFielsandValues As String, Optional ByVal lClause As String = "") As Boolean
        Try
            Dim lStrQuery As String = ""
            If lClause = "" Then
                lStrQuery = "Update " & lTablename & " set " & lFielsandValues & ""
            Else
                lStrQuery = "Update " & lTablename & " set " & lFielsandValues & " Where " & lClause & ""
            End If
            
            If Con.State = ConnectionState.Closed Then Con.Open()
            loledbCommand = New OleDbCommand(lStrQuery, Con)
            Dim i As Integer = loledbCommand.ExecuteNonQuery()
            loledbCommand.Dispose()
            Return True
        Catch ex As Exception
            loledbCommand.Dispose()
            Return False
        End Try
    End Function
    Public Function gfnDeleteRecord(ByVal lTablename As String, Optional ByVal lClause As String = "") As Boolean
        Try
            Dim lStrQuery As String = ""
            If lClause = "" Then
                lStrQuery = "Delete from " & lTablename & ""
            Else
                lStrQuery = "Delete from " & lTablename & " Where " & lClause & ""
            End If
            
            If Con.State = ConnectionState.Closed Then Con.Open()
            loledbCommand = New OleDbCommand(lStrQuery, Con)
            loledbCommand.ExecuteNonQuery()
            loledbCommand.Dispose()
            Return True
        Catch ex As Exception
            loledbCommand.Dispose()
            Return False
        End Try
    End Function

    'This procedure is use to get records from access database
    Public Function gfnSelectQueryDt(ByVal lstrSql As String) As DataTable
        Try
            
            If Con.State = ConnectionState.Closed Then Con.Open()
            loledbCommand = New OleDbCommand(lstrSql, Con)
            gfnSelectQueryDt = New DataTable
            loledbDataAdapter.SelectCommand = loledbCommand
            loledbDataAdapter.Fill(gfnSelectQueryDt)
            loledbDataAdapter.Dispose()
            Return gfnSelectQueryDt
        Catch ex As Exception
            loledbCommand.Dispose()
            gfnSelectQueryDt = Nothing
            Return gfnSelectQueryDt
        End Try
    End Function
    Public Function gfnSelectQueryDtExcel(ByVal lstrSql As String) As DataTable
        Try

            If conExcel.State = ConnectionState.Closed Then conExcel.Open()
            loledbCommand = New OleDbCommand(lstrSql, conExcel)
            gfnSelectQueryDtExcel = New DataTable
            loledbDataAdapter.SelectCommand = loledbCommand
            loledbDataAdapter.Fill(gfnSelectQueryDtExcel)
            loledbDataAdapter.Dispose()
            Return gfnSelectQueryDtExcel
        Catch ex As Exception
            loledbCommand.Dispose()
            gfnSelectQueryDtExcel = Nothing
            Return gfnSelectQueryDtExcel
        End Try
    End Function
    Public Sub EditExcelFile(ByVal sFile As String)
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open(sFile)           ' WORKBOOK TO OPEN THE EXCEL FILE.
        xlWorkSheet = xlWorkBook.Worksheets("Sheet1")
        xlWorkSheet.Cells(1, 1) = 25

        ' circle
        'Dim sc = xlWorkSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeFlowchartConnector, xlWorkSheet.Cells(2, 2).left, xlWorkSheet.Cells(2, 2).Top, 38, 38)
        'sc.Fill.Visible = MsoTriState.msoTrue
        'sc.TextFrame.Characters.Text = 10
        'sc.Fill.ForeColor.RGB = RGB(128, 0, 0)
        'sc.TextFrame.Characters.Font.Name = "MS PGothic"
        'sc.TextFrame.Characters.Font.Size = 10
        'sc.TextFrame.Characters.Font.FontStyle = "Regular"
        'sc.TextFrame.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        'sc.TextFrame.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        'sc.TextFrame.Orientation = MsoTextOrientation.msoTextOrientationHorizontal

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

        xlWorkSheet.Range("B7:C7").MergeCells = True
        xlWorkSheet.Cells("B7") = "S13080"
        xlWorkSheet.Range("B8:D8").MergeCells = True
        ' save and close excel
        xlWorkBook.Save()
        xlWorkBook.Close() : xlApp.Quit()

        ' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook) : xlWorkBook = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet) : xlWorkSheet = Nothing
        MsgBox("Edited Succesfully")
    End Sub

    Public Sub dgvSetting(ByVal dgv As DataGridView)
        dgv.DefaultCellStyle.ForeColor = Color.Black
        dgv.DefaultCellStyle.SelectionForeColor = Color.Black
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
End Module
