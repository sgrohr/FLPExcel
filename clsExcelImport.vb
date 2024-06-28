Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Linq
Imports System.IO
Imports System.Net.WebRequestMethods
Imports System.Security.AccessControl

Public Class clsExcelImport

    Private Shared oXls As Application = Nothing
    Private Shared oWkbExportFile As Workbook = Nothing
    Private Shared Excelzeilen As List(Of List(Of String)) = Nothing

    Public Shared Function openExcelExportFile(filename As String) As List(Of String)
        Dim retval As New List(Of String)
        If oXls Is Nothing Then
            oXls = New Application
            oXls.Visible = False
        End If
        If oWkbExportFile IsNot Nothing Then
            oWkbExportFile.Close()
            oWkbExportFile = Nothing
        End If
        oWkbExportFile = oXls.Workbooks.Open(filename)
        For i As Integer = 1 To oWkbExportFile.Sheets.Count
            retval.Add(oWkbExportFile.Sheets(i).name)
        Next
        Return retval
    End Function

    Public Shared Function ReadExcelfile(Tabellenblatt As String) As List(Of List(Of String))
        Try
            For i As Integer = 1 To oWkbExportFile.Sheets.Count
                If oWkbExportFile.Sheets(i).name = Tabellenblatt Then
                    Dim oSheet1 As Worksheet = oWkbExportFile.Sheets(i)
                    Dim oRg1 As Range = oSheet1.UsedRange
                    Dim array1(,) As Object = oRg1.Value(XlRangeValueDataType.xlRangeValueDefault)
                    Excelzeilen = New List(Of List(Of String))
                    Dim AnzZeilen1 As Integer = array1.GetUpperBound(0)
                    Dim AnzSpalten1 As Integer = array1.GetUpperBound(1)
                    For j As Integer = 1 To AnzZeilen1
                        Dim Spalten As List(Of String) = New List(Of String)
                        Dim strLine As String = ""
                        For k As Integer = 1 To AnzSpalten1
                            If array1(j, k) IsNot Nothing Then
                                Spalten.Add(array1(j, k))
                            Else
                                Spalten.Add("")
                            End If
                        Next
                        Excelzeilen.Add(Spalten)
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
            If oWkbExportFile IsNot Nothing Then oWkbExportFile.Close()
            If oXls IsNot Nothing Then oXls.Quit()
            Return Nothing
        End Try
        If oWkbExportFile IsNot Nothing Then oWkbExportFile.Close()
        If oXls IsNot Nothing Then oXls.Quit()
        Return Excelzeilen
    End Function

    Public Shared Function Excelfile2Datatable(Tabellenblatt As String, Titelzeilevorhanden As Boolean) As System.Data.DataTable
        Dim Excelcontent As List(Of List(Of String)) = ReadExcelfile(Tabellenblatt)
        Dim dt As New System.Data.DataTable
        If Excelcontent IsNot Nothing Then
            If Titelzeilevorhanden Then
                Dim kopfzeile As List(Of String) = Excelcontent(0)
                For Each spalte As String In kopfzeile
                    Tools.modTablesAndRows.CreateColumn(dt, spalte, Tools.modTablesAndRows.eColumnFormat.Zeichen, False, False)
                Next
            Else
                Dim erstezeile As List(Of String) = Excelcontent(0)
                For i As Integer = 0 To erstezeile.Count - 1
                    Tools.modTablesAndRows.CreateColumn(dt, "S" & i + 1, Tools.modTablesAndRows.eColumnFormat.Zeichen, False, False)
                Next
            End If
            For i As Integer = 1 To Excelcontent.Count - 1
                Dim zeile As List(Of String) = Excelcontent(i)
                Tools.modTablesAndRows.writeRow(zeile, dt)
            Next
        End If
        Return dt
    End Function

    Public Shared Sub CloseExcel()
        Try
            If oWkbExportFile IsNot Nothing Then oWkbExportFile.Close()
            If oXls IsNot Nothing Then oXls.Quit()
        Catch ex As Exception

        End Try
    End Sub
End Class

