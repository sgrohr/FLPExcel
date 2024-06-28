Public Class clsExcelSortParameter

    Public Structure tColIdxAscending
        Dim ColumnIndex As Integer
        Dim SortOrder As Microsoft.Office.Interop.Excel.XlSortOrder
        Public Sub New(ByVal Idx As Integer, Optional ByVal Ascending As Boolean = True)
            ColumnIndex = Idx
            SortOrder = IIf(Ascending, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending, Microsoft.Office.Interop.Excel.XlSortOrder.xlDescending)
        End Sub
    End Structure

    Private m_Data As New List(Of tColIdxAscending)

    Private m_Key1 As Integer
    Private m_Ascending1 As Boolean

    Public Function AddSortierung(ByVal SpaltenIndex As Integer, ByVal Aufsteigend As Boolean) As Boolean
        If m_Data.Count >= 3 Then Return (False)
        For Each item As tColIdxAscending In m_Data
            If item.ColumnIndex = SpaltenIndex Then Return (False)
        Next
        m_Data.Add(New tColIdxAscending(SpaltenIndex, Aufsteigend))
        Return (True)
    End Function

    Public Function GetAnzahl() As Integer
        Return (m_Data.Count)
    End Function

    Public Function GetParameter(ByVal Index As Integer) As tColIdxAscending
        Debug.Assert(Index >= 0)
        Debug.Assert(Index < m_Data.Count)
        Return (m_Data(Index))
    End Function
End Class
