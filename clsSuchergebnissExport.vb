Public Class clsSuchergebnissExport
    Public Delegate Sub Export_SuchergebnisDelegate(ByVal Sachdaten As System.Windows.Forms.DataGridView, Pfad As String, Dateiname As String)
    Public Delegate Sub Export_GewaehrleistungsobjekteDelegate(ByVal Sachdaten As List(Of DataRow), Pfad As String, Dateiname As String)

    Public Shared Sub ExportSuchergebnisStart(ByVal Sachdaten As System.Windows.Forms.DataGridView, Pfad As String, Dateiname As String)
        Dim method As Export_SuchergebnisDelegate = AddressOf ExportSuchergebnisAsync
        method.BeginInvoke(Sachdaten, Pfad, Dateiname, Nothing, method)
    End Sub

    Private Shared Sub ExportSuchergebnisAsync(ByVal Sachdaten As System.Windows.Forms.DataGridView, Pfad As String, Dateiname As String)
        Dateiname = clsExportTools.MakeValidDateiname(Dateiname)
        clsExportTools.FastExport(Sachdaten, Pfad, Dateiname)
    End Sub

    Public Shared Sub ExportGewaehrleistungsobjekteStart(ByVal Sachdaten As List(Of DataRow), Pfad As String, Dateiname As String)
        Dim method As Export_GewaehrleistungsobjekteDelegate = AddressOf ExportGewaehrleistungsObjekteAsync
        method.BeginInvoke(Sachdaten, Pfad, Dateiname, Nothing, method)
    End Sub

    Private Shared Sub ExportGewaehrleistungsObjekteAsync(ByVal Sachdaten As List(Of DataRow), Pfad As String, Dateiname As String)
        Dateiname = clsExportTools.MakeValidDateiname(Dateiname)
        clsExportTools.GewaehrleistungsObjekteExport(Sachdaten, Pfad, Dateiname)
    End Sub

End Class
