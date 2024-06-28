Public Class clSachdatenExport
    Public Delegate Sub Export_ObjektSachdatenDelegate(ByVal Objekttyp As String, ByVal Sachdaten As DataRow, AttribTable As DataTable, _
                                                       Pfad As String)

    Public Delegate Sub Export_SeillisteDelegate(ByVal Seildaten As List(Of DataRow), Attributsliste As List(Of String), Pfad As String)

    Public Shared Sub ExportSeilListe(ByVal Seildaten As List(Of DataRow), Attributsliste As List(Of String), Pfad As String)
        Dim method As Export_SeillisteDelegate = AddressOf ExportSeillisteAsync
        method.BeginInvoke(Seildaten, Attributsliste, Pfad, Nothing, method)
    End Sub

    Private Shared Sub ExportSeillisteAsync(seildaten As List(Of DataRow), attributsliste As List(Of String), Dateiname As String)
        Dim swDIN As IO.StreamWriter = Nothing
        Dim swEN As IO.StreamWriter = Nothing
        Dim Dateiliste As New List(Of String)
        For Each row As DataRow In seildaten
            If row("EN") = 1 Then
                If swEN Is Nothing Then
                    swEN = New IO.StreamWriter(Dateiname & "_EN_Seile.csv")
                    Dateiliste.Add(Dateiname & "_EN.csv")
                End If

            Else
                If swDIN Is Nothing Then
                    swDIN = New IO.StreamWriter(Dateiname & "_DIN_Seile.csv")
                    Dateiliste.Add(Dateiname & "_DIN.csv")
                End If
            End If
        Next
        Dim zeile As String = "TAG;"
        For Each spalte As String In attributsliste
            zeile &= spalte & ";"
        Next
        zeile &= "FETTMAUS2;FETTMAUS3;FETTMAUS4"
        If swDIN IsNot Nothing Then swDIN.WriteLine(zeile)
        If swEN IsNot Nothing Then swEN.WriteLine(zeile)
        For Each Data As DataRow In seildaten
            zeile = ""
            If Data("EN") = 1 Then
                zeile = "STEN;"
            Else
                zeile = "STYP;"
            End If
            For Each spalte As String In attributsliste
                If spalte.ToUpper = "QLK" Then
                    zeile &= Mid(Replace(Tools.GetWert(Data, spalte, ""), ",", "."), 1, 9) & ";"
                Else
                    zeile &= Replace(Tools.GetWert(Data, spalte, ""), ",", ".") & ";"
                End If
            Next
            zeile &= ";;"
            If Data("EN") = 1 Then
                swEN.WriteLine(zeile)
            Else
                swDIN.WriteLine(zeile)
            End If
        Next
        If swDIN IsNot Nothing Then swDIN.Close()
        If swEN IsNot Nothing Then swEN.Close()
        swDIN = Nothing
        swEN = Nothing
        Dim Messagestring As String = ""
        If Dateiliste.Count = 1 Then
            Tools.Message.SetHinweis("Seildatei: " & Dateiliste(0) & " erfolgreich geschrieben!")
        ElseIf Dateiliste.Count = 2 Then
            Tools.Message.SetHinweis("Seildateien: " & Dateiliste(0) & " und " & Dateiliste(1) & " erfolgreich geschrieben!")
        End If
    End Sub

    Public Shared Sub ExportObjektSachdatenStart(Objekttyp As String, Sachdaten As DataRow, AttribTable As DataTable, Pfad As String)
        Dim method As Export_ObjektSachdatenDelegate = AddressOf ExportObjektSachdatenAsync
        method.BeginInvoke(Objekttyp, Sachdaten, AttribTable, Pfad, Nothing, method)
    End Sub



    Private Shared Sub ExportObjektSachdatenAsync(Objekttyp As String, Sachdaten As DataRow, AttribTable As DataTable, Pfad As String)
        Dim Arr As String() = New String() {Pfad, "Sachdaten.txt"}
        Dim sw As IO.StreamWriter = New IO.StreamWriter(IO.Path.Combine(Arr))
        Dim Spaltenname As String = ""
        Dim Wert As String = ""
        Dim MaxDisplayNameLenght As Integer = ErmittleMaxTextlengh(AttribTable)
        Dim Zeile As String = ""
        Dim LeerZeile As String = ""
        sw.WriteLine(LeerZeile)
        Zeile = "Sachdaten von Objekt: " & Objekttyp & ", ID=" & Sachdaten("MSLINK") & " erstellt am: " & Now.ToShortDateString
        sw.WriteLine(Zeile)
        Dim strich As String = ""
        For i As Integer = 0 To 80
            strich &= "_"
        Next
        sw.WriteLine(strich)
        sw.WriteLine(LeerZeile)
        For Each r As DataRow In AttribTable.Rows
            Wert = ""
            Spaltenname = ""
            If Not IsDBNull(r("DISPLAY")) Then
                Dim teile() As String = Split(r("DISPLAY"), ";")
                If UBound(teile) = 2 Then
                    For i As Integer = 0 To UBound(teile) - 1
                        Spaltenname = r("COLUMNNAME") & "_" & Split(teile(i), ",")(0)
                        Wert = Wert & " " & GetWert(Sachdaten, Spaltenname)
                    Next
                    Wert = Mid(Wert, 2)
                Else
                    Spaltenname = r("COLUMNNAME") & "_" & Split(r("DISPLAY"), ",")(0)
                    Wert = GetWert(Sachdaten, Spaltenname)
                End If
            Else
                Spaltenname = r("COLUMNNAME")
                Wert = GetWert(Sachdaten, Spaltenname)
            End If
            sw.WriteLine(Auffuellen(r("DISPLAYNAME"), True, MaxDisplayNameLenght + 5) & ":" & Auffuellen(Wert, False, 5))
        Next
        sw.WriteLine(LeerZeile)
        sw.WriteLine(strich)
        sw.Close()
        sw = Nothing

        Dim startInfo As New ProcessStartInfo("Notepad.exe")
        startInfo.WindowStyle = ProcessWindowStyle.Normal
        startInfo.Arguments = IO.Path.Combine(Arr)
        Process.Start(startInfo)
    End Sub

    Private Shared Function GetWert(sachdaten As DataRow, Spaltenname As String) As String
        Dim retVal As String = ""
        If IsDBNull(sachdaten(Spaltenname)) Then
            retVal = ""
        Else
            retVal = sachdaten(Spaltenname)
        End If
        Return (retVal)
    End Function

    Private Shared Function ErmittleMaxTextlengh(Attribtable As DataTable) As Integer
        Dim retVal As Integer = 0
        Dim Spaltenname As String = ""
        For Each r As DataRow In Attribtable.Rows
            If Len(r("DISPLAYNAME")) > retVal Then
                retVal = Len(r("DISPLAYNAME"))
            End If
        Next
        Return (retVal)
    End Function

    Private Shared Function Auffuellen(ByVal Value As String, ByVal Hinten As Boolean, ByVal Stellen As Integer) As String
        Dim RetVal As String = Value
        If Hinten Then
            For i As Integer = 0 To Stellen - Len(Value)
                RetVal &= " "
            Next
        Else
            For i As Integer = 0 To Stellen
                RetVal = " " & RetVal
            Next
        End If
        Return (RetVal)
    End Function

End Class
