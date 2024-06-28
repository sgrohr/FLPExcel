Imports System.Windows.Forms

Public Class clsExportTools

    Public Shared Function MakeValidTabname(ByVal s As String) As String
        For Each c As Char In ":\/?*[]()"
            s = s.Replace(c, "_")
        Next
        Return (s)
    End Function

    Public Shared Function MakeValidDateiname(ByVal s As String) As String
        Dim i As Integer = InStr(s, "\")
        While i > 0
            i = InStr(s, "\")
            s = s.Substring(i, s.Length - i)
        End While
        Return (s)
    End Function

    Public Shared Function IfNullObj(ByVal o As Object, Optional ByVal DefaultValue As String = "") As String
        Dim ret As String = ""
        Try
            If o Is DBNull.Value Then
                ret = DefaultValue
            Else
                ret = o.ToString
            End If
            Return ret
        Catch ex As Exception
            Return ret
        End Try
    End Function

    Public Delegate Sub RaeumlSelExportDelegate(ByRef Data As System.Windows.Forms.DataGridView,
                                                ByVal pDateiname As String,
                                                ByVal KoordSystem As String)

    Public Shared Sub RaeumlSelExportAsync(ByRef Data As System.Windows.Forms.DataGridView,
                                           ByVal pDateiname As String,
                                           ByVal KoordSystem As String)

        Tools.Progress.Messenger.Start()
        Tools.Sperren.Gesperrt = True

        Dim Method As RaeumlSelExportDelegate = AddressOf RaeumlSelExport
        Dim Callback As AsyncCallback = AddressOf RaeumlSelExportCallback
        Method.BeginInvoke(Data, pDateiname, KoordSystem, Callback, Method)

    End Sub

    Private Shared Sub RaeumlSelExportCallback(result As IAsyncResult)
        Tools.Progress.Messenger.Stopp()
        Tools.Sperren.Gesperrt = False
    End Sub


    Public Shared Sub RaeumlSelExport(ByRef Data As System.Windows.Forms.DataGridView,
                                 ByVal pDateiname As String,
                                 ByVal KoordSystem As String)
        Dim m_Xl As New FLPExcel.clsExcel()
        Dim Dateiname As String = clsExportTools.MakeValidTabname(pDateiname)
        Dim Dir As IO.DirectoryInfo = New IO.DirectoryInfo(System.Environment.CurrentDirectory)
        Dim bwr As Tools.clsBoolwithReason = m_Xl.NewDocument(Dateiname, "Mastkoordinaten", Dir)
        Dim FullDateiname As String = ""
        If Not bwr.Result Then
            m_Xl = Nothing
            Return
        Else
            FullDateiname = bwr.Reason
            Dim Endung As String = EXCELExport.C_EXCELDATEI_ENDUNG_XLSX
            If Val(m_Xl.Version) < EXCELExport.C_EXCEL_VERSION_MIT_XLSX Then Endung = EXCELExport.C_EXCELDATEI_ENDUNG
            FullDateiname &= Endung
        End If

        Dim dt As New DataTable("Mastkoordinaten")
        Dim row As DataRow
        Dim TotalDatagridviewColumns As Integer = Data.ColumnCount - 1
        For Each c As DataGridViewColumn In Data.Columns
            Dim idColumn As DataColumn = New DataColumn()
            Dim ColName As String = c.Name
            If c.HeaderText IsNot Nothing AndAlso c.HeaderText.Length > 0 Then
                ColName = c.HeaderText
            End If
            idColumn.ColumnName = ColName
            dt.Columns.Add(idColumn)
        Next
        For Each dr As DataGridViewRow In Data.Rows
            row = dt.NewRow 'Create new row
            For cn As Integer = 0 To TotalDatagridviewColumns
                row.Item(cn) = IfNullObj(dr.Cells(cn).Value) ' falls Zelle keinen Wert hat
                If (Data.Columns(cn).DefaultCellStyle.Format <> String.Empty And IsNumeric(row.Item(cn))) Then
                    Dim tmpZahl As Double = row.Item(cn)
                    Dim tmpValue As String = String.Format("{0:" & Data.Columns(cn).DefaultCellStyle.Format & "}", tmpZahl)
                    row.Item(cn) = tmpValue
                End If
            Next
            dt.Rows.Add(row)
        Next

        Dim SortParams As New FLPExcel.clsExcelSortParameter
        If m_Xl.FillExcel(dt, FLPExcel.eExcelFarbe.C_BLASSBLAU, FLPExcel.eExcelFarbe.C_BLAUGRUEN, SortParams) Then
            Dim AnzahlSpalten As Integer = dt.Columns.Count
            Dim LetzteSpalte As String = "E"
            Select Case AnzahlSpalten
                Case 6
                    LetzteSpalte = "F"
                Case 7
                    LetzteSpalte = "G"
                Case 8
                    LetzteSpalte = "H"
                Case 9
                    LetzteSpalte = "I"
            End Select

            m_Xl.MakeBorderRange("A1", LetzteSpalte & 1)
            m_Xl.InsertRow(1)
            m_Xl.MergeSpalten("A", LetzteSpalte)

            m_Xl.VerdoppleZeilenHoehe(1)
            m_Xl.FillExcelBold("Stand: " & Format$(Now, "dd.MM.yyyy") & vbCrLf & "Alle Koordinaten befinden sich im: " & KoordSystem, FLPExcel.eExcelFarbe.C_WEISS)

            m_Xl.AllFieldsCentered()
            m_Xl.AutoSizeColumns()
            Dim Anzahlzeilen As Integer = dt.Rows.Count + 2
            m_Xl.Druckbereich(Anzahlzeilen, "$A$1:$" & LetzteSpalte)
            m_Xl.CloseDocument()
            m_Xl = Nothing
            If Len(FullDateiname) > 0 Then
                Process.Start(FullDateiname)
            End If
        End If
    End Sub

    Public Shared Sub FastExport(ByRef Data As System.Windows.Forms.DataGridView, _
                                 ByVal Dir As String, _
                                 ByVal pDateiname As String)
        Dim Dateiname As String = clsExportTools.MakeValidTabname(pDateiname)
        If Not Dateiname.ToLower.EndsWith(FLPExcel.EXCELExport.C_EXPORTDATEI_ENDUNG) Then Dateiname &= FLPExcel.EXCELExport.C_EXPORTDATEI_ENDUNG
        Dim Arr As String() = New String() {Dir, Dateiname}
        Dateiname = IO.Path.Combine(Arr)
        'Dateiname = Dir & Dateiname
        Dim ts As System.IO.StreamWriter = Nothing
        Try
            ts = New System.IO.StreamWriter(Dateiname, False, System.Text.Encoding.Default)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation, "Nicht exportiert!")
            Return
        End Try
        Dim s As String = ""
        Dim f As frmProgress = New frmProgress
        f.Show()
        Dim progress As System.Windows.Forms.ProgressBar = f.progress
        progress.Minimum = 0
        progress.Maximum = Data.Rows.Count
        progress.Step = 1
        progress.Value = 0
        For idx As Integer = 0 To Data.Columns.Count - 1
            Dim c As System.Windows.Forms.DataGridViewColumn = Data.Columns(idx)
            If c.Visible Then
                If s.Length = 0 Then
                    s = """" & c.HeaderText & """"
                Else
                    s = s & ";""" & c.HeaderText & """"
                End If
            End If
        Next
        ts.WriteLine(s)
        s = ""
        progress.PerformStep()
        For rowCnt As Integer = 0 To Data.Rows.Count - 1
            For colCnt As Integer = 0 To Data.Columns.Count - 1
                Dim c As System.Windows.Forms.DataGridViewColumn = Data.Columns(colCnt)
                If c.Visible Then
                    Dim Value As String = Data(colCnt, rowCnt).Value.ToString
                    If IsNumeric(Value) Then
                        ' damit ALB-Schluessel und EModul nicht in wiss. Format gewandelt werden
                        Dim numbervalue As Double = Value
                        If ((numbervalue > 0) And (numbervalue < 0.0001)) Or (numbervalue > 10000000.0) Then
                            Value = "[" & Value & "]"
                        End If
                    End If
                    'If c.Name = "ALB-Schlüssel" Then
                    '    Value = "'" & Value
                    'End If
                    If s.Length = 0 Then
                        s = """" & Value & """"
                    Else
                        s = s & ";""" & Value & """"
                    End If
                End If
            Next
            progress.PerformStep()
            ts.WriteLine(s)
            s = ""
        Next
        ts.Close()
        f.Close()
        Dim myProcess As Process = New Process()
        myProcess.StartInfo.FileName = Dateiname
        myProcess.StartInfo.CreateNoWindow = False
        Try
            myProcess.Start()
        Catch
            Tools.Message.SetHinweis("Das Programm mit dem Sie exportieren (EXCEL etc.) wollen lässt sich nicht starten." & _
                                     vbNewLine & "Ändern Sie ggf. die Zuordnung des Programms zum Öffnen der Datei.")
        End Try
    End Sub

    Public Shared Sub FastExport(ByRef Data As DataTable,
                                 ByVal Dir As String,
                                 ByVal pDateiname As String)
        Dim Dateiname As String = clsExportTools.MakeValidTabname(pDateiname)
        If Not Dateiname.ToLower.EndsWith(FLPExcel.EXCELExport.C_EXPORTDATEI_ENDUNG) Then Dateiname &= FLPExcel.EXCELExport.C_EXPORTDATEI_ENDUNG
        Dim Arr As String() = New String() {Dir, Dateiname}
        Dateiname = IO.Path.Combine(Arr)
        'Dateiname = Dir & Dateiname
        Dim ts As System.IO.StreamWriter = Nothing
        Try
            ts = New System.IO.StreamWriter(Dateiname, False, System.Text.Encoding.Default)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation, "Nicht exportiert!")
            Return
        End Try
        Dim s As String = ""
        Dim f As frmProgress = New frmProgress
        f.Show()
        Dim progress As System.Windows.Forms.ProgressBar = f.progress
        progress.Minimum = 0
        progress.Maximum = Data.Rows.Count
        progress.Step = 1
        progress.Value = 0
        For idx As Integer = 0 To Data.Columns.Count - 1
            Dim c As DataColumn = Data.Columns(idx)
            If s.Length = 0 Then
                s = """" & c.ColumnName & """"
            Else
                s = s & ";""" & c.ColumnName & """"
            End If
        Next
        ts.WriteLine(s)
        s = ""
        progress.PerformStep()
        For Each r As DataRow In Data.Rows
            For colCnt As Integer = 0 To Data.Columns.Count - 1
                Dim c As DataColumn = Data.Columns(colCnt)
                Dim Value As String = r(colCnt).ToString
                If IsNumeric(Value) Then
                    ' damit ALB-Schluessel und EModul nicht in wiss. Format gewandelt werden
                    Dim numbervalue As Double = Value
                    If ((numbervalue > 0) And (numbervalue < 0.0001)) Or (numbervalue > 10000000.0) Then
                        Value = "[" & Value & "]"
                    End If
                End If
                'If c.Name = "ALB-Schlüssel" Then
                '    Value = "'" & Value
                'End If
                If s.Length = 0 Then
                    s = """" & Value & """"
                Else
                    s = s & ";""" & Value & """"
                End If
            Next
            progress.PerformStep()
            ts.WriteLine(s)
            s = ""
        Next
        ts.Close()
        f.Close()
        Dim myProcess As Process = New Process()
        myProcess.StartInfo.FileName = Dateiname
        myProcess.StartInfo.CreateNoWindow = False
        Try
            myProcess.Start()
        Catch
            Tools.Message.SetHinweis("Das Programm mit dem Sie exportieren (EXCEL etc.) wollen lässt sich nicht starten." &
                                     vbNewLine & "Ändern Sie ggf. die Zuordnung des Programms zum Öffnen der Datei.")
        End Try
    End Sub

    Public Shared Sub GewaehrleistungsObjekteExport(ByRef Data As List(Of DataRow), _
                                                    ByVal Dir As String, _
                                                    ByVal pDateiname As String)
        Dim Dateiname As String = clsExportTools.MakeValidTabname(pDateiname)
        If Not Dateiname.ToLower.EndsWith(".txt") Then Dateiname &= ".txt"
        Dim Arr As String() = New String() {Dir, Dateiname}
        Dateiname = IO.Path.Combine(Arr)
        Dim ts As System.IO.StreamWriter = Nothing
        Try
            ts = New System.IO.StreamWriter(Dateiname, False, System.Text.Encoding.Default)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation, "Nicht exportiert!")
            Return
        End Try
        Dim f As frmProgress = New frmProgress
        f.Show()
        Dim progress As System.Windows.Forms.ProgressBar = f.progress
        progress.Minimum = 0
        progress.Maximum = Data.Count
        progress.Step = 1
        progress.Value = 0
        Dim s As String = ""

        'Gewährleistungs-ID Index=1
        'Technischer Platz IPS	Index=2	
        'Gewährleistungsbeginn Index=3
        'Gewährleistungsende Index=4
        'Gewährleistungstext Index=5
        'Garantieart Index=6
        'Vertrag Index=7
        'Vertragsposition Index=8
        'Projektnummer Index=9
        'Meldungsrelevant Index=10
        'Bearbeitungskennzeichen Index=11
        'hoffentlich bleibt das so!!!

        s = Tools.FillTextUpToLength("GWID", 10) & _
            Tools.FillTextUpToLength("TPLNR", 30) & _
            Tools.FillTextUpToLength("GWBEG", 8) & _
            Tools.FillTextUpToLength("GWENDE", 8) & _
            Tools.FillTextUpToLength("TEXT", 40) & _
            Tools.FillTextUpToLength("ART", 4) & _
            Tools.FillTextUpToLength("VERTRAG", 12) & _
            Tools.FillTextUpToLength("POS", 4) & _
            Tools.FillTextUpToLength("PROJNR", 24) & _
            Tools.FillTextUpToLength("M", 1) & _
            Tools.FillTextUpToLength("K", 1)
        progress.PerformStep()
        ts.WriteLine(s)
        For Each dr As DataRow In Data
            Dim Meldung As String = ""
            If dr("MELDREL").ToString = 1 Then
                Meldung = "X"
            End If
            s = Tools.FillTextUpToLength(dr("GWID").ToString, 10) & _
                Tools.FillTextUpToLength(dr("TPLNR").ToString, 30) & _
                Tools.FillTextUpToLength(Tools.MakeDateWithoutPoints(dr("GWBEG").ToString), 8) & _
                Tools.FillTextUpToLength(Tools.MakeDateWithoutPoints(dr("GWENDE").ToString), 8) & _
                Tools.FillTextUpToLength(dr("TEXT").ToString, 40) & _
                Tools.FillTextUpToLength(dr("GAART").ToString, 4) & _
                Tools.FillTextUpToLength(dr("VERTRAG").ToString, 12) & _
                Tools.FillTextUpToLength(dr("VERTPOS").ToString, 4) & _
                Tools.FillTextUpToLength(dr("PROJNR").ToString, 24) & _
                Meldung & _
                Tools.FillTextUpToLength(dr("KENNZ").ToString.ToUpper, 1)
            progress.PerformStep()
            ts.WriteLine(s)
            s = ""
        Next
        ts.Close()
        f.Close()
        Dim myProcess As Process = New Process()
        myProcess.StartInfo.FileName = Dateiname
        myProcess.StartInfo.CreateNoWindow = False
        myProcess.Start()

    End Sub

End Class
