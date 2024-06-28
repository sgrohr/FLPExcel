Imports Microsoft.Office.Interop

Public Class EXCELExport
    Implements Georg.IStatusTextSender

    Public Const C_EXCELDATEI_ENDUNG As String = ".xls"
    Public Const C_EXCELDATEI_ENDUNG_XLSX As String = ".xlsx"
    Public Const C_EXPORTDATEI_ENDUNG As String = ".csv"
    Public Const C_EXCEL_VERSION_MIT_XLSX As String = "14"

    Private m_Kundenbezeichnung As String = ""

    Public Delegate Sub Export_ListenAsyncDelegate(ByRef Data As System.Windows.Forms.DataGridView, _
                             ByVal TabName As String, _
                             ByVal MitFarben As Boolean, _
                             ByVal Art As eKopfTyp, _
                             ByVal BLnummer As String, _
                             ByVal sDir As String,
                             ByVal BLleitungsname As String,
                             ByVal BLStandort As String,
                             ByVal LtgPraefix As String)

    Public Delegate Sub Export_SpannweitenListeDelegate(ByRef Data As Windows.Forms.DataGridView, _
                                                        Dateiname As String, Pfad As String)
    Public Property Kundenbezeichnung As String
        Get
            Return m_Kundenbezeichnung
        End Get
        Set(value As String)
            m_Kundenbezeichnung = value
        End Set
    End Property

    Public Shared ReadOnly Property Manager As EXCELExport
        Get
            Return (s_Instance)
        End Get
    End Property

    Private Shared s_Instance As New EXCELExport

    Private m_Statustext As String
    Private m_TemporaererStatustext As String
    Private m_StatusAnzeigezeit As Georg.IStatusTextSender.eStatusAnzeigeZeit = Georg.IStatusTextSender.eStatusAnzeigeZeit.PERMANENT

    Private Sub New()
    End Sub

    Public Enum eKopfTyp
        VERZEICHNIS_AUSFUEHRUNGSPLANUNG
        VERZEICHNIS_GRUENDUNGEN
        VERZEICHNIS_MASTE
        INSTANDHALTUNG_STROMKREISBEZOGEN
        INSTANDHALTUNG_PHASENBEZOGEN
        INSTANDHALTUNG_MAST
        INSTANDHALTUNG_AUFHAENGEPUNKT
        STROMKREISVERFOLGUNG
        KREUZUNGSVERZEICHNIS
        VERZ_RECHTLICHE_SICHERUNG
        VERZEICHNIS_AUSFUEHRUNGSPLANUNG_TEST
        VERZEICHNIS_MASTE_TEST
        SPANNWEITENLISTE
        ALTEMASTTAFEL
        SONSTIGES
        ABSTANDSLISTE
    End Enum

    Public Structure tExport
        Public App As Excel.Application
        Public wSheet As Excel.Worksheet
        Public sDir As String
        Public Dateiname As String
    End Structure

    Public Sub ExportStart(ByRef Data As System.Windows.Forms.DataGridView, _
                             ByVal TabName As String, _
                             ByVal MitFarben As Boolean, _
                             ByVal Art As eKopfTyp, _
                             ByVal BLnummer As String, _
                             ByVal sDir As String,
                             ByVal BLleitungsname As String,
                             ByVal BLStandort As String,
                             ByVal LtgPraefix As String)
        Dim Method As Export_ListenAsyncDelegate = AddressOf Export_ListenAsync
        Method.BeginInvoke(Data, TabName, MitFarben, Art, BLnummer, sDir, BLleitungsname, BLStandort, LtgPraefix, Nothing, Method)
    End Sub

    'Public Sub ExportSpannweitenListeStart(ByRef Data As Windows.Forms.DataGridView, _
    '                                              Dateiname As String, _
    '                                              Pfad As String)
    '    Dim method As Export_SpannweitenListeDelegate = AddressOf ExportSpannweitenlisteAsync
    '    method.BeginInvoke(Data, Dateiname, Pfad, Nothing, method)
    'End Sub

    Private Sub Zeilen_Nachformatieren(ByVal app As Excel.Application,
                                               ByVal begin As Integer,
                                               ByVal ende As Integer,
                                               ByVal name As String,
                                                ByVal font As String,
                                                ByVal size As Integer)
        Dim bereich As String = begin & ":" & ende
        app.Rows(bereich).Select()
        With app.Selection.Font
            .Name = name
            .FontStyle = font
            .Size = size
        End With
    End Sub

    'Private Sub ExportSpannweitenlisteAsync(ByRef Data As Windows.Forms.DataGridView, Dateiname As String, Pfad As String)
    '    Dateiname = clsExportTools.MakeValidTabname(Dateiname)
    '    Tools.Progress.Messenger.Start()
    '    m_TemporaererStatustext = "Excel Export gestartet"
    '    m_StatusAnzeigezeit = Georg.IStatusTextSender.eStatusAnzeigeZeit.TEMPORAER
    '    Georg.Refresh.Manager.DataChanged(Georg.IRefresh.eSendertyp.STATUSTEXT, New Georg.Refresh.StatusTextEventArgs(s_Instance))
    '    m_StatusAnzeigezeit = Georg.IStatusTextSender.eStatusAnzeigeZeit.PERMANENT
    '    Dim d As Object = FLPSettings.ImageManager.GetSpannweitenVorlage()
    '    Dim Name As String = Dateiname & C_EXCELDATEI_ENDUNG
    '    Dim Arr As String() = New String() {Pfad, Name}
    '    Dim XlsPfad As String = IO.Path.Combine(Arr)
    '    Dim fi As IO.FileInfo = New IO.FileInfo(XlsPfad)
    '    If fi.Exists Then fi.Delete()
    '    fi = Nothing
    '    Tools.Bytes2File(d, XlsPfad)
    '    Dim app As Excel.Application = Nothing
    '    Try
    '        app = New Excel.Application
    '    Catch exc As Exception
    '    End Try
    '    If app Is Nothing Then
    '        Tools.Message.SetHinweis(clsExcel.C_KEIN_EXCEL)
    '        Return
    '    End If
    '    app.Visible = False
    '    Dim wBook As Excel.Workbook = app.Workbooks.Open(XlsPfad)
    '    While wBook.Worksheets.Count > 1
    '        wBook.ActiveSheet.delete()
    '    End While
    '    Dim wSheet As Excel.Worksheet = wBook.ActiveSheet
    '    wSheet.Name = Dateiname
    '    Dim xlRow As Integer = 12
    '    Dim Cols As List(Of Integer) = New List(Of Integer)
    '    Cols.AddRange({1, 2, 3, 6, 7, 8, 9})
    '    Dim xlCol As Integer = 1
    '    SpannweitenZeilenAuslesen(Data, wSheet, xlRow, Cols, app, 80)
    '    wBook.Save()
    '    Tools.Progress.Messenger.Stopp()
    '    app.Visible = True
    '    app.UserControl = True
    '    app = Nothing
    'End Sub

    Private Sub Export_ListenAsync(ByRef Data As System.Windows.Forms.DataGridView,
                             ByVal TabName As String,
                             ByVal MitFarben As Boolean,
                             ByVal Art As eKopfTyp,
                             ByVal BLnummer As String,
                             ByVal sDir As String,
                             ByVal BLleitungsname As String,
                             ByVal BLStandort As String,
                             ByVal LtgPraefix As String)
        Dim dt As DataTable = Data.DataSource
        Dim ErsterMast As String = ""
        Dim LetzterMast As String = ""
        If dt.Rows.Count > 0 Then
            For Each col As DataColumn In dt.Columns
                If col.ColumnName.ToLower = "betriebsnummer" Then
                    If Not IsDBNull(dt.Rows(0).Item(col.ColumnName)) Then ErsterMast = dt.Rows(0).Item(col.ColumnName)
                    If Not IsDBNull(dt.Rows(dt.Rows.Count - 1).Item(col.ColumnName)) Then LetzterMast = dt.Rows(dt.Rows.Count - 1).Item(col.ColumnName)
                End If
                If col.ColumnName.ToLower = "nächster mast" Then
                    If Not IsDBNull(dt.Rows(dt.Rows.Count - 1).Item(col.ColumnName)) Then LetzterMast = dt.Rows(dt.Rows.Count - 1).Item(col.ColumnName)
                End If
            Next
        Else
            Return
        End If
        If String.IsNullOrEmpty(ErsterMast) Or String.IsNullOrEmpty(LetzterMast) Then
            For Each col As Windows.Forms.DataGridViewColumn In Data.Columns
                If col.HeaderText.ToLower = "betriebsnummer" Then
                    If Not IsDBNull(Data.Rows.Item(0).Cells(col.Index)) Then ErsterMast = Data.Rows.Item(0).Cells(col.Index).Value
                    If Not IsDBNull(Data.Rows.Item(Data.Rows.Count - 1).Cells(col.Index)) Then LetzterMast = Data.Rows.Item(Data.Rows.Count - 1).Cells(col.Index).Value
                End If
                If col.HeaderText.ToLower = "nächster mast" Then
                    If Not IsDBNull(dt.Rows(dt.Rows.Count - 1).Item(col.Index)) Then LetzterMast = dt.Rows(dt.Rows.Count - 1).Item(col.Index)
                End If
            Next
        End If
        Tools.Progress.Messenger.Start()
        m_TemporaererStatustext = "Excel Export gestartet"
        m_StatusAnzeigezeit = Georg.IStatusTextSender.eStatusAnzeigeZeit.TEMPORAER
        Georg.Refresh.Manager.DataChanged(Georg.IRefresh.eSendertyp.STATUSTEXT, New Georg.Refresh.StatusTextEventArgs(Me))
        m_StatusAnzeigezeit = Georg.IStatusTextSender.eStatusAnzeigeZeit.PERMANENT
        Dim dateiname As String
        If Art = eKopfTyp.ABSTANDSLISTE Then
                dateiname = LtgPraefix
            Else
                dateiname = LtgPraefix & BLnummer & "_" & Replace(TabName, " ", "_")
            End If
            TabName = clsExportTools.MakeValidTabname(TabName)
            Dim app As Excel.Application = Nothing
        Try
            app = New Excel.Application
        Catch exc As Exception
        End Try
        If app Is Nothing Then
            Tools.Message.SetHinweis(clsExcel.C_KEIN_EXCEL)
            Return
        End If
        app.Visible = False
            Dim wBook As Excel.Workbook = app.Workbooks.Add
            While wBook.Worksheets.Count > 1
                wBook.ActiveSheet.delete()
            End While
            Dim wSheet As Excel.Worksheet = wBook.ActiveSheet
            Dim xlRow As Integer = 6
            Dim xlCol As Integer = 1
        Dim max_row As Integer = 6
        Dim Errortext As String = ""
        Try
            If Art = eKopfTyp.VERZEICHNIS_AUSFUEHRUNGSPLANUNG Then
                wSheet.Name = "Verzeichnis_Ausführungsplanung"
                Kopfdaten_VERZ_AUSFUEHRUNGSPLANUNG(app, wBook, wSheet)
                xlRow = 6
                xlCol = 1
                max_row = ZeilenAuslesen(Data, wSheet, xlRow, xlCol, app, 116)
                Formatierung_VERZ_AUSFUEHRUNGSPLANUNG(app, wBook, wSheet, max_row, xlRow, Art)
                Errortext &= Fusszeilen(app, wBook, wSheet, max_row, xlRow, Art, BLnummer, BLStandort, BLleitungsname, ErsterMast, LetzterMast)
                Zeilen_Nachformatieren(app, xlRow, max_row, "Arial", "Standard", 10)
            ElseIf Art = eKopfTyp.VERZEICHNIS_GRUENDUNGEN Then
                wSheet.Name = TabName
                Kopfdaten_VERZ_GRUENDUNGEN(app, wBook, wSheet)
                xlRow = 6
                xlCol = 1
                max_row = ZeilenAuslesen(Data, wSheet, xlRow, xlCol, app, 77)
                Formatierung_VERZ_GRUENDUNGEN(app, wBook, wSheet, max_row, xlRow, Art)
                Errortext &= Fusszeilen(app, wBook, wSheet, max_row, xlRow, Art, BLnummer, BLStandort, BLleitungsname, ErsterMast, LetzterMast)
                Zeilen_Nachformatieren(app, xlRow, max_row, "Arial", "Standard", 10)
            ElseIf Art = eKopfTyp.VERZEICHNIS_MASTE Then
                Kopfdaten_VERZ_DER_MASTE(app, wBook, wSheet)
                wSheet.Name = TabName
                xlRow = 6
                xlCol = 1
                max_row = ZeilenAuslesen(Data, wSheet, xlRow, xlCol, app, 141)
                Formatierung_VERZ_DER_MASTE(app, wBook, wSheet, max_row, xlRow, Art)
                Errortext &= Fusszeilen(app, wBook, wSheet, max_row, xlRow, Art, BLnummer, BLStandort, BLleitungsname, ErsterMast, LetzterMast)
                Zeilen_Nachformatieren(app, xlRow, max_row, "Arial", "Standard", 10)
            ElseIf Art = eKopfTyp.INSTANDHALTUNG_STROMKREISBEZOGEN Then
                wSheet.Name = "Instandhaltung Stromkreisbez"
                Kopfdaten_INSTAND_STROMKREIS_BEZ(app, wBook, wSheet)
                xlRow = 5
                xlCol = 1
                max_row = ZeilenAuslesen(Data, wSheet, xlRow, xlCol, app, 52)
                Formatierung_INSTAND_STROMKREIS_BEZ(app, wBook, wSheet, max_row, xlRow, Art)
                Errortext &= Fusszeilen(app, wBook, wSheet, max_row, xlRow, Art, BLnummer, BLStandort, BLleitungsname, ErsterMast, LetzterMast)
                Zeilen_Nachformatieren(app, xlRow, max_row, "Arial", "Standard", 10)
            ElseIf Art = eKopfTyp.INSTANDHALTUNG_PHASENBEZOGEN Then
                wSheet.Name = "Instandhaltung Phasenbezogen"
                Kopfdaten_INSTAND_PHASEN_BEZ(app, wBook, wSheet)
                xlRow = 5
                xlCol = 1
                max_row = ZeilenAuslesen(Data, wSheet, xlRow, xlCol, app, 51)
                Formatierung_INSTAND_PHASEN_BEZ(app, wBook, wSheet, max_row, xlRow, Art)
                Errortext &= Fusszeilen(app, wBook, wSheet, max_row, xlRow, Art, BLnummer, BLStandort, BLleitungsname, ErsterMast, LetzterMast)
                Zeilen_Nachformatieren(app, xlRow, max_row, "Arial", "Standard", 10)
            ElseIf Art = eKopfTyp.INSTANDHALTUNG_MAST Then
                wSheet.Name = TabName
                Kopfdaten_INSTAND_MASTE(app, wBook, wSheet)
                xlRow = 5
                xlCol = 1
                max_row = ZeilenAuslesen(Data, wSheet, xlRow, xlCol, app, 88)
                formatierung_INSTAND_MASTE(app, wBook, wSheet, max_row, xlRow, Art)
                Errortext &= Fusszeilen(app, wBook, wSheet, max_row, xlRow, Art, BLnummer, BLStandort, BLleitungsname, ErsterMast, LetzterMast)
                Zeilen_Nachformatieren(app, xlRow, max_row, "Arial", "Standard", 10)
            ElseIf Art = eKopfTyp.INSTANDHALTUNG_AUFHAENGEPUNKT Then
                wSheet.Name = TabName
                Kopfdaten_INSTAND_Aufh_bez(app, wBook, wSheet)
                xlRow = 6
                xlCol = 1
                max_row = ZeilenAuslesen(Data, wSheet, xlRow, xlCol, app, 122)
                Formatierung_INSTAND_AUfHBEZ(app, wBook, wSheet, max_row, xlRow, Art)
                Errortext &= Fusszeilen(app, wBook, wSheet, max_row, xlRow, Art, BLnummer, BLStandort, BLleitungsname, ErsterMast, LetzterMast)
                Zeilen_Nachformatieren(app, xlRow, max_row, "Arial", "Standard", 10)
            ElseIf Art = eKopfTyp.STROMKREISVERFOLGUNG Then
                wSheet.Name = "Stromkreisverfolgung"
                Kopfdaten_STROMKREISVERFOLGUNG(app, wBook, wSheet)
                xlRow = 5
                xlCol = 1
                max_row = ZeilenAuslesen(Data, wSheet, xlRow, xlCol, app, 78)
                Formatierung_Stromkreisverfolgung(app, wBook, wSheet, max_row, xlRow, Art)
                Errortext &= Fusszeilen(app, wBook, wSheet, max_row, xlRow, Art, BLnummer, BLStandort, BLleitungsname, ErsterMast, LetzterMast)
                Zeilen_Nachformatieren(app, xlRow, max_row, "Arial", "Standard", 10)
            ElseIf Art = eKopfTyp.KREUZUNGSVERZEICHNIS Then
                Kopfdaten_KREUZUNGSVERZEICHNIS(app, wBook, wSheet)
                xlRow = 4
                xlCol = 1
                max_row = ZeilenAuslesen(Data, wSheet, xlRow, xlCol, app, 78)
                Formatierung_KREUZUGNSVERZEICHNIS(app, wBook, wSheet, max_row, xlRow, Art)
                Errortext &= Fusszeilen(app, wBook, wSheet, max_row, xlRow, Art, BLnummer, BLStandort, BLleitungsname, ErsterMast, LetzterMast)
                Zeilen_Nachformatieren(app, xlRow, max_row, "Arial", "Standard", 10)
            ElseIf Art = eKopfTyp.VERZ_RECHTLICHE_SICHERUNG Then
                Kopfdaten_VERZ_RECHTLICHE_SICHERUNG(app, wBook, wSheet)
                xlRow = 5
                xlCol = 1
                max_row = ZeilenAuslesen(Data, wSheet, xlRow, xlCol, app, 50)
                Formatierung_RECHTLICHE_SICHERUNG(app, wBook, wSheet, max_row, xlRow, Art)
                Errortext &= Fusszeilen(app, wBook, wSheet, max_row, xlRow, Art, BLnummer, BLStandort, BLleitungsname, ErsterMast, LetzterMast)
                Zeilen_Nachformatieren(app, xlRow, max_row, "Arial", "Standard", 10)
            ElseIf Art = eKopfTyp.ABSTANDSLISTE Then
                Kopfdaten_ABSTANDSLISTE(app, wBook, wSheet)
                xlRow = 4
                xlCol = 1
                max_row = ZeilenAuslesen(Data, wSheet, xlRow, xlCol, app, 50)
                Formatierung_ABSTANDSLISTE(app, wBook, wSheet, max_row, xlRow, Art)
                Errortext &= Fusszeilen(app, wBook, wSheet, max_row, xlRow, Art, BLnummer, BLStandort, BLleitungsname, ErsterMast, LetzterMast)
                Zeilen_Nachformatieren(app, xlRow, max_row, "Arial", "Standard", 10)
            ElseIf Art = eKopfTyp.SONSTIGES Then
                Export(Data, TabName, False)
            End If
            wSheet.Cells(2, 1).Select()
        Catch exc As Exception
            Errortext &= exc.Message
            Dim err As New Tools.clsDatenFehler("Export der Tabelle " & TabName & " fehlgeschlagen.", exc)
            Tools.Message.SetToDatenBank(err, "")
            Tools.Message.SetErrorMessage(err, "Prüfen Sie die Bindung für xsl/csv-Dateien.")
        End Try
        If Len(sDir) > 0 Then
            Dim Name As String = ""
            Try
                Dim di As IO.DirectoryInfo = New IO.DirectoryInfo(sDir)
                If Not di.Exists Then di.Create()
                Dim Datum As Date = Date.Now
                Dim DatumStempel As String = Datum.Year & "_" & Datum.Month & "_" & Datum.Day & "_" & Datum.Hour & "_" & Datum.Minute
                Dim Endung As String = EXCELExport.C_EXCELDATEI_ENDUNG_XLSX
                If Val(app.Version) < EXCELExport.C_EXCEL_VERSION_MIT_XLSX Then Endung = EXCELExport.C_EXCELDATEI_ENDUNG
                If Art = eKopfTyp.ABSTANDSLISTE Then
                    Name = TabName & dateiname & Endung
                Else
                    Name = dateiname & "_" & DatumStempel & Endung
                End If
                Dim Arr As String() = New String() {sDir, Name}
                Dim Pfad As String = IO.Path.Combine(Arr)
                wSheet.SaveAs(Pfad)
                app.Quit()
                app = Nothing
                'Process.Start(Pfad)
            Catch ex As System.IO.IOException
                Tools.Progress.Messenger.Stopp()
                Tools.Message.SetErrorMessage(New Tools.clsExcelError("Datei " & Name & " konnte nicht gespeichert werden.", ex), "Prüfen Sie bitte vor dem Export ob die Datei " & vbNewLine & sDir & dateiname & vbNewLine & "ggf. bereits besteht!")
            End Try
        End If
        Tools.Progress.Messenger.Stopp()
        If Errortext.Length > 0 Then Tools.Message.SetHinweis(Errortext)
        'app.Visible = True
        'app.Quit()
        'app = Nothing
    End Sub

    Public Sub Kopfzelle_Format(ByRef Excel As Excel.Application, _
                                       ByRef Position As String, _
                                       ByRef Orientierung As Integer, _
                                       ByRef Merge As Boolean, _
                                       ByRef Inhalt As String, _
                                       Optional ByRef inhalt2 As String = "", _
                                       Optional ByRef inhalt3 As String = "", _
                                       Optional ByRef inhalt4 As String = "", _
                                       Optional ByRef inhalt5 As String = "", _
                                       Optional ByRef inhalt6 As String = "")
        Dim DieseRange As Excel.Range = Excel.Range(Position)
        With Excel.Range(Position).Cells
            .Orientation = Orientierung
            .MergeCells = Merge
        End With
        Excel.Range(Position).Select()
        Dim InhaltZumSetzen As String = Inhalt
        If Len(inhalt6) > 0 Then
            Excel.ActiveCell.FormulaR1C1 = InhaltZumSetzen & Chr(10) & inhalt2 & Chr(10) & inhalt3 & Chr(10) & inhalt4 & Chr(10) & inhalt5 & Chr(10) & inhalt6
        End If
        If Len(inhalt5) > 0 And Len(inhalt6) = 0 Then
            Excel.ActiveCell.FormulaR1C1 = InhaltZumSetzen & Chr(10) & inhalt2 & Chr(10) & inhalt3 & Chr(10) & inhalt4 & Chr(10) & inhalt5
        End If
        If Len(inhalt4) > 0 And Len(inhalt5) = 0 Then
            Excel.ActiveCell.FormulaR1C1 = InhaltZumSetzen & Chr(10) & inhalt2 & Chr(10) & inhalt3 & Chr(10) & inhalt4
        End If
        If Len(inhalt3) > 0 And Len(inhalt4) = 0 Then
            Excel.ActiveCell.FormulaR1C1 = InhaltZumSetzen & Chr(10) & inhalt2 & Chr(10) & inhalt3
        End If
        If Len(inhalt2) > 0 And Len(inhalt3) = 0 Then
            Excel.ActiveCell.FormulaR1C1 = InhaltZumSetzen & Chr(10) & inhalt2
        End If
        If Len(inhalt2) = 0 Then
            Excel.ActiveCell.FormulaR1C1 = InhaltZumSetzen
        End If
    End Sub

    Public Sub Standardformat(ByRef app As Excel.Application, _
                                     ByVal Bereich As String,
                                     ByVal Horizontal As Microsoft.Office.Interop.Excel.Constants,
                                     ByVal Name As String,
                                     ByVal fontstyle As String,
                                     ByVal size As Integer,
                                     Optional ByVal Vertical As Microsoft.Office.Interop.Excel.Constants = Excel.Constants.xlCenter)

        'Standardformat für alle Zellen
        With app.Range(Bereich).Cells
            .HorizontalAlignment = Horizontal
            .VerticalAlignment = Vertical
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False
            .Merge()
            .Font.Bold = True
        End With
        app.Range(Bereich).Select()
        With app.Selection.Font
            .Name = Name
            .FontStyle = fontstyle
            .Size = size
        End With
    End Sub

    Public Sub Formateinteilung(ByVal Bereich As String,
                                       ByVal Borderdef As Microsoft.Office.Interop.Excel.XlBordersIndex,
                                       ByVal weight As Microsoft.Office.Interop.Excel.XlBorderWeight,
                                       ByVal linestyle As Microsoft.Office.Interop.Excel.XlLineStyle,
                                       ByVal colorindex As Microsoft.Office.Interop.Excel.XlColorIndex,
                                       ByVal app As Excel.Application, _
                                       Optional ByVal max_row As String = "")
        Dim Bereich2 As String = Bereich & max_row
        'app.Range(Bereich2).Select()
        With app.Range(Bereich2).Borders(Borderdef)
            .LineStyle = linestyle
            .ColorIndex = colorindex
            .Weight = weight
        End With
    End Sub

    Private Sub FormateinteilungNeu(ByVal Bereich As Excel.Range,
                                           ByVal Borderdef As Microsoft.Office.Interop.Excel.XlBordersIndex,
                                           ByVal weight As Microsoft.Office.Interop.Excel.XlBorderWeight,
                                           ByVal linestyle As Microsoft.Office.Interop.Excel.XlLineStyle,
                                           ByVal colorindex As Microsoft.Office.Interop.Excel.XlColorIndex)
        With Bereich.Borders(Borderdef)
            .LineStyle = linestyle
            .ColorIndex = colorindex
            .Weight = weight
        End With
    End Sub

    Private Const C_SCHUTZVERMERK = "Weitergabe sowie Vervielfältigung dieses Dokumentes, Verwendung und Mit-" & vbNewLine &
                                    "teilung seines Inhaltes sind verboten, soweit nicht ausdrücklich gestattet."


    Private ReadOnly Property Schutzvermerk As String
        Get
            Return C_SCHUTZVERMERK
        End Get
    End Property

    Public Sub Druckbereich_Kopf_Fussdaten(ByVal Bereich As String, ByVal Kopftext_bereich As String,
                                                  ByVal fusszeile_links As String,
                                                  ByVal fusszeile_mitte As String,
                                                  ByVal fusszeile_rechts As String,
                                                  ByVal app As Excel.Application,
                                                  ByVal xlRow As Integer,
                                                  ByVal max_row As String,
                                                  ByVal Art As eKopfTyp)

        If Art = eKopfTyp.KREUZUNGSVERZEICHNIS Then
            max_row = max_row + 7
        Else
            max_row = max_row + 9 'Fussbereich jeder Liste ist 9 Zeilen 
        End If
        Dim printbereich As String = Bereich & "$" & max_row
        'Dim anzseiten As String = app.ExecuteExcel4Macro("Get.Document(50)")
        With app.ActiveSheet.PageSetup
            .PrintTitleRows = Kopftext_bereich
            .PrintArea = printbereich
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = "&8&D"
            .CenterFooter = "&8Copyright by " & Kundenbezeichnung & vbNewLine & "&""Arial,Fett""&10" & Schutzvermerk
            '.CenterFooter = "&8Copyright by " & Kundenbezeichnung
            .RightFooter = "&8&P von &N Seiten"
            .LeftMargin = app.Application.InchesToPoints(0.708661417322835)
            .RightMargin = app.Application.InchesToPoints(0.708661417322835)
            .TopMargin = app.Application.InchesToPoints(0.78740157480315)
            .BottomMargin = app.Application.InchesToPoints(0.78740157480315)
            .HeaderMargin = app.Application.InchesToPoints(0.31496062992126)
            .FooterMargin = app.Application.InchesToPoints(0.31496062992126)
            .PrintHeadings = False
            .PrintGridlines = True
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = Excel.XlPageOrientation.xlLandscape
            .Draft = False
            .PaperSize = Excel.XlPaperSize.xlPaperA3
            .FirstPageNumber = Excel.Constants.xlAutomatic
            .Order = Excel.XlOrder.xlDownThenOver
            .BlackAndWhite = False
            If Art = eKopfTyp.VERZEICHNIS_AUSFUEHRUNGSPLANUNG Then
                ' "Verzeichnis der Ausführungsplanung"
                .Zoom = 45
            ElseIf Art = eKopfTyp.VERZEICHNIS_GRUENDUNGEN Then
                '"Verzeichnis der Gründungen"
                .Zoom = 63
            ElseIf Art = eKopfTyp.VERZEICHNIS_MASTE Then
                '"Verzeichnis der Maste"
                .Zoom = 38
            ElseIf Art = eKopfTyp.INSTANDHALTUNG_STROMKREISBEZOGEN Then
                '"Instandhaltung (Stromkreisbezogen)"
                .Zoom = 90
            ElseIf Art = eKopfTyp.INSTANDHALTUNG_PHASENBEZOGEN Then
                '"Instandhaltung (Phasenbezogen)"
                .Zoom = 91
            ElseIf Art = eKopfTyp.INSTANDHALTUNG_MAST Then
                '"Instandhaltung (Mast)"
                .Zoom = 57
            ElseIf Art = eKopfTyp.INSTANDHALTUNG_AUFHAENGEPUNKT Then
                '"Instandhaltung (Aufhaengepunkt)"
                .Zoom = 43
            ElseIf Art = eKopfTyp.STROMKREISVERFOLGUNG Then
                '"Stromkreisverfolgung"
                .Zoom = 65
            ElseIf Art = eKopfTyp.KREUZUNGSVERZEICHNIS Then
                .Zoom = 70
            ElseIf Art = eKopfTyp.VERZ_RECHTLICHE_SICHERUNG Then
                .Zoom = 25
            ElseIf Art = eKopfTyp.VERZEICHNIS_MASTE_TEST Then
                '"Instandhaltung (Mast)"
                .Zoom = 57
            ElseIf Art = eKopfTyp.VERZEICHNIS_AUSFUEHRUNGSPLANUNG_TEST Then
                '"Instandhaltung (Aufhaengepunkt)"
                .Zoom = 40
            ElseIf Art = eKopfTyp.SONSTIGES Then
                .Zoom = 90
            End If
            .PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed
        End With
    End Sub

    Public Sub Formatierung_zelle(ByVal Schriftart As String,
                                               ByVal fontstyle As String,
                                               ByVal size As Integer,
                                               ByVal underline As Microsoft.Office.Interop.Excel.Constants,
                                               ByVal colorindex As Microsoft.Office.Interop.Excel.Constants,
                                               ByVal Position As String,
                                               ByVal beginn As Integer,
                                               ByVal trennung As Integer,
                                               ByRef app As Excel.Application)
        app.Range(Position).Select()
        Dim max As Integer = app.ActiveCell.Characters.Count()
        With app.ActiveCell.Characters(Start:=beginn, Length:=trennung).Font
            .Name = Schriftart
            .FontStyle = fontstyle
            .Size = size
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = underline
            .ColorIndex = colorindex
        End With
        trennung = trennung + 1
        With app.ActiveCell.Characters(Start:=trennung, Length:=max).Font
            .Name = Schriftart
            .FontStyle = fontstyle
            .Size = size
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone
            .ColorIndex = colorindex
        End With
    End Sub

    Public Sub Kopfdaten_INSTAND_PHASEN_BEZ(ByRef app As Excel.Application,
                                                   ByRef wBook As Excel.Workbook,
                                                   ByRef wSheet As Excel.Worksheet)
        'Standardformat für alle Zellen
        Standardformat(app, "A1:S4", Excel.Constants.xlCenter, "Arial", "Fett", 10)
        'Einzelne Attribute
        Kopfzelle_Format(app, "A2", 0, False, "Bau-", "nummer")
        Kopfzelle_Format(app, "B2", 0, False, "Betriebs-", "nummer")
        Kopfzelle_Format(app, "C1:C2", 0, True, "nächster", " Mast")
        Kopfzelle_Format(app, "D2", 0, False, "Stromkreis")
        Kopfzelle_Format(app, "E2", 0, False, "Netz-", "art")
        Kopfzelle_Format(app, "F2", 0, False, "Phasen-", "bezogen", "(R oder T oder ES)")
        Kopfzelle_Format(app, "G2", 0, False, "Leiter-", "end-", "tem-", "pera-", "tur")
        Kopfzelle_Format(app, "H2", 0, False, "Kurz-", "schluss-", "strom")
        Kopfzelle_Format(app, "I2", 0, False, "Bau-", "jahr", " Leiter")
        Kopfzelle_Format(app, "J2", 0, False, "Bau-", "jahr", " LE TMP")
        Kopfzelle_Format(app, "K2", 0, False, "Funk-", "tion", " des ", "Leiters")
        Kopfzelle_Format(app, "L2", 0, False, "Bezeichnung", "Leiter")
        Kopfzelle_Format(app, "M2", 0, False, "An-", "zahl", "Teil-", "leiter")
        Kopfzelle_Format(app, "N2", 0, False, "Ab-", "stand", "Teil-", "leiter")
        Kopfzelle_Format(app, "O2", 0, False, "Bündel", "anord-", "nung")
        Kopfzelle_Format(app, "P2", 0, False, "Ver-", "binder")
        Kopfzelle_Format(app, "Q2", 0, False, "Flug-", "warn-", "kugeln")
        Kopfzelle_Format(app, "R2", 0, False, "Ver-", "driller")
        'Einheiten und Nummerierung
        Kopfzelle_Format(app, "G3", 0, False, "[°C]")
        Kopfzelle_Format(app, "H3", 0, False, "[A]")
        Kopfzelle_Format(app, "N3", 0, False, "[m]")
        Kopfzelle_Format(app, "A4", 0, False, "1")
        Kopfzelle_Format(app, "B4", 0, False, "2")
        Kopfzelle_Format(app, "C4", 0, False, "3")
        Kopfzelle_Format(app, "D4", 0, False, "4")
        Kopfzelle_Format(app, "E4", 0, False, "5")
        Kopfzelle_Format(app, "F4", 0, False, "6")
        Kopfzelle_Format(app, "G4", 0, False, "7")
        Kopfzelle_Format(app, "H4", 0, False, "8")
        Kopfzelle_Format(app, "I4", 0, False, "9")
        Kopfzelle_Format(app, "J4", 0, False, "10")
        Kopfzelle_Format(app, "K4", 0, False, "11")
        Kopfzelle_Format(app, "L4", 0, False, "12")
        Kopfzelle_Format(app, "M4", 0, False, "13")
        Kopfzelle_Format(app, "N4", 0, False, "14")
        Kopfzelle_Format(app, "O4", 0, False, "15")
        Kopfzelle_Format(app, "P4", 0, False, "16")
        Kopfzelle_Format(app, "Q4", 0, False, "17")
        Kopfzelle_Format(app, "R4", 0, False, "18")
        Kopfzelle_Format(app, "S4", 0, False, "19")
        'Hauptüberschriften erzeugen
        Kopfzelle_Format(app, "A1:B1", 0, True, "Mastnummer")
        Kopfzelle_Format(app, "D1:F1", 0, True, "")
        Kopfzelle_Format(app, "G1:H1", 0, True, "elektrotechn.", "Parameter")
        Kopfzelle_Format(app, "I1:R1", 0, True, "Weitere Angaben zu Leiter")
        Kopfzelle_Format(app, "S1:S2", 0, True, "Bemerkung")
        app.Rows("1:1").RowHeight = 29.25
        app.Rows("2:2").RowHeight = 105
        'Spaltenbeiten
        app.Columns("A:A").ColumnWidth = 20
        app.Columns("B:B").ColumnWidth = 10.71
        app.Columns("C:C").ColumnWidth = 10.71
        app.Columns("D:D").ColumnWidth = 15
        app.Columns("E:E").ColumnWidth = 6
        app.Columns("F:F").ColumnWidth = 9
        app.Columns("G:G").ColumnWidth = 6
        app.Columns("H:H").ColumnWidth = 7
        app.Columns("I:I").ColumnWidth = 6
        app.Columns("J:J").ColumnWidth = 6
        app.Columns("K:K").ColumnWidth = 7
        app.Columns("L:L").ColumnWidth = 25
        app.Columns("M:M").ColumnWidth = 6
        app.Columns("N:N").ColumnWidth = 6
        app.Columns("O:O").ColumnWidth = 7
        app.Columns("P:P").ColumnWidth = 6
        app.Columns("Q:Q").ColumnWidth = 7
        app.Columns("R:R").ColumnWidth = 6
        app.Columns("S:S").ColumnWidth = 33
    End Sub

    Public Sub Kopfdaten_INSTAND_STROMKREIS_BEZ(ByRef app As Excel.Application,
                                                       ByRef wBook As Excel.Workbook,
                                                       ByRef wSheet As Excel.Worksheet)
        'Standardformat für alle Zellen
        Standardformat(app, "A1:Q4", Excel.Constants.xlCenter, "Arial", "Fett", 10)
        'Einzelne Attribute
        Kopfzelle_Format(app, "A2", 0, False, "Baunummer")
        Kopfzelle_Format(app, "B2", 0, False, "Betriebs-", "nummer", "(nur ", "Abspanner)")
        Kopfzelle_Format(app, "C2", 0, False, "anlagen-", "buchhalterische", "Zugehörigkeit")
        Kopfzelle_Format(app, "D2", 0, False, "Betriebs-", "nummer", "(nächster ", "Abspanner)")
        Kopfzelle_Format(app, "E2", 0, False, "An-", "zahl", "Trag-", "maste")
        Kopfzelle_Format(app, "F2", 0, False, "Ein-/", " Aus-", "kreuz-", "ung", "vor-", "handen")
        Kopfzelle_Format(app, "G2", 0, False, "von", "Bahnstrom-", "leitung")
        Kopfzelle_Format(app, "J2", 0, False, "Einfach-", "Seil")
        Kopfzelle_Format(app, "K2", 0, False, "2-Bdl.")
        Kopfzelle_Format(app, "L2", 0, False, "4-Bdl.")
        Kopfzelle_Format(app, "M2", 0, False, "Leiter-", "end-", "Tempe-", "ratur")
        Kopfzelle_Format(app, "N2", 0, False, "Kurz-", "schluss-", "strom")
        'Einheiten und Nummerierung
        Kopfzelle_Format(app, "J3", 0, False, "[m]")
        Kopfzelle_Format(app, "K3", 0, False, "[m]")
        Kopfzelle_Format(app, "L3", 0, False, "[m]")
        Kopfzelle_Format(app, "M3", 0, False, "[°C]")
        Kopfzelle_Format(app, "N3", 0, False, "[A]")
        Kopfzelle_Format(app, "A4", 0, False, "1")
        Kopfzelle_Format(app, "B4", 0, False, "2")
        Kopfzelle_Format(app, "C4", 0, False, "3")
        Kopfzelle_Format(app, "D4", 0, False, "4")
        Kopfzelle_Format(app, "E4", 0, False, "5")
        Kopfzelle_Format(app, "F4", 0, False, "6")
        Kopfzelle_Format(app, "G4", 0, False, "7")
        Kopfzelle_Format(app, "H4", 0, False, "8")
        Kopfzelle_Format(app, "I4", 0, False, "9")
        Kopfzelle_Format(app, "J4", 0, False, "10")
        Kopfzelle_Format(app, "K4", 0, False, "11")
        Kopfzelle_Format(app, "L4", 0, False, "12")
        Kopfzelle_Format(app, "M4", 0, False, "13")
        Kopfzelle_Format(app, "N4", 0, False, "14")
        Kopfzelle_Format(app, "O4", 0, False, "15")
        Kopfzelle_Format(app, "P4", 0, False, "16")
        Kopfzelle_Format(app, "Q4", 0, False, "17")
        'Hauptüberschriften erzeugen
        Kopfzelle_Format(app, "A1:B1", 0, True, "Mastnummer")
        Kopfzelle_Format(app, "C1", 0, True, "Zugehörigkeit")
        Kopfzelle_Format(app, "D1:G1", 0, True, "Verlaufsinformationen")
        Kopfzelle_Format(app, "H1:H2", 0, True, "Stromkreis")
        Kopfzelle_Format(app, "I1:I2", 0, True, "Netz-", "art")
        Kopfzelle_Format(app, "J1:L1", 0, True, "Stromkreislänge")
        Kopfzelle_Format(app, "M1:N1", 0, True, "elektrontechn.", "Parameter")
        Kopfzelle_Format(app, "O1:O2", 0, True, "Bau-", "jahr", "LE", "TMP")
        Kopfzelle_Format(app, "P1:P2", 0, True, "Ver-", "driller")
        Kopfzelle_Format(app, "Q1:Q2", 0, True, "Bemerkung")
        app.Rows("1:1").RowHeight = 29.25
        app.Rows("2:2").RowHeight = 100.25
        'Spaltenbeiten
        app.Columns("A:A").ColumnWidth = 20
        app.Columns("B:B").ColumnWidth = 10.71
        app.Columns("C:C").ColumnWidth = 15
        app.Columns("D:D").ColumnWidth = 10.71
        app.Columns("E:E").ColumnWidth = 6
        app.Columns("F:F").ColumnWidth = 7
        app.Columns("G:G").ColumnWidth = 14
        app.Columns("H:H").ColumnWidth = 22
        app.Columns("I:I").ColumnWidth = 6
        app.Columns("J:J").ColumnWidth = 10.71
        app.Columns("K:K").ColumnWidth = 10.71
        app.Columns("L:L").ColumnWidth = 10.71
        app.Columns("M:M").ColumnWidth = 7
        app.Columns("N:N").ColumnWidth = 7
        app.Columns("O:O").ColumnWidth = 6
        app.Columns("P:P").ColumnWidth = 6
        app.Columns("Q:Q").ColumnWidth = 28

    End Sub

    Public Sub Kopfdaten_INSTAND_MASTE(ByRef app As Excel.Application,
                                              ByRef wBook As Excel.Workbook,
                                              ByRef wSheet As Excel.Worksheet)
        'Standardformat für alle Zellen
        Standardformat(app, "A1:AJ4", Excel.Constants.xlCenter, "Arial", "Fett", 10)
        'Einzelne Attribute
        Kopfzelle_Format(app, "A2", 0, False, "Baunummer")
        Kopfzelle_Format(app, "B2", 0, False, "Betriebs-", "nummer")
        Kopfzelle_Format(app, "C1:C2", 0, True, "anlagen-", "buchhalt-", "erische", "Zugehör-", "igkeit")
        Kopfzelle_Format(app, "D2", 0, False, "Betriebs-", "kenn-", "zahl")
        Kopfzelle_Format(app, "E2", 0, False, "DBEn-", "Standort")
        Kopfzelle_Format(app, "F2", 0, False, "Gemein-", "schafts-", "leitung", "(bei", "Vertrags-", "abschluss)")
        Kopfzelle_Format(app, "G2", 0, False, "Gemein-", "schafts-", "leitung", "(aktuell)")
        Kopfzelle_Format(app, "H2", 0, False, "Typ")
        Kopfzelle_Format(app, "I2", 0, False, "Standort")
        Kopfzelle_Format(app, "J2", 0, False, "Trasse")
        Kopfzelle_Format(app, "K2", 0, False, "Mast")
        Kopfzelle_Format(app, "L2", 0, False, "Be-", "schicht-", "ung")
        Kopfzelle_Format(app, "M2", 0, False, "Funda", "ment")
        Kopfzelle_Format(app, "N2", 0, False, "Werkstatt-", "beschicht-", "ung")
        Kopfzelle_Format(app, "O2", 0, False, "Verzinkt", "Ja /", "Nein /", "teilweise")
        Kopfzelle_Format(app, "P2", 0, False, "Hersteller")
        Kopfzelle_Format(app, "Q2", 0, False, "Applikateur")
        Kopfzelle_Format(app, "R2", 0, False, "Stoff-", "nummer")
        Kopfzelle_Format(app, "S2", 0, False, "Chargen-", "nummer")
        Kopfzelle_Format(app, "T2", 0, False, "Jahr")
        Kopfzelle_Format(app, "U2", 0, False, "Wert")
        Kopfzelle_Format(app, "V2", 0, False, "Ausf. Firma")
        'Einheiten und Spaltennummerierung
        Kopfzelle_Format(app, "U3", 0, False, "[m}")
        Kopfzelle_Format(app, "A4", 0, False, "1")
        Kopfzelle_Format(app, "B4", 0, False, "2")
        Kopfzelle_Format(app, "C4", 0, False, "3")
        Kopfzelle_Format(app, "D4", 0, False, "4")
        Kopfzelle_Format(app, "E4", 0, False, "5")
        Kopfzelle_Format(app, "F4", 0, False, "6")
        Kopfzelle_Format(app, "G4", 0, False, "7")
        Kopfzelle_Format(app, "H4", 0, False, "8")
        Kopfzelle_Format(app, "I4", 0, False, "9")
        Kopfzelle_Format(app, "J4", 0, False, "10")
        Kopfzelle_Format(app, "K4", 0, False, "11")
        Kopfzelle_Format(app, "L4", 0, False, "12")
        Kopfzelle_Format(app, "M4", 0, False, "13")
        Kopfzelle_Format(app, "N4", 0, False, "14")
        Kopfzelle_Format(app, "O4", 0, False, "15")
        Kopfzelle_Format(app, "P4", 0, False, "16")
        Kopfzelle_Format(app, "Q4", 0, False, "17")
        Kopfzelle_Format(app, "R4", 0, False, "18")
        Kopfzelle_Format(app, "S4", 0, False, "19")
        Kopfzelle_Format(app, "T4", 0, False, "20")
        Kopfzelle_Format(app, "U4", 0, False, "21")
        Kopfzelle_Format(app, "V4", 0, False, "22")
        Kopfzelle_Format(app, "W4", 0, False, "23")
        Kopfzelle_Format(app, "X4", 0, False, "24")
        Kopfzelle_Format(app, "Y4", 0, False, "25")
        Kopfzelle_Format(app, "Z4", 0, False, "26")
        'Hauptüberschriften erzeugen
        Kopfzelle_Format(app, "A1:B1", 0, True, "Mastnummer")
        Kopfzelle_Format(app, "D1:E1", 0, True, "Region")
        Kopfzelle_Format(app, "F1:G1", 0, True, "Partner")
        Kopfzelle_Format(app, "H1:I1", 0, True, "Maste")
        Kopfzelle_Format(app, "J1:M1", 0, True, "Baujahre")
        Kopfzelle_Format(app, "N1:S1", 0, True, "weitere Angaben zur Beschichtung")
        Kopfzelle_Format(app, "T1:V1", 0, True, "Masterhöhung")
        Kopfzelle_Format(app, "W1:W2", 0, True, "besondere", "Anbauten")
        Kopfzelle_Format(app, "X1:X2", 0, True, "Natur-", "schutz", "FFH")
        Kopfzelle_Format(app, "Y1:Y2", 0, True, "Zugangswege")
        Kopfzelle_Format(app, "Z1:Z2", 0, True, "Bemerkung")
        Kopfzelle_Format(app, "Z4", 0, True, "26")
        app.Rows("1:1").RowHeight = 29.25
        app.Rows("2:2").RowHeight = 120.25
        app.Columns("A:A").ColumnWidth = 20
        app.Columns("B:B").ColumnWidth = 10.71
        app.Columns("C:C").ColumnWidth = 9
        app.Columns("D:D").ColumnWidth = 8
        app.Columns("E:E").ColumnWidth = 10
        app.Columns("F:F").ColumnWidth = 10.71
        app.Columns("G:G").ColumnWidth = 10.71
        app.Columns("H:H").ColumnWidth = 30
        app.Columns("I:I").ColumnWidth = 20
        app.Columns("J:J").ColumnWidth = 6
        app.Columns("K:K").ColumnWidth = 6
        app.Columns("L:L").ColumnWidth = 7
        app.Columns("M:M").ColumnWidth = 6
        app.Columns("N:N").ColumnWidth = 9
        app.Columns("O:O").ColumnWidth = 9
        app.Columns("P:P").ColumnWidth = 16
        app.Columns("Q:Q").ColumnWidth = 16
        app.Columns("R:R").ColumnWidth = 10.71
        app.Columns("S:S").ColumnWidth = 10.71
        app.Columns("T:T").ColumnWidth = 6
        app.Columns("U:U").ColumnWidth = 6
        app.Columns("V:V").ColumnWidth = 16
        app.Columns("W:W").ColumnWidth = 16
        app.Columns("X:X").ColumnWidth = 6
        app.Columns("Y:Y").ColumnWidth = 16
        app.Columns("Z:Z").ColumnWidth = 33
    End Sub

    Public Sub Kopfdaten_INSTAND_Aufh_bez(ByRef app As Excel.Application,
                                                 ByRef wBook As Excel.Workbook,
                                                 ByRef wSheet As Excel.Worksheet)
        'Standardformat für alle Zellen
        Standardformat(app, "A1:AJ5", Excel.Constants.xlCenter, "Arial", "Fett", 10)
        'Einzelne Attribute
        Kopfzelle_Format(app, "G3", 0, False, "Ketten-", "einbau")
        Kopfzelle_Format(app, "H3", 0, False, "Ketten-", "einbau-", "jahr")
        Kopfzelle_Format(app, "I3", 0, False, "Ebf.Nr.", "Kettenzeichnung")
        Kopfzelle_Format(app, "J3", 0, False, "Ge-", "nehmig-", "ungs-", "datum ", "(zu 8)")
        Kopfzelle_Format(app, "K2:L2", 0, True, "Amatur")
        Kopfzelle_Format(app, "K3", 0, False, "Amatur")
        Kopfzelle_Format(app, "L3", 0, False, "Amatur", " Bau-", "jahr")
        Kopfzelle_Format(app, "M2:O2", 0, True, "Isolator")
        Kopfzelle_Format(app, "M3", 0, False, "Iso-Bezeichnung")
        Kopfzelle_Format(app, "N3", 0, False, "Iso-Hersteller")
        Kopfzelle_Format(app, "O3", 0, False, "Iso-", "Bau-", "jahr")
        Kopfzelle_Format(app, "P2:R2", 0, True, "Schwingungsschutz")
        Kopfzelle_Format(app, "P3", 0, False, "Typ")
        Kopfzelle_Format(app, "Q3", 0, False, "Hersteller")
        Kopfzelle_Format(app, "R3", 0, False, "Bau-", "jahr")
        Kopfzelle_Format(app, "S2:AA2", 0, True, "Stromkreis und Phase")
        Kopfzelle_Format(app, "S3", 0, False, "Stromkreis")
        Kopfzelle_Format(app, "T3", 0, False, "Netzart")
        Kopfzelle_Format(app, "U3", 0, False, "Phasen-", "bezeich-", "nung")
        Kopfzelle_Format(app, "V3", 0, False, "Bezeichnung", "Leiter")
        Kopfzelle_Format(app, "W3", 0, False, "An-", "zahl", "Teil-", "leiter")
        Kopfzelle_Format(app, "X3", 0, False, "Ab-", "stand", "Teil-", "leiter")
        Kopfzelle_Format(app, "Y3", 0, False, "Bündel-", "anord-", "nung")
        Kopfzelle_Format(app, "Z3", 0, False, "Bau-", "jahr", "Leiter")
        Kopfzelle_Format(app, "AA3", 0, False, "Bau-", "jahr", "LE", "TEMP")
        Kopfzelle_Format(app, "AB2:AC2", 0, True, "elektrotechn.", "Parameter")
        Kopfzelle_Format(app, "AB3", 0, False, "Leiter-", "end-", "tem-", "peratur")
        Kopfzelle_Format(app, "AC3", 0, False, "Kurz-", "schlus-", "sstrom")
        Kopfzelle_Format(app, "AD2:AD3", 0, True, "von", " Traverse")
        Kopfzelle_Format(app, "AF2:AF3", 0, True, "zu", " Traverse")
        Kopfzelle_Format(app, "AE2:AE3", 0, True, "von", "Auf-", "hänge-", "punkt")
        Kopfzelle_Format(app, "AG2:AG3", 0, True, "zu", "Auf-", "hänge-", "punkt")
        'Einheiten und Spaltennummerierung
        Kopfzelle_Format(app, "AB4", 0, False, "[°C]")
        Kopfzelle_Format(app, "AC4", 0, False, "[A]")
        Kopfzelle_Format(app, "X4", 0, False, "[m]")
        Kopfzelle_Format(app, "A5", 0, False, "1")
        Kopfzelle_Format(app, "B5", 0, False, "2")
        Kopfzelle_Format(app, "C5", 0, False, "3")
        Kopfzelle_Format(app, "D5", 0, False, "4")
        Kopfzelle_Format(app, "E5", 0, False, "5")
        Kopfzelle_Format(app, "F5", 0, False, "6")
        Kopfzelle_Format(app, "G5", 0, False, "7")
        Kopfzelle_Format(app, "H5", 0, False, "8")
        Kopfzelle_Format(app, "I5", 0, False, "9")
        Kopfzelle_Format(app, "J5", 0, False, "10")
        Kopfzelle_Format(app, "K5", 0, False, "11")
        Kopfzelle_Format(app, "L5", 0, False, "12")
        Kopfzelle_Format(app, "M5", 0, False, "13")
        Kopfzelle_Format(app, "N5", 0, False, "14")
        Kopfzelle_Format(app, "O5", 0, False, "15")
        Kopfzelle_Format(app, "P5", 0, False, "16")
        Kopfzelle_Format(app, "Q5", 0, False, "17")
        Kopfzelle_Format(app, "R5", 0, False, "18")
        Kopfzelle_Format(app, "S5", 0, False, "19")
        Kopfzelle_Format(app, "T5", 0, False, "20")
        Kopfzelle_Format(app, "U5", 0, False, "21")
        Kopfzelle_Format(app, "V5", 0, False, "22")
        Kopfzelle_Format(app, "W5", 0, False, "23")
        Kopfzelle_Format(app, "X5", 0, False, "24")
        Kopfzelle_Format(app, "Y5", 0, False, "25")
        Kopfzelle_Format(app, "Z5", 0, False, "26")
        Kopfzelle_Format(app, "AA5", 0, False, "27")
        Kopfzelle_Format(app, "AB5", 0, False, "28")
        Kopfzelle_Format(app, "Ac5", 0, False, "29")
        Kopfzelle_Format(app, "AD5", 0, False, "30")
        Kopfzelle_Format(app, "AE5", 0, False, "31")
        Kopfzelle_Format(app, "AF5", 0, False, "32")
        Kopfzelle_Format(app, "AG5", 0, False, "33")
        Kopfzelle_Format(app, "AH5", 0, False, "34")
        Kopfzelle_Format(app, "AI5", 0, False, "35")
        Kopfzelle_Format(app, "AJ5", 0, False, "36")
        'Hauptüberschriften erzeugen
        Kopfzelle_Format(app, "A1:B1", 0, True, "Mastnummer")
        Kopfzelle_Format(app, "D1:E1", 0, True, "Traversen")
        Kopfzelle_Format(app, "F1:AC1", 0, True, "Ketten und Phasen vorne (V)", "Ketten und Phasen hinten (Abspanner) (H)")
        Kopfzelle_Format(app, "AD1:AG1", 0, True, "Verdrillung bzw. Auftrennung")
        Kopfzelle_Format(app, "AH1:AH3", 0, True, "Ver-", "binder")
        Kopfzelle_Format(app, "AI1:AI3", 0, True, "Flug-", "warn-", "kugeln")
        Kopfzelle_Format(app, "AJ1:AJ3", 0, True, "Bemerkung")
        'Attribute und Unterteilungen erzeugen
        Kopfzelle_Format(app, "A2:A3", 0, True, "Baunr.")
        Kopfzelle_Format(app, "B2:B3", 0, True, "Betriebs-", "nummer")
        Kopfzelle_Format(app, "C1:C3", 0, True, "anlagen-", "buch-", "halter-", "ische", "Zuge-", "hörigkeit")
        Kopfzelle_Format(app, "D2:D3", 0, True, "Traverse")
        Kopfzelle_Format(app, "E2:E3", 0, True, "Auf-", "hänge-", "punkte")
        Kopfzelle_Format(app, "F2:J2", 0, True, "Ketten")
        Kopfzelle_Format(app, "F3", 0, False, "Ketten-", "typ")
        'Zeilenbreitenkopf
        app.Rows("1:1").RowHeight = 29.25
        app.Rows("2:2").RowHeight = 29.25
        app.Rows("3:3").RowHeight = 101
        'Spaltenbeiten
        app.Columns("A:A").ColumnWidth = 20
        app.Columns("B:B").ColumnWidth = 10.71
        app.Columns("C:C").ColumnWidth = 9.0
        app.Columns("D:D").ColumnWidth = 15
        app.Columns("E:E").ColumnWidth = 7.0
        app.Columns("F:F").ColumnWidth = 6.0
        app.Columns("G:G").ColumnWidth = 10.71
        app.Columns("H:H").ColumnWidth = 7.0
        app.Columns("I:I").ColumnWidth = 22
        app.Columns("J:J").ColumnWidth = 8
        app.Columns("K:K").ColumnWidth = 15
        app.Columns("L:L").ColumnWidth = 7
        app.Columns("M:M").ColumnWidth = 22
        app.Columns("N:N").ColumnWidth = 22
        app.Columns("O:O").ColumnWidth = 6.0
        app.Columns("P:P").ColumnWidth = 22
        app.Columns("Q:Q").ColumnWidth = 22
        app.Columns("R:R").ColumnWidth = 6.0
        app.Columns("S:S").ColumnWidth = 22
        app.Columns("T:T").ColumnWidth = 10.71
        app.Columns("U:U").ColumnWidth = 8
        app.Columns("V:V").ColumnWidth = 22
        app.Columns("W:W").ColumnWidth = 6.0
        app.Columns("X:X").ColumnWidth = 6.0
        app.Columns("Y:Y").ColumnWidth = 7
        app.Columns("Z:Z").ColumnWidth = 6.0
        app.Columns("AA:AA").ColumnWidth = 6.0
        app.Columns("AB:AB").ColumnWidth = 7.0
        app.Columns("AC:AC").ColumnWidth = 7.0
        app.Columns("AD:AD").ColumnWidth = 10.71
        app.Columns("AE:AE").ColumnWidth = 6.0
        app.Columns("AF:AF").ColumnWidth = 10.71
        app.Columns("AG:AG").ColumnWidth = 6.0
        app.Columns("AH:AH").ColumnWidth = 6.0
        app.Columns("AI:AI").ColumnWidth = 7
        app.Columns("AJ:AJ").ColumnWidth = 33
    End Sub

    Public Sub Kopfdaten_STROMKREISVERFOLGUNG(ByRef app As Excel.Application,
                                                     ByRef wBook As Excel.Workbook,
                                                     ByRef wSheet As Excel.Worksheet)
        'Standardformat für alle Zellen
        Standardformat(app, "A1:S4", Excel.Constants.xlCenter, "Arial", "Fett", 10)
        'Einzelne Attribute
        Kopfzelle_Format(app, "A2", 0, False, "Stromkreis")
        Kopfzelle_Format(app, "B2", 0, False, "Stromkreisname")
        Kopfzelle_Format(app, "C2", 0, False, "Spann-", "ungs-", "ebene")
        Kopfzelle_Format(app, "D2", 0, False, "Netz-", "art")
        Kopfzelle_Format(app, "E2", 0, False, "Bahnstrom-", "leitung")
        Kopfzelle_Format(app, "F2", 0, False, "Bahnstromleitungsname")
        Kopfzelle_Format(app, "G2", 0, False, "Startmast")
        Kopfzelle_Format(app, "H2", 0, False, "Zielmast")
        Kopfzelle_Format(app, "I2", 0, False, "Verdriller")
        Kopfzelle_Format(app, "J2", 0, False, "An-", "zahl", "Ab-", "spanner")
        Kopfzelle_Format(app, "K2", 0, False, "An-", "zahl", "Trag-", "maste")
        Kopfzelle_Format(app, "L2", 0, False, "Betriebs-", "kennzahl")
        Kopfzelle_Format(app, "M2", 0, False, "DBEn-", "Standort")
        Kopfzelle_Format(app, "N2", 0, False, "Einfach-", "Seil")
        Kopfzelle_Format(app, "O2", 0, False, "2-Bdl.")
        Kopfzelle_Format(app, "P2", 0, False, "4-Bdl.")
        Kopfzelle_Format(app, "Q2", 0, False, "Leiter-", "end-", "tem-", "peratur")
        Kopfzelle_Format(app, "R2", 0, False, "Kurz-", "schluss-", "strom")
        'Einheiten und Nummerierung
        Kopfzelle_Format(app, "C3", 0, False, "[kV]")
        Kopfzelle_Format(app, "N3", 0, False, "[km]")
        Kopfzelle_Format(app, "O3", 0, False, "[km]")
        Kopfzelle_Format(app, "P3", 0, False, "[km]")
        Kopfzelle_Format(app, "Q3", 0, False, "[°C]")
        Kopfzelle_Format(app, "R3", 0, False, "[A]")
        Kopfzelle_Format(app, "A4", 0, False, "1")
        Kopfzelle_Format(app, "B4", 0, False, "2")
        Kopfzelle_Format(app, "C4", 0, False, "3")
        Kopfzelle_Format(app, "D4", 0, False, "4")
        Kopfzelle_Format(app, "E4", 0, False, "5")
        Kopfzelle_Format(app, "F4", 0, False, "6")
        Kopfzelle_Format(app, "G4", 0, False, "7")
        Kopfzelle_Format(app, "H4", 0, False, "8")
        Kopfzelle_Format(app, "I4", 0, False, "9")
        Kopfzelle_Format(app, "J4", 0, False, "10")
        Kopfzelle_Format(app, "K4", 0, False, "11")
        Kopfzelle_Format(app, "L4", 0, False, "12")
        Kopfzelle_Format(app, "M4", 0, False, "13")
        Kopfzelle_Format(app, "N4", 0, False, "14")
        Kopfzelle_Format(app, "O4", 0, False, "15")
        Kopfzelle_Format(app, "P4", 0, False, "16")
        Kopfzelle_Format(app, "Q4", 0, False, "17")
        Kopfzelle_Format(app, "R4", 0, False, "18")
        Kopfzelle_Format(app, "S4", 0, False, "19")
        'Hauptüberschriften erzeugen
        Kopfzelle_Format(app, "A1:D1", 0, True, "Stromkreise")
        Kopfzelle_Format(app, "E1:K1", 0, True, "Verlaufsinformationen")
        Kopfzelle_Format(app, "L1:M1", 0, True, "Region")
        Kopfzelle_Format(app, "N1:P1", 0, True, "Stromkreislänge")
        Kopfzelle_Format(app, "Q1:R1", 0, True, "elektrontechn.", "Parameter")
        Kopfzelle_Format(app, "S1:S2", 0, True, "Bemerkung")
        app.Rows("1:1").RowHeight = 29.25
        app.Rows("2:2").RowHeight = 92.25
        'Spaltenbeiten
        app.Columns("A:A").ColumnWidth = 30
        app.Columns("B:B").ColumnWidth = 40
        app.Columns("C:C").ColumnWidth = 7
        app.Columns("D:D").ColumnWidth = 10.71
        app.Columns("E:E").ColumnWidth = 10.71
        app.Columns("F:F").ColumnWidth = 34
        app.Columns("G:G").ColumnWidth = 12
        app.Columns("H:H").ColumnWidth = 12
        app.Columns("I:I").ColumnWidth = 10.71
        app.Columns("J:J").ColumnWidth = 8
        app.Columns("K:K").ColumnWidth = 8
        app.Columns("L:L").ColumnWidth = 10.71
        app.Columns("M:M").ColumnWidth = 10.71
        app.Columns("N:N").ColumnWidth = 9
        app.Columns("O:O").ColumnWidth = 9
        app.Columns("P:P").ColumnWidth = 9
        app.Columns("Q:Q").ColumnWidth = 8
        app.Columns("R:R").ColumnWidth = 8
        app.Columns("S:S").ColumnWidth = 40
    End Sub

    Public Sub Kopfdaten_VERZ_DER_MASTE(ByRef app As Excel.Application,
                                               ByRef wBook As Excel.Workbook,
                                               ByRef wSheet As Excel.Worksheet)
        'Standardformat für alle Zellen
        Standardformat(app, "A1:AJ5", Excel.Constants.xlCenter, "Arial", "Fett", 10)
        'Einzelne Attribute
        Kopfzelle_Format(app, "K2", 0, False, "Mast-", "höhe")
        Kopfzelle_Format(app, "L2", 0, False, "Mast-", "breite")
        Kopfzelle_Format(app, "K3", 0, False, "EOK bis", "Erdseil-", "spitze")
        Kopfzelle_Format(app, "L3", 0, False, "bE bei ", "EOK ", "und", "SF - ", "Ecke", " = 0")
        Kopfzelle_Format(app, "M2", 0, False, "Mast-", "gewicht")
        Kopfzelle_Format(app, "M3", 0, False, "Mast-", "gewicht", "EOK bis", "Mast-", "spitze")
        Kopfzelle_Format(app, "O3", 0, False, "A")
        Kopfzelle_Format(app, "P3", 0, False, "B")
        Kopfzelle_Format(app, "Q3", 0, False, "C")
        Kopfzelle_Format(app, "R3", 0, False, "D")
        Kopfzelle_Format(app, "S3", 0, False, "Mastoberteil")
        Kopfzelle_Format(app, "T3", 0, False, "Mastschaftunterteil")
        Kopfzelle_Format(app, "U3", 0, False, "ES-Stütze")
        Kopfzelle_Format(app, "V3", 0, False, "Traverse 1")
        Kopfzelle_Format(app, "W3", 0, False, "Traverse 2")
        Kopfzelle_Format(app, "X3", 0, False, "Schuss 1")
        Kopfzelle_Format(app, "Y3", 0, False, "Schuss 2")
        Kopfzelle_Format(app, "Z3", 0, False, "Schuss 3")
        Kopfzelle_Format(app, "AA3", 0, False, "Schuss 4")
        Kopfzelle_Format(app, "AB3", 0, False, "Schuss 5")
        Kopfzelle_Format(app, "AC3", 0, False, "Schuss 6")
        Kopfzelle_Format(app, "AD3", 0, False, "SF-Ecken")
        'Einheiten und Spaltennummerierung
        Kopfzelle_Format(app, "G4", 0, False, "alt/neu")
        Kopfzelle_Format(app, "H4", 0, False, "[m]")
        Kopfzelle_Format(app, "I4", 0, False, "[gon]")
        Kopfzelle_Format(app, "K4", 0, False, "[m]")
        Kopfzelle_Format(app, "L4", 0, False, "[m]")
        Kopfzelle_Format(app, "M4", 0, False, "[t]")
        Kopfzelle_Format(app, "N4", 0, False, "[m²]")
        Kopfzelle_Format(app, "O4", 0, False, "[m]")
        Kopfzelle_Format(app, "P4", 0, False, "[m]")
        Kopfzelle_Format(app, "Q4", 0, False, "[m]")
        Kopfzelle_Format(app, "R4", 0, False, "[m]")
        Kopfzelle_Format(app, "A5", 0, False, "1")
        Kopfzelle_Format(app, "B5", 0, False, "2")
        Kopfzelle_Format(app, "C5", 0, False, "3")
        Kopfzelle_Format(app, "D5", 0, False, "4")
        Kopfzelle_Format(app, "E5", 0, False, "5")
        Kopfzelle_Format(app, "F5", 0, False, "6")
        Kopfzelle_Format(app, "G5", 0, False, "7")
        Kopfzelle_Format(app, "H5", 0, False, "8")
        Kopfzelle_Format(app, "I5", 0, False, "9")
        Kopfzelle_Format(app, "J5", 0, False, "10")
        Kopfzelle_Format(app, "K5", 0, False, "11")
        Kopfzelle_Format(app, "L5", 0, False, "12")
        Kopfzelle_Format(app, "M5", 0, False, "13")
        Kopfzelle_Format(app, "N5", 0, False, "14")
        Kopfzelle_Format(app, "O5", 0, False, "15")
        Kopfzelle_Format(app, "P5", 0, False, "16")
        Kopfzelle_Format(app, "Q5", 0, False, "17")
        Kopfzelle_Format(app, "R5", 0, False, "18")
        Kopfzelle_Format(app, "S5", 0, False, "19")
        Kopfzelle_Format(app, "T5", 0, False, "20")
        Kopfzelle_Format(app, "U5", 0, False, "21")
        Kopfzelle_Format(app, "V5", 0, False, "22")
        Kopfzelle_Format(app, "W5", 0, False, "23")
        Kopfzelle_Format(app, "X5", 0, False, "24")
        Kopfzelle_Format(app, "Y5", 0, False, "25")
        Kopfzelle_Format(app, "Z5", 0, False, "26")
        Kopfzelle_Format(app, "AA5", 0, False, "27")
        Kopfzelle_Format(app, "AB5", 0, False, "28")
        Kopfzelle_Format(app, "AC5", 0, False, "29")
        Kopfzelle_Format(app, "AD5", 0, False, "30")
        Kopfzelle_Format(app, "AE5", 0, False, "31")
        Kopfzelle_Format(app, "AF5", 0, False, "32")
        'Hauptüberschriften erzeugen
        Kopfzelle_Format(app, "A1:B1", 0, True, "Mastnummer")
        Kopfzelle_Format(app, "D1:AE1", 0, True, "Mast")
        Kopfzelle_Format(app, "AF1:AF3", 0, True, "Bemerkung")
        Kopfzelle_Format(app, "A2:A3", 0, True, "Baunummer")
        'Attribute und Unterteilungen erzeugen
        Kopfzelle_Format(app, "B2:B3", 0, True, "Betriebs-", "nummer")
        Kopfzelle_Format(app, "C1:C3", 0, True, "nächster", " Mast")
        Kopfzelle_Format(app, "D2:D3", 0, True, "Bau-", "jahr")
        Kopfzelle_Format(app, "E2:E3", 0, True, "Windlastzone")
        Kopfzelle_Format(app, "F2:F3", 0, True, "Eislastzone")
        Kopfzelle_Format(app, "G2:G3", 0, True, "Eis-", "last-", "formel")
        Kopfzelle_Format(app, "H2:H3", 0, True, "Spann-", "weite")
        Kopfzelle_Format(app, "I2:I3", 0, True, "Leitungs-", "winkel")
        Kopfzelle_Format(app, "J2:J3", 0, True, "Masttyp")
        Kopfzelle_Format(app, "N2:N3", 0, False, "Anstrichs-", "fläche")

        Kopfzelle_Format(app, "O2:R2", 0, True, "Schrägfußecken")
        Kopfzelle_Format(app, "S2:T2", 0, True, "Zeichnungsnummer Systemzeichnungen")
        Kopfzelle_Format(app, "U2:AD2", 0, True, "Zeichnungsnummer Werkstattzeichnungen")
        Kopfzelle_Format(app, "AE2:AE3", 0, True, "Stahllieferant")
        app.Rows("1:1").RowHeight = 29.25
        app.Rows("2:2").RowHeight = 29.25
        app.Rows("3:3").RowHeight = 92.25
        'Spaltenbeiten
        app.Columns("A:A").ColumnWidth = 20
        app.Columns("B:B").ColumnWidth = 10.71
        app.Columns("C:C").ColumnWidth = 10.71
        app.Columns("D:D").ColumnWidth = 6
        app.Columns("E:E").ColumnWidth = 28
        app.Columns("F:F").ColumnWidth = 28
        app.Columns("G:G").ColumnWidth = 7
        app.Columns("H:H").ColumnWidth = 8
        app.Columns("I:I").ColumnWidth = 8
        app.Columns("J:J").ColumnWidth = 28
        app.Columns("K:K").ColumnWidth = 7
        app.Columns("L:L").ColumnWidth = 7
        app.Columns("M:M").ColumnWidth = 8
        app.Columns("M:M").ColumnWidth = 10
        app.Columns("O:O").ColumnWidth = 6
        app.Columns("P:P").ColumnWidth = 6
        app.Columns("Q:Q").ColumnWidth = 6
        app.Columns("R:R").ColumnWidth = 6
        app.Columns("S:S").ColumnWidth = 28
        app.Columns("T:T").ColumnWidth = 28
        app.Columns("U:U").ColumnWidth = 28
        app.Columns("V:V").ColumnWidth = 18
        app.Columns("W:W").ColumnWidth = 18
        app.Columns("X:X").ColumnWidth = 18
        app.Columns("Y:Y").ColumnWidth = 18
        app.Columns("Z:Z").ColumnWidth = 18
        app.Columns("AA:AA").ColumnWidth = 18
        app.Columns("AB:AB").ColumnWidth = 18
        app.Columns("AC:AC").ColumnWidth = 18
        app.Columns("AD:AD").ColumnWidth = 18
        app.Columns("AE:AE").ColumnWidth = 18
        app.Columns("AF:AF").ColumnWidth = 33

    End Sub

    Public Sub Kopfdaten_VERZ_AUSFUEHRUNGSPLANUNG(ByRef app As Excel.Application,
                                                         ByRef wBook As Excel.Workbook,
                                                         ByRef wSheet As Excel.Worksheet)
        'Standardformat für alle Zellen
        Standardformat(app, "A1:AC5", Excel.Constants.xlCenter, "Arial", "Fett", 10)
        'Einzelne Attribute
        Kopfzelle_Format(app, "L3", 0, False, "Mastoberteil")
        Kopfzelle_Format(app, "M3", 0, False, "Mastunterteil")
        Kopfzelle_Format(app, "N3", 0, False, "EOK bis", "Traversen-", "unterkante", "(unterste", "Traverse)")
        Kopfzelle_Format(app, "O3", 0, False, "EOK", " bis", "Erdseil-", "spitze")
        Kopfzelle_Format(app, "P3", 0, False, "bE bei", "EOK", "und", "SF -", "Ecke ", "= 0")
        Kopfzelle_Format(app, "Q3", 0, False, "bei", "EOK", "ein-", "schließl.", "Fund-", "ament")
        Kopfzelle_Format(app, "R3", 0, False, "A")
        Kopfzelle_Format(app, "S3", 0, False, "B")
        Kopfzelle_Format(app, "T3", 0, False, "C")
        Kopfzelle_Format(app, "U3", 0, False, "D")
        'Einheiten und Spaltennummerierung
        Kopfzelle_Format(app, "D4", 0, False, "[m]")
        Kopfzelle_Format(app, "E4", 0, False, "[m]")
        Kopfzelle_Format(app, "F4", 0, False, "[gon]")
        Kopfzelle_Format(app, "G4", 0, False, "[°]")
        Kopfzelle_Format(app, "N4", 0, False, "[m]")
        Kopfzelle_Format(app, "O4", 0, False, "[m]")
        Kopfzelle_Format(app, "P4", 0, False, "[m]")
        Kopfzelle_Format(app, "Q4", 0, False, "[m]")
        Kopfzelle_Format(app, "R4", 0, False, "[m]")
        Kopfzelle_Format(app, "S4", 0, False, "[m]")
        Kopfzelle_Format(app, "T4", 0, False, "[m]")
        Kopfzelle_Format(app, "U4", 0, False, "[m]")
        Kopfzelle_Format(app, "X4", 0, False, "[m]")
        Kopfzelle_Format(app, "AA4", 0, False, "[m]")
        Kopfzelle_Format(app, "AB4", 0, False, "[m]")
        Kopfzelle_Format(app, "A5", 0, False, "1")
        Kopfzelle_Format(app, "B5", 0, False, "2")
        Kopfzelle_Format(app, "C5", 0, False, "3")
        Kopfzelle_Format(app, "D5", 0, False, "4")
        Kopfzelle_Format(app, "E5", 0, False, "5")
        Kopfzelle_Format(app, "F5", 0, False, "6")
        Kopfzelle_Format(app, "G5", 0, False, "7")
        Kopfzelle_Format(app, "H5", 0, False, "8")
        Kopfzelle_Format(app, "I5", 0, False, "9")
        Kopfzelle_Format(app, "J5", 0, False, "10")
        Kopfzelle_Format(app, "K5", 0, False, "11")
        Kopfzelle_Format(app, "L5", 0, False, "12")
        Kopfzelle_Format(app, "M5", 0, False, "13")
        Kopfzelle_Format(app, "N5", 0, False, "14")
        Kopfzelle_Format(app, "O5", 0, False, "15")
        Kopfzelle_Format(app, "P5", 0, False, "16")
        Kopfzelle_Format(app, "Q5", 0, False, "17")
        Kopfzelle_Format(app, "R5", 0, False, "18")
        Kopfzelle_Format(app, "S5", 0, False, "19")
        Kopfzelle_Format(app, "T5", 0, False, "20")
        Kopfzelle_Format(app, "U5", 0, False, "21")
        Kopfzelle_Format(app, "V5", 0, False, "22")
        Kopfzelle_Format(app, "W5", 0, False, "23")
        Kopfzelle_Format(app, "X5", 0, False, "24")
        Kopfzelle_Format(app, "Y5", 0, False, "25")
        Kopfzelle_Format(app, "Z5", 0, False, "26")
        Kopfzelle_Format(app, "AA5", 0, False, "27")
        Kopfzelle_Format(app, "AB5", 0, False, "28")
        Kopfzelle_Format(app, "AC5", 0, False, "29")
        'Hauptüberschriften erzeugen
        Kopfzelle_Format(app, "A1:C1", 0, True, "Mastnummer")
        Kopfzelle_Format(app, "D1:G1", 0, True, "Leitung")
        Kopfzelle_Format(app, "H1:J1", 0, True, "Maststandort")
        Kopfzelle_Format(app, "K1:U1", 0, True, "Mast")
        Kopfzelle_Format(app, "V1:X1", 0, True, "Isolation")
        Kopfzelle_Format(app, "Y1:AB1", 0, True, "Gründung")
        Kopfzelle_Format(app, "AC1:AC3", 0, True, "Bemerkung")
        'Attribute und Unterteilungen erzeugen
        Kopfzelle_Format(app, "A2:A3", 0, True, "Baunummer")
        Kopfzelle_Format(app, "B2:B3", 0, True, "Betriebs-", "nummer")
        Kopfzelle_Format(app, "C2:C3", 0, True, "nächster", "Mast")
        Kopfzelle_Format(app, "D2:D3", 0, True, "Spann-", "weite")
        Kopfzelle_Format(app, "E2:E3", 0, True, "Abspann-", "abschnitts-", "länge")
        Kopfzelle_Format(app, "F2:G3", 0, True, "Leitungswinkel")
        Kopfzelle_Format(app, "H2:H3", 0, True, "Gemarkung")
        Kopfzelle_Format(app, "I2:I3", 0, True, "Flur")
        Kopfzelle_Format(app, "J2:J3", 0, True, "Flurstücksnummer")
        Kopfzelle_Format(app, "K2:K3", 0, True, "Masttyp")
        Kopfzelle_Format(app, "L2:M2", 0, True, "Zeichnungsnummer", "Systemzeichnung")
        Kopfzelle_Format(app, "N2:O2", 0, True, "Masthöhe")
        Kopfzelle_Format(app, "P2:Q2", 0, True, "Mastbreite")
        Kopfzelle_Format(app, "R2:U2", 0, True, "Schrägfußecken")
        Kopfzelle_Format(app, "V2:V3", 0, True, "Art")
        Kopfzelle_Format(app, "W2:W3", 0, True, "Zeichnungsnummer")
        Kopfzelle_Format(app, "X2:X3", 0, True, "Schwingen-", "höhe")
        Kopfzelle_Format(app, "Y2:Y3", 0, True, "Art")
        Kopfzelle_Format(app, "Z2:Z3", 0, True, "Zeichnungsnummer")
        Kopfzelle_Format(app, "AA2:AA3", 0, True, "Eingrabtiefe /", "Pfahllänge")
        Kopfzelle_Format(app, "AB2:AB3", 0, True, "Grund-", "wasser-", "höhe", "unter", "EOK")
        app.Rows("1:1").RowHeight = 29.25
        app.Rows("2:2").RowHeight = 29.25
        app.Rows("3:3").RowHeight = 110
        app.Columns("A:A").ColumnWidth = 20
        app.Columns("B:B").ColumnWidth = 10.71
        app.Columns("C:C").ColumnWidth = 10.71
        app.Columns("D:D").ColumnWidth = 10.71
        app.Columns("E:E").ColumnWidth = 10.71
        app.Columns("F:F").ColumnWidth = 7
        app.Columns("G:G").ColumnWidth = 7
        app.Columns("H:H").ColumnWidth = 44
        app.Columns("I:I").ColumnWidth = 10.71
        app.Columns("J:J").ColumnWidth = 20
        app.Columns("K:K").ColumnWidth = 30
        app.Columns("L:L").ColumnWidth = 20
        app.Columns("M:M").ColumnWidth = 20
        app.Columns("N:N").ColumnWidth = 10
        app.Columns("O:O").ColumnWidth = 6
        app.Columns("P:P").ColumnWidth = 6
        app.Columns("Q:Q").ColumnWidth = 8
        app.Columns("R:R").ColumnWidth = 6
        app.Columns("S:S").ColumnWidth = 6
        app.Columns("T:T").ColumnWidth = 6
        app.Columns("U:U").ColumnWidth = 6
        app.Columns("V:V").ColumnWidth = 6
        app.Columns("W:W").ColumnWidth = 20
        app.Columns("X:X").ColumnWidth = 11
        app.Columns("Y:Y").ColumnWidth = 19
        app.Columns("Z:Z").ColumnWidth = 20
        app.Columns("AA:AA").ColumnWidth = 20
        app.Columns("AB:AB").ColumnWidth = 7
        app.Columns("AC:AC").ColumnWidth = 33.0
    End Sub

    Public Sub Kopfdaten_VERZ_GRUENDUNGEN(ByRef app As Excel.Application,
                                                 ByRef wBook As Excel.Workbook,
                                                 ByRef wSheet As Excel.Worksheet)
        'Standardformat für alle Zellen
        Standardformat(app, "A1:V5", Excel.Constants.xlCenter, "Arial", "Fett", 10)
        'Einzelne Attribute
        Kopfzelle_Format(app, "J3", 0, False, "laut ", "Bau-", "grund-", "gut-", "achten")
        Kopfzelle_Format(app, "K3", 0, False, "bei ", "Berech-", "nung", "zu-", "grunde", " gelegt")
        Kopfzelle_Format(app, "L3", 0, False, "bei Aus-", "führung", "berück-", "sichtigt")
        Kopfzelle_Format(app, "M3", 0, False, "Beton-", "volumen", " ohne", "Sauber-", "keits-", "schicht")
        Kopfzelle_Format(app, "N3", 0, False, "Festig-", "keits-", "klasse")
        Kopfzelle_Format(app, "O3", 0, False, "besondere", "Eigen-", "schaften")
        Kopfzelle_Format(app, "P3", 0, False, "besondere", "Maß-", "nahmen")
        'Einheiten und Spaltennummerierung
        Kopfzelle_Format(app, "F4", 0, False, "[m]")
        Kopfzelle_Format(app, "G4", 0, False, "[m]")
        Kopfzelle_Format(app, "H4", 0, False, "[m]")
        Kopfzelle_Format(app, "J4", 0, False, "[m]")
        Kopfzelle_Format(app, "K4", 0, False, "[m]")
        Kopfzelle_Format(app, "L4", 0, False, "[m]")
        Kopfzelle_Format(app, "M4", 0, False, "[m³]")
        Kopfzelle_Format(app, "A5", 0, False, "1")
        Kopfzelle_Format(app, "B5", 0, False, "2")
        Kopfzelle_Format(app, "C5", 0, False, "3")
        Kopfzelle_Format(app, "D5", 0, False, "4")
        Kopfzelle_Format(app, "E5", 0, False, "5")
        Kopfzelle_Format(app, "F5", 0, False, "6")
        Kopfzelle_Format(app, "G5", 0, False, "7")
        Kopfzelle_Format(app, "H5", 0, False, "8")
        Kopfzelle_Format(app, "I5", 0, False, "9")
        Kopfzelle_Format(app, "J5", 0, False, "10")
        Kopfzelle_Format(app, "K5", 0, False, "11")
        Kopfzelle_Format(app, "L5", 0, False, "12")
        Kopfzelle_Format(app, "M5", 0, False, "13")
        Kopfzelle_Format(app, "N5", 0, False, "14")
        Kopfzelle_Format(app, "O5", 0, False, "15")
        Kopfzelle_Format(app, "P5", 0, False, "16")
        Kopfzelle_Format(app, "Q5", 0, False, "17")
        Kopfzelle_Format(app, "R5", 0, False, "18")
        Kopfzelle_Format(app, "S5", 0, False, "19")
        Kopfzelle_Format(app, "T5", 0, False, "20")
        Kopfzelle_Format(app, "U5", 0, False, "21")
        Kopfzelle_Format(app, "V5", 0, False, "22")
        'Hauptüberschriften erzeugen
        Kopfzelle_Format(app, "A1:B1", 0, True, "Mastnummer")
        Kopfzelle_Format(app, "C1:Q1", 0, True, "Gründung")
        Kopfzelle_Format(app, "R1", 0, False, "Erdung")
        Kopfzelle_Format(app, "S1:U1", 0, True, "Fundamentertüchtigung")
        Kopfzelle_Format(app, "V1:V3", 0, True, "Bemerkung")
        'Attribute und Unterteilungen erzeugen
        Kopfzelle_Format(app, "A2:A3", 0, True, "Baunummer")
        Kopfzelle_Format(app, "B2:B3", 0, True, "Betriebs-", "nummer")
        Kopfzelle_Format(app, "C2:C3", 0, True, "Bau-", "jahr")
        Kopfzelle_Format(app, "D2:D3", 0, True, "Art")
        Kopfzelle_Format(app, "E2:E3", 0, True, "Zeichnungs-", "nummer")
        Kopfzelle_Format(app, "F2:F3", 0, True, "Mast-", "breite", " bei EOK", "einschl.", "Funda-", "ment")
        Kopfzelle_Format(app, "G2:G3", 0, True, "Breite", " an", "Funda-", "ment-", "sohle")
        Kopfzelle_Format(app, "H2:H3", 0, True, "Eingrabtiefe /", "Pfahllänge")
        Kopfzelle_Format(app, "I2:I3", 0, True, "Bodenart")
        Kopfzelle_Format(app, "J2:L2", 0, True, "Grundwasserhöhe", "unter EOK")
        Kopfzelle_Format(app, "M2:P2", 0, True, "Beton")
        Kopfzelle_Format(app, "Q2:Q3", 0, True, "aus-", "führende", "Firma")
        Kopfzelle_Format(app, "R2:R3", 0, True, "Erdungsart")
        Kopfzelle_Format(app, "S2:S3", 0, True, "Jahr")
        Kopfzelle_Format(app, "T2:T3", 0, True, "ausführende Firma")
        Kopfzelle_Format(app, "U2:U3", 0, True, "Art")
        app.Rows("1:1").RowHeight = 29.25
        app.Rows("2:2").RowHeight = 29.25
        app.Rows("3:3").RowHeight = 103.5
        app.Columns("A:A").ColumnWidth = 20
        app.Columns("B:B").ColumnWidth = 10.71
        app.Columns("C:C").ColumnWidth = 6
        app.Columns("D:D").ColumnWidth = 20
        app.Columns("E:E").ColumnWidth = 10.71
        app.Columns("F:F").ColumnWidth = 8
        app.Columns("G:G").ColumnWidth = 7
        app.Columns("H:H").ColumnWidth = 20
        app.Columns("I:I").ColumnWidth = 20
        app.Columns("J:J").ColumnWidth = 7
        app.Columns("K:K").ColumnWidth = 8
        app.Columns("L:L").ColumnWidth = 8
        app.Columns("M:M").ColumnWidth = 8
        app.Columns("N:N").ColumnWidth = 10.71
        app.Columns("O:O").ColumnWidth = 10.71
        app.Columns("P:P").ColumnWidth = 10.71
        app.Columns("Q:Q").ColumnWidth = 10.71
        app.Columns("R:R").ColumnWidth = 20
        app.Columns("S:S").ColumnWidth = 6
        app.Columns("T:T").ColumnWidth = 20
        app.Columns("U:U").ColumnWidth = 20
        app.Columns("V:V").ColumnWidth = 33
    End Sub

    ''' <summary>
    ''' Bodenrichtwert ist kommentiert da hierfür noch kein Attribut im Grundbuch angelegt ist und der Export
    ''' der Griddaten sonst an dieser Stelle die nachfolgenden Bemerkungen einträgt
    ''' </summary>
    ''' <param name="app"></param>
    ''' <param name="wBook"></param>
    ''' <param name="wSheet"></param>
    ''' <remarks></remarks>
    Public Sub Kopfdaten_VERZ_RECHTLICHE_SICHERUNG(ByRef app As Excel.Application,
                                                          ByRef wBook As Excel.Workbook,
                                                          ByRef wSheet As Excel.Worksheet)
        'Standardformat für alle Zellen
        'Standardformat(app, "A1:AM4", Excel.Constants.xlCenter, "Arial", "Fett", 10)
        Standardformat(app, "A1:AL4", Excel.Constants.xlCenter, "Arial", "Fett", 10)
        'Einzelne Attribute
        Kopfzelle_Format(app, "A2", 0, False, "Mast-", "nummer")
        Kopfzelle_Format(app, "B2", 0, False, "Mast-", "nummer")
        Kopfzelle_Format(app, "C2", 0, False, "Bundesland")
        Kopfzelle_Format(app, "D2", 0, False, "Gemeinde")
        Kopfzelle_Format(app, "E2", 0, False, "Gemarkung")
        Kopfzelle_Format(app, "F2", 0, False, "Flur")
        Kopfzelle_Format(app, "G2", 0, False, "Zähler")
        Kopfzelle_Format(app, "H2", 0, False, "Nenner")
        Kopfzelle_Format(app, "I2", 0, False, "Familienname")
        Kopfzelle_Format(app, "J2", 0, False, "Vorname")
        Kopfzelle_Format(app, "K2", 0, False, "Geburts-", "datum")
        Kopfzelle_Format(app, "L2", 0, False, "PLZ")
        Kopfzelle_Format(app, "M2", 0, False, "Ort")
        Kopfzelle_Format(app, "N2", 0, False, "Straße")
        Kopfzelle_Format(app, "O2", 0, False, "Hausnr.")
        Kopfzelle_Format(app, "P2", 0, False, "Amtsgericht")
        Kopfzelle_Format(app, "Q2", 0, False, "Katasteramt")
        Kopfzelle_Format(app, "R2", 0, False, "Grundbuch", "von")
        Kopfzelle_Format(app, "S2", 0, False, "Grundbuch", "Band")
        Kopfzelle_Format(app, "T2", 0, False, "Grundbuch-", "blatt")
        Kopfzelle_Format(app, "U2", 0, False, "Grundbuch", "Abt. II", "(lfd. Nr.)")
        Kopfzelle_Format(app, "V2", 0, False, "Nutzungsart")
        Kopfzelle_Format(app, "W2", 0, False, "Vertragsart")
        Kopfzelle_Format(app, "X2", 0, False, "Berechtigter")
        Kopfzelle_Format(app, "Y2", 0, False, "Datum", " der", "Bewilli-", "gung")
        Kopfzelle_Format(app, "Z2", 0, False, "Datum", " der", "Eintragung")
        Kopfzelle_Format(app, "AA2", 0, False, "Datum", " der", "Entschädi-", "gung")
        Kopfzelle_Format(app, "AB2", 0, False, "Flurstücks-", "fläche")
        Kopfzelle_Format(app, "AC2", 0, False, "Über-", "spannungs-", "fläche")
        Kopfzelle_Format(app, "AD2", 0, False, "Über-", "spannungs-", "fläche ", "über-", "lappend")
        Kopfzelle_Format(app, "AE2", 0, False, "Über-", "spannungs-", "entschädig-", "ung")
        Kopfzelle_Format(app, "AF2", 0, False, "Kanten-", "länge", "Funda-", "ment")
        Kopfzelle_Format(app, "AG2", 0, False, "Mast-", "austritts-", "fläche", " gesamt")
        Kopfzelle_Format(app, "AH2", 0, False, "Betroffen-", "heit durch", " Mast-", "austritts-", "fläche")
        Kopfzelle_Format(app, "AI2", 0, False, "Anteilige", " Mastent-", "schädigung")
        Kopfzelle_Format(app, "AJ2", 0, False, "Vorübergehend", "beanspruchte", "Fläche")
        Kopfzelle_Format(app, "AK2", 0, False, "LPB", "-Fläche")
        'Kopfzelle_Format(app, "AL2", 0, False, "Boden-", "richtwerte")
        'Einheiten und Spaltennummerierung
        Kopfzelle_Format(app, "AB3", 0, False, "[m²]")
        Kopfzelle_Format(app, "AC3", 0, False, "[m²]")
        Kopfzelle_Format(app, "AD3", 0, False, "[m²]")
        Kopfzelle_Format(app, "AE3", 0, False, "[Euro]")
        Kopfzelle_Format(app, "AF3", 0, False, "[m * m]")
        Kopfzelle_Format(app, "AG3", 0, False, "[m²]")
        Kopfzelle_Format(app, "AH3", 0, False, "[%]")
        Kopfzelle_Format(app, "AI3", 0, False, "[Euro]")
        Kopfzelle_Format(app, "AJ3", 0, False, "[m²]")
        Kopfzelle_Format(app, "AK3", 0, False, "[m²]")
        'Kopfzelle_Format(app, "AL3", 0, False, "[Euro/m²]")
        Kopfzelle_Format(app, "A4", 0, False, "1")
        Kopfzelle_Format(app, "B4", 0, False, "2")
        Kopfzelle_Format(app, "C4", 0, False, "3")
        Kopfzelle_Format(app, "D4", 0, False, "4")
        Kopfzelle_Format(app, "E4", 0, False, "5")
        Kopfzelle_Format(app, "F4", 0, False, "6")
        Kopfzelle_Format(app, "G4", 0, False, "7")
        Kopfzelle_Format(app, "H4", 0, False, "8")
        Kopfzelle_Format(app, "I4", 0, False, "9")
        Kopfzelle_Format(app, "J4", 0, False, "10")
        Kopfzelle_Format(app, "K4", 0, False, "11")
        Kopfzelle_Format(app, "L4", 0, False, "12")
        Kopfzelle_Format(app, "M4", 0, False, "13")
        Kopfzelle_Format(app, "N4", 0, False, "14")
        Kopfzelle_Format(app, "O4", 0, False, "15")
        Kopfzelle_Format(app, "P4", 0, False, "16")
        Kopfzelle_Format(app, "Q4", 0, False, "17")
        Kopfzelle_Format(app, "R4", 0, False, "18")
        Kopfzelle_Format(app, "S4", 0, False, "19")
        Kopfzelle_Format(app, "T4", 0, False, "20")
        Kopfzelle_Format(app, "U4", 0, False, "21")
        Kopfzelle_Format(app, "V4", 0, False, "22")
        Kopfzelle_Format(app, "W4", 0, False, "23")
        Kopfzelle_Format(app, "X4", 0, False, "24")
        Kopfzelle_Format(app, "Y4", 0, False, "25")
        Kopfzelle_Format(app, "Z4", 0, False, "26")
        Kopfzelle_Format(app, "AA4", 0, False, "27")
        Kopfzelle_Format(app, "AB4", 0, False, "28")
        Kopfzelle_Format(app, "AC4", 0, False, "29")
        Kopfzelle_Format(app, "AD4", 0, False, "30")
        Kopfzelle_Format(app, "AE4", 0, False, "31")
        Kopfzelle_Format(app, "AF4", 0, False, "32")
        Kopfzelle_Format(app, "AG4", 0, False, "33")
        Kopfzelle_Format(app, "AH4", 0, False, "34")
        Kopfzelle_Format(app, "AI4", 0, False, "35")
        Kopfzelle_Format(app, "AJ4", 0, False, "36")
        Kopfzelle_Format(app, "AK4", 0, False, "37")
        Kopfzelle_Format(app, "AL4", 0, False, "38")
        'Kopfzelle_Format(app, "AM4", 0, False, "39")
        'Hauptüberschriften erzeugen
        Kopfzelle_Format(app, "A1:B1", 0, True, "Mastnummer")
        Kopfzelle_Format(app, "C1:H1", 0, True, "Politische Grenzen")
        Kopfzelle_Format(app, "I1:O1", 0, True, "Personendaten")
        'Kopfzelle_Format(app, "P1:AL1", 0, True, "Grundbuch")
        Kopfzelle_Format(app, "P1:AK1", 0, True, "Grundbuch")
        'Kopfzelle_Format(app, "AM1:AM2", 0, True, "Bemerkung")
        Kopfzelle_Format(app, "AL1:AL2", 0, True, "Bemerkung")
        app.Rows("1:1").RowHeight = 29.25
        app.Rows("2:2").RowHeight = 92.25
        'Mastnummern
        app.Columns("A:A").ColumnWidth = 11
        app.Columns("B:B").ColumnWidth = 11
        'Flurstück
        app.Columns("C:C").ColumnWidth = 25
        app.Columns("D:D").ColumnWidth = 30
        app.Columns("E:E").ColumnWidth = 30
        app.Columns("F:F").ColumnWidth = 6
        app.Columns("G:G").ColumnWidth = 9
        app.Columns("H:H").ColumnWidth = 9
        'Eigentuemer
        app.Columns("I:I").ColumnWidth = 92
        app.Columns("J:J").ColumnWidth = 43
        app.Columns("K:K").ColumnWidth = 8
        app.Columns("L:L").ColumnWidth = 6
        app.Columns("M:M").ColumnWidth = 30
        app.Columns("N:N").ColumnWidth = 30
        app.Columns("O:O").ColumnWidth = 8
        'Grunbuch
        app.Columns("P:P").ColumnWidth = 26
        app.Columns("Q:Q").ColumnWidth = 51
        app.Columns("R:R").ColumnWidth = 10.71
        app.Columns("S:S").ColumnWidth = 10.71
        app.Columns("T:T").ColumnWidth = 10.71
        app.Columns("U:U").ColumnWidth = 10.71
        app.Columns("V:V").ColumnWidth = 64
        app.Columns("W:W").ColumnWidth = 34
        app.Columns("X:X").ColumnWidth = 12
        app.Columns("Y:Y").ColumnWidth = 10.71
        app.Columns("Z:Z").ColumnWidth = 10.71
        app.Columns("AA:AA").ColumnWidth = 10.71
        app.Columns("AB:AB").ColumnWidth = 8
        app.Columns("AC:AC").ColumnWidth = 10.71
        app.Columns("AD:AD").ColumnWidth = 10.71
        app.Columns("AE:AE").ColumnWidth = 10.71
        app.Columns("AF:AF").ColumnWidth = 8
        app.Columns("AG:AG").ColumnWidth = 8
        app.Columns("AH:AH").ColumnWidth = 10.71
        app.Columns("AI:AI").ColumnWidth = 10.71
        app.Columns("AJ:AJ").ColumnWidth = 8
        app.Columns("AK:AK").ColumnWidth = 8
        'app.Columns("AL:AL").ColumnWidth = 10.71
        app.Columns("AL:AL").ColumnWidth = 40
        'app.Columns("AM:AM").ColumnWidth = 40
    End Sub

    Public Sub Kopfdaten_KREUZUNGSVERZEICHNIS(ByRef app As Excel.Application,
                                                     ByRef wBook As Excel.Workbook,
                                                     ByRef wSheet As Excel.Worksheet)

        'Standardformat für alle Zellen
        Standardformat(app, "A1:L3", Excel.Constants.xlCenter, "Arial", "Fett", 10)
        'Einzelne Attribute
        Kopfzelle_Format(app, "A1", 0, False, "Leitungs-", "nummer")
        Kopfzelle_Format(app, "B1", 0, False, "Leitungsbezeichnung")
        Kopfzelle_Format(app, "C1", 0, False, "linke", "Mastbetriebs-", "nummer")
        Kopfzelle_Format(app, "D1", 0, False, "rechte", "Mastbetriebs-", "nummer")
        Kopfzelle_Format(app, "E1", 0, False, "Kreuzungs-", "objekt-", "nummer")
        Kopfzelle_Format(app, "F1", 0, False, "Objektbeschreibung")
        Kopfzelle_Format(app, "G1", 0, False, "Objekttyp", "lt. DIN VDE 0210")
        Kopfzelle_Format(app, "H1", 0, False, "Abstandstyp")
        Kopfzelle_Format(app, "I1", 0, False, "Eigentümer")
        Kopfzelle_Format(app, "J1", 0, False, "Kreuzungs-", "kilometer")
        Kopfzelle_Format(app, "K1", 0, False, "Netzknoten")
        Kopfzelle_Format(app, "L1", 0, False, "Strecken-", "nummer")
        'Einheiten und Spaltennummerierung
        Kopfzelle_Format(app, "J2", 0, False, "[km]")
        Kopfzelle_Format(app, "L2", 0, False, "[km]")
        Kopfzelle_Format(app, "A3", 0, False, "1")
        Kopfzelle_Format(app, "B3", 0, False, "2")
        Kopfzelle_Format(app, "C3", 0, False, "3")
        Kopfzelle_Format(app, "D3", 0, False, "4")
        Kopfzelle_Format(app, "E3", 0, False, "5")
        Kopfzelle_Format(app, "F3", 0, False, "6")
        Kopfzelle_Format(app, "G3", 0, False, "7")
        Kopfzelle_Format(app, "H3", 0, False, "8")
        Kopfzelle_Format(app, "I3", 0, False, "9")
        Kopfzelle_Format(app, "J3", 0, False, "10")
        Kopfzelle_Format(app, "K3", 0, False, "11")
        Kopfzelle_Format(app, "L3", 0, False, "12")
        'Hauptüberschriften erzeugen
        app.Rows("1:1").RowHeight = 47
        app.Rows("2:2").RowHeight = 13
        app.Rows("3:3").RowHeight = 13
        app.Columns("A:A").ColumnWidth = 10
        app.Columns("B:B").ColumnWidth = 35
        app.Columns("C:C").ColumnWidth = 15
        app.Columns("D:D").ColumnWidth = 15
        app.Columns("E:E").ColumnWidth = 15
        app.Columns("F:F").ColumnWidth = 33
        app.Columns("G:G").ColumnWidth = 20
        app.Columns("H:H").ColumnWidth = 12
        app.Columns("I:I").ColumnWidth = 40
        app.Columns("J:J").ColumnWidth = 13
        app.Columns("K:K").ColumnWidth = 20
        app.Columns("L:L").ColumnWidth = 20

    End Sub

    Public Sub Kopfdaten_ABSTANDSLISTE(ByRef app As Excel.Application,
                                                     ByRef wBook As Excel.Workbook,
                                                     ByRef wSheet As Excel.Worksheet)
        'Standardformat für alle Zellen
        Standardformat(app, "A1:M1", Excel.Constants.xlCenter, "Arial", "Fett", 10)
        'Einzelne Attribute
        Kopfzelle_Format(app, "A1", 0, False, "Mast-", "Mast")
        Kopfzelle_Format(app, "B1", 0, False, "Kreuzungsnr.")
        Kopfzelle_Format(app, "C1", 0, False, "Objekttyp", "lt.VDE")
        Kopfzelle_Format(app, "D1", 0, False, "Objektbeschreibung")
        Kopfzelle_Format(app, "E1", 0, False, "Eigentümer")
        Kopfzelle_Format(app, "F1", 0, False, "Abstand li.", "Mast(Ahpkt.)")
        Kopfzelle_Format(app, "G1", 0, False, "Phase/Stromkreis")
        Kopfzelle_Format(app, "H1", 0, False, "Zustand")
        Kopfzelle_Format(app, "I1", 0, False, "Ist-Seilspannung")
        Kopfzelle_Format(app, "J1", 0, False, "Soll-Seilspanung")
        Kopfzelle_Format(app, "K1", 0, False, "Abstand")
        Kopfzelle_Format(app, "L1", 0, False, "Soll-Abstand")
        Kopfzelle_Format(app, "M1", 0, False, "Mehrabstand")
        'Einheiten und Spaltennummerierung
        Kopfzelle_Format(app, "I2", 0, False, "[N/mm²]")
        Kopfzelle_Format(app, "J2", 0, False, "[N/mm²]")
        Kopfzelle_Format(app, "K2", 0, False, "[m]")
        Kopfzelle_Format(app, "L2", 0, False, "[m]")

        'Kopfzelle_Format(app, "AL3", 0, False, "[Euro/m²]")
        Kopfzelle_Format(app, "A3", 0, False, "1")
        Kopfzelle_Format(app, "B3", 0, False, "2")
        Kopfzelle_Format(app, "C3", 0, False, "3")
        Kopfzelle_Format(app, "D3", 0, False, "4")
        Kopfzelle_Format(app, "E3", 0, False, "5")
        Kopfzelle_Format(app, "F3", 0, False, "6")
        Kopfzelle_Format(app, "G3", 0, False, "7")
        Kopfzelle_Format(app, "H3", 0, False, "8")
        Kopfzelle_Format(app, "I3", 0, False, "9")
        Kopfzelle_Format(app, "J3", 0, False, "10")
        Kopfzelle_Format(app, "K3", 0, False, "11")
        Kopfzelle_Format(app, "L3", 0, False, "12")
        Kopfzelle_Format(app, "M3", 0, False, "13")
        'Hauptüberschriften erzeugen

        app.Columns("A:A").ColumnWidth = 11
        app.Columns("B:B").ColumnWidth = 15
        app.Columns("C:C").ColumnWidth = 15
        app.Columns("D:D").ColumnWidth = 25
        app.Columns("E:E").ColumnWidth = 15
        app.Columns("F:F").ColumnWidth = 15
        app.Columns("G:G").ColumnWidth = 20
        app.Columns("H:H").ColumnWidth = 13
        app.Columns("I:I").ColumnWidth = 20
        app.Columns("J:J").ColumnWidth = 20
        app.Columns("K:K").ColumnWidth = 15
        app.Columns("L:L").ColumnWidth = 20
        app.Columns("M:M").ColumnWidth = 20

    End Sub

    Public Sub Formatierung_INSTAND_PHASEN_BEZ(ByRef app As Excel.Application,
                                                 ByRef wBook As Excel.Workbook,
                                                 ByRef wSheet As Excel.Worksheet,
                                                 ByVal max_row As Integer,
                                                 ByVal xlRow As Integer,
                                                 ByVal Art As eKopfTyp)
        'Umrandung auf dick
        Formateinteilung("A1:S", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:S", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:S", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:S", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:S", Excel.XlBordersIndex.xlInsideHorizontal, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:S", Excel.XlBordersIndex.xlInsideVertical, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        ' Wichtige Spalten auf medium
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:F", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:F", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:F", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:F", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("G1:H", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("G1:H", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("G1:H", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("G1:H", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("I1:R", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("I1:R", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("I1:R", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("I1:R", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("I1:S", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("I1:S", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("I1:S", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("I1:S", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A4:S4", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A5:S5", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Druckbereich_Kopf_Fussdaten("$A$1:$S", "$1:$4", "", "", "", app, xlRow, max_row, Art)
    End Sub

    Public Sub Formatierung_INSTAND_STROMKREIS_BEZ(ByRef app As Excel.Application,
                                                 ByRef wBook As Excel.Workbook,
                                                 ByRef wSheet As Excel.Worksheet,
                                                 ByVal max_row As Integer,
                                                 ByVal xlRow As Integer,
                                                 ByVal Art As eKopfTyp)
        'Umrandung auf dick
        Formateinteilung("A1:Q", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:Q", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:Q", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:Q", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:Q", Excel.XlBordersIndex.xlInsideHorizontal, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:Q", Excel.XlBordersIndex.xlInsideVertical, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        ' Wichtige Spalten auf medium
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:G", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:G", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:G", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:G", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("H1:I", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("H1:I", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("H1:I", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("H1:I", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("J1:L", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("J1:L", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("J1:L", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("J1:L", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("M1:N", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("M1:N", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("M1:N", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("M1:N", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("O1:P", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("O1:P", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("O1:P", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("O1:P", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("O1:Q", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("O1:Q", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("O1:Q", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("O1:Q", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A4:Q4", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A5:Q5", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Druckbereich_Kopf_Fussdaten("$A$1:$Q", "$1:$4", "", "", "", app, xlRow, max_row, Art)
    End Sub

    Public Sub formatierung_INSTAND_MASTE(ByRef app As Excel.Application,
                                                 ByRef wBook As Excel.Workbook,
                                                 ByRef wSheet As Excel.Worksheet,
                                                 ByVal max_row As Integer,
                                                 ByVal xlRow As Integer,
                                                 ByVal Art As eKopfTyp)
        'Umrandung auf dick
        Formateinteilung("A1:Z", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:Z", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:Z", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:Z", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:Z", Excel.XlBordersIndex.xlInsideHorizontal, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:Z", Excel.XlBordersIndex.xlInsideVertical, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        ' Wichtige Spalten auf medium
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:E", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:E", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:E", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:E", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("F1:G", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("F1:G", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("F1:G", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("F1:G", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("H1:I", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("H1:I", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("H1:I", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("H1:I", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("J1:M", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("J1:M", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("J1:M", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("J1:M", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("N1:S", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("N1:S", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("N1:S", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("N1:S", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("T1:V", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("T1:V", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("T1:V", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("T1:V", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("W1:Y", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("W1:Y", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("W1:Y", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("W1:Y", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("W1:Z", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("W1:Z", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("W1:Z", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("W1:Z", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A4:Z4", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A5:Z5", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Druckbereich_Kopf_Fussdaten("$A$1:$Z", "$1:$4", "", "", "", app, xlRow, max_row, Art)
    End Sub

    Public Sub Formatierung_INSTAND_AUfHBEZ(ByRef app As Excel.Application,
                                                 ByRef wBook As Excel.Workbook,
                                                 ByRef wSheet As Excel.Worksheet,
                                                 ByVal max_row As Integer,
                                                 ByVal xlRow As Integer,
                                                 ByVal Art As eKopfTyp)
        'Umrandung auf dick
        Formateinteilung("A1:AJ", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AJ", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AJ", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AJ", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AJ", Excel.XlBordersIndex.xlInsideHorizontal, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AJ", Excel.XlBordersIndex.xlInsideVertical, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        ' Wichtige Spalten auf medium
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:E", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:E", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:E", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:E", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("F1:AC", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("F1:AC", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("F1:AC", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("F1:AC", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("AD1:AG", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("AD1:AG", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("AD1:AG", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("AD1:AG", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("AH1:AI", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("AH1:AI", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("AH1:AI", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("AH1:AI", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("AH1:AJ", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("AH1:AJ", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("AH1:AJ", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("AH1:AJ", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A5:AJ5", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A6:AJ6", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Druckbereich_Kopf_Fussdaten("$A$1:$AJ", "$1:$5", "", "", "", app, xlRow, max_row, Art)
    End Sub

    Public Sub Formatierung_Stromkreisverfolgung(ByRef app As Excel.Application,
                                                     ByRef wBook As Excel.Workbook,
                                                     ByRef wSheet As Excel.Worksheet,
                                                     ByVal max_row As Integer,
                                                     ByVal xlRow As Integer,
                                                     ByVal Art As eKopfTyp)
        'Umrandung auf dick
        Formateinteilung("A1:S", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:S", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:S", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:S", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:S", Excel.XlBordersIndex.xlInsideHorizontal, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:S", Excel.XlBordersIndex.xlInsideVertical, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        ' Wichtige Spalten auf medium
        Formateinteilung("A1:D", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:D", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:D", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:D", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("E1:K", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("E1:K", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("E1:K", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("E1:K", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("L1:M", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("L1:M", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("L1:M", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("L1:M", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("N1:P", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("N1:P", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("N1:P", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("N1:P", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("Q1:R", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("Q1:R", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("Q1:R", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("Q1:R", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("Q1:S", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("Q1:S", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("Q1:S", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("Q1:S", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A4:S4", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A5:S5", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Druckbereich_Kopf_Fussdaten("$A$1:$S", "$1:$4", "", "", "", app, xlRow, max_row, Art)
    End Sub

    Public Sub Formatierung_VERZ_DER_MASTE(ByRef app As Excel.Application,
                                                 ByRef wBook As Excel.Workbook,
                                                 ByRef wSheet As Excel.Worksheet,
                                                 ByVal max_row As Integer,
                                                 ByVal xlRow As Integer,
                                                 ByVal Art As eKopfTyp)
        'Umrandung auf dick
        Dim Range As Excel.Range = app.Range("A1:AF" & max_row)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlInsideHorizontal, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlInsideVertical, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        ' Wichtige Spalten auf medium
        Range = app.Range("A1:B" & max_row)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)

        Range = app.Range("A1:C" & max_row)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)

        Range = app.Range("D1:AE" & max_row)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)

        Range = app.Range("D1:AF" & max_row)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)

        Range = app.Range("A5:AF5" & max_row)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)

        Range = app.Range("A6:AF6" & max_row)
        FormateinteilungNeu(Range, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic)

        Druckbereich_Kopf_Fussdaten("$A$1:$AF", "$1:$5", "", "", "", app, xlRow, max_row, Art)
    End Sub

    Public Sub Formatierung_VERZ_AUSFUEHRUNGSPLANUNG(ByRef app As Excel.Application,
                                                 ByRef wBook As Excel.Workbook,
                                                 ByRef wSheet As Excel.Worksheet,
                                                 ByVal max_row As Integer,
                                                 ByVal xlRow As Integer,
                                                 ByVal Art As eKopfTyp)
        'Umrandung auf dick
        Formateinteilung("A1:AC", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AC", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AC", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AC", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AC", Excel.XlBordersIndex.xlInsideHorizontal, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AC", Excel.XlBordersIndex.xlInsideVertical, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        ' Wichtige Spalten auf medium
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:C", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:G", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:G", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:G", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("D1:G", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("H1:J", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("H1:J", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("H1:J", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("H1:J", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("K1:U", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("K1:U", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("K1:U", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("K1:U", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("V1:X", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("V1:X", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("V1:X", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("V1:X", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("Y1:AB", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("Y1:AB", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("Y1:AB", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("Y1:AB", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("Y1:AC", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("Y1:AC", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("Y1:AC", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("Y1:AC", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A5:AC5", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A6:AC6", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Druckbereich_Kopf_Fussdaten("$A$1:$AC", "$1:$5", "", "", "", app, xlRow, max_row, Art)
    End Sub

    Public Sub Formatierung_VERZ_GRUENDUNGEN(ByRef app As Excel.Application,
                                                 ByRef wBook As Excel.Workbook,
                                                 ByRef wSheet As Excel.Worksheet,
                                                 ByVal max_row As Integer,
                                                 ByVal xlRow As Integer,
                                                 ByVal Art As eKopfTyp)
        'Umrandung auf dick
        Formateinteilung("A1:V", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:V", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:V", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:V", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:V", Excel.XlBordersIndex.xlInsideHorizontal, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:V", Excel.XlBordersIndex.xlInsideVertical, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        ' Wichtige Spalten auf medium
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("C1:Q", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("C1:Q", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("C1:Q", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("C1:Q", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("C1:R", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("C1:R", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("C1:R", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("C1:R", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("S1:U", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("S1:U", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("S1:U", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("S1:U", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("S1:V", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("S1:V", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("S1:V", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("S1:V", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A5:V5", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A6:V6", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Druckbereich_Kopf_Fussdaten("$A$1:$V", "$1:$5", "", "", "", app, xlRow, max_row, Art)
    End Sub

    Public Sub Formatierung_RECHTLICHE_SICHERUNG(ByRef app As Excel.Application,
                                                 ByRef wBook As Excel.Workbook,
                                                 ByRef wSheet As Excel.Worksheet,
                                                 ByVal max_row As Integer,
                                                 ByVal xlRow As Integer,
                                                 ByVal Art As eKopfTyp)
        'Umrandung auf dick
        Formateinteilung("A1:AM", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AM", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AM", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AM", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AM", Excel.XlBordersIndex.xlInsideHorizontal, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AM", Excel.XlBordersIndex.xlInsideVertical, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        ' Wichtige Spalten auf medium
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("C1:H", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("C1:H", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("C1:H", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("C1:H", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("I1:O", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("I1:O", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("I1:O", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("I1:O", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("P1:AL", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("P1:AL", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("P1:AL", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("P1:AL", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("P1:AM", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("P1:AM", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("P1:AM", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("P1:AM", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A4:AM4", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A5:AM5", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Druckbereich_Kopf_Fussdaten("$A$1:$AM", "$1:$4", "", "", "", app, xlRow, max_row, Art)
    End Sub

    Public Sub Formatierung_KREUZUGNSVERZEICHNIS(ByRef app As Excel.Application,
                                                 ByRef wBook As Excel.Workbook,
                                                 ByRef wSheet As Excel.Worksheet,
                                                 ByVal max_row As Integer,
                                                 ByVal xlRow As Integer,
                                                 ByVal Art As eKopfTyp)
        'Umrandung auf dick
        Formateinteilung("A1:L", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:L", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:L", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:L", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:L", Excel.XlBordersIndex.xlInsideHorizontal, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:L", Excel.XlBordersIndex.xlInsideVertical, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        ' Wichtige Spalten auf medium
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:B", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("C1:D", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("C1:D", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("C1:D", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("C1:D", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("E1:L", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("E1:L", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("E1:L", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("E1:L", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A3:L3", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A4:L4", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Druckbereich_Kopf_Fussdaten("$A$1:$L", "$1:$3", "", "", "", app, xlRow, max_row, Art)
    End Sub

    Public Sub Formatierung_ABSTANDSLISTE(ByRef app As Excel.Application,
                                                 ByRef wBook As Excel.Workbook,
                                                 ByRef wSheet As Excel.Worksheet,
                                                 ByVal max_row As Integer,
                                                 ByVal xlRow As Integer,
                                                 ByVal Art As eKopfTyp)
        'Umrandung auf dick
        Formateinteilung("A1:AM", Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AM", Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AM", Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AM", Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AM", Excel.XlBordersIndex.xlInsideHorizontal, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Formateinteilung("A1:AM", Excel.XlBordersIndex.xlInsideVertical, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app, max_row)
        Druckbereich_Kopf_Fussdaten("$A$1:$AM", "$1:$3", "", "", "", app, xlRow, max_row, Art)
    End Sub

    Public Function Fusszeilen(ByRef app As Excel.Application,
                                                ByRef wBook As Excel.Workbook,
                                                ByRef wSheet As Excel.Worksheet,
                                                ByVal max_row As Integer,
                                                ByVal xlRow As Integer,
                                                ByVal Art As eKopfTyp,
                                                ByVal Leitungsnr As String,
                                                ByVal standort As String,
                                                ByVal Leitungsname As String,
                                                ByVal ErsterMast As String,
                                                ByVal LetzterMast As String) As String
        Dim retval As String = ""
        Dim bereich2 As String
        Dim inhalt As String
        Dim inhalt2 As String
        Dim inhalt3 As String
        Dim Abstand As String = "    " '4 Leerzeichen
        Try
            If Art = eKopfTyp.VERZEICHNIS_AUSFUEHRUNGSPLANUNG Then
                ' "Verzeichnis der Ausführungsplanung"
                bereich2 = "A" & max_row + 1 & ":D" & max_row + 9
                inhalt = Abstand & "zu Spalte 22:"
                inhalt2 = Abstand & "EH = Einzelhängekette" & Chr(10) & Abstand & "DH = Doppelhängekette" & Chr(10) & Abstand & "VH = V-Hängekette"
                inhalt3 = Abstand & "WH = Winkelhängekette" & Chr(10) & Abstand & "EA = Einzelabspannkette" & Chr(10) & Abstand & "DA = Doppelhängekette" & Chr(10) & Abstand & "TA = Tragabspannkette"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 8)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt, inhalt2, inhalt3)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlHairline, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formatierung_zelle("Arial", "Standard", 8, Excel.XlUnderlineStyle.xlUnderlineStyleSingle, Excel.Constants.xlAutomatic, bereich2, 5, 17, app)
                bereich2 = "E" & max_row + 1 & ":H" & max_row + 9
                inhalt = ""
                inhalt2 = ""
                inhalt3 = ""
                'Es wird der Text aus der Wertetabelle geschrieben
                'inhalt = Abstand & "zu Spalte 25:"
                'inhalt2 = Abstand & "Bl = Blockgründung" & Chr(10) & Abstand & "St = Stufengründung" & Chr(10) & Abstand & "B = Bohrgründung"
                'inhalt3 = Abstand & "R = Rammpfahlgründung" & Chr(10) & Abstand & "PL = Plattengründung" & Chr(10) & Abstand & "E = Einsetzgründung" & Chr(10) & Abstand & "S = Sondergründung (Details siehe Bemerkungen)"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 8)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt, inhalt2, inhalt3)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlHairline, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formatierung_zelle("Arial", "Standard", 8, Excel.XlUnderlineStyle.xlUnderlineStyleSingle, Excel.Constants.xlAutomatic, bereich2, 5, 17, app)

                bereich2 = "I" & max_row + 1 & ":Z" & max_row + 1
                inhalt = "DB Energie GmbH   " & standort
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlHairline, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "I" & max_row + 2 & ":L" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "I" & max_row + 3 & ":L" & max_row + 9
                inhalt = "aufgestellt"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "M" & max_row + 2 & ":V" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "M" & max_row + 3 & ":V" & max_row + 9
                inhalt = "geprüft"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "W" & max_row + 2 & ":Z" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "W" & max_row + 3 & ":Z" & max_row + 9
                inhalt = "genehmigt"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AA" & max_row + 1 & ":AC" & max_row + 1
                inhalt = "BL" & Leitungsnr & " 04-01 A"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AA" & max_row + 2 & ":AC" & max_row + 3
                inhalt = "BL" & Leitungsnr & " " & Leitungsname
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AA" & max_row + 4 & ":AC" & max_row + 5
                inhalt = "Masttafel der Ausführungsplanung"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AA" & max_row + 6 & ":AC" & max_row + 7
                inhalt = "Mast-Nr.: " & ErsterMast & Abstand & Abstand & Abstand & Abstand & Abstand & Abstand & "bis Mast-Nr.: " & LetzterMast
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AA" & max_row + 8 & ":AC" & max_row + 8
                inhalt = "Ersatz für:"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AA" & max_row + 9 & ":AC" & max_row + 9
                inhalt = "Stand vom:"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)


            ElseIf Art = eKopfTyp.VERZEICHNIS_GRUENDUNGEN Then
                '"Verzeichnis der Gründungen"
                bereich2 = "A" & max_row + 1 & ":H" & max_row + 9
                inhalt = " "
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlHairline, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "I" & max_row + 1 & ":R" & max_row + 1
                inhalt = "DB Energie GmbH   " & standort
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlHairline, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "I" & max_row + 2 & ":L" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "I" & max_row + 3 & ":L" & max_row + 9
                inhalt = "aufgestellt"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "M" & max_row + 2 & ":O" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "M" & max_row + 3 & ":O" & max_row + 9
                inhalt = "geprüft"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "P" & max_row + 2 & ":R" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "P" & max_row + 3 & ":R" & max_row + 9
                inhalt = "genehmigt"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "S" & max_row + 1 & ":V" & max_row + 1
                inhalt = "BL" & Leitungsnr & " 04-01 B"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "S" & max_row + 2 & ":V" & max_row + 3
                inhalt = "BL" & Leitungsnr & " " & Leitungsname
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "S" & max_row + 4 & ":V" & max_row + 5
                inhalt = "Verzeichnis der Gründungen"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "S" & max_row + 6 & ":V" & max_row + 7
                inhalt = "Mast-Nr.: " & ErsterMast & Abstand & Abstand & Abstand & Abstand & Abstand & Abstand & "bis Mast-Nr.: " & LetzterMast
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "S" & max_row + 8 & ":V" & max_row + 8
                inhalt = "Ersatz für:"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "S" & max_row + 9 & ":V" & max_row + 9
                inhalt = "Stand vom:"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

            ElseIf Art = eKopfTyp.VERZEICHNIS_MASTE Then
                '"Verzeichnis der Maste"
                bereich2 = "A" & max_row + 1 & ":P" & max_row + 9
                inhalt = " "
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlHairline, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "Q" & max_row + 1 & ":AC" & max_row + 1
                inhalt = "DB Energie GmbH   " & standort
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlHairline, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "Q" & max_row + 2 & ":T" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "Q" & max_row + 3 & ":T" & max_row + 9
                inhalt = "aufgestellt"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "U" & max_row + 2 & ":Y" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "U" & max_row + 3 & ":Y" & max_row + 9
                inhalt = "geprüft"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "Z" & max_row + 2 & ":AC" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "Z" & max_row + 3 & ":AC" & max_row + 9
                inhalt = "genehmigt"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AD" & max_row + 1 & ":AF" & max_row + 1
                inhalt = "BL" & Leitungsnr & " 04-01 B"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AD" & max_row + 2 & ":AF" & max_row + 3
                inhalt = "BL" & Leitungsnr & " " & Leitungsname
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AD" & max_row + 4 & ":AF" & max_row + 5
                inhalt = "Verzeichnis der Masten"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AD" & max_row + 6 & ":AF" & max_row + 7
                inhalt = "von Mast-Nr.: " & ErsterMast & Abstand & Abstand & Abstand & Abstand & Abstand & Abstand & "bis Mast-Nr.: " & LetzterMast
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AD" & max_row + 8 & ":AF" & max_row + 8
                inhalt = "Ersatz für:"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AD" & max_row + 9 & ":AF" & max_row + 9
                inhalt = "Stand vom:"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
            ElseIf Art = eKopfTyp.SONSTIGES Then
                'TODO EXCEL-export
            ElseIf Art = eKopfTyp.INSTANDHALTUNG_STROMKREISBEZOGEN Then
                '"Instandhaltung (Stromkreisbezogen)"

                bereich2 = "A" & max_row + 1 & ":K" & max_row + 1
                inhalt = "DB Energie GmbH   " & standort
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlHairline, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "A" & max_row + 2 & ":C" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "A" & max_row + 3 & ":C" & max_row + 9
                inhalt = "aufgestellt"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "D" & max_row + 2 & ":G" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "D" & max_row + 3 & ":G" & max_row + 9
                inhalt = "geprüft"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "H" & max_row + 2 & ":K" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "H" & max_row + 3 & ":K" & max_row + 9
                inhalt = "genehmigt"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "L" & max_row + 1 & ":Q" & max_row + 1
                inhalt = "BL" & Leitungsnr & " 04-01 B"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "L" & max_row + 2 & ":Q" & max_row + 3
                inhalt = "BL" & Leitungsnr & " " & Leitungsname
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "L" & max_row + 4 & ":Q" & max_row + 5
                inhalt = "Verzeichnis der Instandhaltung - Stromkreisbezogen"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "L" & max_row + 6 & ":Q" & max_row + 7
                inhalt = "Mast-Nr.: " & ErsterMast & Abstand & Abstand & Abstand & Abstand & Abstand & Abstand & "bis Mast-Nr.: " & LetzterMast
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "L" & max_row + 8 & ":Q" & max_row + 8
                inhalt = "Ersatz für:"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "L" & max_row + 9 & ":Q" & max_row + 9
                inhalt = "Stand vom:"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

            ElseIf Art = eKopfTyp.INSTANDHALTUNG_PHASENBEZOGEN Then
                '"Instandhaltung (Phasenbezogen)"
                bereich2 = "A" & max_row + 1 & ":M" & max_row + 1
                inhalt = "DB Energie GmbH   " & standort
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlHairline, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "A" & max_row + 2 & ":C" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "A" & max_row + 3 & ":C" & max_row + 9
                inhalt = "aufgestellt"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "D" & max_row + 2 & ":H" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "D" & max_row + 3 & ":H" & max_row + 9
                inhalt = "geprüft"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "I" & max_row + 2 & ":M" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "I" & max_row + 3 & ":M" & max_row + 9
                inhalt = "genehmigt"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "N" & max_row + 1 & ":S" & max_row + 1
                inhalt = "BL" & Leitungsnr & " 04-01 B"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "N" & max_row + 2 & ":S" & max_row + 3
                inhalt = "BL" & Leitungsnr & " " & Leitungsname
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "N" & max_row + 4 & ":S" & max_row + 5
                inhalt = "Verzeichnis der Instandhaltung - Phasenbezogen"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "N" & max_row + 6 & ":S" & max_row + 7
                inhalt = "Mast-Nr.: " & ErsterMast & Abstand & Abstand & Abstand & Abstand & Abstand & Abstand & "bis Mast-Nr.: " & LetzterMast
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "N" & max_row + 8 & ":S" & max_row + 8
                inhalt = "Ersatz für:"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "N" & max_row + 9 & ":S" & max_row + 9
                inhalt = "Stand vom:"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

            ElseIf Art = eKopfTyp.INSTANDHALTUNG_MAST Then
                '"Instandhaltung (Mast)"
                bereich2 = "A" & max_row + 1 & ":D" & max_row + 9
                inhalt = Abstand & ""
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlHairline, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formatierung_zelle("Arial", "Standard", 8, Excel.XlUnderlineStyle.xlUnderlineStyleSingle, Excel.Constants.xlAutomatic, bereich2, 5, 17, app)

                bereich2 = "E" & max_row + 1 & ":V" & max_row + 1
                inhalt = "DB Energie GmbH   " & standort
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlHairline, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "E" & max_row + 2 & ":H" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "E" & max_row + 3 & ":H" & max_row + 9
                inhalt = "aufgestellt"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "I" & max_row + 2 & ":P" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "I" & max_row + 3 & ":P" & max_row + 9
                inhalt = "geprüft"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "Q" & max_row + 2 & ":V" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "Q" & max_row + 3 & ":V" & max_row + 9
                inhalt = "genehmigt"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "W" & max_row + 1 & ":Z" & max_row + 1
                inhalt = "BL" & Leitungsnr & " 04-01 B"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "W" & max_row + 2 & ":Z" & max_row + 3
                inhalt = "BL" & Leitungsnr & " " & Leitungsname
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "W" & max_row + 4 & ":Z" & max_row + 5
                inhalt = "Verzeichnis der Instandhaltung - Maste"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "W" & max_row + 6 & ":Z" & max_row + 7
                inhalt = "Mast-Nr.: " & ErsterMast & Abstand & Abstand & Abstand & Abstand & Abstand & Abstand & "bis Mast-Nr.: " & LetzterMast
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "W" & max_row + 8 & ":Z" & max_row + 8
                inhalt = "Ersatz für:"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "W" & max_row + 9 & ":Z" & max_row + 9
                inhalt = "Stand vom:"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

            ElseIf Art = eKopfTyp.INSTANDHALTUNG_AUFHAENGEPUNKT Then
                '"Instandhaltung (Aufhaengepunkt)"
                bereich2 = "A" & max_row + 1 & ":L" & max_row + 9
                inhalt = " "
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formatierung_zelle("Arial", "Standard", 8, Excel.XlUnderlineStyle.xlUnderlineStyleSingle, Excel.Constants.xlAutomatic, bereich2, 5, 17, app)

                bereich2 = "M" & max_row + 1 & ":AC" & max_row + 1
                inhalt = "DB Energie GmbH   " & standort
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "M" & max_row + 2 & ":P" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "M" & max_row + 3 & ":P" & max_row + 9
                inhalt = "aufgestellt"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "Q" & max_row + 2 & ":U" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "Q" & max_row + 3 & ":U" & max_row + 9
                inhalt = "geprüft"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "V" & max_row + 2 & ":AC" & max_row + 2
                inhalt = "Firma"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                bereich2 = "V" & max_row + 3 & ":AC" & max_row + 9
                inhalt = "genehmigt"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10, Excel.Constants.xlTop)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AD" & max_row + 1 & ":AJ" & max_row + 1
                inhalt = "BL" & Leitungsnr & " 04-01 B"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AD" & max_row + 2 & ":AJ" & max_row + 3
                inhalt = "BL" & Leitungsnr & " " & Leitungsname
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AD" & max_row + 4 & ":AJ" & max_row + 5
                inhalt = "Verzeichnis der Instandhaltung - Aufhängepunktbezogen"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AD" & max_row + 6 & ":AJ" & max_row + 7
                inhalt = "Mast-Nr.: " & ErsterMast & Abstand & Abstand & Abstand & Abstand & Abstand & Abstand & "bis Mast-Nr.: " & LetzterMast
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AD" & max_row + 8 & ":AJ" & max_row + 8
                inhalt = "Ersatz für:"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "AD" & max_row + 9 & ":AJ" & max_row + 9
                inhalt = "Stand vom:"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

            ElseIf Art = eKopfTyp.STROMKREISVERFOLGUNG Then
                '"Stromkreisverfolgung"
                bereich2 = "A" & max_row + 1 & ":S" & max_row + 2
                inhalt = Abstand & " "
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 8)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlHairline, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "A" & max_row + 3 & ":S" & max_row + 4
                inhalt = "Übersicht aller Stromkreise der DB Energie im Gesamtdatenbestand und deren Verlauf über die entsprechenden Bahnstromleitungen"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Standard", 8)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "A" & max_row + 5 & ":E" & max_row + 9
                inhalt = "   "
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "J" & max_row + 5 & ":S" & max_row + 9
                inhalt = "   "
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlHairline, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)


                bereich2 = "F" & max_row + 5 & ":I" & max_row + 6
                inhalt = "DB Energie GmbH   " & standort
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "F" & max_row + 7 & ":I" & max_row + 7
                inhalt = "erstellt von:"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "F" & max_row + 8 & ":I" & max_row + 8
                inhalt = "geprüft von:  "
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "F" & max_row + 9 & ":I" & max_row + 9
                inhalt = "Stand vom:  "
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

            ElseIf Art = eKopfTyp.KREUZUNGSVERZEICHNIS Then
                '"Stromkreisverfolgung"
                bereich2 = "A" & max_row + 1 & ":F" & max_row + 7
                inhalt = Abstand & "Es ist zu unterscheiden zwischen Hochspannung über und unter 1kV,"
                inhalt2 = Abstand & "sowie Niederspannungsleitungen (bis 250V gegen Erde),"
                inhalt3 = Abstand & "bei Bundesstraßen ist die Nummer anzugeben."
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Standard", 8)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt, inhalt2, inhalt3)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "G" & max_row + 1 & ":I" & max_row + 1
                inhalt = "DB Energie GmbH"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Standard", 8)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "G" & max_row + 2 & ":I" & max_row + 2
                inhalt = "Niederlassung"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Standard", 8)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)


                bereich2 = "G" & max_row + 3 & ":H" & max_row + 4
                inhalt = "Unternehmen:" & Abstand
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "G" & max_row + 5 & ":H" & max_row + 7
                inhalt = "aufgestellt:"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "I" & max_row + 3 & ":I" & max_row + 4
                inhalt = "Standort:" & Abstand & standort
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "I" & max_row + 5 & ":I" & max_row + 7
                inhalt = "geprüft:"
                Standardformat(app, bereich2, Excel.Constants.xlLeft, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

                bereich2 = "J" & max_row + 1 & ":L" & max_row + 7
                inhalt = "110-kV-Bahnstromleitung" & Abstand & Leitungsnr
                inhalt2 = Leitungsname & Chr(10)
                inhalt3 = "Kreuzungsverzeichnis"
                Standardformat(app, bereich2, Excel.Constants.xlCenter, "Arial", "Fett", 10)
                Kopfzelle_Format(app, bereich2, 0, True, inhalt, inhalt2, inhalt3)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)
                Formateinteilung(bereich2, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBorderWeight.xlMedium, Excel.XlLineStyle.xlContinuous, Excel.XlColorIndex.xlColorIndexAutomatic, app)

            ElseIf Art = eKopfTyp.VERZ_RECHTLICHE_SICHERUNG Then

            End If
        Catch exc As Exception
            retval = exc.Message
            Tools.Message.SetToDatenBank(New Tools.clsExcelError("Fusszeilen", exc), "")
        End Try
        Return retval
    End Function

    Public Function ZeilenAuslesen(ByRef Data As System.Windows.Forms.DataGridView,
                                          ByVal wSheet As Excel.Worksheet,
                                          ByVal xlRow As Integer,
                                          ByVal xlCol As Integer,
                                          ByVal app As Excel.Application,
                                          ByVal Modulo As Integer) As Integer
        Try
            Dim c As System.Windows.Forms.DataGridViewColumn
            Dim versatz As Integer = xlRow
            Dim RowCount As Integer = Data.Rows.Count
            Dim ColumnCount As Integer = Data.Columns.Count
            For rowCnt As Integer = 0 To RowCount - 1
                For colCnt As Integer = 0 To ColumnCount - 1
                    c = Data.Columns(colCnt)
                    If c.Visible Then
                        Dim DieseZelle As Excel.Range = wSheet.Cells(xlRow, xlCol)
                        DieseZelle.NumberFormat = "@"
                        If Data(colCnt, rowCnt) IsNot Nothing And Data(colCnt, rowCnt).Value IsNot Nothing Then
                            DieseZelle.Value = Data(colCnt, rowCnt).FormattedValue
                            DieseZelle.Interior.ColorIndex = 0
                        End If
                        xlCol = xlCol + 1
                    End If
                Next
                'alle x -Zeilen eine Fußzeile mit Seitenumbruch einfügen
                Dim anzahl_zeilen As Integer = xlRow - versatz
                If anzahl_zeilen > 0 Then
                    If anzahl_zeilen Mod Modulo = 0 Then
                        Dim row_umbruch As String = xlRow & ":" & xlRow
                        app.ActiveWindow.SelectedSheets.HPageBreaks.Add(Before:=app.Rows(row_umbruch))
                    End If
                End If

                xlRow = xlRow + 1
                xlCol = 1
                Dim Zaehler As Integer = rowCnt + 1
                m_Statustext = "Eingelesen " & Zaehler & " von " & RowCount & " Zeilen."
                Georg.Refresh.Manager.StatusTextChanged(Me)
            Next
            wSheet.Cells.Interior.PatternColorIndex = 0
            wSheet.Cells.Borders.LineStyle = 0
            wSheet.Cells.HorizontalAlignment = Excel.Constants.xlCenter
            wSheet.Cells.VerticalAlignment = Excel.Constants.xlCenter
        Catch exc As Exception
            Tools.Message.SetToDatenBank(New Tools.clsExcelError("Zeilenauslesen: ", exc), "")
        End Try
        Return xlRow - 1
    End Function

    Public Function SpannweitenZeilenAuslesen(ByRef Data As System.Windows.Forms.DataGridView, _
                                                     ByVal wSheet As Excel.Worksheet, _
                                                     ByVal xlRow As Integer, _
                                                     ByVal cols As List(Of Integer), _
                                                     ByVal app As Excel.Application, _
                                                     ByVal Modulo As Integer) As Integer
        wSheet.Cells(2, 4).value = Environment.UserName
        wSheet.Cells(2, 6).value = Now.ToShortDateString

        Dim c As System.Windows.Forms.DataGridViewColumn
        For rowCnt As Integer = 0 To Data.Rows.Count - 1
            'InsertRow(wSheet, xlRow + 1)
            Dim i As Integer = 1
            For Each ColCnt As Integer In cols
                c = Data.Columns(i - 1)
                If c.Visible Then
                    If Data(i - 1, rowCnt) IsNot Nothing And Data(i - 1, rowCnt).Value IsNot Nothing Then
                        wSheet.Cells(xlRow, ColCnt).value = Data(i - 1, rowCnt).Value.ToString
                    End If
                    i += 1
                End If
            Next
            'alle x -Zeilen ein Fußzeile mit Seitenumbruch einfügen
            If xlRow Mod Modulo = 0 Then
                Dim row_umbruch As String = xlRow & ":" & xlRow
                app.ActiveWindow.SelectedSheets.HPageBreaks.Add(Before:=app.Rows(row_umbruch))
            End If
            xlRow = xlRow + 1
            i = 1
        Next
        Return xlRow - 1
    End Function

    Public Sub Export(ByRef Data As System.Windows.Forms.DataGridView, _
                               ByVal TabName As String, _
                               ByVal MitFarben As Boolean, _
                               Optional ByVal sDir As String = "")
        Dim f As frmProgress = New frmProgress
        f.Show()
        Dim progress As System.Windows.Forms.ProgressBar = f.progress
        progress.Minimum = 0
        progress.Maximum = Data.Rows.Count
        progress.Step = 1
        progress.Value = 0
        TabName = clsExportTools.MakeValidTabname(TabName)
        Dim app As Excel.Application = Nothing
        Try
            app = New Excel.Application
        Catch exc As Exception
        End Try
        If app Is Nothing Then
            Tools.Message.SetHinweis(clsExcel.C_KEIN_EXCEL)
            Return
        End If
        app.Visible = False
        Dim wBook As Excel.Workbook = app.Workbooks.Add
        While wBook.Worksheets.Count > 1
            wBook.ActiveSheet.delete()
        End While
        Dim wSheet As Excel.Worksheet = wBook.ActiveSheet
        'wSheet.Name = MakeValidTabname(TabName)
        wSheet.Name = TabName
        Dim xlRow As Integer = 1
        Dim xlCol As Integer = 1
        For idx As Integer = 0 To Data.Columns.Count - 1
            Dim c As System.Windows.Forms.DataGridViewColumn = Data.Columns(idx)
            If c.Visible Then
                wSheet.Cells(xlRow, xlCol).NumberFormat = "@"
                wSheet.Cells(xlRow, xlCol).value = c.HeaderText
                wSheet.Cells(xlRow, xlCol).font.bold = True
                wSheet.Cells(xlRow, xlCol).HorizontalAlignment = Excel.Constants.xlCenter
                If MitFarben Then
                    wSheet.Cells(xlRow, xlCol).interior.color = System.Drawing.ColorTranslator.ToOle(c.InheritedStyle.BackColor)
                End If
                xlCol = xlCol + 1
            End If
        Next
        xlRow = xlRow + 1
        progress.PerformStep()
        xlCol = 1
        For rowCnt As Integer = 0 To Data.Rows.Count - 1
            For colCnt As Integer = 0 To Data.Columns.Count - 1
                Dim c As System.Windows.Forms.DataGridViewColumn = Data.Columns(colCnt)
                If c.Visible Then
                    wSheet.Cells(xlRow, xlCol).NumberFormat = "@"
                    If Data(colCnt, rowCnt) IsNot Nothing And Data(colCnt, rowCnt).Value IsNot Nothing Then
                        wSheet.Cells(xlRow, xlCol).value = Data(colCnt, rowCnt).Value.ToString
                        If MitFarben Then
                            wSheet.Cells(xlRow, xlCol).interior.color = System.Drawing.ColorTranslator.ToOle(Data(colCnt, rowCnt).InheritedStyle.BackColor)
                        End If
                    End If
                    xlCol = xlCol + 1
                End If
            Next
            progress.PerformStep()
            xlRow = xlRow + 1
            xlCol = 1
        Next

        ' Zeilenumbruch aus
        wSheet.Cells.WrapText = False
        wSheet.Cells.EntireColumn.AutoFit()
        ' jetzt sind die Titelzeilen viel zu breit
        wSheet.Rows(1).cells.wraptext = True
        ' jetzt sind die Titelzeilen umgebrochen
        wSheet.Cells.EntireColumn.AutoFit()
        If Len(sDir) > 0 Then
            Try
                If System.IO.File.Exists(sDir & TabName & C_EXCELDATEI_ENDUNG) Then
                    System.IO.File.Delete(sDir & TabName & C_EXCELDATEI_ENDUNG)
                End If
                wSheet.SaveAs(sDir & TabName & C_EXCELDATEI_ENDUNG)
            Catch ex As System.IO.IOException
                MsgBox("Bitte vor dem Export die Datei " & vbNewLine & sDir & TabName & vbNewLine & " schliessen!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation, "Nicht gespeichert!")
            End Try
        End If
        ' und jetzt sollten die Spalten stimmen
        app.Visible = True
        app.UserControl = True
        f.Close()
    End Sub

#Region "Implementierungen"

    Public ReadOnly Property StatusText As String Implements Georg.IStatusTextSender.StatusText
        Get
            Return (m_Statustext)
        End Get
    End Property

    Public ReadOnly Property StatusTextAktualisierungsTyp As Georg.IStatusTextSender.eStatusTextAktualisierungsTyp Implements Georg.IStatusTextSender.StatusTextAktualisierungsTyp
        Get
            Return (Georg.IStatusTextSender.eStatusTextAktualisierungsTyp.ASSISTENT)
        End Get
    End Property

    Public ReadOnly Property StatusTextSenderTyp As Georg.IStatusTextSender.eStatusTextSenderTyp Implements Georg.IStatusTextSender.StatusTextSenderTyp
        Get
            Return (Georg.IStatusTextSender.eStatusTextSenderTyp.DATENSTATE)
        End Get
    End Property

    Public ReadOnly Property TemporaererStatusText As String Implements Georg.IStatusTextSender.TemporaererStatusText
        Get
            Return (m_TemporaererStatustext)
        End Get
    End Property

    Public ReadOnly Property StatusAnzeigeZeit As Georg.IStatusTextSender.eStatusAnzeigeZeit Implements Georg.IStatusTextSender.StatusAnzeigeZeit
        Get
            Return (m_StatusAnzeigezeit)
        End Get
    End Property
#End Region
End Class

Public Class ExcelBlockExport
    Private m_App As Excel.Application = Nothing
    Private m_wBook As Excel.Workbook = Nothing
    Private m_wsheet As Excel.Worksheet = Nothing
    Private m_Tabname As String = ""

    Public Sub New(ByVal TabName As String)
        Try
            m_App = New Excel.Application
        Catch exc As Exception
        End Try
        If m_App Is Nothing Then
            Tools.Message.SetHinweis(clsExcel.C_KEIN_EXCEL)
        End If
        m_App.Visible = False
        m_Tabname = clsExportTools.MakeValidTabname(TabName)
        m_wBook = m_App.Workbooks.Add
        While m_wBook.Worksheets.Count > 1
            m_wBook.ActiveSheet.delete()
        End While
        m_wsheet = m_wBook.ActiveSheet
        m_wsheet.Name = m_Tabname
    End Sub

    Public Sub New(ByVal Vorlage As String, ByVal Sheet As String, ByVal Dateiname As String)
        m_App = New Excel.Application
        m_App.Visible = False
        Dim fi As IO.FileInfo = New IO.FileInfo(Vorlage)
        If fi.Exists Then
            m_wBook = m_App.Workbooks.Open(Vorlage)
            m_wsheet = m_wBook.Sheets(Sheet)
            m_wsheet.Activate()
        End If
        m_Tabname = Dateiname
    End Sub

    Public Sub Export(ByRef Zeilennummer As Integer, _
                      ByVal Spalte As Integer, _
                      ByVal Überschrift As String, _
                      ByRef Data As System.Windows.Forms.DataGridView)

        'Überschrift das Blockes
        Dim xlRow As Integer = Zeilennummer
        Dim xlCol As Integer = 1 + (Spalte - 1) * 3
        If xlRow <= 1 Then
            m_wsheet.Cells(xlRow, xlCol).NumberFormat = "@"
            m_wsheet.Cells(xlRow, xlCol).value = Überschrift
            m_wsheet.Cells(xlRow, xlCol).font.bold = True
            m_wsheet.Cells(xlRow, xlCol).HorizontalAlignment = Excel.Constants.xlCenter
            xlRow = xlRow + 2
        End If
        'Alle Zeilen in der Dataview
        For rowCnt As Integer = 0 To Data.Rows.Count - 1
            For colCnt As Integer = 0 To Data.Columns.Count - 1
                Dim c As System.Windows.Forms.DataGridViewColumn = Data.Columns(colCnt)
                If c.Visible Then
                    m_wsheet.Cells(xlRow, xlCol).NumberFormat = "@"
                    If Data(colCnt, rowCnt) IsNot Nothing And Data(colCnt, rowCnt).Value IsNot Nothing Then
                        m_wsheet.Cells(xlRow, xlCol).value = Data(colCnt, rowCnt).Value.ToString
                    End If
                    xlCol += 1
                End If
            Next
            xlRow += 1
            xlCol = (1 + (Spalte - 1) * 3)
        Next
        Zeilennummer = xlRow + 2
        ' Zeilenumbruch aus
        m_wsheet.Cells.WrapText = False
        m_wsheet.Cells.EntireColumn.AutoFit()
        ' jetzt sind die Titelzeilen viel zu breit
        m_wsheet.Rows(Zeilennummer).cells.wraptext = True
        ' jetzt sind die Titelzeilen umgebrochen
        m_wsheet.Cells.EntireColumn.AutoFit()
    End Sub

    Public Sub Export_WSW(ByRef Zeilennummer As Integer, _
                      ByRef Data As System.Data.DataRow, _
                      ByVal Objekt As Integer)
        'Unterschiedliche Spaltenzuweisung für unterschiedliche Objekte
        'Objekt 0 = Flurstücksdaten     
        'Spalte 3: Gemarkung, Spalte 4: Flur, Spalte 5: Flurstücksnummer (Zähler/Nenner)
        'Objekt 1 = Eigentümerdaten
        'Spalte 8: Nachname, Spalte 9: Vorname, Spalte 10: Strasse (Strasse + Hausnummer), Spalte 11: PLZ, Spalte 12: Ort
        'Objekt 2 = Pächter
        'Spalte 20: Nachname, Spalte 21: Vorname, Spalte 22: Strasse (Strasse + Hausnummer), Spalte 23: PLZ, Spalte 24: Ort
        'Objekt 3 = Behördendaten
        'Spalte 8: Nachname, Spalte 9: Vorname, Spalte 10: Strasse (Strasse + Hausnummer), Spalte 11: PLZ, Spalte 12: Ort
        'Objekt 4 = Ansprechpartner
        'Spalte 8: Nachname, Spalte 9: Vorname, Spalte 10: Strasse (Strasse + Hausnummer), Spalte 11: PLZ, Spalte 12: Ort
        Dim xlRow As Integer = Zeilennummer
        If Objekt = 0 Then
            Dim zaehler As String = ""
            Dim nenner As String = ""
            Dim alpha As String = ""
            zaehler = Data.Item(0).ToString
            nenner = Data.Item(1).ToString
            alpha = Data.Item(2).ToString
            Dim Flurstuecksnr As String = zaehler
            If alpha > 0 Then
                Flurstuecksnr = Flurstuecksnr & alpha
            End If
            If nenner > 0 Then
                Flurstuecksnr = Flurstuecksnr & "/" & nenner
            End If
            m_wsheet.Cells(xlRow, 5).NumberFormat = "@"
            m_wsheet.Cells(xlRow, 5).value() = Flurstuecksnr
            m_wsheet.Cells(xlRow, 4).NumberFormat = "@"
            m_wsheet.Cells(xlRow, 4).value() = Data.Item(3).ToString
            m_wsheet.Cells(xlRow, 3).value() = Data.Item(4).ToString

        End If
        If Objekt = 1 Then
            m_wsheet.Cells(xlRow, 8).NumberFormat = "@"
            m_wsheet.Cells(xlRow, 8).value() = Data.Item(3).ToString
            m_wsheet.Cells(xlRow, 9).NumberFormat = "@"
            m_wsheet.Cells(xlRow, 9).value() = Data.Item(2).ToString
            m_wsheet.Cells(xlRow, 10).value() = Data.Item(5).ToString
            m_wsheet.Cells(xlRow, 11).NumberFormat = "@"
            m_wsheet.Cells(xlRow, 11).value() = Data.Item(6).ToString
            m_wsheet.Cells(xlRow, 12).value() = Data.Item(7).ToString
        End If
        If Objekt = 2 Then
            m_wsheet.Cells(xlRow, 20).NumberFormat = "@"
            m_wsheet.Cells(xlRow, 20).value() = Data.Item(3).ToString
            m_wsheet.Cells(xlRow, 21).NumberFormat = "@"
            m_wsheet.Cells(xlRow, 21).value() = Data.Item(2).ToString
            m_wsheet.Cells(xlRow, 22).NumberFormat = "@"
            m_wsheet.Cells(xlRow, 22).value() = Data.Item(5).ToString
            m_wsheet.Cells(xlRow, 23).value() = Data.Item(6).ToString
            m_wsheet.Cells(xlRow, 24).value() = Data.Item(7).ToString
        End If
        If Objekt = 3 Then
            m_wsheet.Cells(xlRow, 8).NumberFormat = "@"
            m_wsheet.Cells(xlRow, 8).value() = Data.Item(0).ToString
            m_wsheet.Cells(xlRow, 9).NumberFormat = "@"
            m_wsheet.Cells(xlRow, 9).value() = Data.Item(1).ToString
            m_wsheet.Cells(xlRow, 11).NumberFormat = "@"
            m_wsheet.Cells(xlRow, 11).value() = Data.Item(2).ToString
            m_wsheet.Cells(xlRow, 12).value() = Data.Item(3).ToString
            'm_wsheet.Cells(xlRow, 12).value() = Data.Item(7).ToString
        End If
        If Objekt = 4 Then
            m_wsheet.Cells(xlRow, 8).NumberFormat = "@"
            m_wsheet.Cells(xlRow, 8).value() = Data.Item(3).ToString
            m_wsheet.Cells(xlRow, 9).NumberFormat = "@"
            m_wsheet.Cells(xlRow, 9).value() = Data.Item(2).ToString
            m_wsheet.Cells(xlRow, 10).value() = Data.Item(5).ToString
            m_wsheet.Cells(xlRow, 11).NumberFormat = "@"
            m_wsheet.Cells(xlRow, 11).value() = Data.Item(6).ToString
            m_wsheet.Cells(xlRow, 12).value() = Data.Item(7).ToString
        End If
    End Sub

    Public Sub SaveWorkbook(Optional ByVal sDir As String = "")
        If Len(sDir) > 0 Then
            Try
                If System.IO.File.Exists(sDir & m_Tabname & EXCELExport.C_EXCELDATEI_ENDUNG) Then
                    System.IO.File.Delete(sDir & m_Tabname & EXCELExport.C_EXCELDATEI_ENDUNG)
                End If
                m_wsheet.SaveAs(sDir & m_Tabname & EXCELExport.C_EXCELDATEI_ENDUNG)
            Catch ex As System.IO.IOException
                MsgBox("Bitte vor dem Export die Datei " & vbNewLine & sDir & m_Tabname & vbNewLine & " schliessen!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation, "Nicht gespeichert!")
            End Try
        End If
        ' und jetzt sollten die Spalten stimmen
        m_App.Visible = True
        m_App.UserControl = True
    End Sub

    Public Sub SaveasWorkbook(ByVal sDir As String, ByVal zusatz As String, ByVal flurstuecksnr As String)
        If Len(sDir) > 0 Then
            If Len(zusatz) > 0 Then
                zusatz = "_" & zusatz
            End If
            'm_Tabname = Left(m_Tabname, InStr(1, m_Tabname, ".") - 1)
            Dim Dateiname As String = sDir & m_Tabname & "_" & flurstuecksnr & zusatz & EXCELExport.C_EXCELDATEI_ENDUNG
            Try
                If System.IO.File.Exists(Dateiname) Then
                    System.IO.File.Delete(Dateiname)
                End If
                m_wsheet.SaveAs(Dateiname)
            Catch ex As System.IO.IOException
                MsgBox("Bitte vor dem Export die Datei " & vbNewLine & sDir & flurstuecksnr & m_Tabname & zusatz & EXCELExport.C_EXCELDATEI_ENDUNG & vbNewLine & " schliessen!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation, "Nicht gespeichert!")
            End Try
        End If
        m_App.Visible = True
        m_App.UserControl = True
    End Sub

    Public Sub ActivateSheet(ByVal sheet As String)
        m_wsheet = m_wBook.Sheets(sheet)
        m_wsheet.Activate()
    End Sub
    
End Class
