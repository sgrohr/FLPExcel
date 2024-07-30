Imports Microsoft.Office.Interop

Public Class clsExcel
    Private m_ExcelApp As Excel.Application = Nothing
    Private m_CurWorksheet As String = Nothing
    Private m_RowNumInWorkSheet As Dictionary(Of String, Integer) = Nothing

    Public Const C_KEIN_EXCEL As String = "Excel steht nicht zur Verfügung."

    Public ReadOnly Property Version As String
        Get
            Return m_ExcelApp.Version
        End Get
    End Property


#Region "Auf- und Zumachen Dokument"

    Private Function SonderzeichenBereinigen(ByVal txt As String) As String
        Dim dict As New Dictionary(Of String, String)

        dict.Add(":", "_")
        dict.Add("\", "_")
        dict.Add("/", "_")
        dict.Add("?", "_")
        dict.Add("*", "_")
        dict.Add("[", "_")
        dict.Add("]", "_")

        For Each pair As KeyValuePair(Of String, String) In dict
            If txt.Contains(pair.Key) Then
                txt = Replace(txt, pair.Key, pair.Value)
            End If
        Next

        Return txt
    End Function

    Public Function NewDocument(ByVal Dateiname As String,
                                ByVal ArbeitsblattName As String,
                                ByVal DateiPfad As IO.DirectoryInfo) As Tools.clsBoolwithReason
        Try
            m_ExcelApp = New Excel.Application
        Catch exc As Exception
        End Try
        If m_ExcelApp Is Nothing Then
            Tools.Message.SetHinweis(C_KEIN_EXCEL)
            Return New Tools.clsBoolwithReason(False)
        End If
        Tools.MakeDirIfNeeded(DateiPfad.FullName)
        Dim FullDateiName As String = IO.Path.Combine(DateiPfad.FullName, Dateiname)
        Dim fi As IO.FileInfo = New IO.FileInfo(FullDateiName)
        If fi.Exists Then
            Try
                fi.Delete()
            Catch
                Tools.Message.SetHinweis("Sie haben eine Datei " & Dateiname & " geöffnet. Schließen Sie diese.")
                Return (New Tools.clsBoolwithReason(False))
            End Try
        End If

        ArbeitsblattName = SonderzeichenBereinigen(ArbeitsblattName)
        m_ExcelApp.Visible = False
        m_ExcelApp.Workbooks.Add()
        m_ExcelApp.ActiveWorkbook.SaveAs(FullDateiName)
        While m_ExcelApp.Sheets.Count > 1
            m_ExcelApp.ActiveWorkbook.Sheets(1).Delete()
        End While
        m_ExcelApp.ActiveWorkbook.Sheets(1).Name = ArbeitsblattName
        m_ExcelApp.ActiveWorkbook.Activate()
        m_ExcelApp.ActiveWorkbook.Save()
        m_RowNumInWorkSheet = New Dictionary(Of String, Integer)
        m_RowNumInWorkSheet(ArbeitsblattName) = 1
        m_CurWorksheet = ArbeitsblattName
        Dim retVal As New Tools.clsBoolwithReason(True, FullDateiName)
        Return (retVal)
    End Function

    Public Function CloseDocument() As Tools.clsBoolwithReason
        If m_ExcelApp Is Nothing Then Return (New Tools.clsBoolwithReason(False, "Keine EXCEL-Applikation aktiv"))
        If m_ExcelApp.ActiveWorkbook Is Nothing Then Return (New Tools.clsBoolwithReason(False, "Kein EXCEL-Dokument offen"))
        m_ExcelApp.ActiveWorkbook.Save()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(m_ExcelApp.Worksheets)
        m_ExcelApp.ActiveWorkbook.Close()
        For Each workbook As Excel.Workbook In m_ExcelApp.Workbooks
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(m_ExcelApp.Workbooks)
        m_ExcelApp.EnableEvents = True
        m_ExcelApp.Quit()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(m_ExcelApp)
        m_ExcelApp = Nothing
        m_CurWorksheet = Nothing
        m_RowNumInWorkSheet.Clear()
        m_RowNumInWorkSheet = Nothing
        Return (New Tools.clsBoolwithReason(True))
    End Function

    Public Sub OpenDocument(Filename)
        Dim objXlsx As New Excel.Application
        Dim path As String = Filename
        objXlsx.Workbooks.Open(path)
        objXlsx.Visible = True
    End Sub


    Friend Sub ReleaseComObject(ByVal o As Object)
        Try
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(o)
        Catch
        Finally
            If o IsNot Nothing Then o = Nothing
        End Try
    End Sub


#End Region

#Region "Worksheets aktivieren / anlegen"

    Public Function AddTabellenblatt(ByVal Name As String) As Tools.clsBoolwithReason
        Name = SonderzeichenBereinigen(Name)

        If m_ExcelApp Is Nothing Then Return (New Tools.clsBoolwithReason(False, "Keine EXCEL Applikation aktiv"))
        If BlattnameExistiert(Name) Then Return (New Tools.clsBoolwithReason(False, "Blattname existiert bereits"))
        Dim o As Excel.Worksheet = m_ExcelApp.ActiveWorkbook.Sheets.Add(, m_ExcelApp.ActiveSheet)
        o.Name = Name
        m_RowNumInWorkSheet(Name) = 1
        m_CurWorksheet = Name
        Return (New Tools.clsBoolwithReason(True))
    End Function

    Private Function BlattnameExistiert(ByVal Name As String) As Boolean
        For Each o As Excel.Worksheet In m_ExcelApp.ActiveWorkbook.Sheets
            If o.Name = Name Then Return (True)
        Next
        Return (False)
    End Function

    Public Sub AppInteractive(_istate As Boolean)
        m_ExcelApp.Interactive = _istate
    End Sub
#End Region

#Region "Daten einfuegen"

    Public Sub FillExcel(ByRef Info As Object, _
                         ByVal Farbe As eExcelFarbe,
                         Optional Numberformat As String = Nothing)
        Try
            Dim CurRow As Integer = m_RowNumInWorkSheet(m_CurWorksheet)
            If IsArray(Info) Then
                Dim colCnt As Integer = 1
                For Each o In Info
                    If Not String.IsNullOrEmpty(Numberformat) Then
                        If IsNumeric(o) Then
                            If o.GetType = GetType(System.Single) Or o.GetType = GetType(System.Double) Then
                                m_ExcelApp.Cells(CurRow, colCnt).Numberformat = Numberformat
                            End If
                        End If
                    End If
                    If o.GetType Is GetType(String) Then
                        m_ExcelApp.Cells(CurRow, colCnt).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                        Dim str As String = o
                        If Not String.IsNullOrEmpty(str) AndAlso IsNumeric(str.Substring(0, 1)) Then
                            If Len(str) > 1 Then
                                If Not IsNumeric(str) Then
                                    str = "'" & str
                                End If
                            End If
                        End If
                        m_ExcelApp.Cells(CurRow, colCnt) = str
                    Else
                        m_ExcelApp.Cells(CurRow, colCnt) = o
                    End If
                    colCnt += 1
                Next
            Else
                m_ExcelApp.Cells(CurRow, 1) = Info
            End If
            Faerbe(CurRow, Farbe)
            m_RowNumInWorkSheet(m_CurWorksheet) = CurRow + 1
        Catch exc As Exception
            Tools.AsyncFehlermelder.Melde(Me, New Tools.clsExcelError("clsExcel.FillExcel (ohne SortParameter)", exc))
        End Try
    End Sub

    Public Function FillExcel(ByVal dt As DataTable, _
                         ByVal HeaderFarbe As Integer, _
                         ByVal RowsFarbe As eExcelFarbe, _
                         ByVal SortParams As clsExcelSortParameter) As Boolean
        Try
            Dim ColAnz As Integer = dt.Columns.Count - 1
            Dim Headers(ColAnz) As String
            For i As Integer = 0 To ColAnz
                Headers(i) = dt.Columns(i).ColumnName
            Next
            FillExcelBold(Headers, HeaderFarbe)
            Dim StartRow As Integer = m_RowNumInWorkSheet(m_CurWorksheet)
            Dim CurRow As Integer = StartRow
            For Each row As DataRow In dt.Rows
                For ColCnt As Integer = 0 To ColAnz
                    m_ExcelApp.Cells(CurRow, ColCnt + 1).Numberformat = "@"
                    m_ExcelApp.Cells(CurRow, ColCnt + 1) = row(ColCnt)
                Next
                CurRow = CurRow + 1
            Next
            m_RowNumInWorkSheet(m_CurWorksheet) = CurRow
            ' Sortierungsgejockel
            If SortParams IsNot Nothing Then
                Dim BeginnSpalte As Char = "A"
                Dim EndeSpalte As Char = Chr(Asc("A") + dt.Columns.Count - 1)
                Dim RangeFuerAlles As String = BeginnSpalte & StartRow & ":" & EndeSpalte & CurRow - 1
                m_ExcelApp.Range(RangeFuerAlles).Select()
                Dim r As Excel.Range = m_ExcelApp.Selection
                ' Achtung, es gibt die Konstante Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortRows für die Orientation
                ' aber der Wert ist 2; benoetigt wird fürs Zeilensortieren eine 1
                ' wie im Kindergarten, 1, 2 oder 3
                Select Case SortParams.GetAnzahl
                    Case 0
                        ' nix zu tun
                    Case 1
                        Dim p1 As FLPExcel.clsExcelSortParameter.tColIdxAscending = SortParams.GetParameter(0)
                        Dim k1 As String = Chr(Asc("A") + p1.ColumnIndex) & StartRow
                        Dim r1 As Microsoft.Office.Interop.Excel.Range = m_ExcelApp.Range(k1)
                        r.Sort(Key1:=r1, Order1:=p1.SortOrder, Header:=Excel.XlYesNoGuess.xlNo, Orientation:=1)
                    Case 2
                        Dim p1 As FLPExcel.clsExcelSortParameter.tColIdxAscending = SortParams.GetParameter(0)
                        Dim k1 As String = Chr(Asc("A") + p1.ColumnIndex) & StartRow
                        Dim r1 As Microsoft.Office.Interop.Excel.Range = m_ExcelApp.Range(k1)
                        Dim p2 As FLPExcel.clsExcelSortParameter.tColIdxAscending = SortParams.GetParameter(1)
                        Dim k2 As String = Chr(Asc("A") + p2.ColumnIndex) & StartRow
                        Dim r2 As Microsoft.Office.Interop.Excel.Range = m_ExcelApp.Range(k2)
                        r.Sort(Key1:=r1, Order1:=p1.SortOrder, Key2:=r2, Order2:=p2.SortOrder, Header:=Excel.XlYesNoGuess.xlNo, Orientation:=1)
                    Case 3
                        Dim p1 As FLPExcel.clsExcelSortParameter.tColIdxAscending = SortParams.GetParameter(0)
                        Dim k1 As String = Chr(Asc("A") + p1.ColumnIndex) & StartRow
                        Dim r1 As Microsoft.Office.Interop.Excel.Range = m_ExcelApp.Range(k1)
                        Dim p2 As FLPExcel.clsExcelSortParameter.tColIdxAscending = SortParams.GetParameter(1)
                        Dim k2 As String = Chr(Asc("A") + p2.ColumnIndex) & StartRow
                        Dim r2 As Microsoft.Office.Interop.Excel.Range = m_ExcelApp.Range(k2)
                        Dim p3 As FLPExcel.clsExcelSortParameter.tColIdxAscending = SortParams.GetParameter(2)
                        Dim k3 As String = Chr(Asc("A") + p3.ColumnIndex) & StartRow
                        Dim r3 As Microsoft.Office.Interop.Excel.Range = m_ExcelApp.Range(k3)
                        r.Sort(Key1:=r1, Order1:=p1.SortOrder, Key2:=r2, Order2:=p2.SortOrder, Key3:=r3, Order3:=p3.SortOrder, Header:=Excel.XlYesNoGuess.xlNo, Orientation:=1)
                End Select
                m_ExcelApp.Range("A" & CurRow).Select()
            End If
        Catch exc As Exception
            Tools.AsyncFehlermelder.Melde(Me, New Tools.clsExcelError("clsExcel.FillExcel (mit SortParameter)", exc))
            Return False
        End Try
        Return True
    End Function

    Public Sub FillExcelBold(ByRef Info As Object, _
                             ByVal Farbe As eExcelFarbe)
        ' vor dem Befuellen abfragen !
        Dim CurRow As Integer = m_RowNumInWorkSheet(m_CurWorksheet)
        FillExcel(Info, Farbe)
        ' jetzt ist Dim CurRow As Integer = m_RowNumInWorkSheet(m_CurWorksheet) um 1 groesser
        Zeichensatz(CurRow, m_ExcelApp.Rows(CurRow).font.name, m_ExcelApp.Rows(CurRow).font.size, True)
    End Sub

#End Region

#Region "Bereich zusammenfassen"

    ''' <summary>
    ''' Achtung: Werden Spalten verbunden, und anschliessend die Zeile gefüllt, dann gehen die Informationen 
    ''' in den auf die Ankerspalte folgenden Spalten meldungslos verloren
    ''' Bsp: Merge ("B","E") FillInfo {1,2,3,4,5,6}} liefert 1 2 - - - 6; die in C,D und E gefüülten Werte sind im Nirwana
    ''' </summary>
    ''' <param name="strVonSpalte"></param>
    ''' <param name="strBisSpalte"></param>
    ''' <remarks>MergeSpalten wirkt sich auf die nöchste zu füllende Zeile aus!</remarks>
    Public Sub MergeSpalten(ByVal strVonSpalte As String, ByVal strBisSpalte As String)
        Dim Zeile As Integer = m_RowNumInWorkSheet(m_CurWorksheet)
        m_ExcelApp.Range(strVonSpalte & Zeile & ":" & strBisSpalte & Zeile).Select()
        m_ExcelApp.Selection.MergeCells = True
    End Sub

#End Region

#Region "Format"
    Public Sub AutoSizeColumns()
        m_ExcelApp.Cells.EntireColumn.AutoFit()
    End Sub

    Private Const C_SCHUTZVERMERK = "Weitergabe sowie Vervielfältigung dieses Dokumentes, Verwendung und Mit-" & vbNewLine &
                                    "teilung seines Inhaltes sind verboten, soweit nicht ausdrücklich gestattet."


    Private ReadOnly Property Schutzvermerk As String
        Get
            Return C_SCHUTZVERMERK
        End Get
    End Property

    Public Sub Druckbereich(ByVal Anzahlzeilen As Integer, ByVal LetzteSpalte As String)

        Dim printbereich As String = LetzteSpalte & "$" & Anzahlzeilen
        With m_ExcelApp.ActiveSheet.PageSetup
            .PrintTitleRows = "$1:$3"
            .PrintArea = printbereich
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            '.LeftFooter = "&8&D"
            'Schutzvermerk
            .CenterFooter = "&""Arial,Fett""&8" & Schutzvermerk
            .RightFooter = "&8&P von &N Seiten"
            .LeftMargin = m_ExcelApp.Application.InchesToPoints(0.708661417322835)
            .RightMargin = m_ExcelApp.Application.InchesToPoints(0.708661417322835)
            .TopMargin = m_ExcelApp.Application.InchesToPoints(0.78740157480315)
            .BottomMargin = m_ExcelApp.Application.InchesToPoints(0.78740157480315)
            .HeaderMargin = m_ExcelApp.Application.InchesToPoints(0.31496062992126)
            .FooterMargin = m_ExcelApp.Application.InchesToPoints(0.31496062992126)
            .PrintHeadings = False
            .PrintGridlines = True
            .CenterHorizontally = True
            .CenterVertically = True
            .Orientation = Excel.XlPageOrientation.xlLandscape
            .Draft = False
            .PaperSize = Excel.XlPaperSize.xlPaperA3
            .FirstPageNumber = Excel.Constants.xlAutomatic
            .Order = Excel.XlOrder.xlDownThenOver
            .BlackAndWhite = False
            .Zoom = 100
            .PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed
        End With
        m_ExcelApp.ActiveSheet.Range("A1").select
    End Sub

    Public Sub AllFieldsCentered()
        m_ExcelApp.Cells.Select()
        m_ExcelApp.Selection.HorizontalAlignment = Excel.Constants.xlCenter
    End Sub

    Public Sub MakeBorderRange(ByVal strVonSpalte As String, ByVal strBisSpalte As String)
        Dim excelSheet As Excel.Worksheet = m_ExcelApp.ActiveSheet

        Dim rng As Excel.Range = excelSheet.Range(strVonSpalte, strBisSpalte)
        Dim borders As Excel.Borders = rng.Borders
        borders.LineStyle = Excel.XlLineStyle.xlContinuous
        borders.Weight = Excel.XlBorderWeight.xlMedium
        rng.RowHeight = excelSheet.StandardHeight * 2
        rng.EntireRow.WrapText = True
        rng.Cells.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
    End Sub

    Public Sub VerdoppleZeilenHoehe(ByVal _ZeileCurrent As Integer)
        Dim excelSheet As Excel.Worksheet = m_ExcelApp.ActiveSheet
        Dim rng As Excel.Range = excelSheet.Cells(_ZeileCurrent, 1)
        rng.RowHeight = excelSheet.StandardHeight * 2
        rng.EntireRow.WrapText = True
    End Sub

#End Region

#Region "Einfügen"
    Public Sub InsertRow(ByVal _ZeileCurrent As Integer)
        Dim excelSheet As Excel.Worksheet = m_ExcelApp.ActiveSheet
        Dim rng As Excel.Range = excelSheet.Cells(_ZeileCurrent, 1)
        Dim row As Excel.Range = rng.EntireRow
        row.Insert(Excel.XlInsertShiftDirection.xlShiftDown, False)
        m_RowNumInWorkSheet(m_CurWorksheet) = _ZeileCurrent
    End Sub
#End Region

#Region "Einfärben"

    Private Sub Faerbe(ByVal Row As Integer, ByVal Farbe As eExcelFarbe)
        m_ExcelApp.Rows(Row).Font.Color = GetTextFarbe(Farbe)
        m_ExcelApp.Rows(Row).Interior.Color = GetHintergrundFarbe(Farbe)
    End Sub

    Private Sub Faerbe(ByVal Row As Integer, ByVal Col As Integer, ByVal Farbe As eExcelFarbe)
        m_ExcelApp.Cells(Row, Col).Font.Color = GetTextFarbe(Farbe)
        m_ExcelApp.Cells(Row, Col).Interior.Color = GetHintergrundFarbe(Farbe)
    End Sub

#End Region

#Region "Font"
    Private Sub Zeichensatz(ByVal Row As Integer, ByVal Fontname As String, ByVal FontSize As Integer, ByVal Bold As Boolean)
        With m_ExcelApp.Rows(Row).Font
            .Name = Fontname
            .size = FontSize
            .Bold = Bold
        End With
    End Sub

    Private Sub Zeichensatz(ByVal Row As Integer, ByVal Col As Integer, ByVal Fontname As String, ByVal FontSize As Integer, ByVal Bold As Boolean)
        With m_ExcelApp.Cells(Row, Col).Font
            .Name = Fontname
            .size = FontSize
            .Bold = Bold
        End With
    End Sub

#End Region

End Class
