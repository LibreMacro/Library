REM  *****  BASIC  *****

Option Explicit

' Exporta a planilha (ou um intervalo) para PDF.
' Ex.: ExportToPDF("Planilha1", "/Users/name/relatorio.pdf")
' Ex.: ExportToPDF("Planilha1", "/Users/name/Desktop/relatorio.pdf", "A1:D40")
Sub ExportToPDF(pSheet As String, Optional pPath As String, Optional pRange As String, Optional pLandscape As Boolean, Optional pQuality As Integer, Optional pPageRange As String)

    Dim oDoc As Object, oSheet As Object, oRange As Object
    Dim args(1) As New com.sun.star.beans.PropertyValue
    Dim pdf(5) As New com.sun.star.beans.PropertyValue
    Dim url As String

	If IsMissing(pPath) Then
		pPath = GetUserDocumentsDir & "export.pdf"
	end if
	
    oDoc = ThisComponent
    oSheet = oDoc.Sheets(pSheet)

    ' 1) Orientação (se informado)
    If Not IsMissing(pLandscape) Then
        Dim styles As Object, style As Object
        styles = oDoc.StyleFamilies.getByName("PageStyles")
        style = styles.getByName(oSheet.PageStyle)
        style.IsLandscape = pLandscape
    End If

    ' 2) Se pRange foi passado, pega o intervalo como "Selection"
    If Not IsMissing(pRange) And Len(pRange) > 0 Then
        oRange = oSheet.getCellRangeByName(pRange)
        pdf(0).Name = "Selection" : pdf(0).Value = oRange
    Else
        ' (opcional) se quiser limitar a páginas, use PageRange
        If Not IsMissing(pPageRange) And Len(pPageRange) > 0 Then
            pdf(0).Name = "PageRange" : pdf(0).Value = pPageRange   ' ex: "1-2,4"
        Else
            pdf(0).Name = "ExportFormFields" : pdf(0).Value = False ' placeholder
        End If
    End If

    ' 3) Qualidade de compressão (90 padrão se informado)
    'If Not IsMissing(pQuality) And pQuality >= 1 And pQuality <= 100 Then
    '    pdf(1).Name = "Quality" : pdf(1).Value = pQuality
    'Else
    '    pdf(1).Name = "Quality" : pdf(1).Value = 90
    'End If

    ' (algumas opções úteis e seguras)
    pdf(2).Name = "UseLosslessCompression" : pdf(2).Value = False
    pdf(3).Name = "ReduceImageResolution"  : pdf(3).Value = True
    pdf(4).Name = "MaxImageResolution"     : pdf(4).Value = 300
    pdf(5).Name = "ExportBookmarks"        : pdf(5).Value = False

    ' 4) Monta os argumentos do storeToURL
    args(0).Name = "FilterName" : args(0).Value = "calc_pdf_Export"
    args(1).Name = "FilterData" : args(1).Value = pdf()

    ' 5) Salva
    url = ConvertToURL(pPath)
    oDoc.storeToURL(url, args())
End Sub
