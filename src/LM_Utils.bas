REM  *****  BASIC  *****

Option Explicit

' Exporta a planilha (ou um intervalo) para PDF.
' Ex.: ExportToPDF("Planilha1", "/Users/name/relatorio.pdf")
' Ex.: ExportToPDF("Planilha1", "/Users/name/Desktop/relatorio.pdf", "A1:C3")
Sub ExportToPDF(pSheet As String, Optional pPath As String, Optional pRange As String, Optional pQuality As Variant)

    Dim oDoc As Object, oSheets as Object, oSheet As Object, oRange As Object
    Dim args(1) As New com.sun.star.beans.PropertyValue
    Dim pdf(5) As New com.sun.star.beans.PropertyValue
    Dim url As String
    Dim i as Integer

	If IsMissing(pPath) Then
		pPath = GetUserDocumentsDir &  GetSystemPathSeparator & "export_LibreMacro_PDF.pdf"
	end if
	
    oDoc = ThisComponent

    If Not IsMissing(pRange) And Len(pRange) > 0 Then

        oRange = Cell(pSheet,pRange)
        oDoc.CurrentController.select( oRange )
    Else
        ' Caso contrário, seleciona a planilha inteira.
        oDoc.CurrentController.select( Sheet(pSheet) )
    End If
   
    pdf(0).Name = "Selection" : pdf(0).Value = oDoc.CurrentController.Selection
    pdf(1).Name = "UseLosslessCompression" : pdf(1).Value = False
    pdf(2).Name = "ReduceImageResolution"  : pdf(2).Value = True
    pdf(3).Name = "MaxImageResolution"     : pdf(3).Value = 300
    pdf(4).Name = "ExportBookmarks"        : pdf(4).Value = False

	If IsMissing(pQuality) Then
		pdf(5).Name = "Quality"
        pdf(5).Value = 90 
	else
		pdf(5).Name = "Quality"
        pdf(5).Value = CInt(pQuality)
    end if

    args(0).Name = "FilterName" : args(0).Value = "calc_pdf_Export"
    args(1).Name = "FilterData" : args(1).Value = pdf()


    url = ConvertToURL(pPath)
    oDoc.storeToURL(url, args())
End Sub

