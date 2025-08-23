'GetXMLContent: change the font size of a cell or range of cells
'pUrl: Link of the webservice
'pTag: Tag that will have its content purchased
'Reference: https://ask.libreoffice.org/t/use-of-webservice-within-a-macro-in-libreoffice-6-0-1-1/31030/3
Function GetXMLContent(pUrl as String, pTag as String) As String
    On Error GoTo ErrorHandler

    Dim funtionAccess As Object
    Dim xmlString As String

    functionAccess = createUnoService("com.sun.star.sheet.FunctionAccess")
    xmlString = functionAccess.callFunction("WEBSERVICE",Array(pUrl)) 
    xmlString = functionAccess.callFunction("FILTERXML", array(xmlString, pTag)) 
    GetXMLContent = xmlString

    Exit Function
	ErrorHandler:
    GetXMLContent = "Error " & Err
End Function
