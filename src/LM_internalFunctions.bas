Global vFoundCell as String
Global vFoundRow as Integer

FUNCTION GetColorCode (pColor as String) as Object
dim c$(3)

	if UCase(pColor) = "RED" then
		c(0) = 255
		c(1) = 0
		c(2) = 0
	elseif UCase(pColor) = "BLUE" then
		c(0) = 0
		c(1) = 0
		c(2) = 255	
	elseif UCase(pColor) = "YELLOW" then
		c(0) = 255
		c(1) = 255
		c(2) = 0
	elseif UCase(pColor) = "GREEN" then
		c(0) = 0
		c(1) = 255
		c(2) = 0
	elseif UCase(pColor) = "BLACK" then
		c(0) = 0
		c(1) = 0
		c(2) = 0
		elseif UCase(pColor) = "WHITE" then
		c(0) = 255
		c(1) = 255
		c(2) = 255
	elseif UCase(mid(pColor,1,3)) = "RGB" Then 
	pColor = Replace(pColor, " ", "")
		b = mid(pColor,5,11)
		b2 =  split(b, ")",2)
	 	c = split(b2(0), ",",3)
	 	'c(2) = split( c, ")", 1)
	end if
	
	GetColorCode = c

END FUNCTION

Function CheckIfHasSheet(pSheet as String) as Boolean
	
	If ThisComponent.Sheets.hasByName(pSheet) Then
		 
			CheckIfHasSheet = True
			
		else
		
			MsgBox("Warning: there is no spreadsheet named '" & pSheet & "' !" )
			CheckIfHasSheet = False
				
	end if
		
end Function

function GetCellRef(pRef as String) as Object
dim r$(2)

	if UCase(mid(pRef,1,3)) = "REF" then
		b = mid(pRef,5, len(pRef) - 1  ) 
	 	r = split(b, ",",2)
	end if
	
	r(0) = Cint(r(0)) - 1
	r(1) = Cint(r(1)) - 1
	
	GetCellRef = r

end function

function Ref(pLinha as Integer, pColuna) as String

	Ref = "REF(" & CStr(pLinha) & "," & CStr(pColuna) & ")"

end Function

function rgbColor(pRed as Integer, pGreen As Integer, pBlue As Integer) as String

	rgbColor = "RGB(" & CStr(pRed) & "," & CStr(pGreen) &  "," & CStr(pBlue) & ")"

end function

sub import(pName as String)
		         
	if GlobalScope.BasicLibraries.hasByName(pName) then
	
		if Not GlobalScope.BasicLibraries.isLibraryLoaded(pName) then
	
				GlobalScope.BasicLibraries.loadLibrary(pName)
				
		end if
			
	elseif BasicLibraries.hasByName(pName) then
		
		if Not BasicLibraries.isLibraryLoaded(pName) then
		
			BasicLibraries.loadLibrary(pName)
			
		end if
		
	else
	
		msgbox("Não foi possível importar a biblioteca")
		
	end if

end Sub

Function getRGBfromLong(pEntrada As Long) As String
dim b As Integer
dim g As Integer
dim r As Integer
Dim saldo As Long

	saldo = pEntrada

		If saldo / 65536 > 1 Then
			r = ToFloor(saldo / 65536)
			saldo = saldo - 65536*r
		Else
			r = 0 
		End If
		
		If saldo / 256 > 1 Then
			g = ToFloor(saldo / 256)
			saldo = saldo - 256*g 
		Else
			g = 0
		End If
		
		If saldo > 0 then
			b = saldo
		Else
			b = 0
		End if

	 	getRGBfromLong = "RGB("& Cstr( r ) &","& CStr( g ) & "," & CStr( b ) & ")"

End Function

Private Function GetSystemPathSeparator() As String
    ' Heurística simples: URL do perfil usa "/" no mac/linux; no Windows usamos "\"
    If InStr(1, LCase$(Environ("OS")), "windows") > 0 Then
        GetSystemPathSeparator = "\"
    Else
        GetSystemPathSeparator = "/"
    End If
End Function

Private Function JoinPath(folder As String, name As String) As String
    Dim sep As String : sep = GetSystemPathSeparator()
    If Right$(folder,1) = sep Then
        JoinPath = folder & name
    Else
        JoinPath = folder & sep & name
    End If
End Function


Private Function GetUserDocumentsDir() As String
    Dim home As String, isWin As Boolean
    isWin = InStr(1, LCase$(GetSystemPathSeparator()), "\") > 0

    If isWin Then
        ' Windows: %USERPROFILE%\Documents (padrão)
        home = Environ("USERPROFILE")
        If Len(home) = 0 Then home = Environ("HOMEDRIVE") & Environ("HOMEPATH")
        GetUserDocumentsDir = JoinPath(home, "Documents")
    Else
        ' macOS/Linux: $HOME/Documents
        home = Environ("HOME")
        If Len(home) = 0 Then home = GetHomeFromProfile()
        GetUserDocumentsDir = JoinPath(home, "Documents")
    End If
End Function


' Reference: Adaptation from
' https://ask.libreoffice.org/t/lo-calc-basic-macro-how-to-work-with-integers/23049
function ToFloor( v as Double ) as Long
	dim aux as Long
	aux = v
	if aux > v then 
		ToFloor = aux-1
	else 
		ToFloor = aux
	EndIf

end function	

' Reference:  https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Strings_(Runtime_Library)
Function Replace(Source As String, Search As String, NewPart As String)
  Dim Result As String  
  Result = join(split(Source, Search), NewPart)
  Replace = Result
End Function

Function getFoundCell as String

	getFoundCell = vFoundCell
	
end function


Sub setFoundCell(pEntrada as String) 

	vFoundCell = pEntrada

end Sub

Sub setFoundRow(pEntrada as Integer) 

	vFoundRow = pEntrada

end Sub

Function getFoundRow as Integer

	getFoundRow = vFoundRow
	
end function

Function getOperType(num1 As Integer, num2 As Integer) As String

dim vOperType As String

	If num1 < num2 Then
 		vOperType = "add"
 	ElseIf num1 > num2 then
 		vOperType = "minus"
 	End if	
 	
 	getOperType = vOperType

End Function

Function NumberTransformation(pNum1 As Integer, pNum2 As Integer, pOperType As String, pInc As Integer) As Integer	

	If pOperType = "add" Then
		if pNum1 < pNum2 Then
			pNum1 = pNum1 + pInc
		Else
			pNum1 = pNum2
		End If
	Else 'operType = "minus"
		if pNum1 > pNum2 Then
			pNum1 = pNum1 - pInc
		Else
			pNum1 = pNum2
		End If
	End If
	
	NumberTransformation = pNum1

End Function

Function GetPythonUserScriptsPath As String
    Dim ps As Object
    Dim sPath As String
    
    ' Acessa o objeto de configuração de caminho do usuário
    ps = createUnoService("com.sun.star.util.PathSettings")
    
    ' Obtém o caminho do diretório de usuário e anexa os subdiretórios dos scripts Python
    sPath = ps.UserConfig & "/Scripts/python"
    
    ' Exibe o caminho
    GetPythonUserScriptsPath = sPath
End Function


Function SheetExists(ByVal sName As String) As Boolean
    On Error GoTo Nope
    SheetExists = ThisComponent.Sheets.hasByName(sName)
    Exit Function
Nope:
    SheetExists = False
End Function

'Function ValueWhenParameterIsMissing(pEntrada, pDefault, pValue)
'
'	if IsMissing(pEntrada) Then
'		vRetorno = pDefault
'	Else
'		vRetorno = pValue
'	end If
'	
'	ValueWhenParameterIsMissing = vRetorno
'
'End Function



