REM  *****  BASIC  *****

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
	elseif UCase(mid(pColor,1,3)) = "RGB" then
		b = mid(pColor,5,11)
	 	c = split(b, ",",3)
	end if
	
	GetColorCode = c

END FUNCTION

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

end sub

