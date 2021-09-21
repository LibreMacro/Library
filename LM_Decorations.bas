'ChangeFontSize: change the font size of a cell or range of cells
'pSheet: Nome da planilha (texto)
'pRange: Range das celulas para limpar o conteúdo (texto)
'pUnits: Range das celulas para formatar (texto)  (Ex: "A1:C3", "B5:D11")
Sub ChangeFontSize (pSheet as String, pRange as String, pUnits as Integer)

	Sheet(pSheet).getCellRangeByName(pRange).CharHeight = pUnits

end sub

'ChangeFontColor: change the font color of a cell or range of cells
'pSheet: Nome da planilha (texto)
'pRange: Range das celulas para limpar o conteúdo (texto)
'pColor: Choose one of the options below or enter RGB code.
' 1) red (Default); 2) blue; 3) yellow; 4) green; 5) black
' example of orange color using RGB: pColor = "RGB(255,165,0)"
Sub ChangeFontColor (pSheet as String, pRange as String, Optional pColor as String)
dim c$(3)

	if IsMissing(pColor) then
		' Red (Default)
		c(0) = 255
		c(1) = 0
		c(2) = 0	
	else
		c = getColorCode(pColor)
	end if
	
	Sheet(pSheet).getCellRangeByName(pRange).CharColor = RGB(c(0),c(1),c(2))

end sub

'ChangeCellColor: change the background color of a cell or range of cells
'pSheet: Nome da planilha (texto)
'pRange: Range das celulas para limpar o conteúdo (texto)
'pColor: Choose one of the options below or enter RGB code.
' 1) red (Default); 2) blue; 3) yellow; 4) green; 5) black
' example of orange color using RGB: pColor = "RGB(255,165,0)"
Sub ChangeCellColor (pSheet as String, pRange as String, Optional pColor as Variant)
dim c$(3)

	if IsMissing(pColor) then
		' Red (Default)
		c(0) = 255
		c(1) = 0
		c(2) = 0	
	else
		c = getColorCode(pColor)
	end if
	
	Sheet(pSheet).getCellRangeByName(pRange).CellBackColor = RGB(c(0),c(1),c(2))

end sub



