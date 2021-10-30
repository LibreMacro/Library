
'ChangeFontSize: change the font size of a cell or range of cells
'pSheet: Sheet name (text)
'pRange: Range of cells to generate the effect/formatting (text)
'pUnits: Range das celulas para formatar (texto)  (Ex: "A1:C3", "B5:D11")
Sub ChangeFontSize (pSheet as String, pRange as String, pUnits as Integer)

	'Sheet(pSheet).getCellRangeByName(pRange).CharHeight = pUnits
	Cell(pSheet,pRange).CharHeight = pUnits

end sub

'ChangeFontColor: change the font color of a cell or range of cells
'pSheet: Sheet name (text)
'pRange: Range of cells to generate the effect/formatting (text)
'pColor: Choose one of the options below or enter RGB code.
' 1) red (Default); 2) blue; 3) yellow; 4) green; 5) black
' example of orange color using RGB: pColor = "RGB(255,165,0)"
Sub ChangeFontColor (pSheet as String, pRange as String, Optional pColor as String)
dim c$(3)

	if IsMissing(pColor) then
		c = getColorCode("red") 	' Red (Default)		
	else
		c = getColorCode(pColor)
	end if

	'Sheet(pSheet).getCellRangeByName(pRange).CharColor = RGB(c(0),c(1),c(2))
	Cell(pSheet,pRange).CharColor = RGB(c(0),c(1),c(2))
	
end sub

'ChangeCellColor: change the background color of a cell or range of cells
'pSheet: Sheet name
'pRange: Range of cells to generate the effect/formatting (text)
'pColor: Choose one of the options below or enter RGB code.
' 1) red (Default); 2) blue; 3) yellow; 4) green; 5) black
' example of orange color using RGB: pColor = "RGB(255,165,0)"
Sub ChangeCellColor (pSheet as String, pRange as String, Optional pColor as Variant)
dim c$(3)
dim r$(2)

	if IsMissing(pColor) then
		c = getColorCode("red") 	' Red (Default)	
	else
		c = getColorCode(pColor)
	end if
	
	Cell(pSheet, pRange).CellBackColor = RGB(c(0),c(1),c(2))

end sub

'FormatFont: Highlight text with some specific formatting
'pSheet: Sheet name (text)
'pRange: Range das celulas para limpar o conteúdo (texto)
'pOption: Choose one of the options below.
' "U" - Underline
' "B" - Bold (Default)
' "R" - Red color and bold
' "I"   - Italic
' "N" - Removes formatting: underlines, bold and red color
Sub ChangeFontFormat (pSheet as String, pRange as String, Optional pOption as Variant)
dim vOption as String

	if IsMissing(pOption) then
		Sheet(pSheet).getCellRangeByName(pRange).CharWeight  = 150
	else
		vOption = UCase(pOption)
	end if
	
	if vOption = "B" then
		Cell(pSheet,pRange).CharWeight  = 150 
	elseif  vOption = "U" then
		Cell(pSheet,pRange).CharUnderline  = 1
	elseif vOption = "I" then
		Cell(pSheet,pRange).CharPosture = 2
	elseif vOption = "R" then		
		ChangeFontColor(pSheet,pRange, "red")
		Cell(pSheet,pRange).CharWeight = 150
	elseif vOption = "N" then
			Cell(pSheet,pRange).CharWeight  = 100
			Cell(pSheet,pRange).CharPosture = 0
			Cell(pSheet,pRange).CharUnderline  = 0
			ChangeFontColor(pSheet,pRange, "black")
	end if

end sub

'CreateStripedLines: generates the famous effect of striped lines in spreadsheets
'pSheet: Sheet name (text)
'pRange: Range of cells to generate the effect/formatting (text)
sub CreateStripedLines(pSheet as String, pRange as String)
Dim oSel As Object
Dim num As Long

	oSel = 	Sheet(pSheet).getCellRangeByName(pRange)

	For num = 0 To oSel.getRows.getCount() - 1 
	
		if num mod 2 = 0 then
			oSel.getCellRangeByPosition(0,num,oSel.getColumns.getCount() -1 , num).CellBackColor = RGB( 230,230,230 )
		else
			oSel.getCellRangeByPosition(0,num,oSel.getColumns.getCount() -1 , num).CellBackColor = RGB( 255,255,255 )
		end if

	Next

End Sub

'Under Construction
'Sub CopyFontColor(pSheet as String, pCellToCopy as String, pCellToPaste As String) 
'Dim colorLong As Long
'Dim c$(3)
'	colorLong = Cell(pSheet, pCellToCopy).CharColor
 
' 	c = GetColorCode(  getRGBfromLong( colorLong ) )
 	
' 	ChangeFontColor(pSheet, pCellToPaste,  rgbColor(c(0), c(1), c(2)) )

'End sub

