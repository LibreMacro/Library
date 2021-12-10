
'ChangeFontSize: change the font size of a cell or range of cells
'pSheet: Sheet name (text)
'pRange: Range of cells to generate the effect/formatting (text)
'pSize: Font size
Sub ChangeFontSize (pSheet as String, pRange as String, pSize as Integer)

	If CheckIfHasSheet(pSheet) Then
		Cell(pSheet,pRange).CharHeight = pSize
	end if
	
end sub

'ChangeFontColor: change the font color of a cell or range of cells
'pSheet: Sheet name (text)
'pRange: Range of cells to generate the effect/formatting (text)
'pColor: Choose one of the options below or enter RGB code.
' 1) red (Default); 2) blue; 3) yellow; 4) green; 5) black
' example of orange color using RGB: pColor = "RGB(255,165,0)"
Sub ChangeFontColor (pSheet as String, pRange as String, Optional pColor as String)
dim c$(3)

	If CheckIfHasSheet(pSheet) Then
	
		if IsMissing(pColor) then
			c = getColorCode("red") 	' Red (Default)		
		else
			c = getColorCode(pColor)
		end if
	
		'Sheet(pSheet).getCellRangeByName(pRange).CharColor = RGB(c(0),c(1),c(2))
		Cell(pSheet,pRange).CharColor = RGB(c(0),c(1),c(2))

	end if
		
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

	If CheckIfHasSheet(pSheet) Then

		if IsMissing(pColor) then
			c = getColorCode("red") 	' Red (Default)	
		else
			c = getColorCode(pColor)
		end if
		
		Cell(pSheet, pRange).CellBackColor = RGB(c(0),c(1),c(2))
		
	end if

end sub

'ChangeCellStyle: Change the style of a cell or range of cells
'pSheet: Sheet name (text)
'pRange: Range of cells (text)
'pStyle: Name of the new style to be used (text)
Sub ChangeCellStyle(pSheet as String, pCell as String, Optional pStyle As String)

	if IsMissing(pStyle) then
	
		MsgBox("ChangeCellStyle: Please inform in the third parameter which style should be applied")
	
	else 

		If CheckIfHasSheet(pSheet) Then
			Cell(pSheet, pCell).CellStyle = pStyle
		End if
	
	end if
	
End sub

Sub ChangeRowStyle(pSheet as String, pRow as Integer, Optional pStyle As String)

	if IsMissing(pStyle) then

		MsgBox("ChangeRowStyle: Please inform in the third parameter which style should be applied")

	else
	
		If CheckIfHasSheet(pSheet) Then
			Row(pSheet, pRow).CellStyle = pStyle
		End if
		
	end if
	
End sub

Sub ChangeSheetStyle(pSheet as String, Optional pStyle As String)

	if IsMissing(pStyle) then

		MsgBox("ChangeSheetStyle: Please inform in the second parameter which style should be applied")

	else
	
		If CheckIfHasSheet(pSheet) Then
			Sheet(pSheet).CellStyle = pStyle
		End if
		
	end if
	
End sub

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

	If CheckIfHasSheet(pSheet) Then

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

	end if

end sub

'CreateStripedLines: generates the famous effect of striped lines in spreadsheets
'pSheet: Sheet name (text)
'pRange: Range of cells to generate the effect/formatting (text)
sub CreateStripedLines(pSheet as String, pRange as String)
Dim oSel As Object
Dim num As Long

	If CheckIfHasSheet(pSheet) Then

		oSel = 	Sheet(pSheet).getCellRangeByName(pRange)
	
		For num = 0 To oSel.getRows.getCount() - 1 
		
			if num mod 2 = 0 then
				oSel.getCellRangeByPosition(0,num,oSel.getColumns.getCount() -1 , num).CellBackColor = RGB( 230,230,230 )
			else
				oSel.getCellRangeByPosition(0,num,oSel.getColumns.getCount() -1 , num).CellBackColor = RGB( 255,255,255 )
			end if
	
		Next
		
	end if

End Sub

'ChangeFont: Change the font family to a new one
'pSheet: Sheet name (text)
'pRange: Range of cells (text)
'pFont: Name of the new font that will be used (text)
Sub ChangeFont(pSheet as String, pCell as String, pFont As String)

	If CheckIfHasSheet(pSheet) Then
		Cell(pSheet, pCell).CharFontName	= pFont
	End if
	
End Sub

'Under Construction
'Sub CopyFontColor(pSheet as String, pCellToCopy as String, pCellToPaste As String) 
'Dim colorLong As Long
'Dim c$(3)
'	colorLong = Cell(pSheet, pCellToCopy).CharColor
 
' 	c = GetColorCode(  getRGBfromLong( colorLong ) )
 	
' 	ChangeFontColor(pSheet, pCellToPaste,  rgbColor(c(0), c(1), c(2)) )

'End sub

