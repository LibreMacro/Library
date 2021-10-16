' Sheet: Returns the reference of a worksheet (Object).
' pSheet : worksheet name (text)
FUNCTION Sheet (pSheet as String) as Object

	Sheet = Thiscomponent.Sheets.GetByName(pSheet)

END FUNCTION


rem ************************************************* CELLS

' Cell: Returns a cell reference (Object).
' pSheet :Sheet name (text)
' pCell: Cell or range of cells (text)
FUNCTION Cell (pSheet as String, pCell as String) as Object
	
	if UCase(mid(pCell,1,3)) = "REF" then
	
		r = getCellRef(pCell)
	
		Cell = Sheet(pSheet).getCellByPosition(r(1),r(0))	
	
	else	
	
		Cell = Sheet(pSheet).GetCellRangeByName(pCell)
	
	end if
	
END FUNCTION

 
 ' ActiveSheet: Returns the active sheet reference (Object)
Function ActiveSheet As Object

	ActiveSheet = ThisComponent.getCurrentController.getActiveSheet()
		
End Function

' ActiveSheet: Returns the active sheet reference (text)
Function ActiveSheetName As String

	ActiveSheetName = ThisComponent.getCurrentController.getActiveSheet().getName()
		
End Function


' CreateSheet: Create a new spreadsheet
' pName: Name of the sheet to be created (text)
Sub CreateSheet(pName As String)
	
	Dim spreadsheet As Object
	
	If not ThisComponent.Sheets.hasByName(pName) Then
		spreadsheet = ThisComponent.createInstance("com.sun.star.sheet.Spreadsheet")
		ThisComponent.Sheets.insertByName(pName, spreadsheet)
	End if
	
End Sub


' RemoveSheet: Remove a spreadsheet
' pName: Sheet name to be deleted (text)
Sub RemoveSheet(pName As String)

	Dim spreadsheet As Object
	
	If ThisComponent.Sheets.hasByName(pName) Then
		ThisComponent.Sheets.removeByName(pName)
	End if

End Sub

' FindTextInCell: Search for text within a certain cell
' pText : Text to be searched (text)
' pCell : Cell in which the search takes place (text)
FUNCTION FindTextInCell(pText as String, pCell as String) as Boolean

	if InStr( Cell(pCell).String , pText) <> 0 then
		FindTextInTheCell = true
	else
		FindTextInTheCell = false
	end if

END Function

rem ********************************************** Select one or more cells

' SelectCell: Select one or more cells
' pSheet : Sheet name (text)
' pCellRange: cell name or cell range name (text)
Sub SelectCell(pSheet as String, pCellRange As String)

	Dim Cells As Object
	Cells = Cell(pSheet, pCellRange)
	ThisComponent.getCurrentController.select(Cells)

End Sub


rem ********************************************** INSERT ROWS

'InsertRows: Function for inserting rows at a certain position within a spreadsheet
'pSheet: Sheet name (text)
'pIndex: Insertion position:  Line number (number greater than zero)
'pUnits: quantity to be added (number greater than zero)
Sub InsertRows (pSheet as String, pIndex as Integer, pUnits as Integer)

	Sheet(pSheet).Rows.insertByIndex(pIndex - 1, pUnits)
	
END sub

rem ********************************************** INSERT COLUMNS

'InsertColumns: Insert Columns
'pSheet: Sheet name (text)
'pIndex: Insertion position: Column number (Ex: Column A -> 1, Column B -> 2, Column C -> 3)
'pUnits: Units to be inserted (number greater than zero)
Sub InsertColumns (pSheet as String, pIndex as Integer, pUnits as Integer)

	Sheet(pSheet).Columns.insertByIndex(pIndex - 1, pUnits)
	
END sub

rem ************************************************* DELETE LINES

'DeleteRows: Delete Rows
'pSheet: Desired spreadsheet (text)
'pIndex: Delete position: Row number (number greater than zero)
'pUnits: Units to be inserted (number greater than zero)
Sub DeleteRows (pSheet as String, pIndex as Integer, pUnits as Integer)

	Sheet(pSheet).Rows.removeByIndex(pIndex - 1, pUnits)
	
End sub


rem ************************************************* DELETE COLUMNS

'DeleteColumns: Delete Columns
'pSheet: Desired spreadsheet (text)
'pIndex: Exclusion position: Column number (Ex: Column A -> 1, Column B -> 2, Column C -> 3)
'pUnits: Units to be inserted (number greater than zero)
Sub DeleteColumns (pSheet as String, pIndex as Integer, pUnits as Integer)

	Sheet(pSheet).Columns.removeByIndex(pIndex - 1, pUnits)
	
End sub

'InsertCellNote: Inserts annotation into a cell
'pSheet: Desired spreadsheet (text)
'pCell: Cell where the insertion will be performed (Ex: Cell A2 -> "A2", Cell D3 -> "D3") (text)
'pNote: Text containing the annotation
Sub InsertCellNote(pSheet as String, pCell as String, pNote as String)

	Dim vCellNotes As Object
	Dim vCell as Object
	
	vCell = Cell(pSheet, pCell)
	 
	vCellNotes = Sheet(pSheet).getAnnotations()
	vCellNotes.insertNew(vCell.getCellAddress(), pNote)
	
End Sub

'RemoveCellNote: Inserts note into a cell
'pSheet: Desired spreadsheet (text)
'pCell: Cell where the annotation will be deleted (Ex: Cell A2 -> "A2", Cell D3 -> "D3") (text)
Sub RemoveCellNote(pSheet as String, pCell as Object)

	Dim cellNotes As Object
	Dim oNotas As Object
	Dim oNota As Object
	Dim co1 As Long
	
	oNotas = Sheet(pSheet).getAnnotations()
	oCelda = pCell
	' Reference: https://www.schiavinatto.com/mundolibre/biblioteca/aprendiendo/6.4.6---trabajando-con-notas.html (início)
	If oNotas.getCount() > 0 Then
		For co1 = 0 To oNotas.getCount - 1
			oNota = oNotas.getByIndex( co1 )
			If oNota.getPosition.Column = oCelda.getCellAddress.Column And oNota.getPosition.Row = oCelda.getCellAddress.Row Then
				oNotas.removeByIndex( co1 )
				Exit Sub
			End If
			Next co1
	end if
	' Reference: https://www.schiavinatto.com/mundolibre/biblioteca/aprendiendo/6.4.6---trabajando-con-notas.html (fim)

	cellNotes.RemoveByAddress(pCell.getCellAddress()) 
	
End Sub

rem ************************************************* Cleans the contents of the cells

'ClearContents: Clear existing content in cells
'This version only excludes values, texts and date/time keeping, therefore,
'the formulas.
'pSheet: Sheet name (text)
'pRange: Range of cells to clear content (text)
'pOption (optional):
' small - Erases values, texts and date/time information
' medium - In addition to what content does, it also erases formulas
' all - Erases everything in the cell (formats, annotations, formulas, content, etc.) 
Sub ClearContents (pSheet as String, pRange as String, Optional pOpcao as String)
dim vNumber as Integer
	
	vNumber = 7
	
	if pOpcao = "medium" then
		vNumber = 23
	elseif pOpcao = "large" then
		vNumber = 1023
	end if
	
	'Sheet(pSheet).getCellRangeByName(pRange).ClearContents(vNumber)
	Cell(pSheet,pRange).ClearContents(vNumber)

End sub

rem ********************************************* Ascending Order
'pSheet: sheet name
'pRange: Range of cells to sort (text) (Ex: "A1:C3", "B5:D11")
'pIndex: Reference column number (Ex: 1 - First column, 2 - second column etc.)
Sub SortAsc (pSheet as String, pRange as String, pIndex as Integer)

  Dim oSortFields(0) As New com.sun.star.util.SortField
 
  Dim oSortDesc(0) As New com.sun.star.beans.PropertyValue
   
  oSortFields(0).Field = pIndex - 1
  oSortFields(0).SortAscending = True

  oSortDesc(0).Name = "SortFields"
  oSortDesc(0).Value = oSortFields()

  Sheet(pSheet).getCellRangeByName(pRange).Sort(oSortDesc())
 
End sub

rem ************************************************* Descending order
'pSheet: sheet name (text)
'pRange: Range of cells to sort (text) (Ex: "A1:C3", "B5:D11")
'pIndex: Reference column number (Ex: 1 - First column, 2 - second column etc.)
Sub SortDesc (pSheet as String, pRange as String, pIndex as Integer)

  Dim oSortFields(0) As New com.sun.star.util.SortField
 
  Dim oSortDesc(0) As New com.sun.star.beans.PropertyValue

  oSortFields(0).Field = pIndex
  oSortFields(0).SortAscending = False

  oSortDesc(0).Name = "SortFields"
  oSortDesc(0).Value = oSortFields()
  
  Sheet(pSheet).getCellRangeByName(pRange).Sort(oSortDesc())
 
end sub


