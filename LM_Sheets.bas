REM  *****  BASIC  *****

rem ********************************************** FOLHAS DE CÁLCULO

' Sheet: Retorna a referência de uma folha de cálculo.
' pSheet : nome da folha de cálculo (texto)
FUNCTION Sheet (pSheet as String) as Object

	Sheet = Thiscomponent.Sheets.GetByName(pSheet)

END FUNCTION


rem ********************************************** CÉLULAS

' Cell: Retorna a referência de uma célula.
' pSheet : nome da folha de cálculo (texto)
' pCell: nome da célula (texto)
FUNCTION Cell (pSheet as String, pCell as String) as Object
	
	Cell = Sheet(pSheet).GetCellRangeByName(pCell)

END FUNCTION

 
 ' ActiveSheet: Retorna a planilha ativa (Objeto)
Function ActiveSheet As Object

	ActiveSheet = ThisComponent.getCurrentController.getActiveSheet()
		
End Function


' ActiveSheet: Retorna a planilha ativa (Texto)
Function ActiveSheetName As String

	ActiveSheetName = ThisComponent.getCurrentController.getActiveSheet().getName()
		
End Function


' CreateSheet: Criar uma nova folha de cálculo
' pName: Nome da planilha a ser criada (texto)
Sub CreateSheet(pName As String)
	
	Dim spreadsheet As Object
	
	If not ThisComponent.Sheets.hasByName(pName) Then
		spreadsheet = ThisComponent.createInstance("com.sun.star.sheet.Spreadsheet")
		ThisComponent.Sheets.insertByName(pName, spreadsheet)
	End if
End Sub


' RemoveSheet: Remove uma folha de cálculo
' pName: Nome da planilha a ser excluída
Sub RemoveSheet(pName As String)

	Dim spreadsheet As Object
	
	If ThisComponent.Sheets.hasByName(pName) Then
		ThisComponent.Sheets.removeByName(pName)
	End if

End Sub

' FindTextInCell: Procura texto dentro de determinada célula
' pText : Texto procurado
' pCell : Célula em que se realiza a busca 
FUNCTION FindTextInCell(pText as String, pCell as String) as Boolean

	if InStr( Cell(pCell).String , pText) <> 0 then
		FindTextInTheCell = true
	else
		FindTextInTheCell = false
	end if

END Function

rem ********************************************** SELECIONA UMA OU MAIS CÉLULAS

' SelectCell: Seleciona uma ou mais células
' pSheet : nome da folha de cálculo (texto)
' pCellRange: nome da célula ou do intervalo de células (texto)
Sub SelectCell(pSheet as String, pCellRange As String)

	Dim Cells As Object
	Cells = Cell(pSheet, pCellRange)
	ThisComponent.getCurrentController.select(Cells)

End Sub


rem ********************************************** INSERIR LINHAS

'InsertRows: Inserir Linhas
'IndexL: Index da linha (Começa por 0)
'Units: Unidades a serem inseridas

Sub InsertRows (pSheet as String, IndexL as Integer, Units as Integer)

	Sheet(pSheet).Rows.insertByIndex(IndexL, Units)
	
END sub

rem ********************************************** INSERIR COLUNAS

'InsertColumns: Inserir Colunas
'IndexC: Index da coluna (Começa por 0)
'Units: Unidades a serem inseridas

Sub InsertColumns (pSheet as String, IndexC as Integer, Units as Integer)

	Sheet(pSheet).Columns.insertByIndex(IndexC, Units)
	
END sub

rem ********************************************** EXCLUIR LINHAS

'DeleteRows: Excluir Linhas
'IndexL: Index da linha (Começa por 0)
'Units: Unidades a serem inseridas

Sub DeleteRows (pSheet as String, IndexL as Integer, Units as Integer)

	Sheet(pSheet).Rows.removeByIndex(IndexL, Units)
	
End sub


rem ********************************************** EXCLUIR COLUNAS

'DeleteColumns: Excluir Colunas
'IndexC: Index da coluna (Começa por 0)
'Units: Unidades a serem inseridas

Sub DeleteColumns (pSheet as String, IndexC as Integer, Units as Integer)

	Sheet(pSheet).Columns.removeByIndex(IndexC, Units)
	
End sub


' Referência: https://www.schiavinatto.com/mundolibre/biblioteca/aprendiendo/6.4.6---trabajando-con-notas.html
Sub InsertCellNote(pSheet as String, pCell as Object, pNote as String)

	Dim cellNotes As Object
	cellNotes = Sheet(pSheet).getAnnotations()
	cellNotes.insertNew( pCell.getCellAddress(), pNote)
	
End Sub


Sub RemoveCellNote(pSheet as String, pCell as Object)

	Dim cellNotes As Object
	Dim oNotas As Object
	Dim oNota As Object
	Dim co1 As Long
	
	oNotas = Sheet(pSheet).getAnnotations()
	oCelda = pCell
	' Referência: https://www.schiavinatto.com/mundolibre/biblioteca/aprendiendo/6.4.6---trabajando-con-notas.html (início)
	If oNotas.getCount() > 0 Then
		For co1 = 0 To oNotas.getCount - 1
			oNota = oNotas.getByIndex( co1 )
			If oNota.getPosition.Column = oCelda.getCellAddress.Column And oNota.getPosition.Row = oCelda.getCellAddress.Row Then
				oNotas.removeByIndex( co1 )
				Exit Sub
			End If
			Next co1
	end if
	' Referência: https://www.schiavinatto.com/mundolibre/biblioteca/aprendiendo/6.4.6---trabajando-con-notas.html (fim)


	cellNotes.RemoveByAddress(pCell.getCellAddress())
	'oDirCelda.Column = 2
	'oDirCelda.Row = 10
	'cellNotes.insertNew( oDirCelda, "Teste")  
	
End Sub

rem ********************************************** Limpa Conteudo das Células

'ClearContents: Limpa Conteúdo das células
'pSheet: nome da planilha
'Range: Range das celulas para limpar o conteúdo
'Argumento: 7 = 1 + 2 + 4 (Valor + Texto + Data/Hora)'

Sub ClearContents (pSheet as String, Range as String)

	Sheet(pSheet).getCellRangeByName(Range).ClearContents(7)

End sub

rem ********************************************** Ordem Crescente
'pSheet: nome da planilha
'pRange: Range das celulas para ordenar
'pIndexC: index da coluna de referência para ordenação começa por 0

Sub SortAsc (pSheet as String, pRange as String, pIndexC as Integer)

  Dim oSortFields(0) As New com.sun.star.util.SortField
 
  Dim oSortDesc(0) As New com.sun.star.beans.PropertyValue
   
  oSortFields(0).Field = pIndexC
  oSortFields(0).SortAscending = True

  oSortDesc(0).Name = "SortFields"
  oSortDesc(0).Value = oSortFields()

  Sheet(pSheet).getCellRangeByName(pRange).Sort(oSortDesc())
 
End sub

rem ********************************************** Ordem Decrescente
'pSheet: nome da planilha
'pRange: Range das celulas para ordenar
'pIndexC: index da coluna de referência para ordenação começa por 0


Sub SortDes (pSheet as String, pRange as String, pIndexC as Integer)

  Dim oSortFields(0) As New com.sun.star.util.SortField
 
  Dim oSortDesc(0) As New com.sun.star.beans.PropertyValue

  oSortFields(0).Field = pIndexC
  oSortFields(0).SortAscending = False

  oSortDesc(0).Name = "SortFields"
  oSortDesc(0).Value = oSortFields()
  
  Sheet(pSheet).getCellRangeByName(pRange).Sort(oSortDesc())
 
end sub


