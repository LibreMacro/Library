REM  *****  BASIC  *****

Sub Main

	BasicLibraries.LoadLibrary("LibreMacro")
	
	
	
	'ConfirmDialog(pQuestion as String, Optional pDialogTitle as String) as Boolean
	'QuestionDialog(pQuestion as String, Optional pDialogTitle as String)
	'RetryDialog(pQuestion as String, Optional pDialogTitle as String) as Boolean
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	for i = 2 to 6 step 1
	
		
	IF Cell("Planilha1", REF(i,2) ).Value >=7 THEN
	
		Cell("Planilha1", REF(i,3)).String = "Aprovado"
	
	ELSE
	
		Cell("Planilha1", REF(i,3)).String = "Reprovado"
	
	END IF

	next
End Sub
