REM  *****  BASIC  *****

rem ********************************************** CRIAR JANELAS DE DIÁLOGO

rem **************************** Funções simplificadas
' ConfirmDialog: Caixa de diálogo contendo uma pergunta e duas opções: Ok e Cancelar.
' pQuestion: Pergunta exibida na caixa de diálogo (texto)
' pDialogTitle (Opcional) : Título da caixa de diálogo (texto)
FUNCTION ConfirmDialog(pQuestion as String, Optional pDialogTitle as String) as Boolean

	if Dialog(pQuestion, pDialogTitle, "confirmDialog") = true then
		ConfirmDialog =  true
	else
		ConfirmDialog = false
	end if

END FUNCTION

' RetryDialog: Caixa de diálogo contendo uma pergunta e duas opções: "Tentar de novo" e "Cancelar".
' pQuestion: Pergunta exibida na caixa de diálogo (texto)
' pDialogTitle (Opcional) : Título da caixa de diálogo (texto)
FUNCTION RetryDialog(pQuestion as String, Optional pDialogTitle as String) as Boolean

	if Dialog(pQuestion, pDialogTitle, "retryDialog") = true then
		ConfirmDialog =  true
	else
		ConfirmDialog = false
	end if

END FUNCTION


' QuestionDialog: Caixa de diálogo contendo uma pergunta e duas opções: "Sim" e "Não"
' pQuestion: Pergunta exibida na caixa de diálogo (texto)
' pDialogTitle (Opcional) : Título da caixa de diálogo (texto)
FUNCTION QuestionDialog(pQuestion as String, Optional pDialogTitle as String) as Boolean

	if Dialog(pQuestion, pDialogTitle, "questionDialog") = true then
		ConfirmDialog =  true
	else
		ConfirmDialog = false
	end if

END FUNCTION


rem **************************** Função Geral
' Dialog: Caixa de diálogo contendo uma pergunta e duas opções.
' pQuestion: Pergunta exibida na caixa de diálogo
' pDialogTitle: Título da caixa de diálogo (Opcional)
' -------------
' pDialogType: Tipo de Caixa de diálogo (Opcional). Escolha:
' 	questionDialog -(Botão sim e não) - Valor Padrão
' 	confirmDialog - (Botão ok e cancelar)
' 	retryDialog -   (Botão "Tentar de novo" e cancelar)
' -------------
' pDefaultButton: Escolher o Botão Padrão
' 	firstBtnDefault - Primeiro Botão (Valor Padrão)
' 	secondBtnDefault - Segundo Botão;
' -------------
' pIcon: Escolher o ícone da caixa de diálogo: stopIcon, questionIcon, exclamationIcon, informationIcon
' -------------
' Obs: Para "Caixa de Diálogo" personalizada, consulte:
' https://help.libreoffice.org/6.2/en-US/text/sbasic/shared/03010102.html
FUNCTION Dialog(pQuestion as String, Optional pDialogTitle as String, Optional pDialogType as String, Optional pDefaultButton as String, Optional pIcon as String) as Boolean

	DIM parameters, result, selectedButton, selectedIcon as Integer
	
	If InStr(pDefaultButton , "second") <> 0 then
		selectedButton = 256 
	ELSE 
		selectedButton = 128
	END IF

	If InStr(pIcon , "stop") <> 0 then
		selectedIcon = 16
	ElseIf InStr(pIcon , "exclamation") <> 0 then
		selectedIcon = 48
	ElseIf  InStr(pIcon , "information") <> 0 then
		selectedIcon = 64 	
	Else
		selectedIcon = 32	
	END IF	
	
	If InStr(pDialogType , "confirm") <> 0 then
		parameters =  selectedButton + selectedIcon + 1	
		result = 1	
	ElseIf InStr(pDialogType , "retry") <> 0 then	
		parameters =  selectedButton + selectedIcon + 5	
		result = 4	
	ELSE						'ConfirmDialog	
		parameters =  selectedButton + selectedIcon + 4
		result = 6		
	END IF

	IF MsgBox (pQuestion, parameters, pDialogTitle) = result then 
		Dialog = true
	ELSE 
		Dialog = false
	END IF

END FUNCTION


' Dialog3: Caixa de diálogo contendo uma pergunta e três opções.
' pQuestion: Pergunta exibida na caixa de diálogo
' pDialogTitle: Título da caixa de diálogo (Opcional)
' -------------
' pDialogType: Tipo de Caixa de diálogo (Opcional). Escolha:
' 	confirmDialog - Botões Yes, No, and Cancel buttons  - Valor Padrão
' 	retryDialog   - Botões "Abortar", "Tentar de novo" e "Ignorar"
' -------------
' pDefaultButton: Escolher o Botão Padrão
' 	firstBtnDefault - Primeiro Botão (Valor Padrão)
' 	secondBtnDefault - Segundo Botão;
' -------------
' pIcon: Escolher o ícone da caixa de diálogo: stopIcon, questionIcon, exclamationIcon, informationIcon
' -------------
' Obs: Para "Caixa de Diálogo" personalizada, consulte:
' https://help.libreoffice.org/6.2/en-US/text/sbasic/shared/03010102.html
' ---------------------------
'Exemplo de uso:
'DIM resultado as String
'	
'	resultado = Dialog3("Você tem certeza ?")
'	
'	Select Case resultado
'    Case "No":
'    	'Ação no caso de não
'    Case "Yes"
'    	'Ação no caso de sim
'    Case "Cancel"
'        'Ação no caso de cancelar	
'    End Select

FUNCTION Dialog3(pQuestion as String, Optional pDialogTitle as String, Optional pDialogType as String, Optional pDefaultButton as String, Optional pIcon as String) as String
	DIM parameters, result, selectedButton, selectedIcon, btn1, btn2, btn3 as Integer
	
	If InStr(pDefaultButton , "second") <> 0 then
		selectedButton = 256 
	Elseif  InStr(pDefaultButton , "third") <> 0 then 
		selectedButton = 512
	ELSE 
		selectedButton = 128
	END IF

	If InStr(pIcon , "stop") <> 0 then
		selectedIcon = 16
	ElseIf InStr(pIcon , "exclamation") <> 0 then
		selectedIcon = 48
	ElseIf  InStr(pIcon , "information") <> 0 then
		selectedIcon = 64 	
	Else
		selectedIcon = 32	
	END IF	
	
	If InStr(pDialogType , "retry") <> 0 then	    ' Display Abort, Retry, and Ignore buttons.
		parameters =  selectedButton + selectedIcon + 5	

	Else 'InStr(pDialogType , "confirm") <> 0 then 	' Display Yes, No, and Cancel buttons.
		parameters =  selectedButton + selectedIcon + 3	
		
	END IF
	

	result = MsgBox (pQuestion, parameters, pDialogTitle)
	
	Select Case result
	Case 2: 
		Dialog3 = "Cancel"
	Case 3: 
		Dialog3 = "Abort"
	Case 4: 
		Dialog3 = "Retry"
	Case 5:
		Dialog3 = "Ignore"
	Case 6:
		Dialog3 = "Yes"
	Case 7:
		Dialog3 = "No"
	
	End Select 


END FUNCTION



