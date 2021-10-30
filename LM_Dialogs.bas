'ConfirmDialog: Dialog box containing a question and two options: Ok and Cancel.
'pQuestion: Question displayed in dialog (text)
'pDialogTitle (Optional): Dialog title (text)
FUNCTION ConfirmDialog(pQuestion as String, Optional pDialogTitle as String) as Boolean

	if Dialog(pQuestion, pDialogTitle, "confirmDialog") = true then
		ConfirmDialog =  true
	else
		ConfirmDialog = false
	end if

END FUNCTION

' RetryDialog: Dialog containing a question and two options: "Retry" and "Cancel".
' pQuestion: Question displayed in dialog (text)
' pDialogTitle (Optional) : Dialog title (text)
FUNCTION RetryDialog(pQuestion as String, Optional pDialogTitle as String) as Boolean

	if Dialog(pQuestion, pDialogTitle, "retryDialog") = true then
		RetryDialog =  true
	else
		RetryDialog = false
	end if

END FUNCTION


' QuestionDialog: Dialog box containing a question and two options: "Yes" and "No"
' pQuestion: Question displayed in dialog (text)
' pDialogTitle (Optional) : Dialog title (text)
FUNCTION QuestionDialog(pQuestion as String, Optional pDialogTitle as String) as Boolean

	if Dialog(pQuestion, pDialogTitle, "questionDialog") = true then
		QuestionDialog =  true
	else
		QuestionDialog = false
	end if

END FUNCTION


rem **************************** General Function
' Dialog: Dialog box containing a question and two options.
' pQuestion: Question displayed in dialog
' pDialogTitle: Dialog Title (Optional)
' -------------
' pDialogType: Dialog Type (Optional). Choice:
' questionDialog -(Yes and no button) - Default Value
' confirmDialog - (Ok and cancel button)
' retryDialog - ("Retry" and cancel button)
' -------------
' pDefaultButton: Choose Default Button
' firstBtnDefault - First Button (Default Value)
' secondBtnDefault - Second Button;
' -------------
' pIcon: Choose dialog box icon: stopIcon, questionIcon, exclamationIcon, informationIcon
' -------------
' Note: For custom "Dialog Box", see:
' https://help.libreoffice.org/6.2/en-US/text/sbasic/shared/0310102.html
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

' Dialog3: Dialog box containing one question and three options.
' pQuestion: Question displayed in dialog
' pDialogTitle: Dialog Title (Optional)
' -------------
' pDialogType: Dialog Type (Optional). Choice:
' confirmDialog - Buttons Yes, No, and Cancel buttons - Default Value
' retryDialog - "Abort", "Retry" and "Ignore" buttons
' -------------
' pDefaultButton: Choose Default Button
' firstBtnDefault - First Button (Default Value)
' secondBtnDefault - Second Button;
' -------------
' pIcon: Choose dialog box icon: stopIcon, questionIcon, exclamationIcon, informationIcon
' -------------
' Note: For custom "Dialog Box", see:
' https://help.libreoffice.org/6.2/en-US/text/sbasic/shared/0310102.html
' ---------------------------
'Example of use:
'DIM result as String
'
' result = Dialog3("Are you sure ?")
'
' Select Case result
' Case "No":
' 	Action if not
' Case "Yes"
' 	Action if yes
' Case "Cancel"
' 	Action in case of cancel
' End Select
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
