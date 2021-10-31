'AnimateFontSize: Creates an animation in which a gradual variation of the font size is performed.
'pSheet: Sheet name (text)
'pRange: Range of cells to generate the effect/formatting (text)
'pSize: Final font size after animation ends
'pSpeed: Speed ​​at which the animation is performed ("fast", "medium" or "slow")
Sub AnimateFontSize(pSheet as String, pRange As String, Optional pSize As Integer, Optional pSpeed As String) 

	Dim vInitialSize As Integer
	Dim vFinalSize As Integer
	Dim vTime As Integer

	vInitialSize = Cell(pSheet, pRange).CharHeight
	
	if IsMissing(pSize) Then
		vFinalSize = vInitialSize*1.5	
	Else
		vFinalSize = pSize
	end If
	
	if IsMissing(pSpeed) Then
		vSpeed = "medium"
	Else
		vSpeed = pSpeed
	end If

 	If vSpeed = "fast" then
 		vTime = 0
 	ElseIf  vSpeed = "medium" Then
 		vTime = 20
 	Else 
 		vTime = 50
 	End if
	
	vDiff = abs(vFinalSize - vInitialSize) 
	
	If vFinalSize > vInitialSize then
	
		For i= 1 To vDiff Step 1
			
			ChangeFontSize(pSheet, pRange, vInitialSize+i)
			
			wait vTime
		
		Next
		
	Else
	
		For i= 1 To vDiff Step 1
			
			ChangeFontSize(pSheet, pRange, vInitialSize-i)
			
			wait vTime
		
		Next
	
	End If
	
End Sub

'AnimateFontColor: Creates an animation in which a gradual variation of the font color is performed.
'pSheet: Sheet name (text)
'pRange: Range of cells to generate the effect/formatting (text)
'pColor: Final color at the end of the animation
'pSpeed: Speed ​​at which the animation is performed ("fast", "medium" or "slow")
Sub AnimateFontColor(pSheet as String, pRange As String, Optional pColor As String, Optional pSpeed As String) 
 dim colorLong As Long
 Dim vColor As String
 Dim vSpeed As String
 Dim vTime As Integer
 Dim vInc As Integer
 Dim vOperType1 As String
 Dim vOperType2 As String
 Dim vOperType3 As String
 Dim num1 As Integer
 Dim num2 As Integer
 Dim num3 As Integer
 Dim num4 As Integer
 Dim num5 As Integer
 Dim num6 As Integer
 dim cA$(2) 
 Dim cB$(2) 
 
 	if IsMissing(pSpeed) Then
		vSpeed = "medium"
	Else
		vSpeed = pSpeed
	end If
	
 	If vSpeed = "fast" then
 		vInc = 10
 		vTime = 0
 	ElseIf  vSpeed = "medium" Then
 		vInc = 5
 		vTime = 5
 	Else 
 		vInc = 1
 		vTime = 10
 	End if
 	
 	if IsMissing(pColor) Then
		vColor = "red"
	Else
		vColor = pColor
	end If

 	colorLong = Cell(pSheet, pRange).CharColor
 
 	cA = GetColorCode(  getRGBfromLong( colorLong ) )
 	cB = GetColorCode( vColor )
 	
 	num1 = CInt( cA(0) )
 	num2 = CInt( cB(0) )
 	num3 = CInt( cA(1) )
 	num4 = CInt( cB(1) )
 	num5 = CInt( cA(2) )
 	num6 = CInt( cB(2) )
	                         
 	vOperType1 = getOperType(num1, num2)
 	
 	vOperType2 = getOperType(num3, num4)
 	
 	vOperType3 = getOperType(num5, num6)
 	
	Do While (num1 <> num2 or num3 <> num4 or num5 <> num6)
	   
		num1 = NumberTransformation(num1, num2, vOperType1, vInc)
	   		
		num3 = NumberTransformation(num3, num4, vOperType2, vInc)
		
		num5 = NumberTransformation(num5, num6, vOperType3, vInc)
		
		ChangeFontColor(pSheet, pRange,  rgbColor( num1, num3, num5 ) )   

		wait vTime
		
	Loop

End Sub

' Under Construction
'Sub ToggleCellColor(pSheet as String, pRange As String, Optional pFirstColor As String, Optional pSecondColor As String, Optional pSteps as Integer, Optional pTime As Integer) 

'dim numOfSteps as Integer
'dim vFirstColor As String
'Dim vSecondColor As String

'	if IsMissing(pSteps) Then
'		numOfSteps = 5	
'	else
'		numOfSteps = pSteps	
'	end If
	
'	if IsMissing(pFirstColor) Then
'		vFirstColor = "black"	
'	Else
'		vFirstColor = pFirstColor
'	end If
	
'	c = getColorCode

'	for i = 1 to numOfSteps step 1
	
'		ChangeCellColor(pSheet, pRange, vFirstColor)
		
'		wait 50
		
'		ChangeCellColor(pSheet, pRange, vSecondColor)
		
'		wait 50
				
'	next

'End Sub


