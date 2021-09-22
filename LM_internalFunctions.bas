REM  *****  BASIC  *****

FUNCTION getColorCode (pColor as String) as Object
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
	
	getColorCode = c

END FUNCTION
