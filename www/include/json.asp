<%

' DK, Apr 01, 2008 - version 1.0

'******************************************************************************************
'' @SDESCRIPTION:	STATIC! takes a given string and makes it JSON valid
'' @DESCRIPTION:	all characters which needs to be escaped are beeing replaced by their
''					unicode representation according to the
''					RFC4627#2.5 - http://www.ietf.org/rfc/rfc4627.txt?number=4627
'' @PARAM:			val [string]: value which should be escaped
'' @RETURN:			[string] JSON valid string
'******************************************************************************************
function json_escape(val)
	dim cDoubleQuote, cRevSolidus, cSolidus
	cDoubleQuote = &h22
	cRevSolidus = &h5C
	cSolidus = &h2F

	dim i, currentDigit
	for i = 1 to (len(val))
		currentDigit = mid(val, i, 1)

		if asc(currentDigit) > &h00 and asc(currentDigit) < &h1F then
			currentDigit = json_escapequence(currentDigit)
		elseif asc(currentDigit) >= &h7F and asc(currentDigit) <= &hFF then
			currentDigit = json_escapequence(currentDigit) ' *** ADDED BY DK ***
		elseif asc(currentDigit) >= &hC280 and asc(currentDigit) <= &hC2BF then
			currentDigit = "\u00" + right(json_padLeft(hex(asc(currentDigit) - &hC200), 2, 0), 2)
		elseif asc(currentDigit) >= &hC380 and asc(currentDigit) <= &hC3BF then
			currentDigit = "\u00" + right(json_padLeft(hex(asc(currentDigit) - &hC2C0), 2, 0), 2)
		else
			select case asc(currentDigit)
				case cDoubleQuote: currentDigit = json_escapequence(currentDigit)
				case cRevSolidus: currentDigit = json_escapequence(currentDigit)
				case cSolidus: currentDigit = json_escapequence(currentDigit)
			end select
		end if
		json_escape = json_escape & currentDigit
	next
end function

function json_escapequence(digit)
	json_escapequence = "\u00" + right(json_padLeft(hex(asc(digit)), 2, 0), 2)
end function

function json_padLeft(value, totalLength, paddingChar)
	json_padLeft = right(json_clone(paddingChar, totalLength) & value, totalLength)
end function

function json_clone(byVal str, n)
	dim i
	for i = 1 to n : json_clone = json_clone & str : next
end function

%>
