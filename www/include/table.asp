<%

' DK, Apr 01, 2008 - version 1.0

Class cTablePlaintext

	Public RowID
	Public Extra
	Public Extra2
	Public Name
	Public ReadOnly
	Public DisplayMode ' 0=all (default), 1=just cells, 2=just rowids, 3=rowids and data, -1=json

	Private CountRow
	Private CountCell

	Private Sub Class_Initialize()
		RowID=""
		Extra=""
		Extra2=""
		Name="notset"
		ReadOnly=false
		Multi=false
		DisplayMode=0
		CountRow=0
		CountCell=0
	End Sub

	Public Property Let Multi(flag)
	End Property

	Public Property Let Symbol(text_symbol)
	End Property

	Public Property Let Label(text_label)
	End Property

	Public Property Get TableOpen(nr_of_cols)

		CountRow=0
		CountCell=0

		if DisplayMode<=0 then TableOpen = "{"

	End Property

	Public Property Get TableClose()

		'CountRow=0
		'CountCell=0

		if DisplayMode<=0 then TableClose = "}"

	End Property

	Public Property Get RowOpen()

		if DisplayMode=-1 then
			if CountRow=0 then s="" else s= ","
			RowOpen = s & """" & RowID & """:"""
		elseif DisplayMode=0 then
			RowOpen = "[" & RowID & " " & Extra & ": "
		elseif DisplayMode=2 or DisplayMode=3 then
			if CountRow=0 then s="" else s= ","
			RowOpen = s & RowID
		end if

		CountRow=CountRow+1

	End Property

	Public Property Get RowClose()

		if DisplayMode=-1 then
			RowClose = """"
		elseif DisplayMode=0 then
			RowClose = "]"
		end if

	End Property

	Public Property Get Cell(val)

		if DisplayMode=-1 then
			Cell = JsonEncode(val)
		elseif DisplayMode=0 then
			Cell = "(" & val & ")"
		elseif DisplayMode=1 then
			if CountCell=0 then s="" else s= " • "
			Cell = s & val
		elseif DisplayMode=3 then
			Cell = "|" & val
		end if

		CountCell=CountCell+1

	End Property

	Public Property Get Head(val)
		Head = Cell(val)
	End Property

	Public Function PrependHack(str, row_id, val, check_flag)

		'PrependHack=str
		PrependHack=""

		if DisplayMode=-1 then val=JsonEncode(val)

		str_comma=","
		if str="{}" then str_comma=""

		if check_flag and CountRow<1 then exit function
		tag="{"
		prepend = """" & row_id & """:""" & val & """" & str_comma
		PrependHack = Replace(str, tag, tag & prepend)

	End Function

	Private Function JsonEncode(val)
		'val=Replace(val, """", "\""")
		'val=Replace(val, "/", "\/")
		val=json_escape(val)
		JsonEncode=val
	End Function

End Class


Class cTableListbox

	Public RowID
	Public Extra
	Public Extra2
	Public ExtraAttribs
	Public Name
	Public ReadOnly
	Public FullDisplay

	Private MultiStr

	Private CountRow

	Private Sub Class_Initialize()
		RowID=""
		Extra=""
		Extra2=""
		ExtraAttribs=""
		Name="notset"
		ReadOnly=false
		FullDisplay=true
		Multi=false
		CountRow=0
	End Sub

	Public Property Let Multi(flag)
		MultiStr=""
		if flag then MultiStr="multiple"
	End Property

	Public Property Let Symbol(text_symbol)
	End Property

	Public Property Let Label(text_label)
	End Property

	Public Property Get TableOpen(nr_of_cols)
		CountRow=0
		TableOpen = "<select name=""" & Server.HTMLEncode(Name) & """ " & MultiStr & " id=""" & Server.HTMLEncode(Name) & """ " & ExtraAttribs & "><!-- ROWS_START -->"
	End Property

	Public Property Get TableClose()
		'CountRow=0
		TableClose = "<!-- ROWS_END --></select>"
	End Property

	Public Property Get RowOpen()
		RowOpen = "<option value=""" & Server.HTMLEncode(RowID) & """>"
		if FullDisplay then RowOpen = RowOpen & Server.HTMLEncode(RowID) & " " & Server.HTMLEncode(Extra) & ": "
		CountRow=CountRow+1
	End Property

	Public Property Get RowClose()
		RowClose = "</option>"
	End Property

	Public Property Get Cell(val)
		Cell = Server.HtmlEncode(val)
		if FullDisplay then Cell = "(" & Cell & ")"
	End Property

	Public Property Get Head(val)
		Head = Server.HtmlEncode(val)
		if FullDisplay then Head = "(" & Head & ")"
	End Property

	Public Function PrependHack(str, row_id, val, check_flag)

		'PrependHack=str
		PrependHack=""

		if check_flag and CountRow<1 then exit function
		tag="<!-- ROWS_START -->"
		prepend = "<option value=""" & Server.HTMLEncode(row_id) & """>" & val & "</option>"
		PrependHack = Replace(str, tag, tag & prepend)

	End Function

End Class


Class cTableDatagrid

	Public RowID
	Public Extra
	Public Extra2
	Public Name
	Public ReadOnly

	Private MultiStr

	Private CountRow

	Private Sub Class_Initialize()
		RowID=""
		Extra=""
		Extra2=""
		Name="notset"
		ReadOnly=false
		Multi=false
		CountRow=0
	End Sub

	Public Property Let Multi(flag)
		MultiStr="radio"
		if flag then MultiStr="checkbox"
	End Property

	Public Property Let Symbol(text_symbol)
	End Property

	Public Property Let Label(text_label)
	End Property

	Public Property Get TableOpen(nr_of_cols)
		CountRow=0
		TableOpen = "<table border=""1""><!-- ROWS_START -->"
	End Property

	Public Property Get TableClose()
		TableClose = "<!-- ROWS_END --></table>"
	End Property

	Public Property Get RowOpen()

		styleid=CountRow Mod 2

		RowOpen = "<tr align=""center"" class=""alternate_" & styleid & """><th align=""left"">"
		if not ReadOnly then RowOpen = RowOpen & "<input type=""" & MultiStr & """ name=""" & Server.HTMLEncode(Name) & """ value=""" & Server.HTMLEncode(RowID) & """>"
		br=""
		if Len(RowID)>0 Then RowOpen = RowOpen & Server.HTMLEncode(RowID) : br="<br>"
		if Len(Extra)>0 then RowOpen = RowOpen & br & Server.HTMLEncode(Extra) : br="<br>"
		if Len(Extra2)>0 then RowOpen = RowOpen & br & Server.HTMLEncode(Extra2)
		RowOpen = RowOpen & "</th>"

		CountRow=CountRow+1

	End Property

	Public Property Get RowClose()
		RowClose = "</tr>"
	End Property

	Public Property Get Cell(val)
		Cell = "<td>" & Server.HtmlEncode(val) & "</td>"
	End Property

	Public Property Get HeadSpan(val,n)

		colspan=""
		if n>1 then colspan=" colspan=""" & n & """"

		HeadSpan = "<th" & colspan & ">" & Server.HtmlEncode(val) & "</th>"

	End Property

	Public Property Get Head(val)
		Head=HeadSpan(val, 1)
	End Property

End Class


Class cTableXmlchart

	Public RowID
	Public Extra
	Public Extra2
	Public ReadOnly

	Public ChartTitle
	Public ChartTitle2
	Public ChartSubtitle
	Public ChartColumnUnits
	Public ChartColumnDescs
	Public ChartIgnoredColumns

	Private CountCol
	Private TextLabel
	Private TextSymbol

	Private Sub Class_Initialize()

		RowID=""
		Extra=""
		Extra2=""
		ReadOnly=false

		ChartTitle="Chart Title"
		ChartTitle2="Chart Title2"
		ChartSubtitle="Chart Subtitle"

		ChartColumnUnits=Array("")
		ChartColumnDescs=Array("")
		ChartIgnoredColumns=""

		CountCol=0
		TextLabel=""
		TextSymbol=""

	End Sub

	Public Property Let Multi(flag)
	End Property

	Public Property Let Label(text_label)
		TextLabel=text_label
	End Property

	Public Property Let Symbol(text_symbol)
		if text_symbol<0 then
			TextSymbol=3
		elseif text_symbol>0 then
			TextSymbol=1
		else
			TextSymbol=108 '59
		end if
	End Property

	Public Property Get TableOpen(nr_of_cols)
		TableOpen = ""
		TableOpen = TableOpen & "<?xml version=""1.0"" encoding=""ISO-8859-1"" ?>" & vbNewLine
		TableOpen = TableOpen & "<chartinfo name=""" & XMLEncode(ChartTitle) & """>" & vbNewLine
		TableOpen = TableOpen & "<reportMktPath>" & XMLEncode(ChartTitle2) & "</reportMktPath>" & vbNewLine
		TableOpen = TableOpen & "<BlockQtrType>unused</BlockQtrType>" & vbNewLine
		TableOpen = TableOpen & "<blockname>" & XMLEncode(ChartSubtitle) & "</blockname>" & vbNewLine
		TableOpen = TableOpen & "<columnspecs DoNotChart=""" & XMLEncode(ChartIgnoredColumns) & """>" & vbNewLine
		for i=0 to nr_of_cols-1
			TableOpen = TableOpen & "<column id=""" & i+1 & """ type=""" & XMLEncode(ChartColumnUnits(i)) & """>" & XMLEncode(ChartColumnDescs(i)) & "</column>" & vbNewLine
		next
		TableOpen = TableOpen & "</columnspecs>" & vbNewLine
		TableOpen = TableOpen & "<rowset>" & vbNewLine
	End Property

	Public Property Get TableClose()
		TableClose = "</rowset></chartinfo>"
	End Property

	Public Property Get RowOpen()
		CountCol=0
		RowOpen = "<row label=""" & XMLEncode(RowID) & """ textLabel=""" & XMLEncode(Extra) & """ textLabel2=""" & XMLEncode(Extra2) & """>"
	End Property

	Public Property Get RowClose()
		CountCol=0
		RowClose = "</row>" & vbNewLine
	End Property

	Public Property Get Cell(val)

		CountCol=CountCol+1

		str_label="" : if Len(TextLabel)>0 then str_label=" textLabel=""" & XMLEncode(TextLabel) & """"
		str_symbol="" : if Len(TextSymbol)>0 then str_symbol=" textSymbol=""" & XMLEncode(TextSymbol) & """"

		Cell = "<column id=""" & CountCol & """" & str_label & str_symbol & ">" & XMLEncode(val) & "</column>"

		TextLabel=""
		TextSymbol=""

	End Property


	Public Property Get Head(val)
		Head = ""
	End Property


	Private Function XMLEncode(str)
		str=Replace(str, chr(160), " ")
		XMLEncode=Server.HtmlEncode(str)
	End Function

End Class


%>
