<%@ Language=VBScript %>

<!--#INCLUDE virtual="/lib/asp/adovbs.asp"-->
<!--#INCLUDE virtual="/lib/asp/math.asp"-->
<!--#INCLUDE virtual="/lib/asp/xml_chart_engine.asp"-->
<!--#INCLUDE VIRTUAL="/sys_include/db.asp"-->
<!--#INCLUDE VIRTUAL="/include/xml_db_functions.asp"-->


<%

xmlID=Request("xmlID")
if len(xmlID)<1 then response.redirect "default.asp"
if not isNumeric(xmlID) then response.redirect "default.asp"
xmlID=clng(xmlID)

Dim xmlStr

Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument.3.0")
xmlDoc.async = False
xmlStr=getXMLfromDB(xmlID)
xmlDoc.loadXML(xmlStr)

'Response.Write xmlStr
'Response.End


Set row_nodes = xmlDoc.selectSingleNode("/chartinfo/rowset")
nr_of_rows = row_nodes.childNodes.length
Set row_nodes = nothing

if Request("ID")=1 then
	Colors=Array("EBAB0000", "A1BBE400", "27A25A00", "C4D6FE00")
	nr_of_cols=4
elseif Request("ID")=2 then
	Colors=Array("EBAB0000", "27A25A00", "A1BBE400")
	nr_of_cols=3
elseif Request("ID")=4 then
	Colors=Array("EBAB0000", "A1BBE400", "27A25A00", "C4D6FE00", "88888800")
	nr_of_cols=5
else
	Colors=Array("27A25A00")
	nr_of_cols=1
	' When ChartObj.ChartWidth<0, it doesn't matter what ChartObj.ColumnWidth is
end if


Set ChartObj=new cChart
ChartObj.Stack=False
ChartObj.NoNegativeValues=False
ChartObj.MultipleScales=False
ChartObj.ChartShadows=False
ChartObj.BackgroundGradient=True
ChartObj.ChartWidth=680*1.19
ChartObj.ChartHeight=340*1.1
ChartObj.VerticalGridlines=3
ChartObj.FooterLegend=""
ChartObj.DoNotChartAttribute = "DoNotChart"
ChartObj.VerticalGridStep=1

ChartObj.ColumnWidth=0

ChartObj.ColumnsInGroup=nr_of_cols
ChartObj.ScaleFactorForColumnsInGroup=0.85

ChartObj.VerticalGridStrategyForLabelsPrimary=1
ChartObj.VerticalGridStrategyForLabelsSecondary=2
if Request("chartID")=10 or Request("chartID")=11 or Request("chartID")=12 or Request("chartID")=13 then ChartObj.VerticalGridStrategyForLabelsSecondary=1
ChartObj.VerticalGridDensityForStrategySwitch=8
ChartObj.VerticalGridDensityForLabelDisplay=12

ChartObj.BottomToTopRows=false
if Request("ID")=3 then ChartObj.BottomToTopRows=true

if nr_of_rows<2 then
	ChartObj.LegendOffsetY=-40
elseif nr_of_rows>8 then
	ChartObj.LegendOffsetY=70
end if


ReDim ColumnConfig(nr_of_cols)

for i=0 to nr_of_cols-1
	Set ColumnConfig(i)=new cChartColumn
	'-----------------------------------
	' Column 1
	ColumnConfig(i).RGBT=Colors(i)
	ColumnConfig(i).ScalePosition=0
	ColumnConfig(i).TextForScale=""
	ColumnConfig(i).GridlineType=1
	ColumnConfig(i).GridlineRGBT="FF660000"
	ColumnConfig(i).DrawType=2
	'ColumnConfig(i).MinMethod=-1 ' make the minimum equal to zero on this column
	ColumnConfig(i).Symbol=-1
	ColumnConfig(i).SymbolWidth=9
	ColumnConfig(i).SymbolHeight=9
	ColumnConfig(i).SymbolYOffset=-5
next



' Create and initialize PowerDoc object
Set PD = Server.CreateObject("PDOC.PowerDoc")
PD.Initialize PD_INI_SHARE, Application("pdEngine")
' Specify the page to render
PD.Page = 0


If Request("format")="2" or Request("format")="3" Then
	PD.Zoom = 4.166
	' Set clipping area
	PD.Clip 0, 0, 850*1.19*PD.Zoom, 660*1.1*PD.Zoom
Else
	' Set zoom factor
	PD.Zoom = 1
	' Set clipping area
	PD.Clip 0, 0, 850*1.19*PD.Zoom, 660*1.1*PD.Zoom
End If

' Make document
ChartObj.XMLToPDC xmlDoc, PD, ColumnConfig


If Request("format")="2" Then

	' Render the page as a JPEG image
	Response.ContentType = "image/jpeg"
	Response.BinaryWrite PD.MakeJPG(,,,,,95)

Else

	' Render the page as a PNG image
	Response.ContentType = "image/png"
	Response.BinaryWrite PD.MakePNG

End If


Set PD = nothing
Set xmlDoc = nothing

%>
