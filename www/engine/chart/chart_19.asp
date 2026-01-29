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


secondary_labels_strategy=2
draw_type=1
rgbt_0="EBAB0000"
rgbt_1="A1BBE400"
rgbt_2="27A25A00"
rgbt_3="B4474200"
if Request("ID")=1 then
	if Request("chartID")=4 or Request("chartID")=6 then draw_type=0 : rgbt_0="XXXXXXXX" : rgbt_2="718BC400" : rgbt_1="A1BBE4BB"
	if Request("chartID")=5 or Request("chartID")=7 then draw_type=0 : rgbt_0="XXXXXXXX" : rgbt_2="CB7B0000" : rgbt_1="EBAB00BB"
end if

nr_of_cols=3
if Request("ID")=2 then nr_of_cols=1


Set ChartObj=new cChart
ChartObj.Stack=False
ChartObj.NoNegativeValues=False
ChartObj.MultipleScales=True
ChartObj.ChartShadows=False
ChartObj.BackgroundGradient=True
ChartObj.ChartWidth=680*1.19
ChartObj.ChartHeight=340*1.1
ChartObj.VerticalGridlines=3
ChartObj.FooterLegend=""
ChartObj.DoNotChartAttribute = "DoNotChart"
ChartObj.VerticalGridStep=1
ChartObj.VerticalGridStrategyForLabelsPrimary=0
ChartObj.VerticalGridStrategyForLabelsSecondary=secondary_labels_strategy
ChartObj.VerticalGridDensityForStrategySwitch=4
ChartObj.VerticalGridDensityForLabelDisplay=12

ReDim ColumnConfig(nr_of_cols)

if nr_of_cols=1 then

	Set ColumnConfig(0)=new cChartColumn

	'----------------------------------
	ColumnConfig(0).RGBT=rgbt_3
	ColumnConfig(0).ScalePosition=1
	ColumnConfig(0).TextForScale="[AUTO]"
	ColumnConfig(0).GridlineType=1
	ColumnConfig(0).ZerolineRGBT="000000FF"
	ColumnConfig(0).GridlineRGBT="FF000000"
	ColumnConfig(0).DrawType=1
	ColumnConfig(0).MinMethod=3
	ColumnConfig(0).MaxMethod=3
	'ColumnConfig(0).NoNegativeValues=1
	'-----------------------------------

else

	Set ColumnConfig(0)=new cChartColumn
	Set ColumnConfig(1)=new cChartColumn
	Set ColumnConfig(2)=new cChartColumn
	Set ColumnConfig(3)=new cChartColumn
	'-----------------------------------
	ColumnConfig(0).RGBT=rgbt_0
	ColumnConfig(0).ScalePosition=1
	ColumnConfig(0).TextForScale="[AUTO]"
	ColumnConfig(0).GridlineType=1
	ColumnConfig(0).GridlineRGBT="88888800"
	ColumnConfig(0).DrawType=draw_type
	ColumnConfig(0).MinMaxArray="0,1,2"
	ColumnConfig(0).TextPrefixForScale=" ("
	ColumnConfig(0).MinMethod=2
	ColumnConfig(0).MaxMethod=2
	'-----------------------------------
	ColumnConfig(1).RGBT=rgbt_1
	ColumnConfig(1).ScalePosition=0
	ColumnConfig(1).TextForScale=""
	ColumnConfig(1).GridlineType=0
	ColumnConfig(1).GridlineRGBT="88888800"
	ColumnConfig(1).DrawType=draw_type
	ColumnConfig(1).MinMaxArray="0,1,2"
	ColumnConfig(1).MinMethod=2
	ColumnConfig(1).MaxMethod=2
	'-----------------------------------
	ColumnConfig(2).RGBT=rgbt_2
	ColumnConfig(2).ScalePosition=0
	ColumnConfig(2).TextForScale=""
	ColumnConfig(2).GridlineType=0
	ColumnConfig(2).GridlineRGBT="88888800"
	ColumnConfig(2).DrawType=1
	ColumnConfig(2).MinMaxArray="0,1,2"
	ColumnConfig(2).MinMethod=2
	ColumnConfig(2).MaxMethod=2
	'----------------------------------
	ColumnConfig(3).RGBT=rgbt_3
	ColumnConfig(3).ScalePosition=2
	ColumnConfig(3).TextForScale="[AUTO]"
	ColumnConfig(3).GridlineType=4
	ColumnConfig(3).ZerolineRGBT="000000FF"
	ColumnConfig(3).GridlineRGBT="FF000000"
	ColumnConfig(3).DrawType=1
	ColumnConfig(3).MinMethod=3
	ColumnConfig(3).MaxMethod=3
	'ColumnConfig(3).NoNegativeValues=1
	'-----------------------------------

end if


' Create and initialize PowerDoc object
Set PD = Server.CreateObject("PDOC.PowerDoc")
PD.Initialize PD_INI_SHARE, Application("pdEngine")
' Specify the page to render
PD.Page = 0


If Request("format")="2" or Request("format")="3" Then
	PD.Zoom = 4.166
	' Set clipping area
	PD.Clip 0, 0, 850*1.19*PD.Zoom, 640*1.1*PD.Zoom
Else
	' Set zoom factor
	PD.Zoom = 1
	' Set clipping area
	PD.Clip 0, 0, 850*1.19*PD.Zoom, 640*1.1*PD.Zoom
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
