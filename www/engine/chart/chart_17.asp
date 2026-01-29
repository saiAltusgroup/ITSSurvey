<%@ Language=VBScript %>

<!--#INCLUDE virtual="/lib/asp/adovbs.asp"-->
<!--#INCLUDE virtual="/lib/asp/math.asp"-->
<!--#INCLUDE virtual="/lib/asp/xml_chart_engine.asp"-->
<!--#INCLUDE VIRTUAL="/sys_include/db.asp"-->
<!--#INCLUDE VIRTUAL="/include/xml_db_functions.asp"-->
<!--#INCLUDE virtual="/include/language.asp"-->


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


nr_of_cols=3
if Request("chartID")=2 and Request("ID")=1 then nr_of_cols=1
if Request("chartID")=49 then nr_of_cols=1
if Request("chartID")=4 or Request("chartID")=5 or Request("chartID")=6 or Request("chartID")=7 then nr_of_cols=1
if Request("chartID")=16 or Request("chartID")=23 or Request("chartID")=26 or Request("chartID")=27 then nr_of_cols=2
if Request("chartID")=8 or Request("chartID")=19 or Request("chartID")=24 or Request("chartID")=25 or Request("chartID")=30 or Request("chartID")=43 then nr_of_cols=4
if Request("chartID")=22 or Request("chartID")=47 or Request("chartID")=48 or Request("chartID")=52 or Request("chartID")=53 then nr_of_cols=5
if Request("chartID")=45 or Request("chartID")=46 then nr_of_cols=6
if Request("chartID")=8 and Request("ID")=1 then nr_of_cols=2


Set ChartObj=new cChart
ChartObj.Stack=False
ChartObj.NoNegativeValues=False
ChartObj.MultipleScales=False
ChartObj.ChartShadows=False
ChartObj.BackgroundGradient=True
ChartObj.ChartWidth=680*1.19
ChartObj.ChartHeight=340*1.1
ChartObj.VerticalGridlines=3
ChartObj.FooterLegend=Lang("chart_legend_a")
if Request("chartID")=46 then ChartObj.FooterLegend=Lang("chart_asset_value") & "  -  " & Lang("chart_legend_a")
ChartObj.FooterLegend2=Lang("chart_legend_c")
ChartObj.DoNotChartAttribute = "DoNotChart"
ChartObj.VerticalGridStep=1

ChartObj.ColumnWidth=0

ChartObj.ColumnsInGroup=nr_of_cols*1
ChartObj.ScaleFactorForColumnsInGroup=0.85

ChartObj.VerticalGridStrategyForLabelsPrimary=0
ChartObj.VerticalGridStrategyForLabelsSecondary=1
ChartObj.VerticalGridDensityForStrategySwitch=4
ChartObj.VerticalGridDensityForLabelDisplay=12

'ChartObj.ColumnMapings=Array(0,0,0, 1,1,1, 2,2,2, 3,3,3) ' this supports up to 4 columns
'ChartObj.ColumnMapings=Array(0,0,0, 1,1,1, 2,2,2, 3,3,3, 4,4,4) ' this supports up to 5 columns
ChartObj.ColumnMapings=Array(0,0,0, 1,1,1, 2,2,2, 3,3,3, 4,4,4, 5,5,5) ' this supports up to 6 columns
ChartObj.BottomToTopRows=false



color1="EBAB0000"
if Request("chartID")=4 or Request("chartID")=6 then color1="A1BBE400"
if Request("chartID")=5 or Request("chartID")=7 then color1="EBAB0000"

'Colors=Array(color1, "A1BBE400", "27A25A00", "B4474200") ' this supports up to 4 columns
'Colors=Array(color1, "A1BBE400", "27A25A00", "B4474200", "88888800") ' this supports up to 5 columns
Colors=Array(color1, "A1BBE400", "27A25A00", "B4474200", "9999cc00", "99990000") ' this supports up to 6 columns

ReDim ColumnConfig(nr_of_cols*3-1)

j=0
for i=0 to nr_of_cols-1

	Set ColumnConfig(j+0)=new cChartColumn
	Set ColumnConfig(j+1)=new cChartColumn
	Set ColumnConfig(j+2)=new cChartColumn

	'-----------------------------------
	' MAX Column
	ColumnConfig(j).RGBT="XXXXXXXX"
	'ColumnConfig(j).RGBT=Colors(i)
	ColumnConfig(j).ScalePosition=0
	ColumnConfig(j).TextForScale=""
	ColumnConfig(j).GridlineType=1
	ColumnConfig(j).GridlineRGBT="33339980"
	ColumnConfig(j).DrawType=2
	'ColumnConfig(j).MinMethod=-1 ' make the minimum equal to zero on this column
	j=j+1
	'-----------------------------------
	' MIN Column
	ColumnConfig(j).RGBT=Colors(i)
	ColumnConfig(j).ScalePosition=0
	ColumnConfig(j).TextForScale=""
	ColumnConfig(j).GridlineType=1
	ColumnConfig(j).GridlineRGBT="33339980"
	ColumnConfig(j).DrawType=2
	'ColumnConfig(j).MinMethod=-1 ' make the minimum equal to zero on this column
	j=j+1
	'-----------------------------------
	' AVG Column
	ColumnConfig(j).RGBT="FFFFFFFF"
	ColumnConfig(j).ScalePosition=0
	ColumnConfig(j).TextForScale=""
	ColumnConfig(j).GridlineType=1
	ColumnConfig(j).GridlineRGBT="FF660000"
	ColumnConfig(j).DrawType=2
	'ColumnConfig(j).MinMethod=-1 ' make the minimum equal to zero on this column
	ColumnConfig(j).Symbol=-1
	ColumnConfig(j).SymbolWidth=9
	ColumnConfig(j).SymbolHeight=9
	ColumnConfig(j).SymbolYOffset=-5
	j=j+1
	'-----------------------------------

next



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
