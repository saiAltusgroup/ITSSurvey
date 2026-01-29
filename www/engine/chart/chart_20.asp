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


Set ChartObj=new cChart
ChartObj.Stack=True
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
ChartObj.VerticalGridStrategyForLabelsPrimary=0
ChartObj.VerticalGridStrategyForLabelsSecondary=2
ChartObj.VerticalGridDensityForStrategySwitch=10
ChartObj.VerticalGridDensityForLabelDisplay=12

Dim ColumnConfig(5)
Set ColumnConfig(0)=new cChartColumn
Set ColumnConfig(1)=new cChartColumn
Set ColumnConfig(2)=new cChartColumn
Set ColumnConfig(3)=new cChartColumn
Set ColumnConfig(4)=new cChartColumn
'-----------------------------------
ColumnConfig(0).RGBT="27A25A00"
ColumnConfig(0).ScalePosition=0
ColumnConfig(0).TextForScale=""
ColumnConfig(0).GridlineType=1
ColumnConfig(0).GridlineRGBT="33339980"
ColumnConfig(0).DrawType=0
ColumnConfig(0).MinMethod=-1 ' make the minimum equal to zero on this column
'-----------------------------------
ColumnConfig(1).RGBT="B4474200"
ColumnConfig(1).ScalePosition=0
ColumnConfig(1).TextForScale=""
ColumnConfig(1).GridlineType=1
ColumnConfig(1).GridlineRGBT="33339980"
ColumnConfig(1).DrawType=0
ColumnConfig(1).MinMethod=-1 ' make the minimum equal to zero on this column
'-----------------------------------
ColumnConfig(2).RGBT="EBAB0000"
ColumnConfig(2).ScalePosition=0
ColumnConfig(2).TextForScale=""
ColumnConfig(2).GridlineType=1
ColumnConfig(2).GridlineRGBT="FF660000"
ColumnConfig(2).DrawType=0
ColumnConfig(2).MinMethod=-1 ' make the minimum equal to zero on this column
'-----------------------------------
ColumnConfig(3).RGBT="88000000"
ColumnConfig(3).ScalePosition=0
ColumnConfig(3).TextForScale=""
ColumnConfig(3).GridlineType=1
ColumnConfig(3).GridlineRGBT="FF660000"
ColumnConfig(3).DrawType=0
ColumnConfig(3).MinMethod=-1 ' make the minimum equal to zero on this column
'-----------------------------------
ColumnConfig(4).RGBT="00880000"
ColumnConfig(4).ScalePosition=0
ColumnConfig(4).TextForScale=""
ColumnConfig(4).GridlineType=1
ColumnConfig(4).GridlineRGBT="FF660000"
ColumnConfig(4).DrawType=0
ColumnConfig(4).MinMethod=-1 ' make the minimum equal to zero on this column
'-----------------------------------




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
