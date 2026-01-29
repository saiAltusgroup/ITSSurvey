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


Set ChartObj=new cChartPie
ChartObj.ChartWidth=680*1.19
ChartObj.ChartHeight=480*1.1
ChartObj.ChartShadows=False
If Request("ID")<>"0" Then ChartObj.FooterLegend=Lang("chart_legend_b")
ChartObj.StartAngle=270
'ChartObj.SystemColors=Array("27A25A00", "0000FF00", "EBAB0000", "88000000", "00880000", "00008800", "88880000", "00888800", "88008800", "00000000")
ChartObj.SystemColors=Array("EBAB0000", "B4474200", "27A25A00", "88000000", "00880000", "00008800", "88880000", "00888800", "88008800", "00000000")


' Create and initialize PowerDoc object
Set PD = Server.CreateObject("PDOC.PowerDoc")
PD.Initialize PD_INI_SHARE, Application("pdEngine")
' Specify the page to render
PD.Page = 0

If Request("format")="4" Then

	PD.Zoom = 0.45
	' Set clipping area
	PD.Clip 0, 0, 682*1.19*PD.Zoom, 590*1.1*PD.Zoom

	PD.OffsetX=80*1.19*PD.Zoom
	PD.OffsetY=0*1.1*PD.Zoom

ElseIf Request("format")="2" or Request("format")="3" Then
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
ChartObj.XMLToPDC xmlDoc, PD, 0


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
