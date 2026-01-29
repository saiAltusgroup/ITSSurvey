<%

' DK, Apr 01, 2008 - version 1.0

Function SafeDivide(a, b)

		SafeDivide=0
		if (b<>0) then SafeDivide=a/b

End Function


Function Sort(a, a2, n)

	dim c, c2, i, j

	for i = 0 to n-1
		c=a(i): c2=a2(i)
		j=i

		do while (j>0)
			if (a(j-1)<c) then exit do
			a(j)=a(j-1): a2(j)=a2(j-1)
			j=j-1
		loop

		a(j)=c:	a2(j)=c2
	next

End Function


Function InArray(a, s)

	Dim i

	for i = 0 to UBound(a)

		if a(i) = s Then

			InArray = true
			Exit Function

		end if

	next

	InArray = false

End Function


Function HtmlToPdf(html)

	Set theDoc=Server.CreateObject("ABCpdf8.Doc")

	Call theDoc.Rect.Inset(10, 10)
	theDoc.Page = theDoc.AddPage()

	'theDoc.SetInfo 0, "HostWebBrowser", "0"
	' Use the Gecko engine
	theDoc.HtmlOptions.Engine = 1

	theID = theDoc.AddImageHtml(html)

	for i=0 to 100
		if Not theDoc.Chainable(theID) then exit for
		theDoc.Page = theDoc.AddPage()
		theID = theDoc.AddImageToChain(theID)
	next

	for i=1 to theDoc.PageCount
		theDoc.PageNumber = i
		theDoc.Flatten()
	next

	'theDoc.Save(Server.MapPath("/files/pdf/temp.pdf"))

	HtmlToPdf=theDoc.GetData()

	theDoc.Clear()

	Set theDoc=Nothing

End Function

%>
