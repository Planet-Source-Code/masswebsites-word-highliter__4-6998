<div align="center">

## Word Highliter


</div>

### Description

Ever see a search engine that highlights your search results? This code does that. It takes a querystring, and compares it to a second variable. The second variable will be written to the page, and any matching words will be in highlighted and bolded. See screenshot for details.
 
### More Info
 
request.querystring("qs")

Assumes your querystring and compare variable are delimited by spaces. This can easily be changed. You could also replace the querystring input with a value from a database, an xml file, or anything else. Pleasse comment.

Returns a string with highlighted words that are similar.

Currently Unknown. This has also not been user load-tested.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[masswebsites](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/masswebsites.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/masswebsites-word-highliter__4-6998/archive/master.zip)

### API Declarations

Have fun.


### Source Code

```
<%
' hit highliter
' this was not user load tested
' have fun - masswebsites@yahoo.com
option explicit
Response.Expires=0
%>
<html>
<head>
<title></title>
</head>
<body>
	<%
	' check querystring
	if request("qs") = "" then
		Response.Write "No querystring supplied. <a href='hitliter.asp?qs=asp%20loop%20example'>Try this one</a>"
	else
		dim sQueryString
		dim sSummary
		dim sDisplay
		sQueryString = request("qs")
		' example summary - could also be from a database or other data source
		sSummary = "I hope you find this ASP code sample useful. It uses nested for next loops , strcomp , arrays , join , split , and some other fun things. If you can comment, criticize, or optimize this code example, please post a reply."
		' display with matching words highlighted. Could also set a variable = hitLite(sSummary,sQueryString)
		Response.Write hitLite(sSummary,sQueryString) & "<br><br>"
	end if
	%>
</body>
</html>
<%
' **************************************
' put it in a global include file
' **************************************
function hitLite (sSummary,sQueryString)
	dim arrQueryString
	dim iQs
	dim i
	dim arrSummary
	dim sArrQueryStringTmp
	dim iComp
	dim sArrSummaryTmp
	dim sRoot
	dim sLastChar
	dim sLength
	' break the variable we want to compare querystring to into an array
	arrSummary = split(sSummary," ")
	' break querystring into an array, use space for delimiter
	arrQueryString = split(sQuerystring," ")
	' for every word in the querystring
	for iQs = 0 to ubound(arrQuerystring)
		' assign the value to a temp variable
		sArrQueryStringTmp = arrQuerystring(iQs)
		' don't include common search words, you can take this out. I was using this for a search engine
		if (sArrQueryStringTmp <> "and") and (sArrQueryStringTmp <> "or") and (sArrQueryStringTmp <> "+") then
			' for each element in the variable array, replace querystring word in varaible array with the word plus bgcolor ( highlight )
			for i = 0 to ubound(arrSummary)
				sArrSummaryTmp = arrSummary(i)
				' if the 2 strings compare, stick the display value in a span with style!
				if strcomp(sArrSummaryTmp,sArrQueryStringTmp,1) = 0 then
					arrSummary(i) = "<span style=background:yellow;font-weight:bold;>" & sArrSummaryTmp & "</span>"
				else
					' check "s", comma, period
					' must end "s" or comma or period AND be greater than 1 character
					sLastChar = right(sArrSummaryTmp,1)
					sLength = len(sArrSummaryTmp)
					if (sLastChar = "s") or (sLastChar = ".") or (sLastChar = ",") and (sLength > 1) then
						'the word minus last letter
						sRoot = left(sArrSummaryTmp,sLength-1)
						' if the root comapres to the querystring
						' replace that element of the array with the root highlited and the last character in regular display
						if strcomp(sRoot,sArrQueryStringTmp,1) = 0 then
							' don't include the comma or period in the lite
							arrSummary(i) = "<span style=background:yellow;font-weight:bold;>" & sRoot & "</span>" & sLastChar & "&nbsp;"
						end if
					end if
				end if
			next
		end if
	next
	hitLite = join(arrSummary, " ")
end function
%>
```

