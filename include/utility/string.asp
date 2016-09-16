<%
Dim charCr, charLf, charCrLf, htmlBr
charCr=Chr(13)
charLf=Chr(10)
charCrLf=charCr & charLf
htmlBr="<br/>"

Function HtmlEncode(byref text)
	if Instr(text,"&")>0 then
		text=Replace(text,"&","&amp;")
	end if
	if Instr(text,"<")>0 then
		text=Replace(text,"<","&lt;")
	end if
	if Instr(text,">")>0 then
		text=Replace(text,">","&gt;")
	end if
	if Instr(text,"""")>0 then
		text=Replace(text,"""","&quot;")
	end if
	if Instr(text,"'")>0 then
		text=Replace(text,"'","&#x27;")
	end if

	HtmlEncode=text
End Function

Function HtmlDecode(byref text)
	if Instr(text,"&quot;")>0 then
		text = Replace(text, "&quot;", """")
	end if
	if Instr(text,"&apos;")>0 then
		text = Replace(text, "&apos;", "'")
	end if
	if Instr(text,"&lt;")>0 then
		text = Replace(text, "&lt;"  , "<")
	end if
	if Instr(text,"&gt;")>0 then
		text = Replace(text, "&gt;"  , ">")
	end if
	if Instr(text,"&amp;")>0 then
		text = Replace(text, "&amp;" , "&")
	end if
	if Instr(text,"&nbsp;")>0 then
		text = Replace(text, "&nbsp;", " ")
	end if

	HtmlDecode = text
End Function

function textfilter(byref str,byval filterScript)
if str="" then
	textfilter=""
else
	dim re,strContent
	set re=new RegExp
	re.Global=true
	re.IgnoreCase=false
	strContent=str

	if filterScript then
		re.Pattern="([jJ][aA][vV][aA][sS][cC][rR][iI][pP][tT]):"
		strContent=re.Replace(strContent,"$1 :")
		re.Pattern="([vV][bB][sS][cC][rR][iI][pP][tT]):"
		strContent=re.Replace(strContent,"$1 :")
		re.Pattern="(:[eE][xX][pP][rR][eE][sS][sS][iI][oO][nN])\("
		strContent=re.Replace(strContent,"$1 (")
		re.Pattern="([bB][eE][hH][aA][vV][iI][oO][rR]\s*):"
		strContent=re.Replace(strContent,"$1-ban:")
	end if

	if InStr(strContent,charCrLf)>0 then
		strContent=replace(strContent,charCrLf," ")
	end if
	if InStr(strContent,charCr)>0 then
		strContent=replace(strContent,charCr," ")
	end if
	if InStr(strContent,charLf)>0 then
		strContent=replace(strContent,charLf," ")
	end if

	textfilter=strContent
end if
end function

Sub NewLineToHtmlBr(byref str)
	if InStr(str,charCrLf)>0 then
		str=replace(str,charCrLf,htmlBr)
	end if
	if InStr(str,charCr)>0 then
		str=replace(str,charCr,htmlBr)
	end if
	if InStr(str,charLf)>0 then
		str=replace(str,charLf,htmlBr)
	end if
End Sub
%>