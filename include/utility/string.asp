<%
Dim charCr, charLf, charCrLf
charCr=Chr(13)
charLf=Chr(10)
charCrLf=charCr & charLf

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

Function HtmlNewLineEncode(byref text)
	if Instr(text,charCrLf)>0 then
		text=Replace(text,charCrLf,"&#xd;&#xa;")
	end if
	if Instr(text,charCr)>0 then
		text=Replace(text,charCr,"&#xd;")
	end if
	if Instr(text,charLf)>0 then
		text=Replace(text,charLf,"&#xa;")
	end if

	HtmlNewLineEncode=text
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
	re.IgnoreCase=true
	strContent=str

	if filterScript then
		re.pattern="(?:javascript|vbscript)\s*:"
		strContent=re.replace(strContent,"")

		re.pattern="(style=.*?):expression"
		strContent=re.replace(strContent,"$1")

		re.pattern="(<.*?style\s*=.*?)behavior:([^>]*>)"
		strContent=re.replace(strContent,"$1$2")
	end if

	strContent=replace(strContent,chr(13)&chr(10)," ")
	textfilter=strContent
end if
end function
%>