<%
function textfilter(byref str,byval filterScript)
if str="" then
	textfilter=""
else
	dim re,strContent
	set re=new RegExp
	re.IgnoreCase=true
	strContent=str

	if filterScript then
		on error resume next

		re.pattern="(?:javascript|vbscript)\s*:"
		if err.number=0 then strContent=re.replace(strContent,"")

		re.pattern="(style=.*?):expression"
		if err.number=0 then strContent=re.replace(strContent,"$1")

		re.pattern="(<.*?style\s*=.*?)behavior:(.*?>)"
		if err.number=0 then strContent=re.replace(strContent,"$1$2")

		on error goto 0
	end if

	strContent=replace(strContent,chr(13)&chr(10)," ")
	textfilter=strContent
end if
end function
%>