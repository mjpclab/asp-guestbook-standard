<%
const CharAmp="&"
const CharLt="<"
const CharGt=">"
const CharQuot=""""
const CharApos="'"

const HtmlAmp="&amp;"
const HtmlLt="&lt;"
const HtmlGt="&gt;"
const HtmlQuot="&quot;"
const HtmlApos="&apos;"
const HtmlAposCode="&#39;"
const HtmlNbsp="&nbsp;"

const CharSpace=" "
const CharEmpty=""

const htmlBr="<br/>"

Function HtmlEncode(byref text)
	if Instr(text,CharAmp)>0 then
		text=Replace(text,CharAmp,HtmlAmp)
	end if
	if Instr(text,CharLt)>0 then
		text=Replace(text,CharLt,HtmlLt)
	end if
	if Instr(text,CharGt)>0 then
		text=Replace(text,CharGt,HtmlGt)
	end if
	if Instr(text,CharQuot)>0 then
		text=Replace(text,CharQuot,HtmlQuot)
	end if
	if Instr(text,CharApos)>0 then
		text=Replace(text,CharApos,HtmlAposCode)
	end if

	HtmlEncode=text
End Function

Function HtmlDecode(byref text)
	if Instr(text,HtmlAmp)>0 then
		text = Replace(text, HtmlAmp, CharAmp)
	end if
	if Instr(text,HtmlLt)>0 then
		text = Replace(text, HtmlLt, CharLt)
	end if
	if Instr(text,HtmlGt)>0 then
		text = Replace(text, HtmlGt, CharGt)
	end if
	if Instr(text,HtmlQuot)>0 then
		text = Replace(text, HtmlQuot, CharQuot)
	end if
	if Instr(text,HtmlApos)>0 then
		text = Replace(text, HtmlApos, CharApos)
	end if
	if Instr(text,HtmlAposCode)>0 then
		text = Replace(text, HtmlAposCode, CharApos)
	end if
	if Instr(text,HtmlNbsp)>0 then
		text = Replace(text, HtmlNbsp, CharSpace)
	end if

	HtmlDecode = text
End Function

function textfilter(byref str,byval filterScript)
if str=CharEmpty then
	textfilter=CharEmpty
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

	if InStr(strContent,vbCrLf)>0 then
		strContent=replace(strContent,vbCrLf,CharSpace)
	end if
	if InStr(strContent,vbCr)>0 then
		strContent=replace(strContent,vbCr,CharSpace)
	end if
	if InStr(strContent,vbLf)>0 then
		strContent=replace(strContent,vbLf,CharSpace)
	end if

	textfilter=strContent
end if
end function

Sub NewLineToHtmlBr(byref str)
	if InStr(str,vbCrLf)>0 then
		str=replace(str,vbCrLf,htmlBr)
	end if
	if InStr(str,vbCr)>0 then
		str=replace(str,vbCr,htmlBr)
	end if
	if InStr(str,vbLf)>0 then
		str=replace(str,vbLf,htmlBr)
	end if
End Sub
%>