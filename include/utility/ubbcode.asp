<%
function UBBCode(byref strContent,byval allUbbFlags)
if Len(strContent)=0 then
	UBBCode=""
else
	const embed_prefix="<div class=""ubb-wrapper"">"
	const embed_postfix="</div>"
	dim re,i_count,originalStr
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True

	if UbbFlag_image or allUbbFlags then
		re.Pattern="\[img\](.[^\[]*)\[\/img\]"
		strContent=re.Replace(strContent,embed_prefix & "<img src=""$1""/>" & embed_postfix)
		re.Pattern="\[img=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/img\]"
		strContent=re.Replace(strContent,embed_prefix & "<img src=""$3"" style=""width:$1px;height:$2px;""/>" & embed_postfix)
	end if

	if UbbFlag_url or allUbbFlags then
		re.Pattern="\[URL\](\w+:\/\/[^\[]*)\[\/URL\]"
		strContent= re.Replace(strContent,"<a href=""$1"" target=""_blank"">$1</a>")
		re.Pattern="\[URL\]([^\[]*)\[\/URL\]"
		strContent= re.Replace(strContent,"<a href=""http://$1"" target=""_blank"">$1</a>")

		re.Pattern="\[URL=(\w+:\/\/[^\[]*)\]([^\[]*)(\[\/URL\])"
		strContent= re.Replace(strContent,"<a href=""$1"" target=""_blank"">$2</a>")
		re.Pattern="\[URL=([^\[]*)\]([^\[]*)(\[\/URL\])"
		strContent= re.Replace(strContent,"<a href=""http://$1"" target=""_blank"">$2</a>")

		re.Pattern="\[EMAIL\]mailto:([^\[]*)\[\/EMAIL\]"
		strContent= re.Replace(strContent,"<a href=""mailto:$1"" target=""_blank"">$1</a>")
		re.Pattern="\[EMAIL\](.[^\[]*)\[\/EMAIL\]"
		strContent= re.Replace(strContent,"<a href=""mailto:$1"" target=""_blank"">$1</a>")
	end if

	if UbbFlag_player or allUbbFlags then
		'Flash
		re.Pattern="\[FLASH\]([^\[]*)\[\/FLASH\]"
		strContent= re.Replace(strContent,embed_prefix & "<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000""  codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"">" & _
											"<param name=""movie"" value=""$1"" />" & _
											"<param name=""quality"" value=""high"" />" & _
											"<embed src=""$1"" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash""></embed>" & _
											"</object>" & embed_postfix)

		re.Pattern="\[FLASH=*([0-9]*),*([0-9]*)\]([^\[]*)\[\/FLASH\]"
		strContent= re.Replace(strContent,embed_prefix & "<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"" style=""width:$1px;height:$2px;"">" & _
											"<param name=""movie"" value=""$3"" />" & _
											"<param name=""quality"" value=""high"" />" & _
											"<embed src=""$3"" style=""width:$1px;height:$2px;"" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash""></embed>" & _
											"</object>" & embed_postfix)

		'Media Player
		re.Pattern="\[MP\]([^\[]*)\[\/MP\]"
		strContent=re.Replace(strContent,embed_prefix & "<object classid=""clsid:22d6f312-b0f6-11d0-94ab-0080c74c7e95""   codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=6,4,5,715"">" & _
											"<param name=""Filename"" value=""$1"">" & _
											"<param name=""ShowStatusBar"" value=""true"" />" & _
											"<param name=""AutoStart"" value=""false"" />" & _
											"<embed src=""$1"" type=""application/x-oleobject"" codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701"" flename=""mp""></embed>" & _
											"</object>" & embed_postfix)

		re.Pattern="\[MP=*([0-9]*),*([0-9]*)\]([^\[]*)\[\/MP]"
		strContent=re.Replace(strContent,embed_prefix & "<object classid=""clsid:22d6f312-b0f6-11d0-94ab-0080c74c7e95""  codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=6,4,5,715"" style=""width:$1px;height:$2px;"">" & _
											"<param name=""Filename"" value=""$3"">" & _
											"<param name=""ShowStatusBar"" value=""true"" />" & _
											"<param name=""AutoStart"" value=""false"" />" & _
											"<embed src=""$3"" style=""width:$1px;height:$2px;"" type=""application/x-oleobject"" codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701"" flename=""mp""></embed>" & _
											"</object>" & embed_postfix)

		'Real Player
		re.Pattern="\[RM\]([^\[]*)\[\/RM\]"
		strContent=re.Replace(strContent,embed_prefix & "<object classid=""clsid:cfcdaa03-8be4-11cf-b84b-0020afbbccfa"" style=""width:320px;height:200px;"">" & _
											"<param name=""src"" value=""$1"" />" & _
											"<param name=""console"" value=""RaPlayer"" />" & _
											"<param name=""controls"" value=""Imagewindow,StatusBar,ControlPanel"">" & _
											"<param name=""autostart"" value=""false"">" & _
											"<embed src=""$1"" type=""application/vnd.rn-realmedia"" console=""RaPlayer"" controls=""Imagewindow,StatusBar,ControlPanel"" autostart=""false""></embed>" & _
											"</object>" & embed_postfix)

		re.Pattern="\[RM=*([0-9]*),*([0-9]*)\]([^\[]*)\[\/RM]"
		strContent=re.Replace(strContent,embed_prefix & "<object classid=""clsid:cfcdaa03-8be4-11cf-b84b-0020afbbccfa"" style=""width:$1px;height:$2px;"">" & _
											"<param name=""src"" value=""$3"" />" & _
											"<param name=""console"" value=""RaPlayer"" />" & _
											"<param name=""controls"" value=""Imagewindow,StatusBar,ControlPanel"">" & _
											"<param name=""autostart"" value=""false"">" & _
											"<embed src=""$3"" style=""width:$1px;height:$2px;"" type=""application/vnd.rn-realmedia"" console=""RaPlayer"" controls=""Imagewindow,StatusBar,ControlPanel"" autostart=""false""></embed>" & _
											"</object>" & embed_postfix)
	end if

	if UbbFlag_autourl or allUbbFlags then
		'http://
		re.Pattern = "(^|[^>=""])((?:https|http|sftp|ftp|rtsp|mms|ed2k)://([A-Za-z0-9\./=\?%\-{}&#_|~`@':+!])+)"
		strContent = re.Replace(strContent, "$1<a target=""_blank"" href=""$2"">$2</a>")

		'www
		re.Pattern = "(^|[^>=""/])(www\.[A-Za-z0-9\./=\?%\-{}&#_|~`@':+!]+)"
		strContent = re.Replace(strContent,"$1<a target=""_blank"" href=""http://$2"">$2</a>")
	end if

	if UbbFlag_face or allUbbFlags then
		if isnumeric(SmallFaceCount) and SmallFaceCount<>"" then
			re.Pattern="\[face(\d{1," & len(cstr(SmallFaceCount)) & "})\]"
			strContent=re.Replace(strContent,"<img src=""asset/smallface/$1.gif"" />")
		end if
	end if

	for i_count=1 to 5
		originalStr=strContent

		if UbbFlag_paragraph or allUbbFlags then
			re.Pattern="\[QUOTE\](.[^\[]*)\[\/QUOTE\]"
			strContent=re.Replace(strContent,"<blockquote>$1</blockquote>")
			re.Pattern="\[li\](.[^\[]*)\[\/li\]"
			strContent=re.Replace(strContent,"<li>$1</li>")
		end if

		if UbbFlag_fontstyle or allUbbFlags then
			re.Pattern="\[font=(.[^\[\]]*)\](.[^\[]*)\[\/font\]"
			strContent=re.Replace(strContent,"<span style=""font-family:$1"">$2</span>")
			re.Pattern="\[i\](.[^\[]*)\[\/i\]"
			strContent=re.Replace(strContent,"<span style=""font-style:italic"">$1</span>")
			re.Pattern="\[b\](.[^\[]*)\[\/b\]"
			strContent=re.Replace(strContent,"<span style=""font-weight:bold"">$1</span>")
			re.Pattern="\[u\](.[^\[]*)\[\/u\]"
			strContent=re.Replace(strContent,"<span style=""text-decoration:underline"">$1</span>")
			re.Pattern="\[strike\](.[^\[]*)\[\/strike\]"
			strContent=re.Replace(strContent,"<span style=""text-decoration:line-through"">$1</span>")
		end if

		if UbbFlag_fontcolor or allUbbFlags then
			re.Pattern="\[color=(.[^\[\]]*)\](.[^\[]*)\[\/color\]"
			strContent=re.Replace(strContent,"<span style=""color:$1"">$2</span>")
			re.Pattern="\[bgcolor=\s*(.[^\[\]]*)\](.[^\[]*)\[\/bgcolor\]"
			strContent=re.Replace(strContent,"<span style=""background-color:$1"">$2</span>")
		end if

		if UbbFlag_alignment or allUbbFlags then
			re.Pattern="\[align=(.[^\[\]]*)\](.[^\[]*)\[\/align\]"
			strContent=re.Replace(strContent,"<span style=""width:100%; display:block; text-align:$1"">$2</span>")
			re.Pattern="\[center\](.[^\[]*)\[\/center\]"
			strContent=re.Replace(strContent,"<span style=""width:100%; display:block; text-align:center"">$1</span>")
		end if

		'if UbbFlag_movement or allUbbFlags then
		'	re.Pattern="\[fly\](.[^\[]*)\[\/fly\]"
		'	strContent=re.Replace(strContent,"<marquee behavior=""alternate"" direction=""left"" scrollamount=""3"">$1</marquee>")
		'	re.Pattern="\[move\](.[^\[]*)\[\/move\]"
		'	strContent=re.Replace(strContent,"<marquee direction=""left"" scrollamount=""3"">$1</marquee>")
		'end if

		'if UbbFlag_cssfilter or allUbbFlags then
		'	re.Pattern="\[GLOW=\s*([0-9]*),\s*(#*[a-z0-9]*),\s*([0-9]*)\](.[^\[]*)\[\/GLOW]"
		'	strContent=re.Replace(strContent,"<table style=""width:$1px; filter:glow(color=$2, strength=$3)""><tr><td>$4</td></tr></table>")
		'	re.Pattern="\[SHADOW=\s*([0-9]*),\s*(#*[a-z0-9]*),\s*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
		'	strContent=re.Replace(strContent,"<table style=""width:$1px; filter:shadow(color=$2, strength=$3)""><tr><td>$4</td></tr></table>")
		'end if

		if originalStr=strContent then exit for
	next

	re.pattern="<(.*?)(?:javascript|vbscript)\s*:(.*?)>"
	strContent=re.replace(strContent,"<$1$2>")
	re.pattern="(<.*?style\s*=.*?):expression(.*?>)"
	strContent=re.replace(strContent,"$1$2")
	set re=Nothing

	strContent=replace(strContent,"width:px;","width:auto;")
	strContent=replace(strContent,"height:px;","height:auto;")

	strContent=replace(strContent,chr(13)&chr(10),chr(10))
	strContent=replace(strContent,chr(10),"<br/>" &chr(13)&chr(10))
	UBBCode=strContent
end if
end function

function convertstr(byref str,byval htmlflag,byval allUbbFlags)
	dim tHTML,tUBB,tNewline
	tHTML=CBool(htmlflag and 1)
	tUBB=CBool(htmlflag and 2)
	tNewline=CBool(htmlflag and 4)

	if tHTML then
		str=Replace(str,"<script","&lt;script")
		str=Replace(str,"</script>","&lt;/script>")
	else
		str=server.HTMLEncode(str)
		str=replace(str,chr(9),"        ")
		'str=replace(str," ","&nbsp;")
	end if
	if tUBB=true then str=ubbcode(str,allUbbFlags)
	if tHTML=false and tUBB=false then
		if tNewline=true then
			str=replace(str,chr(13)&chr(10),"<br/>")
		else
			str=replace(str,chr(13)&chr(10)," ")
		end if
	end if
end function
%>