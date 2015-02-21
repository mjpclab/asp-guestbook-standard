<%
function UBBCode(byref strContent,byval allflags)
if strContent="" then
	UBBCode=""
else
	const embed_prefix1="<table celpadding=""0"" cellspacing=""0"" style=""display:inline;""><tr><td><a href=""$1"" target=""_blank"">"
	const embed_prefix3="<table celpadding=""0"" cellspacing=""0"" style=""display:inline;""><tr><td><a href=""$3"" target=""_blank"">"
	const embed_postfix="</a></td></tr></table>"
	dim re,i_count,originalStr
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True

	for i_count=1 to 5
		originalStr=strContent

		if UbbFlag_image or allflags then
			re.Pattern="\[img\](.[^\[]*)\[\/img\]"
			strContent=re.Replace(strContent,embed_prefix1 & "<img src=""$1"" style=""border-width:0px;"" />" & embed_postfix)
			re.Pattern="\[img=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/img\]"
			strContent=re.Replace(strContent,embed_prefix3 & "<img src=""$3"" style=""width:$1px; height:$2px; border-width:0px;"" />" & embed_postfix)
		end if
		
		if UbbFlag_url or allflags then
			re.Pattern="\[URL\](http:\/\/.[^\[]*)\[\/URL\]"
			strContent= re.Replace(strContent,"<a href=""$1"" target=""_blank"">$1</a>")
			re.Pattern="\[URL\](.[^\[]*)\[\/URL\]"
			strContent= re.Replace(strContent,"<a href=""http://$1"" target=""_blank"">$1</a>")

			re.Pattern="\[URL=(http:\/\/.[^\[]*)\](.[^\[]*)(\[\/URL\])"
			strContent= re.Replace(strContent,"<a href=""$1"" target=""_blank"">$2</a>")
			re.Pattern="\[URL=(.[^\[\]]*)\](.[^\[]*)(\[\/URL\])"
			strContent= re.Replace(strContent,"<a href=""http://$1"" target=""_blank"">$2</a>")

			re.Pattern="\[EMAIL\]mailto:(.[^\[]*)\[\/EMAIL\]"
			strContent= re.Replace(strContent,"<a href=""mailto:$1"" target=""_blank"">$1</a>")
			re.Pattern="\[EMAIL\](.[^\[]*)\[\/EMAIL\]"
			strContent= re.Replace(strContent,"<a href=""mailto:$1"" target=""_blank"">$1</a>")
		end if
		
		if UbbFlag_player or allflags then
			'Flash
			re.Pattern="\[FLASH\](.[^\[]*)\[\/FLASH\]"
			strContent= re.Replace(strContent,embed_prefix1 & "<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000""  codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"">" & _
												"<param name=""movie"" value=""$1"" />" & _
												"<param name=""quality"" value=""high"" />" & _
												"<embed src=""$1"" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash""></embed>" & _
												"</object>" & embed_postfix)
												
			re.Pattern="\[FLASH=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/FLASH\]"
			strContent= re.Replace(strContent,embed_prefix3 & "<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"" style=""width:$1px; height:$2px;"">" & _
												"<param name=""movie"" value=""$3"" />" & _
												"<param name=""quality"" value=""high"" />" & _
												"<embed src=""$3"" width=""$1"" height=""$2"" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash""></embed>" & _
												"</object>" & embed_postfix)
			
			'Media Player
			re.Pattern="\[MP\](.[^\[]*)\[\/MP\]"
			strContent=re.Replace(strContent,embed_prefix1 & "<object classid=""clsid:22d6f312-b0f6-11d0-94ab-0080c74c7e95""   codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=6,4,5,715"">" & _
												"<param name=""Filename"" value=""$1"">" & _
												"<param name=""ShowStatusBar"" value=""true"" />" & _
												"<param name=""AutoStart"" value=""false"" />" & _
												"<embed src=""$1"" type=""application/x-oleobject"" codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701"" flename=""mp""></embed>" & _
												"</object>" & embed_postfix)
			
			re.Pattern="\[MP=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/MP]"
			strContent=re.Replace(strContent,embed_prefix3 & "<object classid=""clsid:22d6f312-b0f6-11d0-94ab-0080c74c7e95""  codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=6,4,5,715"" style=""width:$1px; height:$2px;"">" & _
												"<param name=""Filename"" value=""$3"">" & _
												"<param name=""ShowStatusBar"" value=""true"" />" & _
												"<param name=""AutoStart"" value=""false"" />" & _
												"<embed src=""$3"" width=""$1"" height=""$2"" type=""application/x-oleobject"" codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701"" flename=""mp""></embed>" & _
												"</object>" & embed_postfix)

			'Real Player
			re.Pattern="\[RM\](.[^\[]*)\[\/RM\]"
			strContent=re.Replace(strContent,embed_prefix1 & "<object classid=""clsid:cfcdaa03-8be4-11cf-b84b-0020afbbccfa"" style=""width:320px;height:200px;"">" & _
												"<param name=""src"" value=""$1"" />" & _
												"<param name=""console"" value=""RaPlayer"" />" & _
												"<param name=""controls"" value=""Imagewindow,StatusBar,ControlPanel"">" & _
												"<param name=""autostart"" value=""false"">" & _
												"<embed src=""$1"" type=""application/vnd.rn-realmedia"" console=""RaPlayer"" controls=""Imagewindow,StatusBar,ControlPanel"" autostart=""false""></embed>" & _
												"</object>" & embed_postfix)

			re.Pattern="\[RM=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/RM]"
			strContent=re.Replace(strContent,embed_prefix3 & "<object classid=""clsid:cfcdaa03-8be4-11cf-b84b-0020afbbccfa"" style=""width:$1px; height:$2px;"">" & _
												"<param name=""src"" value=""$3"" />" & _
												"<param name=""console"" value=""RaPlayer"" />" & _
												"<param name=""controls"" value=""Imagewindow,StatusBar,ControlPanel"">" & _
												"<param name=""autostart"" value=""false"">" & _
												"<embed src=""$3"" width=""$1"" height=""$2"" type=""application/vnd.rn-realmedia"" console=""RaPlayer"" controls=""Imagewindow,StatusBar,ControlPanel"" autostart=""false""></embed>" & _
												"</object>" & embed_postfix)
		end if

		if UbbFlag_autourl or allflags then
			'http://
			re.Pattern = "^((?:http|ftp|rtsp|mms|ed2k)://((?!&nbsp;|\s|&quot;|&apos;)[A-Za-z0-9\./=\?%\-{}&#_|~`@':+!])+)"
			strContent = re.Replace(strContent,"<a target=""_blank"" href=""$1"">$1</a>")
			re.Pattern = "([^>=""])((?:http|ftp|rtsp|mms|ed2k)://((?!&nbsp;|\s|&quot;|&apos;)[A-Za-z0-9\./=\?%\-{}&#_|~`@':+!])+)"
			strContent = re.Replace(strContent, "$1<a target=""_blank"" href=""$2"">$2</a>")

			'www
			re.Pattern = "^(www\.((?!&nbsp;|\s|&quot;|&apos;)[A-Za-z0-9\./=\?%\-{}&#_|~`@':+!])+)"
			strContent = re.Replace(strContent,"<a target=""_blank"" href=""http://$1"">$1</a>")
			re.Pattern = "([^A-Za-z>=/\.\\""])(www\.((?!&nbsp;|\s|&quot;|&apos;)[A-Za-z0-9\./=\?%\-{}&#_|~`@':+!])+)"
			strContent = re.Replace(strContent, "$1<a target=""_blank"" href=""http://$2"">$2</a>")
		end if
		
		if UbbFlag_paragraph or allflags then
			re.Pattern="\[QUOTE\](.[^\[]*)\[\/QUOTE\]"
			strContent=re.Replace(strContent,"<blockquote>$1</blockquote>")
			re.Pattern="\[li\](.[^\[]*)\[\/li\]"
			strContent=re.Replace(strContent,"<li>$1</li>")
		end if
		
		if UbbFlag_fontstyle or allflags then
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
		
		if UbbFlag_fontcolor or allflags then
			re.Pattern="\[color=(.[^\[\]]*)\](.[^\[]*)\[\/color\]"
			strContent=re.Replace(strContent,"<span style=""color:$1"">$2</span>")
			re.Pattern="\[bgcolor=\s*(.[^\[\]]*)\](.[^\[]*)\[\/bgcolor\]"
			strContent=re.Replace(strContent,"<span style=""background-color:$1"">$2</span>")
		end if

		if UbbFlag_alignment or allflags then
			re.Pattern="\[align=(.[^\[\]]*)\](.[^\[]*)\[\/align\]"
			strContent=re.Replace(strContent,"<span style=""width:100%; display:block; text-align:$1"">$2</span>")
			re.Pattern="\[center\](.[^\[]*)\[\/center\]"
			strContent=re.Replace(strContent,"<span style=""width:100%; display:block; text-align:center"">$1</span>")
		end if
		
		'if UbbFlag_movement or allflags then
		'	re.Pattern="\[fly\](.[^\[]*)\[\/fly\]"
		'	strContent=re.Replace(strContent,"<marquee behavior=""alternate"" direction=""left"" scrollamount=""3"">$1</marquee>")
		'	re.Pattern="\[move\](.[^\[]*)\[\/move\]"
		'	strContent=re.Replace(strContent,"<marquee direction=""left"" scrollamount=""3"">$1</marquee>")
		'end if
		
		'if UbbFlag_cssfilter or allflags then
		'	re.Pattern="\[GLOW=\s*([0-9]*),\s*(#*[a-z0-9]*),\s*([0-9]*)\](.[^\[]*)\[\/GLOW]"
		'	strContent=re.Replace(strContent,"<table style=""width:$1px; filter:glow(color=$2, strength=$3)""><tr><td>$4</td></tr></table>")
		'	re.Pattern="\[SHADOW=\s*([0-9]*),\s*(#*[a-z0-9]*),\s*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
		'	strContent=re.Replace(strContent,"<table style=""width:$1px; filter:shadow(color=$2, strength=$3)""><tr><td>$4</td></tr></table>")
		'end if

		if UbbFlag_face or allflags then
			if isnumeric(SmallFaceCount) and SmallFaceCount<>"" then
				re.Pattern="\[face(\d{1," & len(cstr(SmallFaceCount)) & "})\]"
				strContent=re.Replace(strContent,"<img src=""" & SmallFacePath & "$1.gif"" />")
			end if
		end if
		
		on error resume next
			re.pattern="<(.*?)(?:javascript|vbscript)\s*:(.*?)>"
			if err.number=0 then strContent=re.replace(strContent,"<$1$2>")

			re.pattern="(<.*?style\s*=.*?):expression(.*?>)"
			if err.number=0 then strContent=re.replace(strContent,"$1$2")
		on error goto 0
		
		if originalStr=strContent then exit for
	next	

	set re=Nothing
	strContent=replace(strContent,chr(13)&chr(10),chr(10))
	strContent=replace(strContent,chr(10),"<br/>" &chr(13)&chr(10))
	UBBCode=strContent
end if
end function
%>