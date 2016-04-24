<%
function UBBCode(byref strContent,byval allUbbFlags)
if strContent="" then
	UBBCode=""
else
	const embed_prefix="<div class=""ubb-wrapper"">"
	const embed_postfix="</div>"
	dim i_count,originalStr

	Dim reCase
	Set reCase=new RegExp
	reCase.IgnoreCase=false
	reCase.Global=True

	Dim NeedSecureCheck
	NeedSecureCheck=false

	Dim UbbQuoteStart
	UbbQuoteStart=Instr(strContent,"[")
	if UbbQuoteStart > 0 then
		if Instr(UbbQuoteStart+1, strContent,"[/") > 0 then
			NeedSecureCheck=true

			if UbbFlag_image or allUbbFlags then
				reCase.Pattern="\[[iI][mM][gG]\]([^\[]+)\[\/[iI][mM][gG]\]"
				strContent=reCase.Replace(strContent,embed_prefix & "<img src=""$1""/>" & embed_postfix)
				reCase.Pattern="\[[iI][mM][gG]=(\d+),(\d+)\]([^\[]+)\[\/[iI][mM][gG]\]"
				strContent=reCase.Replace(strContent,embed_prefix & "<img src=""$3"" style=""width:$1px;height:$2px;""/>" & embed_postfix)
			end if

			if UbbFlag_url or allUbbFlags then
				reCase.Pattern="\[[uU][rR][lL]\](\w+://[^\[]+)\[\/[uU][rR][lL]\]"
				strContent= reCase.Replace(strContent,"<a href=""$1"" target=""_blank"">$1</a>")
				reCase.Pattern="\[[uU][rR][lL]\]([^\[]+)\[\/[uU][rR][lL]\]"
				strContent= reCase.Replace(strContent,"<a href=""http://$1"" target=""_blank"">$1</a>")

				reCase.Pattern="\[[uU][rR][lL]=(\w+://[^\[]+)\]([^\[]+)(\[\/[uU][rR][lL]\])"
				strContent= reCase.Replace(strContent,"<a href=""$1"" target=""_blank"">$2</a>")
				reCase.Pattern="\[[uU][rR][lL]=([^\[]+)\]([^\[]+)(\[\/[uU][rR][lL]\])"
				strContent= reCase.Replace(strContent,"<a href=""http://$1"" target=""_blank"">$2</a>")

				reCase.Pattern="\[[eE][mM][aA][iI][lL]\]mailto:([^\[]+)\[\/[eE][mM][aA][iI][lL]\]"
				strContent= reCase.Replace(strContent,"<a href=""mailto:$1"" target=""_blank"">$1</a>")
				reCase.Pattern="\[[eE][mM][aA][iI][lL]\]([^\[]+)\[\/[eE][mM][aA][iI][lL]\]"
				strContent= reCase.Replace(strContent,"<a href=""mailto:$1"" target=""_blank"">$1</a>")
			end if

			if UbbFlag_player or allUbbFlags then
				'Flash
				reCase.Pattern="\[[fF][lL][aA][sS][hH]\]([^\[]+)\[\/[fF][lL][aA][sS][hH]\]"
				strContent= reCase.Replace(strContent,embed_prefix & "<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000""  codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"">" & _
						"<param name=""movie"" value=""$1"" />" & _
						"<param name=""quality"" value=""high"" />" & _
						"<embed src=""$1"" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash""></embed>" & _
						"</object>" & embed_postfix)

				reCase.Pattern="\[[fF][lL][aA][sS][hH]=(\d+),(\d+)\]([^\[]+)\[\/[fF][lL][aA][sS][hH]\]"
				strContent= reCase.Replace(strContent,embed_prefix & "<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"" style=""width:$1px;height:$2px;"">" & _
						"<param name=""movie"" value=""$3"" />" & _
						"<param name=""quality"" value=""high"" />" & _
						"<embed src=""$3"" style=""width:$1px;height:$2px;"" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash""></embed>" & _
						"</object>" & embed_postfix)

				'Media Player
				reCase.Pattern="\[[mM][pP]\]([^\[]+)\[\/[mM][pP]\]"
				strContent=reCase.Replace(strContent,embed_prefix & "<object classid=""clsid:22d6f312-b0f6-11d0-94ab-0080c74c7e95""   codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=6,4,5,715"">" & _
						"<param name=""Filename"" value=""$1"">" & _
						"<param name=""ShowStatusBar"" value=""true"" />" & _
						"<param name=""AutoStart"" value=""false"" />" & _
						"<embed src=""$1"" type=""application/x-oleobject"" codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701"" flename=""mp""></embed>" & _
						"</object>" & embed_postfix)

				reCase.Pattern="\[[mM][pP]=(\d+),(\d+)\]([^\[]+)\[\/[mM][pP]]"
				strContent=reCase.Replace(strContent,embed_prefix & "<object classid=""clsid:22d6f312-b0f6-11d0-94ab-0080c74c7e95""  codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=6,4,5,715"" style=""width:$1px;height:$2px;"">" & _
						"<param name=""Filename"" value=""$3"">" & _
						"<param name=""ShowStatusBar"" value=""true"" />" & _
						"<param name=""AutoStart"" value=""false"" />" & _
						"<embed src=""$3"" style=""width:$1px;height:$2px;"" type=""application/x-oleobject"" codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701"" flename=""mp""></embed>" & _
						"</object>" & embed_postfix)

				'Html5 Video
				reCase.Pattern="\[video\]([^\[]+)\[\/video\]"
				strContent=reCase.Replace(strContent,embed_prefix & "<video src=""$1"" controls=""controls""></video>" &  embed_postfix)

				reCase.Pattern="\[video=(\d+),(\d+)\]([^\[]+)\[\/video]"
				strContent=reCase.Replace(strContent,embed_prefix & "<video src=""$3"" style=""width:$1px;height:$2px;"" controls=""controls""></video>" &  embed_postfix)
			end if

			for i_count=1 to 5
				originalStr=strContent

				if UbbFlag_paragraph or allUbbFlags then
					reCase.Pattern="\[[qQ][uU][oO][tT][eE]\]([^\[]+)\[\/[qQ][uU][oO][tT][eE]\]"
					strContent=reCase.Replace(strContent,"<blockquote>$1</blockquote>")

					reCase.Pattern="\[\/[lL][iI]\]\s*\[[lL][iI]\]"
					strContent=reCase.Replace(strContent,"[/li][li]")
					reCase.Pattern="\[[lL][iI]\]([^\[]+)\[\/[lL][iI]\]"
					strContent=reCase.Replace(strContent,"<li>$1</li>")
				end if

				if UbbFlag_fontstyle or allUbbFlags then
					reCase.Pattern="\[[fF][oO][nN][tT]=([^\[\]]+)\]([^\[]+)\[\/[fF][oO][nN][tT]\]"
					strContent=reCase.Replace(strContent,"<span style=""font-family:$1"">$2</span>")
					reCase.Pattern="\[[iI]\]([^\[]+)\[\/[iI]\]"
					strContent=reCase.Replace(strContent,"<span style=""font-style:italic"">$1</span>")
					reCase.Pattern="\[[bB]\]([^\[]+)\[\/[bB]\]"
					strContent=reCase.Replace(strContent,"<span style=""font-weight:bold"">$1</span>")
					reCase.Pattern="\[[uU]\]([^\[]+)\[\/[uU]\]"
					strContent=reCase.Replace(strContent,"<span style=""text-decoration:underline"">$1</span>")
					reCase.Pattern="\[[sS][tT][rR][iI][kK][eE]\]([^\[]+)\[\/[sS][tT][rR][iI][kK][eE]\]"
					strContent=reCase.Replace(strContent,"<span style=""text-decoration:line-through"">$1</span>")
				end if

				if UbbFlag_fontcolor or allUbbFlags then
					reCase.Pattern="\[[cC][oO][lL][oO][rR]=([^\[\]]+)\]([^\[]+)\[\/[cC][oO][lL][oO][rR]\]"
					strContent=reCase.Replace(strContent,"<span style=""color:$1"">$2</span>")
					reCase.Pattern="\[[bB][gG][cC][oO][lL][oO][rR]=([^\[\]]+)\]([^\[]+)\[\/[bB][gG][cC][oO][lL][oO][rR]\]"
					strContent=reCase.Replace(strContent,"<span style=""background-color:$1"">$2</span>")
				end if

				if UbbFlag_alignment or allUbbFlags then
					reCase.Pattern="\[[aA][lL][iI][gG][nN]=([^\[\]]+)\]([^\[]+)\[\/[aA][lL][iI][gG][nN]\]"
					strContent=reCase.Replace(strContent,"<span style=""display:block; text-align:$1"">$2</span>")
					reCase.Pattern="\[[cC][eE][nN][tT][eE][rR]\]([^\[]+)\[\/[cC][eE][nN][tT][eE][rR]\]"
					strContent=reCase.Replace(strContent,"<span style=""display:block; text-align:center"">$1</span>")
				end if

				if Len(originalStr)=Len(strContent) then
					if originalStr=strContent then exit for
				end if
			next
		end if

		if UbbFlag_face or allUbbFlags then
			reCase.Pattern="\[[fF][aA][cC][eE](\d+)\]"
			strContent=reCase.Replace(strContent,"<img src=""asset/smallface/$1.gif"" />")
		end if
	end if

	if UbbFlag_autourl or allUbbFlags then
		NeedSecureCheck=true

		'http://
		reCase.Pattern = "(^|[^>=""])((?:[hH][tT][tT][pP][sS]?|[sS][fF][tT][pP]|[fF][tT][pP]|[rR][tT][sS][pP]|[mM][mM][sS]|[eE][dD]2[kK])://([A-Za-z0-9\./=\?%\-()\[\]{}&#_|~`@':+!])+)"
		strContent = reCase.Replace(strContent, "$1<a target=""_blank"" href=""$2"">$2</a>")

		'www
		reCase.Pattern = "(^|[^>=""/])([wW][wW][wW]\.[A-Za-z0-9\./=\?%\-()\[\]{}&#_|~`@':+!]+)"
		strContent = reCase.Replace(strContent,"$1<a target=""_blank"" href=""http://$2"">$2</a>")
	end if

	if NeedSecureCheck then
		reCase.Pattern="([jJ][aA][vV][aA][sS][cC][rR][iI][pP][tT]):"
		strContent=reCase.Replace(strContent,"$1 :")
		reCase.Pattern="([vV][bB][sS][cC][rR][iI][pP][tT]):"
		strContent=reCase.Replace(strContent,"$1 :")
		reCase.Pattern="(:[eE][xX][pP][rR][eE][sS][sS][iI][oO][nN])\("
		strContent=reCase.Replace(strContent,"$1 (")
		reCase.Pattern="([bB][eE][hH][aA][vV][iI][oO][rR]\s*):"
		strContent=reCase.Replace(strContent,"$1-ban:")
	end if

	set re=Nothing

	if InStr(strContent,charCrLf)>0 then
		strContent=replace(strContent,charCrLf,charLf)
	end if
	if InStr(strContent,charCr)>0 then
		strContent=replace(strContent,charCr,charLf)
	end if
	if InStr(strContent,charLf)>0 then
		strContent=replace(strContent,charLf,"<br/>")
	end if

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
		str=HtmlEncode(str)
	end if

	if tUBB then
		str=ubbcode(str,allUbbFlags)
	end if

	if Not tHTML and Not tUBB and tNewline then
		if InStr(str,charCrLf)>0 then
			str=replace(str,charCrLf,charLf)
		end if
		if InStr(str,charCr)>0 then
			str=replace(str,charCr,charLf)
		end if
		if InStr(str,charLf)>0 then
			str=replace(str,charLf,"<br/>")
		end if
	end if
end function
%>