<%
dim admin_name,admin_faceid,admin_faceurl,admin_email,admin_qqid,admin_msnid,admin_homepage,p_rs
'===================
sub getadmininfo()
	set rs3=server.CreateObject("ADODB.Recordset")
	rs3.open sql_common2_getadmininfo,cn,0,1,1

	admin_name="" & rs3.fields("name") & ""
	admin_faceid="" & rs3.fields("faceid") & ""
	admin_faceurl="" & rs3.fields("faceurl") & ""
	admin_email="" & rs3.fields("email") & ""
	admin_qqid="" & rs3.fields("qqid") & ""
	admin_msnid="" & rs3.fields("msnid") & ""
	admin_homepage="" & rs3.fields("homepage") & ""
	
	rs3.close
	set rs3=nothing
	
	if clng(admin_faceid)<>0 and admin_faceurl="" then admin_faceurl=FacePath & admin_faceid  & ".gif"
end sub
'===================
sub showadminicons()
	if ReplyInWord=false then
		Response.Write admin_name & "<br/>"
		if admin_faceurl<>"" then Response.Write "<img alt="""" src=""" & admin_faceurl & """ class=""face"" />"
							
		if admin_email<>"" or admin_qqid<>"" or admin_msnid<>"" or admin_homepage<>"" then
			Response.Write "<br/>"
		end if
	end if

	if admin_email<>"" then Response.Write "<a href=""mailto:" &admin_email& """ title=""版主邮箱：" &admin_email& """><img alt="""" src=""image/icon_mail.gif"" class=""imgicon"" /></a> "
	if admin_qqid<>"" then Response.Write "<img src=""image/icon_qq.gif"" alt=""版主QQ：" &admin_qqid& """ class=""imgicon""" &js_status& " /> "
	if admin_msnid<>"" then Response.Write "<img src=""image/icon_msn.gif"" alt=""版主MSN：" &admin_msnid& """ class=""imgicon""" &js_status& " /> "
	if admin_homepage<>"" then Response.Write "<a href=""" &admin_homepage& """ target=""_blank"" title=""版主主页：" &admin_homepage& """><img alt="""" src=""image/icon_homepage.gif"" class=""imgicon"" /></a> "

end sub
'===================
sub showguestreplyicons(byval follow_id,byval parent_id,byval show_reply)
dim url,out_str

url="leaveword.asp?follow=" & follow_id
if left(pagename,8)="showword" then url=url & "&return=showword"
url=Server.HTMLEncode(url)

out_str="<div style=""width:100%; text-align:right; padding-bottom:5px;"">"
if parent_id<0 then
	out_str=out_str & "<span style=""padding-left:15px;""><img alt="""" src=""image/icon_toplocked.gif"" class=""imgicon"" />(置顶)</span>"
end if
if show_reply then
	out_str=out_str & "<span style=""padding-left:15px;""><a href=""" &url& """><img alt="""" src=""image/icon_reply.gif"" class=""imgicon"" />[回复]</a></span>"
end if
out_str=out_str & "</div>"

Response.Write out_str
end sub
'===================
sub innerreply(byref rs2)
	%>
	<tr>
		<td class="embedbox">
			<div class="embedbox">
				<table class="embedbox" cellpadding="0" cellspacing="0"><tr><td>
					<span style="font-weight:bold; color:<%=TableReplyColor%>;">版主<%if admin_name<>"" then response.write "(" & admin_name & ")"%>回复：</span>
					<%
					response.write "<span style=""color:" &TableReplyColor& """>(" & rs2("replydate") & ")</span> "
									
					showadminicons()
					response.write "<hr size=""1"" />"
					if admin_faceurl<>"" then Response.Write "<img alt="""" src=""" & admin_faceurl & """ class=""face"" style=""float:left;margin:0px 5px 5px 0px;"" />"
									
					response.write "<span style=""color:" & TableReplyContentColor & """>"
					if encrypted=true and pagename<>"showword" and left(pagename,5)<>"admin" then
						response.Write "<img alt="""" src=""image/icon_key.gif"" class=""imgicon"" />(需要预设的密码才能查看...)[<a href=""showword.asp?id=" &rs2("id")& """>点击这里验证...</a>]"
					else
						reply_htmlflag=rs2("htmlflag")
						reply_txt=rs2("reinfo")
													
						convertstr reply_txt,reply_htmlflag,true
						Response.Write reply_txt
					end if
					response.write "</span>"
					%>
				</td></tr></table>								
			</div>
		</td>
	</tr>
	<%
end sub
'===================
sub outerreply(byref rs2)
	%>
	<tr>
	  <td colspan="2" class="replytitle">[版主回复：]</td>
	</tr>
	<tr>
	  	<td class="tableleft">
	  		<%showadminicons()%><br/>
	  		<%=rs2("replydate")%>
	  	</td>
	  	<td class="wordscontent">
	  		<%
	  		response.write "<span style=""color:" & TableReplyContentColor & """>"
	  		if encrypted=true and pagename<>"showword" and left(pagename,5)<>"admin" then
	  			response.Write "<img alt="""" src=""image/icon_key.gif"" class=""imgicon"" />(需要预设的密码才能查看...)[<a href=""showword.asp?id=" &rs2("id")& """>点击这里验证...</a>]"
	  		else
	  			reply_htmlflag=rs2("htmlflag")
	  			reply_txt=rs2("reinfo")
										
	  			convertstr reply_txt,reply_htmlflag,true
	  			Response.Write reply_txt
	  		end if
	  		response.write "</span>"
	  		%>
	  	</td>
	</tr>
	<%
end sub
'===================
sub inneraudit()
	%>
	<tr>
		<td class="embedbox">
			<div class="embedbox">
				<table class="embedbox" cellpadding="0" cellspacing="0"><tr><td>
					<span style="color:<%=TableTitleColor%>;">(留言待审核...)</span>
					<hr size="1" />
					<img alt="" src="image/icon_wait2pass.gif" class="imgicon" />(留言待审核...)
				</td></tr></table>								
			</div>
		</td>
	</tr>
	<%
end sub
'===================
sub outeraudit()
	%>
	<tr>
		<td class="tableleft">
			<%
			fid=rs("faceid")
			if isnumeric(fid) and StatusShowHead=true then
				if fid>=1 and fid<=FaceCount then
					%>
					<img alt="" src="<%=FacePath & cstr(fid)& ".gif"%>" class="face" />
					<%
				end if
			end if
			Response.Write "<br/>"&rs("logdate")
			%>
		</td>
	    <td class="tableright">
			<table cellpadding="2" cellspacing="0" style="width:100%; border-width:0px; border-collapse:collapse;">
				<tr>
					<td class="wordstitle"><div style="width:100%; height:21px; line-height:21px; vertical-align:middle; overflow:hidden;">(留言待审核...)</div></td>
				</tr>
				<tr>
					<td class="wordscontent">
						<%if rs.Fields("parent_id")<=0 and left(pagename,5)<>"admin" then showguestreplyicons rs.Fields("id"),rs.Fields("parent_id"),StatusWrite and StatusGuestReply and clng(rs.Fields("guestflag") and 512)=0%>
						<img alt="" src="image/icon_wait2pass.gif" class="imgicon" />(留言待审核...)
					</td>
				</tr>
			</table>
	    </td>
	</tr>
	<%
end sub
'===================
sub innerword(byref t_rs)
	%>
	<tr>
		<td class="embedbox">
			<div class="embedbox">
				<table class="embedbox" cellpadding="0" cellspacing="0"><tr><td>
					<span style="color:<%=TableTitleColor%>;">
						<%=t_rs.Fields("name") & "： (" & t_rs.Fields("logdate") & ") "%>
	  					<%
	  					if (iswhisper=false and clng(guestflag and 256)=0) or (pagename="showword" and needverify) or left(pagename,5)="admin" then
	  						if t_rs("email")<>"" then Response.Write "<a href=""mailto:" &t_rs("email")& """ title=""作者邮箱：" &t_rs("email")& """><img alt="""" src=""image/icon_mail.gif"" class=""imgicon"" /></a> "
	  						if t_rs("qqid")<>"" then Response.Write "<img src=""image/icon_qq.gif"" alt=""作者QQ：" &t_rs("qqid")& """ class=""imgicon""" &js_status& " /> "
	  						if t_rs("msnid")<>"" then Response.Write "<img src=""image/icon_msn.gif"" alt=""作者MSN：" &t_rs("msnid")& """ class=""imgicon""" &js_status& " /> "
	  						if t_rs("homepage")<>"" then Response.Write "<a href=""" &t_rs("homepage")& """ target=""_blank"" title=""作者主页：" &t_rs("homepage")& """><img alt="""" src=""image/icon_homepage.gif"" class=""imgicon"" /></a> "
		  					
		  					if left(pagename,5)="admin" then
								if AdminShowIP>0 then Response.Write "<img src=""image/icon_ip.gif"" alt=""IP：" &getip(t_rs("ipaddr"),cint(AdminShowIP))& """ class=""ipicon""" &js_status& " /> "
								if AdminShowOriginalIP>0 and t_rs("originalip")<>"" then Response.Write "<img src=""image/icon_ip2.gif"" alt=""原始IP：" &getip(t_rs("originalip"),cint(AdminShowOriginalIP))& """ class=""ipicon""" &js_status& " /> "
		  					else
		  						if ShowIP>0 then Response.Write "<img src=""image/icon_ip.gif"" alt=""IP：" &getip(t_rs("ipaddr"),cint(ShowIP))& """ class=""ipicon""" &js_status& " /> "
		  					end if
	  					end if
	  					%>
						<br/><%if iswhisper=true and pagename<>"showword" and left(pagename,5)<>"admin" then response.write "(给版主的悄悄话...)" else response.write t_rs("title")%>
					</span>
					<hr size="1" />
					
					<%
					if StatusShowHead and t_rs.Fields("faceid")>=1 and t_rs.Fields("faceid")<=FaceCount then Response.Write "<img alt="""" src=""" & FacePath & t_rs.Fields("faceid") & ".gif"" class=""innerface""/>"
	  				if iswhisper=true and pagename<>"showword" and left(pagename,5)<>"admin" then
	  					response.Write "<img alt="""" src=""image/icon_whisper.gif"" class=""imgicon"" />(给版主的悄悄话...)"
	  				elseif ishidden=true and left(pagename,5)<>"admin" then
	  					response.Write "<img alt="""" src=""image/icon_hide.gif"" class=""imgicon"" />(留言被管理员隐藏...)"
	  				else
	  					guest_txt="" & t_rs("article") & ""
	  					if left(pagename,5)="admin" and AdminViewCode=true then
	  						guest_txt=server.htmlEncode(guest_txt)
	  					else
	  						convertstr guest_txt,guestflag,false
	  					end if
	  					
	  					if guest_txt<>"" then
	  						Response.Write guest_txt
	  					else
	  						Response.Write "(无内容)"
	  					end if
	  				end if
					%>
				</td></tr></table>								
			</div>
		</td>
	</tr>
	<%
end sub
'===================
sub outerword(byref t_rs)
	%>
	<tr>
		<td class="tableleft">
	  		<%=t_rs("name")%><br/>
	  		<%fid=t_rs("faceid")
	  		if isnumeric(fid) and StatusShowHead=true then
	  			if fid>=1 and fid<=FaceCount then
	  				%>
	  				<img alt="" src="<%=FacePath & cstr(fid)& ".gif"%>" class="face" />
	  				<%
	  			end if
	  		end if

	  		if (iswhisper=false and clng(guestflag and 256)=0) or (pagename="showword" and needverify) or left(pagename,5)="admin" then
	  			if t_rs("email")<>"" or t_rs("qqid")<>"" or t_rs("msnid")<>"" or t_rs("homepage")<>"" or (left(pagename,5)<>"admin" and ShowIP>0) or (left(pagename,5)="admin" and AdminShowIP>0) or (left(pagename,5)="admin" and AdminShowOriginalIP>0 and t_rs("originalip")<>"") then
	  				Response.Write "<br/>"
	  			end if
	  			if t_rs("email")<>"" then Response.Write "<a href=""mailto:" &t_rs("email")& """ title=""作者邮箱：" &t_rs("email")& """><img alt="""" src=""image/icon_mail.gif"" class=""imgicon"" /></a> "
	  			if t_rs("qqid")<>"" then Response.Write "<img src=""image/icon_qq.gif"" alt=""作者QQ：" &t_rs("qqid")& """ class=""imgicon""" &js_status& " /> "
	  			if t_rs("msnid")<>"" then Response.Write "<img src=""image/icon_msn.gif"" alt=""作者MSN：" &t_rs("msnid")& """ class=""imgicon""" &js_status& " /> "
	  			if t_rs("homepage")<>"" then Response.Write "<a href=""" &t_rs("homepage")& """ target=""_blank"" title=""作者主页：" &t_rs("homepage")& """><img alt="""" src=""image/icon_homepage.gif"" class=""imgicon"" /></a> "
		  		
		  		if left(pagename,5)="admin" then
					if AdminShowIP>0 then Response.Write "<img src=""image/icon_ip.gif"" alt=""IP：" &getip(t_rs("ipaddr"),cint(AdminShowIP))& """ class=""ipicon""" &js_status& " /> "
					if AdminShowOriginalIP>0 and t_rs("originalip")<>"" then Response.Write "<img src=""image/icon_ip2.gif"" alt=""原始IP：" &getip(t_rs("originalip"),cint(AdminShowOriginalIP))& """ class=""ipicon""" &js_status& " /> "
		  		else
		  			if ShowIP>0 then Response.Write "<img src=""image/icon_ip.gif"" alt=""IP：" &getip(t_rs("ipaddr"),cint(ShowIP))& """ class=""ipicon""" &js_status& " /> "
		  		end if
	  		end if

	  		Response.Write "<br/>"&t_rs("logdate")
	  		%>
		</td>
		<td class="tableright"<%if ReplyInWord then response.write " style=""padding-bottom:12px;"""%>>
			<table cellpadding="2" cellspacing="0" style="width:100%; border-width:0px; border-collapse:collapse;">
	  			<tr>
	  				<td class="wordstitle"><div style="width:100%; height:21px; line-height:21px; vertical-align:middle; overflow:hidden;"><%if iswhisper=true and pagename<>"showword" and left(pagename,5)<>"admin" then response.write "(给版主的悄悄话...)" else response.write t_rs("title")%></div></td>
	  			</tr>
	  			<tr>
	  				<td class="wordscontent">
	  					<%
	  					if t_rs.Fields("parent_id")<=0 and left(pagename,5)<>"admin" then showguestreplyicons t_rs.Fields("id"),t_rs.Fields("parent_id"),StatusWrite and StatusGuestReply and clng(t_rs.Fields("guestflag") and 512)=0
	  					
	  					if iswhisper=true and pagename<>"showword" and left(pagename,5)<>"admin" then
	  						response.Write "<img alt="""" src=""image/icon_whisper.gif"" class=""imgicon"" />(给版主的悄悄话...)"
	  					elseif ishidden=true and left(pagename,5)<>"admin" then
	  						response.Write "<img alt="""" src=""image/icon_hide.gif"" class=""imgicon"" />(留言被管理员隐藏...)"
	  					else
	  						guest_txt="" & t_rs("article") & ""
	  						if left(pagename,5)="admin" and AdminViewCode=true then
	  							guest_txt=server.htmlEncode(guest_txt)
	  						else
	  							convertstr guest_txt,guestflag,false
	  						end if
	  						
	  						if guest_txt<>"" then
	  							Response.Write guest_txt
	  						else
	  							Response.Write "(无内容)"
	  						end if
	  					end if
	  					%>
	  				</td>
	  			</tr>
				<%if clng(t_rs.Fields("replied") AND 1)<>0 and ReplyInWord=true then innerreply(t_rs)	'内嵌%>
				<%if (pagename="admin" or pagename="admin_search" or pagename="admin_showword") and ReplyInWord=true then%>
					<%set p_rs=t_rs%>
					<!-- #include file="admincontrols.inc" -->
				<%end if%>
				
				<%
				if t_rs.Fields("parent_id")<=0 and clng(t_rs.Fields("replied") AND 2)<>0 and (encrypted=false or pagename="showword" or left(pagename,5)="admin") and ReplyInWord=true then
					dim hidden_condition
					hidden_condition=""
					if left(pagename,5)<>"admin" then hidden_condition=GetHiddenWordCondition()
				
					set rs1=Server.CreateObject("ADODB.Recordset")
					rs1.Open Replace(Replace(sql_common2_guestreply,"{0}",t_rs.Fields("id")),"{1}",hidden_condition),cn,0,1,1
					while not rs1.eof
						guestflag=rs1("guestflag")
						if clng(guestflag and 40)=8 then ishidden=true else ishidden=false
						if clng(guestflag and 32)<>0 then iswhisper=true else iswhisper=false
						if clng(guestflag and 96)=96 then encrypted=true else encrypted=false
					
						if clng(guestflag and 16)<>0 and left(pagename,5)<>"admin" then		'待审核
							inneraudit()
						else
							innerword(rs1)
							if rs1.Fields("replied") then innerreply(rs1)
						end if
						
						if pagename="admin" or pagename="admin_search" or pagename="admin_showword" then
							set p_rs=rs1
							%><!-- #include file="admincontrols.inc" --><%
						end if
						rs1.movenext
					wend
					rs1.close
					set rs1=nothing
				end if
				%>
			</table>
		</td>
	</tr>
	<%
end sub
'===================
function getpuretext(byref instr)
dim re,outstr

set re=new Regexp
re.ignoreCase=true
re.global=true

re.pattern="<[^>]*>"
outstr=re.replace(instr,"")

re.pattern="\[[^\[]*\]"
outstr=re.replace(outstr,"")

outstr=replace(outstr,"&lt;","<")
outstr=replace(outstr,"&gt;",">")
'outstr=replace(outstr,"&nbsp;"," ")
outstr=replace(outstr,"&quot;","""")
outstr=replace(outstr,"&apos;","'")
outstr=replace(outstr,"&#39;","'")

getpuretext=outstr
end function
'===================
function geturlpath()
dim host,url
host="http://"
if Request.ServerVariables("SERVER_NAME")<>"" then
	host=host & Request.ServerVariables("SERVER_NAME")
else
	host=host & Request.ServerVariables("HTTP_HOST")
end if
if Request.ServerVariables("PATH_INFO")<>"" then
	host=host & Request.ServerVariables("PATH_INFO")
else
	host=host & Request.ServerVariables("URL")
end if
url=left(host,InStrRev(host,"/"))

geturlpath=url
end function
'===================
function isemail(byref mailstr)
if mailstr<>"" then
	dim pos1,pos2
	pos1=instr(mailstr,"@")
	pos2=instrrev(mailstr,".")

	if pos1>1 and pos2>pos1+1 and len(cstr(mailstr))>pos2 and mid(mailstr,pos1+1,1)<>"." then
		isemail=true
	else
		isemail=false
	end if
else
	isemail=false
end if
end function
'===================
sub newinform()
	if isemail(MailReceive) and MailSmtpServer<>"" then
		dim status_str,ertn,subject_str,body_str
		status_str=""
		ertn=chr(13) & chr(10)
		if clng(guestflag and 16)<>0 then status_str = status_str & "待审核 "
		if clng(guestflag and 32)<>0 then status_str = status_str & "悄悄话"
		if clng(guestflag and 32)<>0 and clng(guestflag and 64)<>0 then status_str = status_str & "(已加密)"
		if status_str="" then status_str="普通"

		subject_str=ltrim(HomeName & " 留言本 新留言通知")
		body_str=HomeName & " 留言本 已添加一条新留言" & ertn & _
				"时间：" & cstr(logdate1) & ertn & _
				"姓名：" & getpuretext(name1) & ertn & _
				"标题：" & getpuretext(title1) & ertn & _
				"内容：" & ertn & getpuretext(content1) & ertn & ertn & _
				"状态：" & status_str & ertn & _
				"打开留言本：" & geturlpath & "index.asp"
		
		if MailComponent="jmail" then
			sendbyjmail MailReceive,subject_str,body_str
		elseif MailComponent="cdo" then
			sendbycdo MailReceive,subject_str,body_str
		end if
	end if
end sub
'===================
sub replyinform()
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	CreateConn cn,dbtype
	rs.open sql_common2_replyinform & request.form("mainid"),cn,0,1,1
	if not rs.eof then
		if isemail(rs.Fields("email")) and clng(rs.Fields("guestflag") and 128)<>0 and MailSmtpServer<>"" then
				
			dim ertn,id,subject_str,body_str
			ertn=chr(13) & chr(10)
			if rs.Fields("parent_id")>0 then
				id=cstr(rs.Fields("parent_id"))
			else
				id=request.form("mainid")
			end if
			
			subject_str=ltrim(HomeName & " 留言本 回复通知")
			body_str=HomeName & " 留言本：您的留言已回复" & ertn & _
					"内容(" & cstr(rs.Fields("logdate")) & ")：" & ertn & getpuretext(rs.Fields("article")) & ertn & ertn & _
					"回复(" & cstr(replydate1) & ")：" & ertn & getpuretext(Request.Form("rcontent")) & ertn & ertn & _
					"查看留言：" & geturlpath & "showword.asp?id=" & id

			if MailComponent="jmail" then
				sendbyjmail rs.Fields("email"),subject_str,body_str
			elseif MailComponent="cdo" then
				sendbycdo rs.Fields("email"),subject_str,body_str
			end if
		end if
	end if

	rs.close : cn.close : set rs=nothing : set cn=nothing
end sub
'===================
sub sendbyjmail(byref receiver,byref subject_str,byref body_str)
	on error resume next
	set jmail = Server.CreateObject("JMAIL.Message")
		
	if err.number=0 then
		err.clear
		on error goto 0
					
		jmail.silent = true
		jmail.logging = false
		jmail.Charset = "GB2312"
		jmail.ContentType = "text/plain"
		jmail.From = MailFrom
		jmail.MailServerUserName = MailUserId
		jmail.MailServerPassword = MailUserPass
		jmail.Priority = MailLevel			'邮件的紧急程度，1 为最快，5 为最慢， 3 为默认值
		jmail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")    

		jmail.AddRecipient receiver
		jmail.Subject = subject_str
		jmail.Body = body_str
					
		jmail.Send(MailSmtpServer)
		jmail.Close()
		set jmail=nothing
	end if
end sub
'===================
sub sendbycdo(byref receiver,byref subject_str,byref body_str)
	on error resume next
	set cdocon=server.CreateObject("CDO.Configuration")
	set cdomail=server.CreateObject("CDO.Message")
				
	if err.number=0 then
		err.Clear
		on error goto 0
					
		cdocon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")=2		'a remote smtp server
		cdocon.Fields("http://schemas.microsoft.com/cdo/configuration/languagecode")=2052	'0x804
		cdocon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")=MailSmtpServer
		cdocon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=1	'cdoBasic
		cdocon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername")=MailUserId
		cdocon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword")=MailUserPass
		'cdocon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
		if MailLevel<3 then
			cdocon.Fields("urn:schemas:httpmail:priority")=1
		elseif MailLevel=3 then
			cdocon.Fields("urn:schemas:httpmail:priority")=0
		else
			cdocon.Fields("urn:schemas:httpmail:priority")=-1
		end if
		cdocon.Fields.Update
					
		set cdomail.Configuration=cdocon
		cdomail.From=MailFrom
		cdomail.To=receiver
		cdomail.Subject=subject_str
		cdomail.TextBody=body_str
		cdomail.Send
		
		set cdomail=nothing
		set cdocon=nothing
	end if
end sub
%>