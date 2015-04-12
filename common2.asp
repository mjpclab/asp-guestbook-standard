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
sub showadminicons()%>
	<%if admin_email<>"" then%><a class="icon" href="mailto:<%=admin_email%>" title="版主邮箱：<%=admin_email%>"><img src="image/icon_mail.gif"/></a><%end if%>
	<%if admin_qqid<>"" then%><span class="icon" title="版主QQ：<%=admin_qqid%>"><img src="image/icon_qq.gif"/></span><%end if%>
	<%if admin_msnid<>"" then%><span class="icon"><img src="image/icon_skype.gif" alt="版主Skype：<%=admin_msnid%>"/></span><%end if%>
	<%if admin_homepage<>"" then%><a class="icon" href="<%=admin_homepage%>" target="_blank" title="版主主页：<%=admin_homepage%>"><img src="image/icon_homepage.gif"/></a><%end if%>
<%end sub
'===================
sub showguestreplyicons(byval follow_id,byval parent_id,byval show_reply)
dim url

url="leaveword.asp?follow=" & follow_id
if left(pagename,8)="showword" then url=url & "&return=showword"
url=Server.HTMLEncode(url)%>
<div class="guest-tools">
	<%if parent_id<0 then%><span class="tool"><img src="image/icon_toplocked.gif"/>(置顶)</span><%end if%>
	<%if show_reply then%><span class="tool"><a href="<%=url%>"><img src="image/icon_reply.gif"/>[回复]</a></span><%end if%>
</div>
<%end sub
'===================
sub showguestinfoicons(t_rs)%>
	<%if t_rs("email")<>"" then%><a class="icon" href="mailto:<%=t_rs("email")%>" title="作者邮箱：<%=t_rs("email")%>"><img src="image/icon_mail.gif"/></a><%end if%>
	<%if t_rs("qqid")<>"" then%><span class="icon" title="作者QQ：<%=t_rs("qqid")%>"><img src="image/icon_qq.gif"/></span><%end if%>
	<%if t_rs("msnid")<>"" then%><span class="icon" title="作者Skype：<%=t_rs("msnid")%>"><img src="image/icon_skype.gif"/></span><%end if%>
	<%if t_rs("homepage")<>"" then%><a class="icon" href="<%=t_rs("homepage")%>" target="_blank" title="作者主页：<%=t_rs("homepage")%>"><img src="image/icon_homepage.gif" /></a><%end if%>

	<%if left(pagename,5)="admin" then%>
		<%if AdminShowIPv4>0 and Len(t_rs.Fields("ipv4addr"))>0 then%><span class="icon" title="IP：<%=GetIPv4(t_rs.Fields("ipv4addr"),AdminShowIPv4)%>"><img src="image/icon_ip.gif"/></span><%end if%>
		<%if AdminShowIPv6>0 and Len(t_rs.Fields("ipv6addr"))>0 then%><span class="icon" title="IP：<%=GetIPv6(t_rs.Fields("ipv6addr"),AdminShowIPv6)%>"><img src="image/icon_ip.gif"/></span><%end if%>
		<%if AdminShowOriginalIPv4>0 and Len(t_rs.Fields("originalipv4"))>0 then%><span class="icon" title="原始IP：<%=GetIPv4(t_rs.Fields("originalipv4"),AdminShowOriginalIPv4)%>"><img src="image/icon_ip2.gif"/></span><%end if%>
		<%if AdminShowOriginalIPv6>0 and Len(t_rs.Fields("originalipv6"))>0 then%><span class="icon" title="原始IP：<%=GetIPv6(t_rs.Fields("originalipv6"),AdminShowOriginalIPv6)%>"><img src="image/icon_ip2.gif"/></span><%end if%>
	<%else%>
		<%if ShowIPv4>0 and Len(t_rs.Fields("ipv4addr"))>0 then%><span class="icon" title="IP：<%=GetIPv4(t_rs.Fields("ipv4addr"),ShowIPv4)%>"><img src="image/icon_ip.gif"/></span><%end if%>
		<%if ShowIPv6>0 and Len(t_rs.Fields("ipv6addr"))>0 then%><span class="icon" title="IP：<%=GetIPv6(t_rs.Fields("ipv6addr"),ShowIPv6)%>"><img src="image/icon_ip.gif"/></span><%end if%>
	<%end if%>
<%end sub
'===================
sub inneradminreply(byref rs2)%>
<div class="message inner-message admin-message">
	<div class="summary">
		<div class="name">版主<%if admin_name<>"" then response.write "(" & admin_name & ")"%>回复：</div>
		<div class="date">(<%=rs2("replydate")%>)</div>
		<div class="icons"><%showadminicons()%></div>
	</div>
	<%if admin_faceurl<>"" then%><img class="face" src="<%=admin_faceurl%>"/><%end if%>
	<div class="words">
		<%if encrypted=true and pagename<>"showword" and left(pagename,5)<>"admin" then%>
			<span class="inner-hint"><img src="image/icon_key.gif"/>(需要预设的密码才能查看...)[<a href="showword.asp?id=<%=rs2("id")%>">点击这里验证...</a>]</span>
		<%else
			reply_htmlflag=rs2("htmlflag")
			reply_txt=rs2("reinfo")

			convertstr reply_txt,reply_htmlflag,true
			Response.Write reply_txt
		end if%>
	</div>
</div>
<%end sub
'===================
sub outeradminreply(byref rs2)%>
<div class="message outer-message admin-message">
	<h2 class="title">[版主回复：]</h2>
	<div class="info">
		<%if admin_faceurl<>"" then%><img class="face" src="<%=admin_faceurl%>"/><%end if%>
		<%if admin_email<>"" or admin_qqid<>"" or admin_msnid<>"" or admin_homepage<>"" then%>
			<div class="icons"><%showadminicons()%></div>
		<%end if%>
		<div class="date"><%=rs2("replydate")%></div>
	</div>
	<div class="detail">
		<div class="words">
			<%if encrypted=true and pagename<>"showword" and left(pagename,5)<>"admin" then%>
				<span class="inner-hint"><img src="image/icon_key.gif"/>(需要预设的密码才能查看...)[<a href="showword.asp?id=<%=rs2("id")%>">点击这里验证...</a>]</span>
			<%else
				reply_htmlflag=rs2("htmlflag")
				reply_txt=rs2("reinfo")

				convertstr reply_txt,reply_htmlflag,true
				Response.Write reply_txt
			end if%>
		</div>
	</div>
</div>
<%end sub
'===================
sub inneraudit()%>
<div class="message inner-message guest-message auditing-message">
	<div class="summary">
		<h2 class="title">(留言待审核...)</h2>
	</div>
	<div class="words">
		<span class="inner-hint"><img src="image/icon_wait2pass.gif"/>(留言待审核...)</span>
	</div>
</div>
<%end sub
'===================
sub outeraudit(t_rs)%>
<div class="message outer-message guest-message auditing-message">
	<div class="info">
		<%fid=t_rs("faceid")
		if isnumeric(fid) and StatusShowHead=true then
			if fid>=1 and fid<=FaceCount then%>
				<img class="face" src="<%=FacePath & cstr(fid)%>.gif" />
			<%end if
		end if%>

		<div class="date"><%=t_rs("logdate")%></div>
	</div>
	<div class="detail">
		<h2 class="title">(留言待审核...)</h2>
		<%if rs.Fields("parent_id")<=0 and left(pagename,5)<>"admin" then showguestreplyicons rs.Fields("id"),rs.Fields("parent_id"),StatusWrite and StatusGuestReply and clng(rs.Fields("guestflag") and 512)=0%>
		<div class="words">
			<span class="inner-hint"><img src="image/icon_wait2pass.gif" />(留言待审核...)</span>
		</div>
	</div>
</div>
<%end sub
'===================
sub innerword(byref t_rs)%>
<div class="message inner-message guest-message">
	<div class="summary">
		<div class="name"><%=t_rs("name")%>：</div>
		<div class="date">(<%=t_rs("logdate")%>)</div>
			<%if (iswhisper=false and clng(guestflag and 256)=0) or (pagename="showword" and needverify) or left(pagename,5)="admin" then%>
				<div class="icons"><%showguestinfoicons(t_rs)%></div>
			<%end if%>
		<h2 class="title"><%if iswhisper=true and pagename<>"showword" and left(pagename,5)<>"admin" then response.write "(给版主的悄悄话...)" else response.write t_rs("title")%></h2>
	</div>
	<%if StatusShowHead and t_rs.Fields("faceid")>=1 and t_rs.Fields("faceid")<=FaceCount then%><img src="<%=FacePath & t_rs.Fields("faceid")%>.gif"/><%end if%>
	<div class="words">
		<%if iswhisper=true and pagename<>"showword" and left(pagename,5)<>"admin" then%>
			<span class="inner-hint"><img src="image/icon_whisper.gif"/>(给版主的悄悄话...)</span>
		<%elseif ishidden=true and left(pagename,5)<>"admin" then%>
			<span class="inner-hint"><img src="image/icon_hide.gif"/>(留言被管理员隐藏...)</span>
		<%else
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
		end if%>
	</div>
</div>
<%end sub
'===================
sub outerword(byref t_rs)%>
<div class="message outer-message guest-message">
	<div class="info">
		<div class="name"><%=t_rs("name")%></div>

		<%fid=t_rs("faceid")
		if isnumeric(fid) and StatusShowHead=true then
			if fid>=1 and fid<=FaceCount then%>
				<img class="face" src="<%=FacePath & cstr(fid)& ".gif"%>" />
			<%end if
		end if%>

		<%if (iswhisper=false and clng(guestflag and 256)=0) or (pagename="showword" and needverify) or left(pagename,5)="admin" then%>
			<div class="icons"><%showguestinfoicons(t_rs)%></div>
		<%end if%>

		<div class="date"><%=t_rs("logdate")%></div>
	</div>
	<div class="detail">
		<h2 class="title"><%if iswhisper=true and pagename<>"showword" and left(pagename,5)<>"admin" then response.write "(给版主的悄悄话...)" else response.write t_rs("title")%></h2>
		<%if t_rs.Fields("parent_id")<=0 and left(pagename,5)<>"admin" then showguestreplyicons t_rs.Fields("id"),t_rs.Fields("parent_id"),StatusWrite and StatusGuestReply and clng(t_rs.Fields("guestflag") and 512)=0%>
		<div class="words">
			<%if iswhisper=true and pagename<>"showword" and left(pagename,5)<>"admin" then%>
				<span class="inner-hint"><img src="image/icon_whisper.gif"/>(给版主的悄悄话...)</span>
			<%elseif ishidden=true and left(pagename,5)<>"admin" then%>
				<span class="inner-hint"><img src="image/icon_hide.gif"/>(留言被管理员隐藏...)</span>
			<%else
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
		</div>

		<%if clng(t_rs.Fields("replied") AND 1)<>0 and ReplyInWord=true then inneradminreply(t_rs)	'内嵌%>

		<%if (pagename="admin" or pagename="admin_search" or pagename="admin_showword") and ReplyInWord=true then%>
			<%set p_rs=t_rs%>
			<!-- #include file="admintools.inc" -->
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
				ishidden=(clng(guestflag and 40)=8)
				iswhisper=(clng(guestflag and 32)<>0)
				encrypted=(clng(guestflag and 96)=96)

				if clng(guestflag and 16)<>0 and left(pagename,5)<>"admin" then		'待审核
					inneraudit()
				else
					innerword(rs1)
					if rs1.Fields("replied") then inneradminreply(rs1)
				end if

				if pagename="admin" or pagename="admin_search" or pagename="admin_showword" then
					set p_rs=rs1
					%><!-- #include file="admintools.inc" --><%
				end if
				rs1.movenext
			wend
			rs1.close
			set rs1=nothing
		end if
		%>
	</div>
</div>
<%end sub
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
dim host,url,buffer,port,httpHost
host="http://"
httpHost=Request.ServerVariables("HTTP_HOST")
if Len(httpHost)>0 then
	host=host & httpHost
else
	buffer=Request.ServerVariables("SERVER_NAME")
	if IsIPv6(buffer) then
		buffer = "[" & buffer & "]"
	end if
	host=host & buffer

	port=Request.ServerVariables("SERVER_PORT")
	if Len(port)>0 and port<>"80" then
		host=host & ":" & port
	end if
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