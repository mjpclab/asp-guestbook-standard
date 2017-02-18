<%
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

sub newinform()
	if isemail(MailReceive) and MailSmtpServer<>"" then
		dim status_str,book_prefix,subject_str,body_str
		status_str=""
		if clng(guestflag and 16)<>0 then status_str = status_str & "待审核 "
		if clng(guestflag and 32)<>0 then status_str = status_str & "悄悄话"
		if clng(guestflag and 32)<>0 and clng(guestflag and 64)<>0 then status_str = status_str & "(已加密)"
		if status_str="" then status_str="普通"

		book_prefix = ltrim(HomeName & " 留言本")
		subject_str = book_prefix & " 新留言通知"
		body_str=book_prefix & " 已添加一条新留言" & vbCrLf & _
				"时间：" & cstr(logdate1) & vbCrLf & _
				"姓名：" & getpuretext(name1) & vbCrLf & _
				"标题：" & getpuretext(title1) & vbCrLf & _
				"内容：" & vbCrLf & getpuretext(content1) & vbCrLf & vbCrLf & _
				"状态：" & status_str & vbCrLf & _
				"打开留言本：" & geturlpath & "index.asp"
		
		if MailComponent="jmail" then
			sendbyjmail MailReceive,subject_str,body_str
		elseif MailComponent="cdo" then
			sendbycdo MailReceive,subject_str,body_str
		end if
	end if
end sub

sub replyinform()
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	Call CreateConn(cn)
	rs.open sql_common2_replyinform & request.form("mainid"),cn,0,1,1
	if not rs.eof then
		if isemail(rs.Fields("email")) and clng(rs.Fields("guestflag") and 128)<>0 and MailSmtpServer<>"" then

			dim id,book_prefix,subject_str,body_str
			if rs.Fields("parent_id")>0 then
				id=cstr(rs.Fields("parent_id"))
			else
				id=request.form("mainid")
			end if

			book_prefix = ltrim(HomeName & " 留言本")
			subject_str=book_prefix & " 回复通知"
			body_str=book_prefix & " 您的留言已回复" & vbCrLf & _
					"内容(" & cstr(rs.Fields("logdate")) & ")：" & vbCrLf & getpuretext(rs.Fields("article")) & vbCrLf & vbCrLf & _
					"回复(" & cstr(replydate1) & ")：" & vbCrLf & getpuretext(Request.Form("rcontent")) & vbCrLf & vbCrLf & _
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
%>