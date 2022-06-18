<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/write.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/mail.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="tips.asp" -->
<!-- #include file="error.asp" -->
<%
'======================================================
sub wordsbaned
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	if StatusStatistics then call addstat("banned")
	Call ErrorPage(4)
	Response.End
end sub
'======================================================
sub floodbaned
	rs.Close : rs2.Close : cn.Close : set rs=nothing : set rs2=nothing : set cn=nothing
	if StatusStatistics then call addstat("banned")
	Call ErrorPage(7)
	Response.End
end sub
'======================================================
sub bancheck(byref re,byref field,byval tfiltermode,byval bitflag)
	if CBool(tfiltermode and bitflag) then
		if CBool(tfiltermode AND 256) then  'pure text
			Dim tcompareflag
			if re.IgnoreCase then tcompareflag=1 else tcompareflag=0 end if
			if Instr(1, field, re.Pattern, tcompareflag)>0 then Call wordsbaned
		else
			if re.Test(field) then Call wordsbaned
		end if
	end if
end sub
'======================================================
sub filtercheck(byref re,byref field,byval tfiltermode,byval bitflag,byref treplacestr,byref filtered)
	if CBool(tfiltermode and bitflag) then
		Dim tmatchcase
		tmatchcase=CBool(tfiltermode AND 8192)

		if CBool(tfiltermode AND 256) then  'pure text
			Dim tcompareflag
			if re.IgnoreCase then tcompareflag=1 else tcompareflag=0 end if
			if Instr(1, field, re.Pattern, tcompareflag)>0 then    'matched
				if CBool(tfiltermode and 4096) then     'wait to audit
					guestflag=guestflag OR 16
				else
					filtered=true
					field=Replace(field, re.Pattern, treplacestr, 1, -1, tcompareflag)
				end if
			end if
		else
			if re.Test(field) then  'matched
				if CBool(tfiltermode and 4096) then     'wait to audit
					guestflag=guestflag OR 16
				else
					filtered=true
					field=re.Replace(field,treplacestr)
				end if
			end if
		end if
	end if
end sub
'======================================================
Function getLeaveWordUrl
	Dim qstr
	qstr=Request.Form("qstr")
	if qstr<>"" then qstr="?" & qstr
	getLeaveWordUrl = "leaveword.asp" & qstr
End Function
'======================================================

Response.Expires=-1
if checkIsBannedIP() then
	Call ErrorPage(1)
	Response.End
elseif Not StatusOpen then
	Call ErrorPage(2)
	Response.End
elseif Not StatusWrite then
	Call ErrorPage(3)
	Response.End
elseif flood_minwait>0 and isdate(Session(InstanceName & "_wrote_time")) then
	if datediff("s",Session(InstanceName & "_wrote_time"),now())<=flood_minwait then
		if StatusStatistics then call addstat("banned")
		Call ErrorPage(6)
		Response.End
	end if
elseif flood_minwait>0 and isdate(Request.Cookies("wrote_time")) then
	if datediff("s",Request.Cookies("wrote_time"),now())<=flood_minwait then
		if StatusStatistics then call addstat("banned")
		Call ErrorPage(6)
		Response.End
	end if
end if
if Request.Form("iname")="" or Request.Form("ititle")="" then Response.Redirect("index.asp")

Session(InstanceName & "_ititle")=Request.Form("ititle")
Session(InstanceName & "_icontent")=Request.Form("icontent")

if WriteVcodeCount>0 and (Request.Form("ivcode")<>Session(InstanceName & "_vcode_write") or Session(InstanceName & "_vcode_write")="") then
	Session(InstanceName & "_vcode_write")=""
	Call TipsPage("验证码错误。", getLeaveWordUrl)
	Response.End
else
	Session(InstanceName & "_vcode_write")=""
end if
'===================================================================

SetTimelessCookie "iname",Request.form("iname")
SetTimelessCookie "imail",Request.form("imail")
SetTimelessCookie "iqq",Request.form("iqq")
SetTimelessCookie "imsn",Request.form("imsn")
SetTimelessCookie "ihomepage",Request.form("ihomepage")
SetTimelessCookie "ihead",Request.form("ihead")

'===================================================================

Dim cn,rs,sql

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
set rs2=server.CreateObject("ADODB.Recordset")

dim content1,name1,title1,email1,qqid1,msnid1,homepage1,ipv4addr1,ipv6addr1,originalipv41,originalipv61,head1,guestflag,whisperpwd
dim tmpAddr
'---------------------
name1=Request.form("iname")
title1=request.form("ititle")
email1=request.form("imail")
qqid1=request.form("iqq")
msnid1=request.form("imsn")
homepage1=request.form("ihomepage")

tmpAddr=CStr(request.ServerVariables("REMOTE_ADDR"))
if IsIPv4(tmpAddr) then
	ipv4addr1=Left(tmpAddr,15)
elseif IsIPv6(tmpAddr) then
	ipv6addr1=expandIPv6(tmpAddr,false)
end if
tmpAddr=CStr(request.ServerVariables("HTTP_X_FORWARDED_FOR"))
if IsIPv4(tmpAddr) then
	originalipv41=Left(tmpAddr,15)
elseif IsIPv6(tmpAddr) then
	originalipv61=expandIPv6(tmpAddr,false)
end if
face1=request.form("ihead")

content1=request.form("icontent")
if WordsLimit<>0 and len(content1)>WordsLimit then content1=left(content1,WordsLimit)

guestflag=guestlimit
whisperpwd=""
if StatusNeedAudit then guestflag=guestflag OR 16
if StatusWhisper and Request.Form("chk_whisper")="1" then guestflag=guestflag OR 32
if StatusEncryptWhisper and Request.Form("chk_encryptwhisper")="1" and Request.Form("iwhisperpwd")<>"" then
	guestflag=guestflag OR 64
	whisperpwd=md5(Request.Form("iwhisperpwd"),32)
end if
if Request.Form("imailreplyinform")="1" then guestflag=guestflag OR 128
if Request.Form("hidecontact")="1" then guestflag=guestflag OR 256
'-------------------------
Dim tregexp,tfiltermode,treplacestr,filtered
set re=new RegExp
filtered=false

Call CreateConn(cn)
rs.Open sql_write_filter,cn,0,1,1
while Not rs.EOF
	tregexp=rs("regexp")
	tfiltermode=rs("filtermode")
	treplacestr=rs("replacestr")

	if tregexp<>"" then
		re.Global=true
		re.Pattern=tregexp
		re.IgnoreCase=Not CBool(tfiltermode AND 8192)
		if CBool(tfiltermode AND 512) then  'wildcard
			re.Multiline=false
			re.Pattern=Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(re.Pattern,"(","\("),")","\)"),"[","\["),"]","\]"),"{","\{"),"}","\}"),"<","\<"),">","\>"),"^","\^"),"$","\$"),"+","\+"),"\","\\"),".","\."),"?","."),"*",".*?")
		elseif CBool(tfiltermode AND 1024) then     'regexp
			re.Multiline=false
		elseif CBool(tfiltermode AND 2048) then     'regexp multiline
			re.Multiline=true
		end if

		if clng(tfiltermode and 16384)<>0 then    'baned
			bancheck re,name1,tfiltermode,1
			bancheck re,email1,tfiltermode,2
			bancheck re,qqid1,tfiltermode,4
			bancheck re,msnid1,tfiltermode,8
			bancheck re,homepage1,tfiltermode,16
			bancheck re,title1,tfiltermode,32
			bancheck re,content1,tfiltermode,64
		else
			filtercheck re,name1,tfiltermode,1,treplacestr,filtered
			filtercheck re,email1,tfiltermode,2,treplacestr,filtered
			filtercheck re,qqid1,tfiltermode,4,treplacestr,filtered
			filtercheck re,msnid1,tfiltermode,8,treplacestr,filtered
			filtercheck re,homepage1,tfiltermode,16,treplacestr,filtered
			filtercheck re,title1,tfiltermode,32,treplacestr,filtered
			filtercheck re,content1,tfiltermode,64,treplacestr,filtered
		end if
	end if
	rs.MoveNext
wend
rs.Close
set re=nothing
if filtered and StatusStatistics then addstat("filtered")
'-------------------------
if name1="" or title1="" then Response.Redirect "index.asp"

name1=HtmlEncode(name1)
name1=textfilter(name1,false)
name1=left(name1,20)

title1=HtmlEncode(title1)
title1=textfilter(title1,false)
title1=left(title1,30)

email1=replace(HtmlEncode(email1)," ","")
email1=textfilter(email1,true)
email1=left(email1,50)

qqid1=replace(HtmlEncode(qqid1)," ","")
qqid1=textfilter(qqid1,false)
qqid1=left(qqid1,16)

msnid1=replace(HtmlEncode(msnid1)," ","")
msnid1=textfilter(msnid1,false)
msnid1=left(msnid1,50)

content1=replace(content1,"<%","< %")

logdate1=ServerTimeToUTC(now())

homepage1=Trim(homepage1)
if homepage1<>"" then
	if InStr(homepage1,"://")=0 then homepage1="http://" & homepage1
	homepage1=textfilter(homepage1,true)
	homepage1=left(homepage1,127)
end if

if isnumeric(face1)=false or face1="" then
	face1=0
elseif len(cstr(face1))>len(cstr(FaceCount))then
	face1=0
elseif clng(face1)>clng(FaceCount) then
	face1=0
else
	face1=abs(cbyte(face1))
end if

'flood check
if flood_searchrange>0 and (flood_sfnewword or flood_sfnewreply) and (flood_include or flood_equal) and (flood_sititle or flood_sicontent) then
	if isnumeric(Request.Form("follow")) and Request.Form("follow")<>"" and StatusGuestReply then
		rs2.Open sql_write_flood_ids & Request.Form("follow"),cn,,,1
	else
		rs2.Open sql_write_flood_idnull,cn,,,1
	end if

	if flood_sfnewword and rs2.EOF then
		sql=sql_write_flood_head

		if flood_sititle then
			if flood_include and title1<>"" then sql=sql & Replace(sql_write_flood_titlelike,"{0}",FilterSqlStr(title1))
			if flood_equal or title1="" then sql=sql & Replace(sql_write_flood_titleequal,"{0}",FilterSqlStr(title1))
			if flood_sicontent then sql=sql & " OR"
		end if

		if flood_sicontent then
			if flood_include and content1<>"" then sql=sql & Replace(sql_write_flood_articlelike,"{0}",FilterSqlStr(content1))
			if flood_equal or content1="" then sql=sql & Replace(sql_write_flood_articleequal,"{0}",FilterSqlStr(content1))
		end if
		
		sql=sql & Replace(sql_write_flood_wordstail,"{0}",flood_searchrange)
		rs.Open sql,cn,,,1
		if not rs.EOF then floodbaned()
		rs.Close
	end if

	if flood_sfnewreply and (not rs2.EOF) then
		sql=sql_write_flood_head

		if flood_sititle and title1<>"Re:" then
			if flood_include and title1<>"" then sql=sql & Replace(sql_write_flood_titlelike,"{0}",FilterSqlStr(title1))
			if flood_equal or title1="" then sql=sql & Replace(sql_write_flood_titleequal,"{0}",FilterSqlStr(title1))
			if flood_sicontent then sql=sql & " OR"
		end if

		if flood_sicontent then
			if flood_include and content1<>"" then sql=sql & Replace(sql_write_flood_articlelike,"{0}",FilterSqlStr(content1))
			if flood_equal or content1="" then sql=sql & Replace(sql_write_flood_articleequal,"{0}",FilterSqlStr(content1))
		end if
		
		sql=sql & Replace(Replace(sql_write_flood_replytail,"{0}",flood_searchrange),"{1}",rs2.Fields("id"))
		rs.Open sql,cn,,,1
		if not rs.EOF then floodbaned()
		rs.Close
	end if
	
	rs2.Close
end if

'------------------------
rs.Open sql_write_idnull,cn,0,3,1
rs.AddNew
rs("name")=name1
rs("title")=title1
rs("email")=email1
rs("qqid")=qqid1
rs("msnid")=msnid1
rs("homepage")=homepage1
rs("logdate")=logdate1
rs("lastupdated")=logdate1
rs("ipv4addr")=ipv4addr1
rs("ipv6addr")=ipv6addr1
rs("originalipv4")=originalipv41
rs("originalipv6")=originalipv61
rs("faceid")=face1
rs("guestflag")=guestflag
rs("whisperpwd")=whisperpwd
rs("article")=content1
rs("parent_id")=0
if isnumeric(Request.Form("follow")) and Request.Form("follow")<>"" and StatusGuestReply then
	rs2.Open sql_write_verify_repliable & Request.Form("follow"),cn,0,1,1
	if not rs2.EOF then
		rs("parent_id")=Request.Form("follow")
		cn.Execute Replace(Replace(sql_write_updatelastupdated,"{0}",logdate1),"{1}",Request.Form("follow")),,129
		cn.Execute sql_write_updateparentflag & Request.Form("follow"),,129
	end if
	rs2.Close : set rs2=nothing
end if

if IsAccess then	'update root_id info
	if rs("parent_id")<=0 then
		rs("root_id")=rs("id")
	else
		rs("root_id")=rs("parent_id")
	end if
end if
rs.Update

rs.Close : cn.Close : set rs=nothing : set cn=nothing

if StatusStatistics then call addstat("written")

Dim NowTime
NowTime=Now
SetTimelessCookie "wrote_time",NowTime
Session(InstanceName & "_wrote_time")=NowTime
Session(InstanceName & "_ititle")=""
Session(InstanceName & "_icontent")=""
Session(InstanceName & "_guestflag")=guestflag

if MailNewInform then newinform()

if Request.Form("return")="showword" and isnumeric(Request.Form("follow")) and Request.Form("follow")<>"" then
	Response.Redirect "showword.asp?id=" & Request.Form("follow")
else
	Response.Redirect "index.asp"
end if
%>