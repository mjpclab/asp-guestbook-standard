<!-- #include file="config.asp" -->
<!-- #include file="md5.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if checkIsBannedIP then
	Response.Redirect "err.asp?number=1"
	Response.End
elseif StatusOpen=false then
	Response.Redirect "err.asp?number=2"
	Response.End
end if
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> 留言本 浏览留言</title>
	<!-- #include file="inc_stylesheet.asp" -->

	<script type="text/javascript">
	function setFocus()
	{
		if(document && document.form5 && document.form5.ipass)document.form5.ipass.focus();
	}
	</script>
</head>

<body<%=bodylimit%> onload="<%=framecheck%>setFocus();">

<%

'<ShowArticle>
Dim cn,rs,ItemsCount
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

CreateConn cn,dbtype
rs.Open sql_showword_count,cn,0,1,1
ItemsCount=rs(0)
rs.Close
%>

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"浏览留言"%>

	<!-- #include file="func_guest.inc" -->
	<!-- #include file="topbulletin.inc" -->
	<%if StatusSearch and ShowTopSearchBox then%><!-- #include file="searchbox_guest.inc" --><%end if%>

	<%
	dim showbox,showstr,needverify,cantverify,idexists
	showbox=false
	showstr=""
	needverify=false
	cantverify=false
	idexists=false
	if isnumeric(Request("id")) and Request("id")<>"" then
		rs.Open sql_showword & Request("id"),cn,,,1
		if not rs.EOF then
			idexists=true
			if clng(rs.Fields("guestflag") and 64)<>0 then
				needverify=true
			elseif clng(rs.Fields("guestflag") and 32)<>0 then
				cantverify=true
			end if
		end if
	end if
	
	if Request.Form("ispostback")="1" then
		if VcodeCount>0 and (Request.Form("ivcode")<>session("vcode") or session("vcode")="") then
			showbox=true
			showstr="验证码错误。"
			session("vcode")=""
			session("id")=""
			session("guest_pwd")=""
		else
			session("vcode")=""
			session("id")=cstr(Request.Form("id"))
			session("guest_pwd")=md5(Request.Form("ipass"),32)
			Response.Redirect "showword.asp?id=" & Request.Form("id")
		end if
	end if

	if not showbox then
		if not idexists then
			showbox=true
			showstr="留言不存在。"
		elseif needverify then
			if session("id")<>cstr(rs("id")) then
				showbox=true
			elseif session("guest_pwd")<>rs("whisperpwd") then
				showbox=true
				showstr="密码不正确。"
			end if
		end if
	end if

	if VcodeCount>0 and showbox=true then
		session("vcode")=getvcode(VcodeCount)
	else
		session("vcode")=""
	end if
	%>

	<%if showbox=true then%>
	<%if rs.State=1 then rs.Close%>
		<div class="region form-region">
			<h3 class="title">验证已加密留言</h3>
			<div class="content">
				<form method="post" action="showword.asp" name="form5" onsubmit="if(this.ipass.value==''){alert('请输入密码。');this.ipass.focus();return false;} else if(this.ivcode && this.ivcode.value==''){alert('请输入验证码。');this.ivcode.focus(); return false;} else this.submit1.disabled=true;">
				<input type="hidden" name="ispostback" value="1" />
				<input type="hidden" name="id" value="<%=request("id")%>" />

				<div class="field">
					<span class="label">密码：</span>
					<span class="value"><input type="password" name="ipass" maxlength="16" autofocus="autofocus" /></span>
				</div>

				<%if VcodeCount>0 then%>
				<div class="field">
					<span class="label">验证码：</span>
					<span class="value"><input type="text" name="ivcode" autocomplete="off"><img class="captcha" src="show_vcode.asp" /></span>
				</div>
				<%end if%>

				<div class="command">
					<input value="确定" type="submit" name="submit1" />
				</div>

				</form>
			</div>
		</div>
		<%if showstr<>"" then Response.Write "<br/><p style=""text-align:center"">" &showstr& "</p>"%>
	<%else
		dim ItemsPerPage,pagename
		ItemsPerPage=1
		if cantverify then
			pagename="showword_cantverify"
		else
			pagename="showword"
		end if%>
		<!-- #include file="listword_guest.inc" -->
		<%rs.Close%>
	<%end if%>
	<%cn.Close : set rs=nothing : set cn=nothing%>

	<!-- #include file="func_guest.inc" -->

	<%if StatusSearch and ShowBottomSearchBox then%><!-- #include file="searchbox_guest.inc" --><%end if%>
</div>

<!-- #include file="bottom.asp" -->
<%if StatusStatistics and session("gotclientinfo")<>true then%><script type="text/javascript" src="getclientinfo.asp" defer="defer" async="async"></script><%end if%>
</body>
</html>