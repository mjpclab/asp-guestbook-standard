<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/common2.asp" -->
<!-- #include file="include/sql/topbulletin.asp" -->
<!-- #include file="include/sql/showword.asp" -->
<!-- #include file="include/sql/guest_listword.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/ubbcode.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="include/utility/message.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="error.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if checkIsBannedIP() then
	Call ErrorPage(1)
	Response.End
elseif Not StatusOpen then
	Call ErrorPage(2)
	Response.End
end if

Dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)

Dim showWord, showVerify, showMessage, pagename
showWord=false : showVerify=false : showMessage=""

Dim id
id=FilterKeyword(Trim(Request.QueryString("id")))
if id="" Or Not Isnumeric(id) then
	Response.Status="404 Not Found"
	showMessage="留言不存在。"
Else
	rs.Open sql_showword & id,cn,,,1

	if rs.EOF then
		Response.Status="404 Not Found"
		showMessage="留言不存在。"
	elseif Not CBool(rs.Fields("guestflag") and 32) then
		showWord=true
		pagename="showword"
	elseif Not CBool(rs.Fields("guestflag") and 64) then
		showWord=true
		pagename="showword_cannot_verify"
	elseif IsEmpty(Request.Form("ispostback")) And Session(InstanceName & "_id")=id And Session(InstanceName & "_guest_pwd")=rs.Fields("whisperpwd") then
		showWord=true
		pagename="showword"
	elseif IsEmpty(Request.Form("ispostback")) then
		showVerify=true
		if VcodeCount>0 then
			Session(InstanceName & "_vcode")=getvcode(VcodeCount)
		else
			Session(InstanceName & "_vcode")=""
		end if
	elseif VcodeCount>0 and (Request.Form("ivcode")<>Session(InstanceName & "_vcode") or Session(InstanceName & "_vcode")="") then
		showVerify=true
		showMessage="验证码错误。"
		Session(InstanceName & "_vcode")=getvcode(VcodeCount)
	elseif md5(Request.Form("ipass"),32)<>rs.Fields("whisperpwd") then
		showVerify=true
		showMessage="密码错误。"
		if VcodeCount>0 then
			Session(InstanceName & "_vcode")=getvcode(VcodeCount)
		else
			Session(InstanceName & "_vcode")=""
		end if
	else
		showWord=true
		pagename="showword"
		Session(InstanceName & "_vcode")=""
		Session(InstanceName & "_id")=id
		Session(InstanceName & "_guest_pwd")=rs.Fields("whisperpwd")
	end if
End If
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 浏览留言<%if showWord and pagename="showword" then Response.Write " " & rs.Fields("title")%></title>
	<!-- #include file="inc_stylesheet.asp" -->

	<script type="text/javascript">
	function setFocus()
	{
		if(document && document.form5 && document.form5.ipass)document.form5.ipass.focus();
	}
	</script>
</head>

<body<%=bodylimit%> onload="<%=framecheck%>setFocus();">
<div id="outerborder" class="outerborder">

	<%if ShowTitle then%><%Call InitHeaderData("浏览留言")%><!-- #include file="include/template/header.inc" --><%end if%>

	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/guest_func.inc" -->
	<!-- #include file="include/template/topbulletin.inc" -->
	<%if StatusSearch and ShowTopSearchBox then%><!-- #include file="include/template/guest_searchbox.inc" --><%end if%>

	<%if showVerify then%>
		<div class="region form-region">
			<h3 class="title">验证已加密留言</h3>
			<div class="content">
				<form method="post" action="showword.asp?id=<%=id%>" name="form5" onsubmit="if(this.ipass.value==''){alert('请输入密码。');this.ipass.focus();return false;} else if(this.ivcode && this.ivcode.value==''){alert('请输入验证码。');this.ivcode.focus(); return false;} else this.submit1.disabled=true;">
				<input type="hidden" name="ispostback" />

				<div class="field">
					<span class="label">密码：</span>
					<span class="value"><input type="password" name="ipass" maxlength="16" autofocus="autofocus" /></span>
				</div>

				<%if VcodeCount>0 then%>
				<div class="field">
					<span class="label">验证码：</span>
					<span class="value"><input type="text" name="ivcode" autocomplete="off"><img id="captcha" class="captcha" src="show_vcode.asp?t=0" /></span>
				</div>
				<%end if%>

				<div class="command">
					<input value="确定" type="submit" name="submit1" />
				</div>

				</form>
			</div>
		</div>
	<%elseif showWord then
		dim ItemsPerPage
		ItemsPerPage=1%>
		<!-- #include file="include/template/guest_listword.inc" -->
	<%end if%>

	<%
	if rs.State=1 then rs.Close
	cn.Close : set rs=nothing : set cn=nothing
	%>

	<%if showMessage<>"" then Response.Write "<br/><br/><div class=""centertext"">" & showMessage & "</div><br/><br/>"%>

	<!-- #include file="include/template/guest_func.inc" -->

	<%if StatusSearch and ShowBottomSearchBox then%><!-- #include file="include/template/guest_searchbox.inc" --><%end if%>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>

<script type="text/javascript">
	<!-- #include file="asset/js/refresh-captcha.js" -->
</script>
<!-- #include file="include/template/getclientinfo.inc" -->
</body>
</html>