<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/const.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/common2.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_reply.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/ubbcode.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="include/utility/message.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
Response.Expires=-1
if isnumeric(Request.QueryString("id"))=false or Request.QueryString("id")="" then
	Response.Redirect "admin.asp"
	Response.End 
end if
%>

<!-- #include file="include/template/dtd.inc" -->
<html lang="zh-CN">
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 回复留言</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->

	<script type="text/javascript">
	function submitcheck(cobject)
	{
		if (cobject.rcontent.value.length===0) {alert('请输入回复内容。'); cobject.rcontent.focus(); return (false);}
		cobject.submit1.disabled=true;
		return (true);
	}
	function sfocus()
	{
		if (!form3.rcontent.modified)
		{
			form3.rcontent.focus();
			form3.rcontent.select();
		}
	}
	</script>
</head>

<body<%=bodylimit%> onload="sfocus();<%=framecheck%>">
<%
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

Call CreateConn(cn)
rs.Open sql_adminreply_reply &Request("id"),cn,,,1
	
if Not rs.EOF then
	t_html=rs("htmlflag")
	c_old="" & rs("reinfo") & ""
else
	t_html=adminlimit
	c_old=""
end if

rs.Close
cn.close
%>

<div id="outerborder" class="outerborder">

<%if ShowTitle then%><%Call InitHeaderData("管理")%><!-- #include file="include/template/header.inc" --><%end if%>
<div id="mainborder" class="mainborder">
<!-- #include file="include/template/admin_mainmenu.inc" -->
<div class="region">
	<h3 class="title">回复留言</h3>
	<div class="content">
		<form method="post" action="admin_savereply.asp" onsubmit="return submitcheck(this)" name="form3">
		回复内容：<br/>
		<textarea name="rcontent" id="rcontent" onkeydown="if(!this.modified)this.modified=true; var e=event?event:arguments[0]; if(e && e.ctrlKey && e.keyCode==13 && this.form.submit1)this.form.submit1.click();" rows="<%=ReplyTextHeight%>"><%=c_old%></textarea>
		<!-- #include file="include/template/ubbtoolbar.inc" -->
		<%ShowUbbToolBar(true)%>
		<input type="hidden" name="rootid" value="<%=request.QueryString("rootid")%>" />
		<input type="hidden" name="mainid" value="<%=request.QueryString("id")%>" />
		<input type="hidden" name="page" value="<%=request.QueryString("page")%>" />
		<input type="hidden" name="type" value="<%=request.QueryString("type")%>" />
		<input type="hidden" name="searchtxt" value="<%=request.QueryString("searchtxt")%>" />
		<p>
			<input type="checkbox" name="html1" id="html1" value="1"<%if cint(t_html and 1)<>0 then Response.Write " checked=""checked"""%> /><label for="html1">支持HTML标记</label><br/>
			<input type="checkbox" name="ubb1" id="ubb1" value="1"<%if cint(t_html and 2)<>0 then Response.Write " checked=""checked"""%> /><label for="ubb1">支持UBB标记</label><br/>
			<input type="checkbox" name="newline1" id="newline1" value="1"<%if cint(t_html and 4)<>0 then Response.Write " checked=""checked"""%> /><label for="newline1">不支持HTML和UBB标记时允许回车换行</label><br/>
			<br/>
			<input type="checkbox" name="lock2top" id="lock2top" value="1" /><label for="lock2top">回复后置顶留言</label><br/>
			<input type="checkbox" name="bring2top" id="bring2top" value="1" /><label for="bring2top">回复后提前留言</label>
		</p>
		<div class="command"><input type="submit" value="发表回复" name="submit1" id="submit1" /></div>
		</form>
	</div>
</div>

<%
iPage=Request.QueryString("page")

Call CreateConn(cn)
rs.Open sql_adminreply_words & Request.QueryString("id"),cn,,,1

if Not rs.EOF then
	dim pagename, inAdminPage
	pagename="admin_reply"
	inAdminPage=true
	%><!-- #include file="include/template/admin_listword.inc" --><%
end if
rs.Close : cn.Close : set rs=nothing : set cn=nothing	
%>
</div>

<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>