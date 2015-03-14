<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if isnumeric(Request.QueryString("id"))=false or Request.QueryString("id")="" then
	Response.Redirect "admin.asp"
	Response.End
end if

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

CreateConn cn,dbtype
rs.Open sql_adminedit & Request.QueryString("id"),cn,,,1
	
if rs.EOF=false then
	guestflag=rs("guestflag")
	guest_txt="" & rs("article") & ""
	guest_txt=replace(server.htmlEncode(guest_txt),chr(13)&chr(10),"&#13;&#10;")
else
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.Redirect "admin.asp"
	Response.End
end if
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> 留言本 编辑留言</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<link rel="stylesheet" type="text/css" href="adminstyle.css"/>
	<!-- #include file="style.asp" -->
	<!-- #include file="adminstyle.asp" -->

	<script type="text/javascript">
	function submitcheck(cobject)
	{
		if (cobject.etitle.value.length===0) {alert('请输入标题。'); cobject.etitle.focus(); return false;}
		cobject.submit1.disabled=true;
		return (true);
	}
	function sfocus()
	{
		if (!form3.modified)
		{
			form3.econtent.focus();
			form3.econtent.select();
		}
	}
	</script>
</head>

<body<%=bodylimit%> onload="sfocus();<%=framecheck%>">

<div id="outerborder" class="outerborder">

<%if ShowTitle=true then show_book_title 3,"管理"%>
<!-- #include file="admincontrols.inc" -->

<div class="region">
	<h3 class="title">编辑留言</h3>
	<div class="content">
		<form method="post" action="admin_saveedit.asp" onsubmit="return submitcheck(this)" name="form3">
		<input type="hidden" name="rootid" value="<%=request.QueryString("rootid")%>" />
		<input type="hidden" name="mainid" value="<%=request.QueryString("id")%>" />
		<input type="hidden" name="page" value="<%=request.QueryString("page")%>" />
		<input type="hidden" name="type" value="<%=request.QueryString("type")%>" />
		<input type="hidden" name="searchtxt" value="<%=request.QueryString("searchtxt")%>" />
		<div class="field">
			<span class="row">标题：</span>
			<span class="row"><input type="text" name="etitle" onkeydown="if(!this.form.modified)this.form.modified=true;" size="<%=ReplyTextWidth%>" maxlength="30" value="<%=rs.Fields("title")%>"></span>
		</div>
		<div class="field">
			<span class="row">内容：</span>
			<span class="row"><textarea name="econtent" id="econtent" onkeydown="if(!this.form.modified)this.form.modified=true; var e=event?event:arguments[0]; if(e && e.ctrlKey && e.keyCode==13 && this.form.submit1)this.form.submit1.click();" cols="<%=ReplyTextWidth%>" rows="<%=ReplyTextHeight%>"><%=guest_txt%></textarea></span>
			<span class="row">
				<!-- #include file="ubbtoolbar.inc" -->
				<%ShowUbbToolBar(true)%>
			</span>
			<span class="row">
				<input type="checkbox" name="html1" id="html1" value="1"<%if cint(guestflag and 1)<>0 then Response.Write " checked=""checked"""%> /><label for="html1">支持HTML标记</label><br/>
				<input type="checkbox" name="ubb1" id="ubb1" value="1"<%if cint(guestflag and 2)<>0 then Response.Write " checked=""checked"""%> /><label for="ubb1">支持UBB标记</label><br/>
				<input type="checkbox" name="newline1" id="newline1" value="1"<%if cint(guestflag and 4)<>0 then Response.Write " checked=""checked"""%> /><label for="newline1">不支持HTML和UBB标记时允许回车换行</label>
			</span>
		</div>
		<div class="command"><input type="submit" value="保存留言" name="submit1" id="submit1" /></div>
		</form>
	</div>
</div>

<%dim pagename
pagename="admin_edit"%>
<!-- #include file="listword_admin.inc" -->
<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>