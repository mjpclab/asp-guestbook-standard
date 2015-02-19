<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if isnumeric(Request.QueryString("id"))=false or Request.QueryString("id")="" then
	Response.Redirect "admin.asp"
	Response.End 
end if
%>

<!-- #include file="inc_dtd.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-cn">
<head>
	<title><%=HomeName%> 留言本 回复留言</title>

	<script type="text/javascript">
	//<![CDATA[
	
	function submitcheck(cobject)
	{
		if (cobject.rcontent.value=="") {alert('请输入回复内容。'); cobject.rcontent.focus(); return (false);}
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
	
	//]]>
	</script>
	
	<!-- #include file="style.asp" -->
</head>

<body<%=bodylimit%> onload="sfocus();<%=framecheck%>">
<%
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

CreateConn cn,dbtype
rs.Open sql_adminreply_reply &Request("id"),cn,,,1
	
if rs.EOF=false then 
	t_html=rs("htmlflag")
	c_old="" & rs("reinfo") & ""
	c_old=replace(server.htmlEncode(c_old),chr(13)&chr(10),"&#13;&#10;")
else
	t_html=adminlimit
	c_old=""
end if

rs.Close
cn.close
%>

<div id="outerborder" class="outerborder">

<%if ShowTitle=true then show_book_title 3,"管理"%>
<!-- #include file="admintool.inc" -->

<table cellpadding="2" class="generalwindow" ID="Table1">
	<tr>
		<td class="centertitle">回复留言</td>
	</tr>
	<tr>
		<td class="wordscontent" style="text-align:center; padding:20px 2px;">
			<form method="post" action="admin_savereply.asp" onsubmit="return submitcheck(this)" name="form3">
			回复内容：<br/>
			<textarea name="rcontent" id="rcontent" onkeydown="if(!this.modified)this.modified=true; var e=event?event:arguments[0]; if(e && e.ctrlKey && e.keyCode==13 && this.form.submit1)this.form.submit1.click();" cols="<%=ReplyTextWidth%>" rows="<%=ReplyTextHeight%>"><%=c_old%></textarea>
			<!-- #include file="ubbtoolbar.inc" -->
			<%ShowUbbToolBar(true)%>
			<input type="hidden" name="rootid" value="<%=request.QueryString("rootid")%>" />
			<input type="hidden" name="mainid" value="<%=request.QueryString("id")%>" />
			<input type="hidden" name="page" value="<%=request.QueryString("page")%>" />
			<input type="hidden" name="type" value="<%=request.QueryString("type")%>" />
			<input type="hidden" name="searchtxt" value="<%=request.QueryString("searchtxt")%>" />
			<p style="text-align:left;">
				<input type="checkbox" name="html1" id="html1" value="1"<%if cint(t_html and 1)<>0 then Response.Write " checked=""checked"""%> /><label for="html1">支持HTML标记</label><br/>
				<input type="checkbox" name="ubb1" id="ubb1" value="1"<%if cint(t_html and 2)<>0 then Response.Write " checked=""checked"""%> /><label for="ubb1">支持UBB标记</label><br/>
				<input type="checkbox" name="newline1" id="newline1" value="1"<%if cint(t_html and 4)<>0 then Response.Write " checked=""checked"""%> /><label for="newline1">不支持HTML和UBB标记时允许回车换行</label><br/>
				<br/>
				<input type="checkbox" name="lock2top" id="lock2top" value="1" /><label for="lock2top">回复后置顶留言</label><br/>
				<input type="checkbox" name="bring2top" id="bring2top" value="1" /><label for="bring2top">回复后提前留言</label>
			</p>
			<input type="submit" value="发表回复" name="submit1" id="submit1" />
			</form>
		</td>
	</tr>
</table>

<%
CreateConn cn,dbtype
rs.Open sql_adminreply_words & Request.QueryString("id"),cn,,,1

if rs.EOF=false then
	dim pagename
	pagename="admin_reply"
	%><!-- #include file="listword_admin.inc" --><%
end if
rs.Close : cn.Close : set rs=nothing : set cn=nothing	
%>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>