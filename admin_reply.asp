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
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> ���Ա� �ظ�����</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->

	<script type="text/javascript">
	function submitcheck(cobject)
	{
		if (cobject.rcontent.value=="") {alert('������ظ����ݡ�'); cobject.rcontent.focus(); return (false);}
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

<%if ShowTitle=true then show_book_title 3,"����"%>
<!-- #include file="admincontrols.inc" -->

<div class="region">
	<h3 class="title">�ظ�����</h3>
	<div class="content">
		<form method="post" action="admin_savereply.asp" onsubmit="return submitcheck(this)" name="form3">
		�ظ����ݣ�<br/>
		<textarea name="rcontent" id="rcontent" onkeydown="if(!this.modified)this.modified=true; var e=event?event:arguments[0]; if(e && e.ctrlKey && e.keyCode==13 && this.form.submit1)this.form.submit1.click();" rows="<%=ReplyTextHeight%>"><%=c_old%></textarea>
		<!-- #include file="ubbtoolbar.inc" -->
		<%ShowUbbToolBar(true)%>
		<input type="hidden" name="rootid" value="<%=request.QueryString("rootid")%>" />
		<input type="hidden" name="mainid" value="<%=request.QueryString("id")%>" />
		<input type="hidden" name="page" value="<%=request.QueryString("page")%>" />
		<input type="hidden" name="type" value="<%=request.QueryString("type")%>" />
		<input type="hidden" name="searchtxt" value="<%=request.QueryString("searchtxt")%>" />
		<p>
			<input type="checkbox" name="html1" id="html1" value="1"<%if cint(t_html and 1)<>0 then Response.Write " checked=""checked"""%> /><label for="html1">֧��HTML���</label><br/>
			<input type="checkbox" name="ubb1" id="ubb1" value="1"<%if cint(t_html and 2)<>0 then Response.Write " checked=""checked"""%> /><label for="ubb1">֧��UBB���</label><br/>
			<input type="checkbox" name="newline1" id="newline1" value="1"<%if cint(t_html and 4)<>0 then Response.Write " checked=""checked"""%> /><label for="newline1">��֧��HTML��UBB���ʱ����س�����</label><br/>
			<br/>
			<input type="checkbox" name="lock2top" id="lock2top" value="1" /><label for="lock2top">�ظ����ö�����</label><br/>
			<input type="checkbox" name="bring2top" id="bring2top" value="1" /><label for="bring2top">�ظ�����ǰ����</label>
		</p>
		<div class="command"><input type="submit" value="����ظ�" name="submit1" id="submit1" /></div>
		</form>
	</div>
</div>

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