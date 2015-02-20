<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->
<!-- #include file="common2.asp" -->

<%
Response.Expires=-1
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> 留言本 生成调用代码</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<!-- #include file="style.asp" -->

	<script type="text/javascript">
	function swIframe(jschk)
	{
		var span=document.getElementById('iframeSettings');
		if(jschk.checked)
			span.style.visibility='hidden';
		else
			span.style.visibility='visible';
	}
	
	function urlEncode(url)
	{
		return url.replace('%','%25').replace('?','%3F').replace(' ','%20').replace('#','%23').replace('&','%26').replace('\'','%27').replace('/','%2F').replace('<','%3C').replace('>','%3E');
	}
	
	function generateCallCode()
	{
		var frm_n=document.getElementById('frm_n');
		var frm_len=document.getElementById('frm_len');
		var frm_code=document.getElementById('frm_code');
	
		var url=document.getElementById('url').value;
		var n=document.getElementById('frm_n').value;
		var target=document.getElementById('frm_target').value;
		var prefix=document.getElementById('frm_prefix').value;
		var len=document.getElementById('frm_len').value;
		var nobr=document.getElementById('frm_nobr').checked;
		var js=document.getElementById('frm_js').checked;
		var width=document.getElementById('frm_width').value;
		var height=document.getElementById('frm_height').value;
		
		if(n=='')
		{
			alert('请输入显示条数。');
			frm_n.focus();
			return false;
		}
		else if(isNaN(n))
		{
			alert('显示条数有误，请检查。');
			frm_n.select();
			return false;
		}
		else if(isNaN(len))
		{
			alert('字数限制有误，请检查。');
			frm_len.select();
			return false;
		}
		else
		{
			var temp;
			if(js) temp='<' + 'script type="text/javascript" language="javascript" src="{0}"><\/script>';
			else temp='<iframe width="{w}" height="{h}" src="{0}" frameborder="0"><\/iframe>';
			
			var src=
				url +
				'?n=' + n +
				(js ? '&js=yes' : '') +
				(target!='' ? '&target='+target : '') +
				(prefix!='' ? '&prefix='+urlEncode(prefix) : '') +
				(len!='' ? '&len='+len : '') +
				(nobr ? '&nobr=yes' : '');
			
			frm_code.value=temp.replace('{w}',width).replace('{h}',height).replace('{0}',src);
			return false;
		}
	}
	</script>

</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

<%if ShowTitle=true then show_book_title 3,"管理"%>
<!-- #include file="admintool.inc" -->

<table border="1" cellpadding="2" class="generalwindow">
	<tr>
		<td class="centertitle">生成调用代码</td>
	</tr>
	<tr>
		<td class="wordscontent" style="padding:20px 2px;">
			<form method="post" action="" onsubmit="return generateCallCode();">
			<input type="hidden" name="IsPostBack" id="IsPostBack" value="1" />
			<input type="hidden" name="url" id="url" value="<%=geturlpath%>tlist.asp" />
			
			此页用于生成留言本留言标题调用代码。请输入调用代码所需的参数，其中“显示条数”为必填项：<br/><br/>
			
			显示条数：　<input type="text" name="frm_n" id="frm_n" size="10" maxlength="10" value="<%=Request.Form("frm_n")%>" />*<br/>
			打开窗口：　<select name="frm_target" id="frm_target">
				<option value="" <%=seled(Request.Form("frm_target")="")%>>(默认)</option>
				<option value="_blank" <%=seled(Request.Form("frm_target")="_blank")%>>打开新页面</option>
				<option value="_self" <%=seled(Request.Form("frm_target")="_self")%>>相同窗口或框架窗口</option>
				<option value="_top" <%=seled(Request.Form("frm_target")="_top")%>>整个浏览器窗口</option>
				<option value="_parent" <%=seled(Request.Form("frm_target")="_parent")%>>父框架窗口</option>
			</select><br/>
			标题前缀：　<input type="text" name="frm_prefix" id="frm_prefix" size="10" value="<%=Request.Form("frm_prefix")%>" /><br/>
			字数限制：　<input type="text" name="frm_len" id="frm_len" size="10" value="<%=Request.Form("frm_len")%>" /><br/>
			不输出断行：<input type="checkbox" name="frm_nobr" id="frm_nobr" value="1" <%=cked(Request.Form("frm_nobr")="1")%> /><br/>
			使用JS模式：<input type="checkbox" name="frm_js" id="frm_js" value="1" <%=cked(Request.Form("frm_js")="1")%> onclick="swIframe(this);" /><br/>
			<span id="iframeSettings" style="<%if Request.Form("frm_js")="1" then Response.Write "visibility:hidden;"%>">
			窗口宽度：　<input type="text" name="frm_width" id="frm_width" size="10" value="<%=Request.Form("frm_width")%>" /><br/>
			窗口高度：　<input type="text" name="frm_height" id="frm_height" size="10" value="<%=Request.Form("frm_height")%>" /><br/>
			</span>
			<br/><input type="submit" name="submit1" id="submit1" value="生成调用代码" /><br/><br/>
			
			
			已生成的调用代码：<br/>
			<textarea readonly="readonly" name="frm_code" id="frm_code" cols="<%=ReplyTextWidth%>" rows="<%=ReplyTextHeight%>"></textarea>
			</form>
		</td>
	</tr>
</table>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>