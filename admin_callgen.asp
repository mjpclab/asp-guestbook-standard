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
	<title><%=HomeName%> ���Ա� ���ɵ��ô���</title>
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
			alert('��������ʾ������');
			frm_n.focus();
			return false;
		}
		else if(isNaN(n))
		{
			alert('��ʾ�����������顣');
			frm_n.select();
			return false;
		}
		else if(isNaN(len))
		{
			alert('���������������顣');
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

<%if ShowTitle=true then show_book_title 3,"����"%>
<!-- #include file="admintool.inc" -->

<table border="1" cellpadding="2" class="generalwindow">
	<tr>
		<td class="centertitle">���ɵ��ô���</td>
	</tr>
	<tr>
		<td class="wordscontent" style="padding:20px 2px;">
			<form method="post" action="" onsubmit="return generateCallCode();">
			<input type="hidden" name="IsPostBack" id="IsPostBack" value="1" />
			<input type="hidden" name="url" id="url" value="<%=geturlpath%>tlist.asp" />
			
			��ҳ�����������Ա����Ա�����ô��롣��������ô�������Ĳ��������С���ʾ������Ϊ�����<br/><br/>
			
			��ʾ��������<input type="text" name="frm_n" id="frm_n" size="10" maxlength="10" value="<%=Request.Form("frm_n")%>" />*<br/>
			�򿪴��ڣ���<select name="frm_target" id="frm_target">
				<option value="" <%=seled(Request.Form("frm_target")="")%>>(Ĭ��)</option>
				<option value="_blank" <%=seled(Request.Form("frm_target")="_blank")%>>����ҳ��</option>
				<option value="_self" <%=seled(Request.Form("frm_target")="_self")%>>��ͬ���ڻ��ܴ���</option>
				<option value="_top" <%=seled(Request.Form("frm_target")="_top")%>>�������������</option>
				<option value="_parent" <%=seled(Request.Form("frm_target")="_parent")%>>����ܴ���</option>
			</select><br/>
			����ǰ׺����<input type="text" name="frm_prefix" id="frm_prefix" size="10" value="<%=Request.Form("frm_prefix")%>" /><br/>
			�������ƣ���<input type="text" name="frm_len" id="frm_len" size="10" value="<%=Request.Form("frm_len")%>" /><br/>
			��������У�<input type="checkbox" name="frm_nobr" id="frm_nobr" value="1" <%=cked(Request.Form("frm_nobr")="1")%> /><br/>
			ʹ��JSģʽ��<input type="checkbox" name="frm_js" id="frm_js" value="1" <%=cked(Request.Form("frm_js")="1")%> onclick="swIframe(this);" /><br/>
			<span id="iframeSettings" style="<%if Request.Form("frm_js")="1" then Response.Write "visibility:hidden;"%>">
			���ڿ�ȣ���<input type="text" name="frm_width" id="frm_width" size="10" value="<%=Request.Form("frm_width")%>" /><br/>
			���ڸ߶ȣ���<input type="text" name="frm_height" id="frm_height" size="10" value="<%=Request.Form("frm_height")%>" /><br/>
			</span>
			<br/><input type="submit" name="submit1" id="submit1" value="���ɵ��ô���" /><br/><br/>
			
			
			�����ɵĵ��ô��룺<br/>
			<textarea readonly="readonly" name="frm_code" id="frm_code" cols="<%=ReplyTextWidth%>" rows="<%=ReplyTextHeight%>"></textarea>
			</form>
		</td>
	</tr>
</table>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>