<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
Response.Expires=-1
%>

<!-- #include file="include/template/dtd.inc" -->
<html lang="zh-CN">
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 生成调用代码</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->

	<script type="text/javascript">
	function generateCallCode()
	{
		var frm_n=document.getElementById('frm_n');
		var frm_len=document.getElementById('frm_len');
		var frm_code=document.getElementById('frm_code');

		var mode=document.getElementById('frm_mode').value;
		var baseurl=document.getElementById('frm_baseurl').value;
		var n=frm_n.value;
		var len=frm_len.value;
		var prefix=document.getElementById('frm_prefix').value;
		var target=document.getElementById('frm_target').value;

		if(n.length===0)
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

		if(baseurl.substr(-1)!=='/') {
			baseurl+='/';
		}

		var url=baseurl + 'tlist.asp?mode=' + mode + '&n=' + n;
		if(len) {
			url+='&len='+len;
		}
		if(prefix) {
			url+='&prefix='+encodeURIComponent(prefix)
		}
		if(target) {
			url+='&target='+target;
		}

		if(mode==='iframe') {
			frm_code.value='<iframe src="' +url+ '" frameborder="0"><\/iframe>';
		}
		else if(mode==='js') {
			frm_code.value='<' + 'script type="text/javascript" src="' +url+ '" charset="utf-8"><\/script>';
		}
		else if(mode==='json') {
			frm_code.value='请用GET方式访问以下URL：\n' + url;
		}
	}
	</script>
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

<%if ShowTitle then%><%Call InitHeaderData("管理")%><!-- #include file="include/template/header.inc" --><%end if%>
<div id="mainborder" class="mainborder">
<!-- #include file="include/template/admin_mainmenu.inc" -->
<div class="region region-longtext region-callgen">
	<h3 class="title">生成调用代码</h3>
	<div class="content">
		<p>此页用于生成留言本留言标题调用代码。请输入调用代码所需的参数，其中“显示条数”为必填项：</p>

		<form>
		<div class="field field-mode">
			<span class="label">使用模式</span>
			<span class="value">
				<select name="frm_mode" id="frm_mode">
					<option value="iframe">iframe</option>
					<option value="js">JS输出</option>
					<option value="json">JSON数据</option>
				</select>
			</span>
		</div>
		<div class="field field-base-url">
			<span class="label">留言本根URL</span>
			<span class="value"><input type="text" name="frm_baseurl" id="frm_baseurl" value="<%=geturlpath%>" /></span>
		</div>
		<div class="field field-n">
			<span class="label">显示条数<span class="required">*</span></span>
			<span class="value"><input type="text" name="frm_n" id="frm_n" maxlength="10" value="10"/></span>
		</div>
		<div class="field field-len">
			<span class="label">标题字数限制</span>
			<span class="value"><input type="text" name="frm_len" id="frm_len"/></span>
		</div>
		<div class="field field-prefix">
			<span class="label">标题前缀</span>
			<span class="value"><input type="text" name="frm_prefix" id="frm_prefix"/></span>
		</div>
		<div class="field field-target">
			<span class="label">打开窗口</span>
			<span class="value">
				<select name="frm_target" id="frm_target">
					<option value="">(默认)</option>
					<option value="_blank">打开新页面(_blank)</option>
					<option value="_self">相同窗口或框架窗口(_self)</option>
					<option value="_top">整个浏览器窗口(_top)</option>
					<option value="_parent">父框架窗口(_parent)</option>
				</select>
			</span>
		</div>
		<div class="field-command">
			<input type="button" name="btn_generate" id="btn_generate" value="生成调用代码" onclick="generateCallCode();" />
		</div>
		<div class="field">
			<span class="row">已生成的调用代码：</span>
			<span class="value"><textarea readonly="readonly" name="frm_code" id="frm_code" rows="<%=ReplyTextHeight%>"></textarea></span>
		</div>
		</form>
	</div>
</div>
</div>

<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>