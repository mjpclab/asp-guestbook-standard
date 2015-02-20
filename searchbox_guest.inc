<script type="text/javascript">
function searchsubmitcheck(frm)
{
	if (frm.searchtxt.value=='') {alert('请输入搜索内容！');frm.searchtxt.focus();return false;}
	frm.searchsubmit.disabled=true;
	return true;
}
</script>

<table cellpadding="2" class="generalwindow">
	<tr>
		<td class="centertitle">[留言搜索]</td>
	</tr>
	<tr>
		<td class="wordscontent" style="text-align:center; padding:20px 0px;">
			<form method="post" action="search.asp" onsubmit="return searchsubmitcheck(this);">
			搜索：<input type="text" name="searchtxt" size="<%=SearchTextWidth%>" value="<%=request("searchtxt")%>" />
			<input type="submit" value="搜索" name="searchsubmit" />
			<select name="type" size="1" onchange="searchtxt.focus();">
				<option value="name"<%=seled(request("type")="name")%>>按姓名搜索</option>
				<option value="title"<%=seled(request("type")="title")%>>按标题搜索</option>
				<option value="article"<%=seled(request("type")="article" or request("type")="")%>>按留言内容搜索</option>
				<option value="reply"<%=seled(request("type")="reply")%>>按版主回复搜索</option>
			</select>
			</form>
		</td>
	</tr>
</table>