<script type="text/javascript">
function searchsubmitcheck(frm)
{
	if (frm.searchtxt.value.length===0) {alert('�������������ݣ�');frm.searchtxt.focus();return false;}
	frm.searchsubmit.disabled=true;
	return true;
}
</script>

<div class="region form-region center-region">
	<h3 class="title">[��������]</h3>
	<div class="content">
		<form method="post" action="search.asp" onsubmit="return searchsubmitcheck(this);">
			������<input type="text" name="searchtxt" size="<%=SearchTextWidth%>" value="<%=request("searchtxt")%>" />
			<input type="submit" value="����" name="searchsubmit" />
			<select name="type" size="1" onchange="searchtxt.focus();">
				<option value="name"<%=seled(request("type")="name")%>>����������</option>
				<option value="title"<%=seled(request("type")="title")%>>����������</option>
				<option value="article"<%=seled(request("type")="article" or request("type")="")%>>��������������</option>
				<option value="reply"<%=seled(request("type")="reply")%>>�������ظ�����</option>
			</select>
		</form>
	</div>
</div>
