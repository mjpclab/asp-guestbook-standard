<div class="guest-functions">
	<div class="main">
		<a class="function function-multi-del" href="javascript:void 0;"><img src="asset/image/icon_mdel.gif"/>删除选定留言</a>
		<a class="function function-multi-pass" href="javascript:void 0;"><img src="asset/image/icon_mpass.gif"/>通过审核选定留言</a>
		<%if RPage<>"" then%>
			<%if AdminDisplayMode="book" then%>
				<a class="function" href="setdisplaymode.asp?mode=forum&modeflag=admin&rpage=<%=RPage%><%=GetRequests()%>"><img alt="" src="asset/image/icon_mode_forum.gif" style="border-width:0px;" />切换到标题模式</a>
			<%elseif AdminDisplayMode="forum" then%>
				<a class="function" href="setdisplaymode.asp?mode=book&modeflag=admin&rpage=<%=RPage%><%=GetRequests()%>"><img alt="" src="asset/image/icon_mode_book.gif" style="border-width:0px;" />切换到完整模式</a>
			<%end if%>
		<%end if%>
	</div>
</div>
<script type="text/javascript">
	(function multiDelConfirm(){
		var $multiDelLink = $('.function-multi-del:last');
		$multiDelLink.click(function () {
			var delSelTip = <%=LCase(CStr(DelSelTip))%>;
			var $selected = $('input.seltodel:checked');
			if (!$selected.length) {
				alert('请先选定要删除的留言。');
			} else if (!delSelTip || confirm('确实要删除选定留言吗？')) {
				form7.submit();
			}
		});
	}());
	(function multiPassConfirm() {
		var $multiPassLink = $('.function-multi-pass:last');
		$multiPassLink.click(function () {
			var passSelTip = <%=LCase(CStr(PassSelTip))%>;
			var $selected = $('input.seltodel:checked');
			if (!$selected.length) {
				alert('请先选定要通过审核的留言。');
			} else if (!passSelTip || confirm('确实要通过审核选定留言吗？')) {
				form7.action = 'admin_mpassaudit.asp';
				form7.submit();
			}
		});
	}());
</script>