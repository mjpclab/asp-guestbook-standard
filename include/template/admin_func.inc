<div class="guest-functions">
	<div class="main">
		<a class="function function-multi-del" href="javascript:void 0;"><img src="asset/image/icon_mdel.gif"/>ɾ��ѡ������</a>
		<a class="function function-multi-pass" href="javascript:void 0;"><img src="asset/image/icon_mpass.gif"/>ͨ�����ѡ������</a>
		<%if RPage<>"" then%>
			<%if AdminDisplayMode="book" then%>
				<a class="function" href="setdisplaymode.asp?mode=forum&modeflag=admin&rpage=<%=RPage%><%=GetRequests()%>"><img alt="" src="asset/image/icon_mode_forum.gif" style="border-width:0px;" />�л�������ģʽ</a>
			<%elseif AdminDisplayMode="forum" then%>
				<a class="function" href="setdisplaymode.asp?mode=book&modeflag=admin&rpage=<%=RPage%><%=GetRequests()%>"><img alt="" src="asset/image/icon_mode_book.gif" style="border-width:0px;" />�л�������ģʽ</a>
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
				alert('����ѡ��Ҫɾ�������ԡ�');
			} else if (!delSelTip || confirm('ȷʵҪɾ��ѡ��������')) {
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
				alert('����ѡ��Ҫͨ����˵����ԡ�');
			} else if (!passSelTip || confirm('ȷʵҪͨ�����ѡ��������')) {
				form7.action = 'admin_mpassaudit.asp';
				form7.submit();
			}
		});
	}());
</script>