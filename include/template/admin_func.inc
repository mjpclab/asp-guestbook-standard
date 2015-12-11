<div class="guest-functions">
	<div class="main">
		<a class="function function-multi-del" href="javascript:for(var i=0;i<=form7.elements.seltodel.length-1;i++)if(form7.elements.seltodel[i].checked){<%if DelSelTip=true then Response.Write "if (confirm('确实要删除选定留言吗？')==true)"%>form7.submit();break;}else if(i==form7.elements.seltodel.length-1)alert('请先选定要删除的留言。');"><img src="image/icon_mdel.gif"/>删除选定留言</a>
		<a class="function function-multi-pass" href="javascript:for(var i=0;i<=form7.elements.seltodel.length-1;i++)if(form7.elements.seltodel[i].checked){<%if PassSelTip=true then Response.Write "if (confirm('确实要通过审核选定留言吗？')==true)"%>{form7.action='admin_mpassaudit.asp';form7.submit();}break;}else if(i==form7.elements.seltodel.length-1)alert('请先选定要通过审核的留言。');"><img src="image/icon_mpass.gif"/>通过审核选定留言</a>
		<%if RPage<>"" then%>
			<%if AdminDisplayMode="book" then%>
				<a class="function" href="setdisplaymode.asp?mode=forum&modeflag=admin&rpage=<%=RPage%><%=GetRequests()%>"><img alt="" src="image/icon_mode_forum.gif" style="border-width:0px;" />切换到标题模式</a>
			<%elseif AdminDisplayMode="forum" then%>
				<a class="function" href="setdisplaymode.asp?mode=book&modeflag=admin&rpage=<%=RPage%><%=GetRequests()%>"><img alt="" src="image/icon_mode_book.gif" style="border-width:0px;" />切换到完整模式</a>
			<%end if%>
		<%end if%>
	</div>
</div>
