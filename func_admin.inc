<div class="guest-functions">
	<div class="main">
		<a class="function" href="javascript:for(var i=0;i<=form7.elements.length-1;i++)if(form7.elements[i].name=='seltodel' && form7.elements[i].checked){<%if DelSelTip=true then Response.Write "if (confirm('确实要删除选定留言吗？')==true)"%>form7.submit();break;}else if(i==form7.elements.length-1)alert('请先选定要删除的留言。');" onmouseover="return true;"><img src="image/icon_mdel.gif"/>删除选定留言</a>
		<a class="function" href="javascript:for(var i=0;i<=form7.elements.length-1;i++)if(form7.elements[i].name=='seltodel' && form7.elements[i].checked){<%if PassSelTip=true then Response.Write "if (confirm('确实要通过审核选定留言吗？')==true)"%>{form7.action='admin_mpass.asp';form7.submit();}break;}else if(i==form7.elements.length-1)alert('请先选定要通过审核的留言。');" onmouseover="return true;"><img src="image/icon_mpass.gif"/>通过审核选定留言</a>
		<%if RPage<>"" then%>
			<%if AdminDisplayMode="book" then%>
				<a class="function" href="setdisplaymode.asp?mode=forum&modeflag=admin&rpage=<%=RPage%><%=GetRequests()%>"><img alt="" src="image/icon_mode_forum.gif" style="border-width:0px;" />切换到论坛模式</a>
			<%elseif AdminDisplayMode="forum" then%>
				<a class="function" href="setdisplaymode.asp?mode=book&modeflag=admin&rpage=<%=RPage%><%=GetRequests()%>"><img alt="" src="image/icon_mode_book.gif" style="border-width:0px;" />切换到留言本模式</a>
			<%end if%>
		<%end if%>
	</div>
</div>
