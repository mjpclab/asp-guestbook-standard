<div class="guest-functions">
	<div class="main">
		<%if StatusWrite=true then%><a class="function" href="leaveword.asp" title="ǩд����"><img src="asset/image/icon_write.gif"/>ǩд����</a><%end if%>
		<%if RPage<>"" then%>
			<%if GuestDisplayMode="book" then%>
				<a class="function" href="setdisplaymode.asp?mode=forum&modeflag=guest&rpage=<%=RPage%><%=GetRequests()%>"><img src="asset/image/icon_mode_forum.gif"/>�л�������ģʽ</a>
			<%elseif GuestDisplayMode="forum" then%>
				<a class="function" href="setdisplaymode.asp?mode=book&modeflag=guest&rpage=<%=RPage%><%=GetRequests()%>"><img src="asset/image/icon_mode_book.gif"/>�л�������ģʽ</a>
			<%end if%>
		<%end if%>
	</div>
	<div class="aside">
		<a class="function" href="admin.asp">����</a>
	</div>
</div>