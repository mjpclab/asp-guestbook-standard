<table class="topic-list">
<thead>
	<tr>
		<th class="icon"></th><th class="title">标题</th><th class="name">姓名</th><th class="date">留言时间</th>
	</tr>
</thead>
<tbody>
<%
while Not rs.eof
	guestflag=rs("guestflag")
	isauditting=CBool(guestflag AND 16)
	ishidden=((guestflag AND 40)=8)
	iswhisper=CBool(guestflag AND 32)
	encrypted=((guestflag AND 96)=96)
	%>
	<tr class="message<%if isauditting then%> auditting-message<%end if%>">
		<td class="icon">
			<%if isauditting then%><img src="asset/image/icon_wait2pass.gif" title="待审核" alt="待审核"/>
			<%elseif ishidden then%><img src="asset/image/icon_hide.gif" title="被隐藏" alt="被隐藏"/>
			<%elseif iswhisper then%><img src="asset/image/icon_whisper.gif" title="悄悄话" alt="悄悄话"/>
			<%elseif rs.Fields("parent_id")<0 then%><img src="asset/image/icon_toplocked.gif" title="置顶" alt="置顶"/>
			<%end if%>
		</td>
		<td class="title">
			<%if isauditting then%><a href="showword.asp?id=<%=rs.Fields("id")%>">(留言待审核...)</a>
			<%elseif ishidden then%><a href="showword.asp?id=<%=rs.Fields("id")%>">(留言被管理员隐藏...)</a>
			<%elseif iswhisper then%><a href="showword.asp?id=<%=rs.Fields("id")%>">(给版主的悄悄话...)</a>
			<%elseif Trim(rs.Fields("title"))="" then%><a href="showword.asp?id=<%=rs.Fields("id")%>">(无标题)</a>
			<%else%><a href="showword.asp?id=<%=rs.Fields("id")%>" title="<%=rs.Fields("title")%>"><%=rs.Fields("title")%></a>
			<%end if%>
		</td>
		<td class="name"><%if Not isauditting then Response.Write rs.Fields("name")%></td>
		<td class="date"><%=UTCToDisplayTime(rs.Fields("logdate"))%></td>
	</tr>
	<%rs.MoveNext
wend
%>
</tbody>
</table>
