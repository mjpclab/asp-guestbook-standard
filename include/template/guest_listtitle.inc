<table class="topic-list">
<thead>
	<tr>
		<th class="icon"></th><th class="title">����</th><th class="name">����</th><th class="date">����ʱ��</th>
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
			<%if isauditting then%><img src="asset/image/icon_wait2pass.gif" title="�����" alt="�����"/>
			<%elseif ishidden then%><img src="asset/image/icon_hide.gif" title="������" alt="������"/>
			<%elseif iswhisper then%><img src="asset/image/icon_whisper.gif" title="���Ļ�" alt="���Ļ�"/>
			<%elseif rs.Fields("parent_id")<0 then%><img src="asset/image/icon_toplocked.gif" title="�ö�" alt="�ö�"/>
			<%end if%>
		</td>
		<td class="title">
			<%if isauditting then%><a href="showword.asp?id=<%=rs.Fields("id")%>">(���Դ����...)</a>
			<%elseif ishidden then%><a href="showword.asp?id=<%=rs.Fields("id")%>">(���Ա�����Ա����...)</a>
			<%elseif iswhisper then%><a href="showword.asp?id=<%=rs.Fields("id")%>">(�����������Ļ�...)</a>
			<%elseif Trim(rs.Fields("title"))="" then%><a href="showword.asp?id=<%=rs.Fields("id")%>">(�ޱ���)</a>
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
