<%
param_url="?rootid=" & p_rs.Fields("root_id") & "&id=" & p_rs.Fields("id") & "&page=" & ipage
if request("type")<>"" and request("searchtxt")<>"" then
	param_url=param_url & "&type=" & server.URLEncode(request("type")) & "&searchtxt=" & server.URLEncode(request("searchtxt"))
end if
param_url=Server.HtmlEncode(param_url)
%>

<div class="admin-tools">
	<div class="group">
		<div class="name">ѡ�����ԣ�</div>
		<div class="tools">
			<span class="tool"><input type="checkbox" name="seltodel" id="c<%=p_rs.Fields("id")%>" value="<%=p_rs("id")%>" /><label for="c<%=p_rs.Fields("id")%>">(ѡ��)</label></span>
		</div>
	</div>
	<div class="group">
		<div class="name">�ÿ����ԣ�</div>
		<div class="tools">
			<%if clng(guestflag and 16)<>0 then%><a class="tool" href="admin_passaudit.asp<%=param_url%>" title="ͨ�����"<%if PassAuditTip=true then response.Write " onclick=""return confirm('ȷʵҪͨ�������');"""%>><img src="image/icon_pass.gif" />[ͨ�����]</a><%end if%>
			<a class="tool" href="admin_edit.asp<%=param_url%>" title="�༭�ÿ�����"><img src="image/icon_edit.gif" />[�༭����]</a>
			<%if clng(guestflag and 32)<>0 then%><a class="tool" href="admin_pubwhisper.asp<%=param_url%>" title="�������Ļ�"<%if PubWhisperTip=true then response.Write " onclick=""return confirm('ȷʵҪ�������Ļ���');"""%>><img src="image/icon_pub.gif" />[�������Ļ�]</a><%end if%>
			<%if clng(guestflag and 256)=0 then%><a class="tool" href="admin_hidecontact.asp<%=param_url%>" title="���طÿ���ϵ��ʽ"><img src="image/icon_hidecontact.gif" />[������ϵ]</a><%end if%>
			<%if clng(guestflag and 256)<>0 then%><a class="tool" href="admin_unhidecontact.asp<%=param_url%>" title="�����ÿ���ϵ��ʽ"><img src="image/icon_unhidecontact.gif" />[������ϵ]</a><%end if%>
			<%if clng(guestflag and 40)=0 then%><a class="tool" href="admin_hideword.asp<%=param_url%>" title="���طÿ���������"><img src="image/icon_hide.gif" />[��������]</a><%end if%>
			<%if clng(guestflag and 40)=8 then%><a class="tool" href="admin_unhideword.asp<%=param_url%>" title="�����ÿ���������"><img src="image/icon_unhide.gif" />[��������]</a><%end if%>
		</div>
	</div>
	<div class="group">
		<div class="name">�������ܣ�</div>
		<div class="tools">
			<a class="tool" href="admin_reply.asp<%=param_url%>" title="<%if clng(guestflag and 16)<>0 then response.write "ͨ����˲�"%>�ظ�������"<%if clng(guestflag and 16)<>0 and PassAuditTip=true then response.Write " onclick=""return confirm('ȷʵҪͨ�������');"""%>><img src="image/icon_reply.gif" />[<%if clng(guestflag and 16)<>0 then response.write "ͨ����˲�"%>�ظ�����]</a>
			<a class="tool" href="admin_del.asp<%=param_url%>" title="ɾ������(�����ظ�)"<%if DelTip=true then Response.Write " onclick=""return confirm('ȷʵҪɾ��������');"""%>><img src="image/icon_del.gif" />[ɾ������]</a>
			<%if clng(p_rs.Fields("replied") AND 1)<>0 then %><a class="tool" href="admin_delreply.asp<%=param_url%>" title="ɾ���ظ�"<%if DelReTip=true then Response.Write " onclick=""return confirm('ȷʵҪɾ���ظ���');"""%>><img src="image/icon_delreply.gif" />[ɾ���ظ�]</a><%end if%>
		</div>
	</div>
	<div class="group">
		<div class="name">���Կ��ƣ�</div>
		<div class="tools">
			<%if p_rs.Fields("parent_id")=0 then%><a class="tool" href="admin_lock2top.asp<%=param_url%>" title="������ʼ����ʾ�ڶ���"<%if Lock2TopTip=true then Response.Write " onclick=""return confirm('ȷʵҪ�ö�������');"""%>><img src="image/icon_toplocked.gif" />[�ö�����]</a><%end if%>
			<%if p_rs.Fields("parent_id")<0 then%><a class="tool" href="admin_unlock2top.asp<%=param_url%>" title="ȡ�������ö�"<%if Lock2TopTip=true then Response.Write " onclick=""return confirm('ȷʵҪ�ö�������');"""%>><img src="image/icon_untoplocked.gif" />[ȡ���ö�]</a><%end if%>
			<a class="tool" href="admin_bring2top.asp<%=param_url%>" title="��ǰ���Ե���ǰ"<%if Bring2TopTip=true then Response.Write " onclick=""return confirm('ȷʵҪ��ǰ������');"""%>><img src="image/icon_top.gif" />[��ǰ����]</a>
			<%if p_rs.Fields("parent_id")<0 or p_rs.Fields("logdate")<>p_rs.Fields("lastupdated") then%><a class="tool" href="admin_reorder.asp<%=param_url%>" title="ʹ���Իָ���ԭʼ����λ��"<%if ReorderTip=true then Response.Write " onclick=""return confirm('ȷʵҪ��������˳����');"""%>><img src="image/icon_reorder.gif" />[����˳��]</a><%end if%>
			<%if clng(p_rs.Fields("guestflag") and 512)=0 then%><a class="tool" href="admin_lockreply.asp<%=param_url%>" title="�����ÿͻظ�"><img src="image/icon_lockreply.gif" />[�����ظ�]</a><%end if%>
			<%if clng(p_rs.Fields("guestflag") and 512)<>0 then%><a class="tool" href="admin_unlockreply.asp<%=param_url%>" title="�����ÿͻظ�"><img src="image/icon_reply.gif" />[�����ظ�]</a><%end if%>
		</div>
	</div>
	<%if clng(guestflag and 952)<>0 OR p_rs.Fields("parent_id")<0 then%>
	<div class="group">
		<div class="name">����״̬��</div>
		<div class="tools">
			<%if clng(guestflag and 16)<>0 then%><span class="tool"><img src="image/icon_wait2pass.gif" />�ȴ����</span><%end if%>
			<%if clng(guestflag and 32)<>0 then%><span class="tool"><img src="image/icon_whisper.gif" />���Ļ�<%if clng(guestflag and 64)<>0 then response.write ",�Ѽ���"%></span><%end if%>
			<%if p_rs.Fields("parent_id")<0 then%><span class="tool"><img src="image/icon_toplocked.gif" />�������ö�</span><%end if%>
			<%if clng(guestflag and 512)<>0 and p_rs.Fields("parent_id")<=0 then%><span class="tool"><img src="image/icon_lockreply.gif" />�ظ�������</span><%end if%>
			<%if clng(guestflag and 256)<>0 then%><span class="tool"><img src="image/icon_hidecontact.gif" />��ϵ������</span><%end if%>
			<%if clng(guestflag and 40)=8 then%><span class="tool"><img src="image/icon_hide.gif" />����������</span><%end if%>
			<%if clng(guestflag and 128)<>0 then%><span class="tool"><img src="image/icon_mail.gif" />�ظ�֪ͨ<%if MailReplyInform=false then response.write ",�ѽ���"%></span><%end if%>
		</div>
	</div>
	<%end if%>
</div>