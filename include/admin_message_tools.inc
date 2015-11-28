<%
param_url="?rootid=" & p_rs.Fields("root_id") & "&id=" & p_rs.Fields("id") & "&page=" & ipage
if request("type")<>"" and request("searchtxt")<>"" then
	param_url=param_url & "&type=" & server.URLEncode(request("type")) & "&searchtxt=" & server.URLEncode(request("searchtxt"))
end if
param_url=Server.HtmlEncode(param_url)
%>

<div class="admin-message-tools">
	<div class="group">
		<div class="name">选定留言：</div>
		<div class="tools">
			<span class="tool"><input type="checkbox" name="seltodel" id="c<%=p_rs.Fields("id")%>" value="<%=p_rs("id")%>" /><label for="c<%=p_rs.Fields("id")%>">(选定)</label></span>
		</div>
	</div>
	<div class="group">
		<div class="name">访客留言：</div>
		<div class="tools">
			<%if clng(guestflag and 16)<>0 then%><a class="tool" href="admin_passaudit.asp<%=param_url%>" title="通过审核"<%if PassAuditTip=true then response.Write " onclick=""return confirm('确实要通过审核吗？');"""%>><img src="image/icon_pass.gif" />[通过审核]</a><%end if%>
			<a class="tool" href="admin_edit.asp<%=param_url%>" title="编辑访客留言"><img src="image/icon_edit.gif" />[编辑留言]</a>
			<%if clng(guestflag and 32)<>0 then%><a class="tool" href="admin_pubwhisper.asp<%=param_url%>" title="公开悄悄话"<%if PubWhisperTip=true then response.Write " onclick=""return confirm('确实要公开悄悄话吗？');"""%>><img src="image/icon_pub.gif" />[公开悄悄话]</a><%end if%>
			<%if clng(guestflag and 256)=0 then%><a class="tool" href="admin_hidecontact.asp<%=param_url%>" title="隐藏访客联系方式"><img src="image/icon_hidecontact.gif" />[隐藏联系]</a><%end if%>
			<%if clng(guestflag and 256)<>0 then%><a class="tool" href="admin_unhidecontact.asp<%=param_url%>" title="公开访客联系方式"><img src="image/icon_unhidecontact.gif" />[公开联系]</a><%end if%>
			<%if clng(guestflag and 40)=0 then%><a class="tool" href="admin_hideword.asp<%=param_url%>" title="隐藏访客留言内容"><img src="image/icon_hide.gif" />[隐藏内容]</a><%end if%>
			<%if clng(guestflag and 40)=8 then%><a class="tool" href="admin_unhideword.asp<%=param_url%>" title="公开访客留言内容"><img src="image/icon_unhide.gif" />[公开内容]</a><%end if%>
		</div>
	</div>
	<div class="group">
		<div class="name">版主功能：</div>
		<div class="tools">
			<a class="tool" href="admin_reply.asp<%=param_url%>" title="<%if clng(guestflag and 16)<>0 then response.write "通过审核并"%>回复此留言"<%if clng(guestflag and 16)<>0 and PassAuditTip=true then response.Write " onclick=""return confirm('确实要通过审核吗？');"""%>><img src="image/icon_reply.gif" />[<%if clng(guestflag and 16)<>0 then response.write "通过审核并"%>回复留言]</a>
			<a class="tool" href="admin_del.asp<%=param_url%>" title="删除留言(包括回复)"<%if DelTip=true then Response.Write " onclick=""return confirm('确实要删除留言吗？');"""%>><img src="image/icon_del.gif" />[删除留言]</a>
			<%if clng(p_rs.Fields("replied") AND 1)<>0 then %><a class="tool" href="admin_delreply.asp<%=param_url%>" title="删除回复"<%if DelReTip=true then Response.Write " onclick=""return confirm('确实要删除回复吗？');"""%>><img src="image/icon_delreply.gif" />[删除回复]</a><%end if%>
		</div>
	</div>
	<div class="group">
		<div class="name">留言控制：</div>
		<div class="tools">
			<%if p_rs.Fields("parent_id")=0 then%><a class="tool" href="admin_lock2top.asp<%=param_url%>" title="将留言始终显示在顶端"<%if Lock2TopTip=true then Response.Write " onclick=""return confirm('确实要置顶留言吗？');"""%>><img src="image/icon_toplocked.gif" />[置顶留言]</a><%end if%>
			<%if p_rs.Fields("parent_id")<0 then%><a class="tool" href="admin_unlock2top.asp<%=param_url%>" title="取消留言置顶"<%if Lock2TopTip=true then Response.Write " onclick=""return confirm('确实要置顶留言吗？');"""%>><img src="image/icon_untoplocked.gif" />[取消置顶]</a><%end if%>
			<a class="tool" href="admin_bring2top.asp<%=param_url%>" title="提前留言到最前"<%if Bring2TopTip=true then Response.Write " onclick=""return confirm('确实要提前留言吗？');"""%>><img src="image/icon_top.gif" />[提前留言]</a>
			<%if p_rs.Fields("parent_id")<0 or p_rs.Fields("logdate")<>p_rs.Fields("lastupdated") then%><a class="tool" href="admin_reorder.asp<%=param_url%>" title="使留言恢复到原始排序位置"<%if ReorderTip=true then Response.Write " onclick=""return confirm('确实要重置留言顺序吗？');"""%>><img src="image/icon_reorder.gif" />[重置顺序]</a><%end if%>
			<%if clng(p_rs.Fields("guestflag") and 512)=0 then%><a class="tool" href="admin_lockreply.asp<%=param_url%>" title="锁定访客回复"><img src="image/icon_lockreply.gif" />[锁定回复]</a><%end if%>
			<%if clng(p_rs.Fields("guestflag") and 512)<>0 then%><a class="tool" href="admin_unlockreply.asp<%=param_url%>" title="允许访客回复"><img src="image/icon_reply.gif" />[允许回复]</a><%end if%>
		</div>
	</div>
	<%if clng(guestflag and 952)<>0 OR p_rs.Fields("parent_id")<0 then%>
	<div class="group">
		<div class="name">其它状态：</div>
		<div class="tools">
			<%if clng(guestflag and 16)<>0 then%><span class="tool"><img src="image/icon_wait2pass.gif" />等待审核</span><%end if%>
			<%if clng(guestflag and 32)<>0 then%><span class="tool"><img src="image/icon_whisper.gif" />悄悄话<%if clng(guestflag and 64)<>0 then response.write ",已加密"%></span><%end if%>
			<%if p_rs.Fields("parent_id")<0 then%><span class="tool"><img src="image/icon_toplocked.gif" />留言已置顶</span><%end if%>
			<%if clng(guestflag and 512)<>0 and p_rs.Fields("parent_id")<=0 then%><span class="tool"><img src="image/icon_lockreply.gif" />回复已锁定</span><%end if%>
			<%if clng(guestflag and 256)<>0 then%><span class="tool"><img src="image/icon_hidecontact.gif" />联系已隐藏</span><%end if%>
			<%if clng(guestflag and 40)=8 then%><span class="tool"><img src="image/icon_hide.gif" />内容已隐藏</span><%end if%>
			<%if clng(guestflag and 128)<>0 then%><span class="tool"><img src="image/icon_mail.gif" />回复通知<%if MailReplyInform=false then response.write ",已禁用"%></span><%end if%>
		</div>
	</div>
	<%end if%>
</div>
