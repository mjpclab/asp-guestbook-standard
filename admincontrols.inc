<%
param_url="?rootid=" & p_rs.Fields("root_id") & "&id=" & p_rs.Fields("id") & "&page=" & ipage
if request("type")<>"" and request("searchtxt")<>"" then
	param_url=param_url & "&type=" & server.URLEncode(request("type")) & "&searchtxt=" & server.URLEncode(request("searchtxt"))
end if
param_url=Server.HtmlEncode(param_url)
%>
<tr>
	<td colspan="2" style="width:100%; border:solid <%=TableBorderColor%>; border-width:1px 0px 1px 0px; padding:3px; color:<%=TableContentColor%>; background-color:<%=TableContentBGC%>;">
		<table cellpadding="0" cellspacing="1" style="width:100%; border:0px; border-collapse:collapse;">
		<tr>
			<td style="width:80px;">
				选定留言：
			</td>
			<td>
				<input type="checkbox" name="seltodel" id="c<%=p_rs.Fields("id")%>" value="<%=p_rs("id")%>" /><label for="c<%=p_rs.Fields("id")%>">(选定)</label>
			</td>
		</tr>
		<tr>
			<td style="width:80px;">
				访客留言：
			</td>
			<td>
				<%if clng(guestflag and 16)<>0 then%><a href="admin_passaudit.asp<%=param_url%>" title="通过审核"<%if PassAuditTip=true then response.Write " onclick=""return confirm('确实要通过审核吗？');"""%>><img alt="" src="image/icon_pass.gif" class="imgicon" />[通过审核]</a> <%end if%>
				<a href="admin_edit.asp<%=param_url%>" title="编辑访客留言"><img alt="" src="image/icon_edit.gif" class="imgicon" />[编辑留言]</a> 
				<%if clng(guestflag and 32)<>0 then%><a href="admin_pubwhisper.asp<%=param_url%>" title="公开悄悄话"<%if PubWhisperTip=true then response.Write " onclick=""return confirm('确实要公开悄悄话吗？');"""%>><img alt="" src="image/icon_pub.gif" class="imgicon" />[公开悄悄话]</a>  <%end if%>
				<%if clng(guestflag and 256)=0 then%><a href="admin_hidecontact.asp<%=param_url%>" title="隐藏访客联系方式"><img alt="" src="image/icon_hidecontact.gif" class="imgicon" />[隐藏联系]</a>  <%end if%>
				<%if clng(guestflag and 256)<>0 then%><a href="admin_unhidecontact.asp<%=param_url%>" title="公开访客联系方式"><img alt="" src="image/icon_unhidecontact.gif" class="imgicon" />[公开联系]</a>  <%end if%>
				<%if clng(guestflag and 40)=0 then%><a href="admin_hideword.asp<%=param_url%>" title="隐藏访客留言内容"><img alt="" src="image/icon_hide.gif" class="imgicon" />[隐藏内容]</a>  <%end if%>
				<%if clng(guestflag and 40)=8 then%><a href="admin_unhideword.asp<%=param_url%>" title="公开访客留言内容"><img alt="" src="image/icon_unhide.gif" class="imgicon" />[公开内容]</a>  <%end if%>
			</td>
		</tr>
		<tr>
			<td style="width:80px;">
				版主功能：
			</td>
			<td>
				<a href="admin_reply.asp<%=param_url%>" title="<%if clng(guestflag and 16)<>0 then response.write "通过审核并"%>回复此留言"<%if clng(guestflag and 16)<>0 and PassAuditTip=true then response.Write " onclick=""return confirm('确实要通过审核吗？');"""%>><img alt="" src="image/icon_reply.gif" class="imgicon" />[<%if clng(guestflag and 16)<>0 then response.write "通过审核并"%>回复留言]</a> 
				<a href="admin_del.asp<%=param_url%>" title="删除留言(包括回复)"<%if DelTip=true then Response.Write " onclick=""return confirm('确实要删除留言吗？');"""%>><img alt="" src="image/icon_del.gif" class="imgicon" />[删除留言]</a> 
				<%if clng(p_rs.Fields("replied") AND 1)<>0 then %><a href="admin_delreply.asp<%=param_url%>" title="删除回复"<%if DelReTip=true then Response.Write " onclick=""return confirm('确实要删除回复吗？');"""%>><img alt="" src="image/icon_delreply.gif" class="imgicon" />[删除回复]</a> <%end if%>
			</td>
		</tr>
		<%if p_rs.Fields("parent_id")<=0 then%>
		<tr>
			<td style="width:80px;">
				留言控制：
			</td>
			<td>
				<%if p_rs.Fields("parent_id")=0 then%><a href="admin_lock2top.asp<%=param_url%>" title="将留言始终显示在顶端"<%if Lock2TopTip=true then Response.Write " onclick=""return confirm('确实要置顶留言吗？');"""%>><img alt="" src="image/icon_toplocked.gif" class="imgicon" />[置顶留言]</a> <%end if%>
				<%if p_rs.Fields("parent_id")<0 then%><a href="admin_unlock2top.asp<%=param_url%>" title="取消留言置顶"<%if Lock2TopTip=true then Response.Write " onclick=""return confirm('确实要置顶留言吗？');"""%>><img alt="" src="image/icon_untoplocked.gif" class="imgicon" />[取消置顶]</a> <%end if%>
				<a href="admin_bring2top.asp<%=param_url%>" title="提前留言到最前"<%if Bring2TopTip=true then Response.Write " onclick=""return confirm('确实要提前留言吗？');"""%>><img alt="" src="image/icon_top.gif" class="imgicon" />[提前留言]</a> 
				<%if p_rs.Fields("parent_id")<0 or p_rs.Fields("logdate")<>p_rs.Fields("lastupdated") then%><a href="admin_reorder.asp<%=param_url%>" title="使留言恢复到原始排序位置"<%if ReorderTip=true then Response.Write " onclick=""return confirm('确实要重置留言顺序吗？');"""%>><img alt="" src="image/icon_reorder.gif" class="imgicon" />[重置顺序]</a> <%end if%>
				<%if clng(p_rs.Fields("guestflag") and 512)=0 then%><a href="admin_lockreply.asp<%=param_url%>" title="锁定访客回复"><img alt="" src="image/icon_lockreply.gif" class="imgicon" />[锁定回复]</a> <%end if%>
				<%if clng(p_rs.Fields("guestflag") and 512)<>0 then%><a href="admin_unlockreply.asp<%=param_url%>" title="允许访客回复"><img alt="" src="image/icon_reply.gif" class="imgicon" />[允许回复]</a> <%end if%>
			</td>
		</tr>
		<%end if%>
		<%if clng(guestflag and 952)<>0 OR p_rs.Fields("parent_id")<0 then%>
		<tr>
			<td style="width:80px;">
				其它状态：
			</td>
			<td>
				<%if clng(guestflag and 16)<>0 then%><img alt="" src="image/icon_wait2pass.gif" class="imgicon" />等待审核  <%end if%>
				<%if clng(guestflag and 32)<>0 then%><img alt="" src="image/icon_whisper.gif" class="imgicon" />悄悄话<%if clng(guestflag and 64)<>0 then response.write ",已加密"%>  <%end if%>
				<%if p_rs.Fields("parent_id")<0 then%><img alt="" src="image/icon_toplocked.gif" class="imgicon" />留言已置顶  <%end if%>
				<%if clng(guestflag and 512)<>0 and p_rs.Fields("parent_id")<=0 then%><img alt="" src="image/icon_lockreply.gif" class="imgicon" />回复已锁定  <%end if%>
				<%if clng(guestflag and 256)<>0 then%><img alt="" src="image/icon_hidecontact.gif" class="imgicon" />联系已隐藏  <%end if%>
				<%if clng(guestflag and 40)=8 then%><img alt="" src="image/icon_hide.gif" class="imgicon" />内容已隐藏  <%end if%>
				<%if clng(guestflag and 128)<>0 then%><img alt="" src="image/icon_mail.gif" class="imgicon" />回复通知<%if MailReplyInform=false then response.write ",已禁用"%>  <%end if%>
			</td>
		</tr>
		<%end if%>
		</table>
	</td>
</tr>
