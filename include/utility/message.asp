<%
sub get_divided_page(byref cn,byref rs,byval pk,byval countsql,byval sql,byval keyword,byval request_page,byval ItemsEachPage,byref ItemsCount,byref PagesCount,byref CurrentItemsCount,byref ipage)
	if clng(ItemsEachPage)>0 then
		ItemsEachPage=clng(ItemsEachPage)
	else
		ItemsEachPage=10
	end if

	rs.Open countsql,cn,0,1,1
	ItemsCount=rs(0)
	rs.Close

	if ItemsCount mod ItemsEachPage <>0 then
		PagesCount=ItemsCount \ ItemsEachPage +1
	else
		PagesCount=ItemsCount \ ItemsEachPage
	end if

	if isnumeric(request_page) and request_page<>"" then
		if request_page>2147483647 then
			ipage=PagesCount
		elseif clng(request_page)>PagesCount then
			ipage=PagesCount
		elseif clng(request_page)>0 then
			ipage=clng(request_page)
		else
			ipage=1
		end if
	else
		ipage=1
	end if

	if ipage=PagesCount then
		CurrentItemsCount=ItemsCount - ItemsEachPage*(ipage -1)
	else
		CurrentItemsCount=ItemsEachPage
	end if

	if ItemsCount>0 then
		sql_from=mid(sql,instr(1,sql,"FROM",1))

		sql_innerfields=replace(replace(keyword," INC","")," DEC","")
		if instr(sql_innerfields,pk & ",")=0 AND instr(sql_innerfields,"," & pk)=0 AND sql_innerfields<>pk then sql_innerfields=sql_innerfields & "," & pk

		where_and=""
		if instr(1,sql,"where",1)>0 then
			where_and=" AND "
		else
			where_and=" WHERE"
		end if

		rs.Open Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sql_common_dividepage _
			,"{0}",CurrentItemsCount) _
			,"{1}",ItemsEachPage*(ipage-1)+CurrentItemsCount) _
			,"{2}",sql_innerfields) _
			,"{3}",sql_from) _
			,"{4}",replace(replace(keyword,"INC","ASC"),"DEC","DESC")) _
			,"{5}",replace(replace(keyword,"INC","DESC"),"DEC","ASC")) _
			,"{6}",sql) _
			,"{7}",where_and) _
			,"{8}",pk) _
		,cn,0,1,1
	end if
end sub

sub show_page_list(byval CurPage,byval PagesCount,byval filename,byval pagetitle,byval param)
	Dim arr_param_names, arr_param_values, param_name, param_value
	arr_param_names=split(param,",")
	ReDim arr_param_values(UBound(arr_param_names))
	Dim url_param
	url_param=""
	for i=0 to ubound(arr_param_names)
		param_name=arr_param_names(i)
		if(param_name)<>"" then
			param_value=Request(param_name)
			arr_param_values(i)=param_value
			if url_param<>"" then url_param = url_param & "&"
			url_param = url_param & param_name & "=" & Server.UrlEncode(param_value)
		end if
	next

	if ShowAdvPageList then
		txt_align="center"
		if CurPage-Int(AdvPageListCount/2)<1 then start_page=1 else start_page=CurPage-Int(AdvPageListCount/2)
		if start_page+AdvPageListCount-1>PagesCount then end_page=PagesCount else end_page=start_page+AdvPageListCount-1
		if end_page-start_page+1<AdvPageListCount and start_page>1 then
			if end_page-AdvPageListCount+1>1 then
				start_page=end_page-AdvPageListCount+1
			else
				start_page=1
			end if
		end if
	else
		txt_align="left"
		start_page=1
		end_page=PagesCount
	end if

	'计算各分页控件目标页号
	if ShowAdvPageList and PagesCount>AdvPageListCount then
		first_page_no=1
		last_page_no=PagesCount
		if CurPage+1>PagesCount then next_page_no=PagesCount else next_page_no=CurPage+1
		if CurPage-1<1 then prev_page_no=1 else prev_page_no=CurPage-1
		if CurPage+AdvPageListCount>PagesCount then largenext_page_no=PagesCount else largenext_page_no=CurPage+AdvPageListCount
		if CurPage-AdvPageListCount<1 then largeprev_page_no=1 else largeprev_page_no=CurPage-AdvPageListCount

		str_first_page=		"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &first_page_no& param & """><img src=""asset/image/icon_page_first.gif"" title=""第" &first_page_no& "页"" /></a>"
		str_largeprev_page=	"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &largeprev_page_no& param & """><img src=""asset/image/icon_page_largeprev.gif"" title=""第" &largeprev_page_no& "页"" /></a>"
		str_prev_page=		"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &prev_page_no& param & """><img src=""asset/image/icon_page_prev.gif"" title=""第" &prev_page_no& "页"" /></a>"

		str_last_page=		"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &last_page_no& param & """><img src=""asset/image/icon_page_last.gif"" title=""第" &last_page_no& "页"" /></a>"
		str_largenext_page=	"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &largenext_page_no& param & """><img src=""asset/image/icon_page_largenext.gif"" title=""第" &largenext_page_no& "页"" /></a>"
		str_next_page=		"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &next_page_no& param & """><img src=""asset/image/icon_page_next.gif"" title=""第" &next_page_no& "页"" /></a>"

		str_first2_page=	"<a class=""js-page-control"" name=""js-page-control"" onmouseup=""stop_refresh_pagenum && stop_refresh_pagenum();"" onmousedown=""stop_refresh_pagenum && stop_refresh_pagenum(); refresh_pagenum && refresh_pagenum(" &(-PagesCount+1)& ");"" title=""卷至首页""><img src=""asset/image/icon_page_first2.gif"" /></a>"
		str_largeprev2_page="<a class=""js-page-control"" name=""js-page-control"" onmouseup=""stop_refresh_pagenum && stop_refresh_pagenum();"" onmousedown=""stop_refresh_pagenum && stop_refresh_pagenum(); refresh_pagenum && refresh_pagenum(" &(-AdvPageListCount)& ");"" title=""上卷" &AdvPageListCount& "页""><img src=""asset/image/icon_page_largeprev2.gif"" /></a>"
		str_prev2_page=		"<a class=""js-page-control"" name=""js-page-control"" onmouseup=""stop_refresh_pagenum && stop_refresh_pagenum();"" onmousedown=""stop_refresh_pagenum && stop_refresh_pagenum(); refresh_pagenum && refresh_pagenum(" &(-1)& ");"" title=""上卷1页""><img src=""asset/image/icon_page_prev2.gif"" /></a>"

		str_last2_page=		"<a class=""js-page-control"" name=""js-page-control"" onmouseup=""stop_refresh_pagenum && stop_refresh_pagenum();"" onmousedown=""stop_refresh_pagenum && stop_refresh_pagenum(); refresh_pagenum && refresh_pagenum(" &(PagesCount-1)& ");"" title=""卷至末页""><img src=""asset/image/icon_page_last2.gif"" /></a>"
		str_largenext2_page="<a class=""js-page-control"" name=""js-page-control"" onmouseup=""stop_refresh_pagenum && stop_refresh_pagenum();"" onmousedown=""stop_refresh_pagenum && stop_refresh_pagenum(); refresh_pagenum && refresh_pagenum(" &(AdvPageListCount)& ");"" title=""下卷" &AdvPageListCount& "页""><img src=""asset/image/icon_page_largenext2.gif"" /></a>"
		str_next2_page=		"<a class=""js-page-control"" name=""js-page-control"" onmouseup=""stop_refresh_pagenum && stop_refresh_pagenum();"" onmousedown=""stop_refresh_pagenum && stop_refresh_pagenum(); refresh_pagenum && refresh_pagenum(" &(1)& ");"" title=""下卷1页""><img src=""asset/image/icon_page_next2.gif"" /></a>"
	end if%>

	<div class="region page-list">
		<h3 class="title"><%=pagetitle%></h3>
		<div class="content">
			<%if ShowAdvPageList and PagesCount>AdvPageListCount then%>
				<div class="nav backward-nav"><%=str_first_page & str_largeprev_page & str_prev_page & str_first2_page & str_largeprev2_page & str_prev2_page%></div>
				<div class="nav forward-nav"><%=str_next_page & str_largenext_page & str_last_page & str_next2_page & str_largenext2_page & str_last2_page%></div>
			<%end if%>
			<form method="get" action="<%=filename%>">
			<div class="pagenum-list">
				<%for j=start_page to end_page%><a name="pagenum" class="pagenum<%if j=CurPage then Response.Write " pagenum-current"%>" href="<%=filename%>?page=<%=j & "&" & url_param%>"><%=j%></a><%next%>
			</div>
			<div class="goto">(共<%=PagesCount%>页)　转到页数<input type="text" name="page" class="page" maxlength="10" /> <input type="submit" class="submit" value="GO" /></div>
			<%
			for i=0 to ubound(arr_param_names)
				param_name=arr_param_names(i)
				if param_name<>"" then
				param_value=arr_param_values(i)
				%><input type="hidden" name="<%=param_name%>" value="<%=param_value%>" /><%
				end if
			next
			%>
			</form>
		</div>
	</div>
	<script type="text/javascript">
		var ghandle;
		var PagesCount=<%=PagesCount%>;
		var PageListCount=<%=AdvPageListCount%>;
		var CurrentPage=<%=CurPage%>;
	</script>
	<script type="text/javascript" src="asset/js/pagecontrols.js" defer="defer" async="async"></script>
<%end sub


dim admin_name,admin_faceid,admin_faceurl,admin_email,admin_qqid,admin_msnid,admin_homepage,p_rs

sub getadmininfo()
	set rs3=server.CreateObject("ADODB.Recordset")
	rs3.open sql_common2_getadmininfo,cn,0,1,1

	admin_name="" & rs3.fields("name") & ""
	admin_faceid="" & rs3.fields("faceid") & ""
	admin_faceurl="" & rs3.fields("faceurl") & ""
	admin_email="" & rs3.fields("email") & ""
	admin_qqid="" & rs3.fields("qqid") & ""
	admin_msnid="" & rs3.fields("msnid") & ""
	admin_homepage="" & rs3.fields("homepage") & ""

	rs3.close
	set rs3=nothing

	if clng(admin_faceid)<>0 and admin_faceurl="" then admin_faceurl="asset/face/" & admin_faceid  & ".gif"
end sub

function GetHiddenWordCondition()
	dim t_condition
	t_condition=""
	if HideHidden then t_condition=t_condition & sql_condition_hidehidden
	if HideAudit then t_condition=t_condition & sql_condition_hideaudit
	if HideWhisper then t_condition=t_condition & sql_condition_hidewhisper

	GetHiddenWordCondition=t_condition
end function

sub showAdminIcons()
	if admin_email<>"" then%><a class="icon" href="mailto:<%=admin_email%>" title="版主邮箱：<%=admin_email%>"><img src="asset/image/icon_mail.gif"/></a><%end if
	if admin_qqid<>"" then%><span class="icon" title="版主QQ：<%=admin_qqid%>"><img src="asset/image/icon_qq.gif"/></span><%end if
	if admin_msnid<>"" then%><span class="icon"><img src="asset/image/icon_skype.gif" alt="版主Skype：<%=admin_msnid%>"/></span><%end if
	if admin_homepage<>"" then%><a class="icon" href="<%=admin_homepage%>" target="_blank" title="版主主页：<%=admin_homepage%>"><img src="asset/image/icon_homepage.gif"/></a><%end if
end sub

sub showAdminMessageTools(byref rs)
param_url="?rootid=" & rs.Fields("root_id") & "&id=" & rs.Fields("id") & "&page=" & ipage
if request("type")<>"" and request("searchtxt")<>"" then
	param_url=param_url & "&type=" & server.URLEncode(request("type")) & "&searchtxt=" & server.URLEncode(request("searchtxt"))
end if
param_url=Server.HtmlEncode(param_url)
%>

<div class="admin-message-tools">
	<div class="group">
		<div class="name">选定留言：</div>
		<div class="tools">
			<span class="tool"><input type="checkbox" name="seltodel" class="seltodel checkbox" id="c<%=rs.Fields("id")%>" value="<%=rs("id")%>" /><label for="c<%=rs.Fields("id")%>">(选定)</label></span>
		</div>
	</div>
	<div class="group">
		<div class="name">访客留言：</div>
		<div class="tools">
			<%if CBool(guestflag AND 16) then%><a class="tool" href="admin_passaudit.asp<%=param_url%>" title="通过审核"<%if PassAuditTip then response.Write " onclick=""return confirm('确实要通过审核吗？');"""%>><img src="asset/image/icon_pass.gif" />[通过审核]</a><%end if%>
			<a class="tool" href="admin_edit.asp<%=param_url%>" title="编辑访客留言"><img src="asset/image/icon_edit.gif" />[编辑留言]</a>
			<%if CBool(guestflag AND 32) then%><a class="tool" href="admin_pubwhisper.asp<%=param_url%>" title="公开悄悄话"<%if PubWhisperTip then response.Write " onclick=""return confirm('确实要公开悄悄话吗？');"""%>><img src="asset/image/icon_pub.gif" />[公开悄悄话]</a><%end if%>
			<%if Not CBool(guestflag AND 256) then%><a class="tool" href="admin_hidecontact.asp<%=param_url%>" title="隐藏访客联系方式"><img src="asset/image/icon_hidecontact.gif" />[隐藏联系]</a><%end if%>
			<%if CBool(guestflag AND 256) then%><a class="tool" href="admin_unhidecontact.asp<%=param_url%>" title="公开访客联系方式"><img src="asset/image/icon_unhidecontact.gif" />[公开联系]</a><%end if%>
			<%if Not CBool(guestflag AND 40) then%><a class="tool" href="admin_hideword.asp<%=param_url%>" title="隐藏访客留言内容"><img src="asset/image/icon_hide.gif" />[隐藏内容]</a><%end if%>
			<%if (guestflag AND 40)=8 then%><a class="tool" href="admin_unhideword.asp<%=param_url%>" title="公开访客留言内容"><img src="asset/image/icon_unhide.gif" />[公开内容]</a><%end if%>
		</div>
	</div>
	<div class="group">
		<div class="name">版主功能：</div>
		<div class="tools">
			<a class="tool" href="admin_reply.asp<%=param_url%>" title="<%if CBool(guestflag AND 16) then response.write "通过审核并"%>回复此留言"<%if CBool(guestflag AND 16) and PassAuditTip then response.Write " onclick=""return confirm('确实要通过审核吗？');"""%>><img src="asset/image/icon_reply.gif" />[<%if CBool(guestflag AND 16) then response.write "通过审核并"%>回复留言]</a>
			<a class="tool" href="admin_del.asp<%=param_url%>" title="删除留言(包括回复)"<%if DelTip then Response.Write " onclick=""return confirm('确实要删除留言吗？');"""%>><img src="asset/image/icon_del.gif" />[删除留言]</a>
			<%if CBool(rs.Fields("replied") AND 1) then %><a class="tool" href="admin_delreply.asp<%=param_url%>" title="删除回复"<%if DelReTip then Response.Write " onclick=""return confirm('确实要删除回复吗？');"""%>><img src="asset/image/icon_delreply.gif" />[删除回复]</a><%end if%>
		</div>
	</div>
	<%if rs.Fields("parent_id")<=0 then%>
	<div class="group">
		<div class="name">留言控制：</div>
		<div class="tools">
			<%if rs.Fields("parent_id")=0 then%><a class="tool" href="admin_lock2top.asp<%=param_url%>" title="将留言始终显示在顶端"<%if Lock2TopTip then Response.Write " onclick=""return confirm('确实要置顶留言吗？');"""%>><img src="asset/image/icon_toplocked.gif" />[置顶留言]</a><%end if%>
			<%if rs.Fields("parent_id")<0 then%><a class="tool" href="admin_unlock2top.asp<%=param_url%>" title="取消留言置顶"<%if Lock2TopTip then Response.Write " onclick=""return confirm('确实要置顶留言吗？');"""%>><img src="asset/image/icon_untoplocked.gif" />[取消置顶]</a><%end if%>
			<a class="tool" href="admin_bring2top.asp<%=param_url%>" title="提前留言到最前"<%if Bring2TopTip then Response.Write " onclick=""return confirm('确实要提前留言吗？');"""%>><img src="asset/image/icon_top.gif" />[提前留言]</a>
			<%if rs.Fields("parent_id")<0 or rs.Fields("logdate")<>rs.Fields("lastupdated") then%><a class="tool" href="admin_reorder.asp<%=param_url%>" title="使留言恢复到原始排序位置"<%if ReorderTip then Response.Write " onclick=""return confirm('确实要重置留言顺序吗？');"""%>><img src="asset/image/icon_reorder.gif" />[重置顺序]</a><%end if%>
			<%if CBool(guestflag AND 512) then%>
				<a class="tool" href="admin_unlockreply.asp<%=param_url%>" title="允许访客回复"><img src="asset/image/icon_reply.gif" />[允许回复]</a>
			<%else%>
				<a class="tool" href="admin_lockreply.asp<%=param_url%>" title="锁定访客回复"><img src="asset/image/icon_lockreply.gif" />[锁定回复]</a>
			<%end if%>
		</div>
	</div>
	<%end if%>
	<%if CBool(guestflag AND 952) OR rs.Fields("parent_id")<0 then%>
	<div class="group">
		<div class="name">其它状态：</div>
		<div class="tools">
			<%if CBool(guestflag AND 16) then%><span class="tool"><img src="asset/image/icon_wait2pass.gif" />等待审核</span><%end if%>
			<%if CBool(guestflag AND 32) then%><span class="tool"><img src="asset/image/icon_whisper.gif" />悄悄话<%if CBool(guestflag AND 64) then response.write ",已加密"%></span><%end if%>
			<%if rs.Fields("parent_id")<0 then%><span class="tool"><img src="asset/image/icon_toplocked.gif" />留言已置顶</span><%end if%>
			<%if CBool(guestflag AND 512) and rs.Fields("parent_id")<=0 then%><span class="tool"><img src="asset/image/icon_lockreply.gif" />回复已锁定</span><%end if%>
			<%if CBool(guestflag AND 256) then%><span class="tool"><img src="asset/image/icon_hidecontact.gif" />联系已隐藏</span><%end if%>
			<%if (guestflag AND 40)=8 then%><span class="tool"><img src="asset/image/icon_hide.gif" />内容已隐藏</span><%end if%>
			<%if CBool(guestflag AND 128) then%><span class="tool"><img src="asset/image/icon_mail.gif" />回复通知<%if Not MailReplyInform then response.write ",已禁用"%></span><%end if%>
		</div>
	</div>
	<%end if%>
</div><%
end sub

sub showGuestMessageTools(byval follow_id,byval parent_id,byval show_reply)
dim url
url="leaveword.asp?follow=" & follow_id
if left(pagename,8)="showword" then url=url & "&return=showword"
url=Server.HTMLEncode(url)%>
<div class="guest-message-tools">
	<%if parent_id<0 then%><span class="tool"><img src="asset/image/icon_toplocked.gif"/>(置顶)</span><%end if%>
	<%if show_reply then%><span class="tool"><a href="<%=url%>"><img src="asset/image/icon_reply.gif"/>[回复]</a></span><%end if%>
</div>
<%end sub

sub showGuestIcons(rs)
	Dim email,qqid,msnid,homepage,ipv4addr,ipv6addr,originalipv4,originalipv6

	email=rs.Fields("email")
	if email<>"" then%><a class="icon" href="mailto:<%=email%>" title="作者邮箱：<%=email%>"><img src="asset/image/icon_mail.gif"/></a><%end if
	qqid=rs.Fields("qqid")
	if qqid<>"" then%><span class="icon" title="作者QQ：<%=qqid%>"><img src="asset/image/icon_qq.gif"/></span><%end if
	msnid=rs.Fields("msnid")
	if msnid<>"" then%><span class="icon" title="作者Skype：<%=msnid%>"><img src="asset/image/icon_skype.gif"/></span><%end if
	homepage=rs.Fields("homepage")
	if homepage<>"" then%><a class="icon" href="<%=homepage%>" target="_blank" title="作者主页：<%=homepage%>"><img src="asset/image/icon_homepage.gif" /></a><%end if

	if inAdminPage then
		ipv4addr=rs.Fields("ipv4addr")
		if AdminShowIPv4>0 and ipv4addr<>"" then
			ipv4addr=GetIPv4WithMask(ipv4addr,AdminShowIPv4)%><span class="icon-entry"><span class="icon" title="IP：<%=ipv4addr%>"><img src="asset/image/icon_ip.gif"/></span></span><%end if
		ipv6addr=rs.Fields("ipv6addr")
		if AdminShowIPv6>0 and ipv6addr<>"" then
			ipv6addr=GetIPv6WithMask(ipv6addr,AdminShowIPv6)%><span class="icon-entry"><span class="icon" title="IP：<%=ipv6addr%>"><img src="asset/image/icon_ip.gif"/></span></span><%end if
		originalipv4=rs.Fields("originalipv4")
		if AdminShowOriginalIPv4>0 and originalipv4<>"" then
			originalipv4=GetIPv4WithMask(originalipv4,AdminShowOriginalIPv4)%><span class="icon-entry"><span class="icon" title="原始IP：<%=originalipv4%>"><img src="asset/image/icon_ip2.gif"/></span></span><%end if
		originalipv6=rs.Fields("originalipv6")
		if AdminShowOriginalIPv6>0 and originalipv6<>"" then
			originalipv6=GetIPv6WithMask(originalipv6,AdminShowOriginalIPv6)%><span class="icon-entry"><span class="icon" title="原始IP：<%=originalipv6%>"><img src="asset/image/icon_ip2.gif"/></span></span><%end if
	else
		ipv4addr=rs.Fields("ipv4addr")
		if ShowIPv4>0 and ipv4addr<>"" then
			ipv4addr=GetIPv4WithMask(ipv4addr,ShowIPv4)%><span class="icon" title="IP：<%=ipv4addr%>"><img src="asset/image/icon_ip.gif"/></span><%end if
		ipv6addr=rs.Fields("ipv6addr")
		if ShowIPv6>0 and ipv6addr<>"" then
			ipv6addr=GetIPv6WithMask(ipv6addr,ShowIPv6)%><span class="icon" title="IP：<%=ipv6addr%>"><img src="asset/image/icon_ip.gif"/></span><%end if
	end if
end sub

sub inneradminreply(byref rs2)%>
<div class="message inner-message admin-message">
	<div class="summary">
		<div class="name">版主<%if admin_name<>"" then response.write "(" & admin_name & ")"%>回复：</div>
		<div class="date">(<%=UTCToDisplayTime(rs2("replydate"))%>)</div>
		<div class="icons"><%showAdminIcons()%></div>
	</div>
	<%if admin_faceurl<>"" then%><img class="face" src="<%=admin_faceurl%>"/><%end if%>
	<div class="words">
		<%if encrypted and pagename<>"showword" and Not inAdminPage then%>
			<span class="inner-hint"><img src="asset/image/icon_key.gif"/>(需要预设的密码才能查看...)[<a href="showword.asp?id=<%=rs2("id")%>">点击这里验证...</a>]</span>
		<%else
			reply_htmlflag=rs2("htmlflag")
			reply_txt=rs2("reinfo")

			convertstr reply_txt,reply_htmlflag,true
			Response.Write reply_txt
		end if%>
	</div>
</div>
<%end sub

sub outeradminreply(byref rs2)%>
<div class="message outer-message admin-message">
	<h2 class="title">[版主回复：]</h2>
	<div class="info">
		<%if admin_faceurl<>"" then%><img class="face" src="<%=admin_faceurl%>"/><%end if%>
		<%if admin_email<>"" or admin_qqid<>"" or admin_msnid<>"" or admin_homepage<>"" then%>
			<div class="icons"><%showAdminIcons()%></div>
		<%end if%>
		<div class="date"><%=UTCToDisplayTime(rs2("replydate"))%></div>
	</div>
	<div class="detail">
		<div class="words">
			<%if encrypted and pagename<>"showword" and Not inAdminPage then%>
				<span class="inner-hint"><img src="asset/image/icon_key.gif"/>(需要预设的密码才能查看...)[<a href="showword.asp?id=<%=rs2("id")%>">点击这里验证...</a>]</span>
			<%else
				reply_htmlflag=rs2("htmlflag")
				reply_txt=rs2("reinfo")

				convertstr reply_txt,reply_htmlflag,true
				Response.Write reply_txt
			end if%>
		</div>
	</div>
</div>
<%end sub

sub inneraudit()%>
<div class="message inner-message guest-message auditing-message">
	<div class="summary">
		<h2 class="title">(留言待审核...)</h2>
	</div>
	<div class="words">
		<span class="inner-hint"><img src="asset/image/icon_wait2pass.gif"/>(留言待审核...)</span>
	</div>
</div>
<%end sub

sub outeraudit(t_rs)%>
<div class="message outer-message guest-message auditing-message">
	<div class="info">
		<%fid=t_rs("faceid")
		if isnumeric(fid) and StatusShowHead then
			if fid>=1 and fid<=FaceCount then%>
				<img class="face" src="asset/face/<%=fid%>.gif" />
			<%end if
		end if%>

		<div class="date"><%=UTCToDisplayTime(t_rs("logdate"))%></div>
	</div>
	<div class="detail">
		<h2 class="title">(留言待审核...)</h2>
		<%if rs.Fields("parent_id")<=0 and Not inAdminPage then showGuestMessageTools rs.Fields("id"),rs.Fields("parent_id"),StatusWrite and StatusGuestReply and Not CBool(rs.Fields("guestflag") AND 512)%>
		<div class="words">
			<span class="inner-hint"><img src="asset/image/icon_wait2pass.gif" />(留言待审核...)</span>
		</div>
	</div>
</div>
<%end sub

sub innerword(byref t_rs)%>
<div class="message inner-message guest-message<%if isauditting then%> auditting-message<%end if%>">
	<div class="summary">
		<div class="name"><%=t_rs("name")%>：</div>
		<div class="date">(<%=UTCToDisplayTime(t_rs("logdate"))%>)</div>
			<%if (Not iswhisper and Not CBool(guestflag AND 256)) or (pagename="showword" and needverify) or inAdminPage then%>
				<div class="icons"><%showGuestIcons(t_rs)%></div>
			<%end if%>
		<h2 class="title"><%if iswhisper and pagename<>"showword" and Not inAdminPage then response.write "(给版主的悄悄话...)" else response.write t_rs("title")%></h2>
	</div>
	<%if StatusShowHead and t_rs.Fields("faceid")>=1 and t_rs.Fields("faceid")<=FaceCount then%><img class="face" src="asset/face/<%=t_rs.Fields("faceid")%>.gif"/><%end if%>
	<div class="words">
		<%if iswhisper and pagename<>"showword" and Not inAdminPage then%>
			<span class="inner-hint"><img src="asset/image/icon_whisper.gif"/>(给版主的悄悄话...)</span>
		<%elseif ishidden and Not inAdminPage then%>
			<span class="inner-hint"><img src="asset/image/icon_hide.gif"/>(留言被管理员隐藏...)</span>
		<%else
			guest_txt="" & t_rs("article") & ""
			if inAdminPage and AdminViewCode then
				guest_txt=server.htmlEncode(guest_txt)
			else
				convertstr guest_txt,guestflag,false
			end if

			if guest_txt<>"" then
				Response.Write guest_txt
			else
				Response.Write "(无内容)"
			end if
		end if%>
	</div>
</div>
<%end sub

sub outerword(byref rs)%>
<div class="message outer-message guest-message<%if isauditting then%> auditting-message<%end if%>">
	<div class="info">
		<div class="name"><%=rs("name")%></div>

		<%fid=rs("faceid")
		if isnumeric(fid) and StatusShowHead then
			if fid>=1 and fid<=FaceCount then%>
				<img class="face" src="asset/face/<%=fid%>.gif" />
			<%end if
		end if%>

		<%if (Not iswhisper and Not CBool(guestflag AND 256)) or (pagename="showword" and needverify) or inAdminPage then%>
			<div class="icons"><%showGuestIcons(rs)%></div>
		<%end if%>

		<div class="date"><%=UTCToDisplayTime(rs("logdate"))%></div>
	</div>
	<div class="detail">
		<h2 class="title"><%if iswhisper and pagename<>"showword" and Not inAdminPage then response.write "(给版主的悄悄话...)" else response.write rs("title")%></h2>
		<%if rs.Fields("parent_id")<=0 and Not inAdminPage then showGuestMessageTools rs.Fields("id"),rs.Fields("parent_id"),StatusWrite and StatusGuestReply and Not CBool(rs.Fields("guestflag") and 512)%>
		<div class="words">
			<%if iswhisper and pagename<>"showword" and Not inAdminPage then%>
				<span class="inner-hint"><img src="asset/image/icon_whisper.gif"/>(给版主的悄悄话...)</span>
			<%elseif ishidden and Not inAdminPage then%>
				<span class="inner-hint"><img src="asset/image/icon_hide.gif"/>(留言被管理员隐藏...)</span>
			<%else
				guest_txt="" & rs("article") & ""
				if inAdminPage and AdminViewCode then
					guest_txt=server.htmlEncode(guest_txt)
				else
					convertstr guest_txt,guestflag,false
				end if

				if guest_txt<>"" then
					Response.Write guest_txt
				else
					Response.Write "(无内容)"
				end if
			end if
			%>
		</div>

		<%
		if CBool(rs.Fields("replied") AND 1) and ReplyInWord then inneradminreply(rs)	'内嵌

		if inAdminPage and ReplyInWord then
			Call showAdminMessageTools(rs)
		end if

		if rs.Fields("parent_id")<=0 and CBool(rs.Fields("replied") AND 2) and (Not encrypted or pagename="showword" or inAdminPage) and ReplyInWord then
			dim hidden_condition
			hidden_condition=""
			if Not inAdminPage then hidden_condition=GetHiddenWordCondition()

			set rs1=Server.CreateObject("ADODB.Recordset")
			rs1.Open Replace(Replace(sql_common2_guestreply,"{0}",rs.Fields("id")),"{1}",hidden_condition),cn,0,1,1
			while not rs1.eof
				guestflag=rs1("guestflag")
				isauditting=CBool(guestflag AND 16)
				ishidden=((guestflag AND 40)=8)
				iswhisper=(CBool(guestflag AND 32))
				encrypted=(clng(guestflag and 96)=96)

				if isauditting and Not inAdminPage then
					inneraudit()
				else
					innerword(rs1)
					if rs1.Fields("replied") then inneradminreply(rs1)
				end if

				if inAdminPage then
					Call showAdminMessageTools(rs1)
				end if
				rs1.movenext
			wend
			rs1.close
			set rs1=nothing
		end if
		%>
	</div>
</div>
<%end sub
%>