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
	if param<>"" then
		if left(param,1)<>"&" then param="&" & param
	end if

	arr_param=split(param,"&")

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

	'�������ҳ�ؼ�Ŀ��ҳ��
	if ShowAdvPageList then
		first_page_no=1
		last_page_no=PagesCount
		if CurPage+1>PagesCount then next_page_no=PagesCount else next_page_no=CurPage+1
		if CurPage-1<1 then prev_page_no=1 else prev_page_no=CurPage-1
		if CurPage+AdvPageListCount>PagesCount then largenext_page_no=PagesCount else largenext_page_no=CurPage+AdvPageListCount
		if CurPage-AdvPageListCount<1 then largeprev_page_no=1 else largeprev_page_no=CurPage-AdvPageListCount

		str_first_page=		"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &first_page_no& param & """><img src=""asset/image/icon_page_first.gif"" title=""��" &first_page_no& "ҳ"" /></a>"
		str_largeprev_page=	"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &largeprev_page_no& param & """><img src=""asset/image/icon_page_largeprev.gif"" title=""��" &largeprev_page_no& "ҳ"" /></a>"
		str_prev_page=		"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &prev_page_no& param & """><img src=""asset/image/icon_page_prev.gif"" title=""��" &prev_page_no& "ҳ"" /></a>"

		str_last_page=		"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &last_page_no& param & """><img src=""asset/image/icon_page_last.gif"" title=""��" &last_page_no& "ҳ"" /></a>"
		str_largenext_page=	"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &largenext_page_no& param & """><img src=""asset/image/icon_page_largenext.gif"" title=""��" &largenext_page_no& "ҳ"" /></a>"
		str_next_page=		"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &next_page_no& param & """><img src=""asset/image/icon_page_next.gif"" title=""��" &next_page_no& "ҳ"" /></a>"

		str_first2_page=	"<a class=""js-page-control"" name=""js-page-control""><img onmouseup=""stop_refresh_pagenum && stop_refresh_pagenum();"" onmousedown=""refresh_pagenum && refresh_pagenum(" &(-PagesCount+1)& ");"" src=""asset/image/icon_page_first2.gif"" title=""������ҳ"" /></a>"
		str_largeprev2_page="<a class=""js-page-control"" name=""js-page-control""><img onmouseup=""stop_refresh_pagenum && stop_refresh_pagenum();"" onmousedown=""refresh_pagenum && refresh_pagenum(" &(-AdvPageListCount)& ");"" src=""asset/image/icon_page_largeprev2.gif"" class=""pageicon"" title=""�Ͼ�" &AdvPageListCount& "ҳ"" /></a>"
		str_prev2_page=		"<a class=""js-page-control"" name=""js-page-control""><img onmouseup=""stop_refresh_pagenum && stop_refresh_pagenum();"" onmousedown=""refresh_pagenum && refresh_pagenum(" &(-1)& ");"" src=""asset/image/icon_page_prev2.gif"" class=""pageicon"" title=""�Ͼ�1ҳ"" /></a>"

		str_last2_page=		"<a class=""js-page-control"" name=""js-page-control""><img onmouseup=""stop_refresh_pagenum && stop_refresh_pagenum();"" onmousedown=""refresh_pagenum && refresh_pagenum(" &(PagesCount-1)& ");"" src=""asset/image/icon_page_last2.gif"" class=""pageicon"" title=""����ĩҳ"" /></a>"
		str_largenext2_page="<a class=""js-page-control"" name=""js-page-control""><img onmouseup=""stop_refresh_pagenum && stop_refresh_pagenum();"" onmousedown=""refresh_pagenum && refresh_pagenum(" &(AdvPageListCount)& ");"" src=""asset/image/icon_page_largenext2.gif"" class=""pageicon"" title=""�¾�" &AdvPageListCount& "ҳ"" /></a>"
		str_next2_page=		"<a class=""js-page-control"" name=""js-page-control""><img onmouseup=""stop_refresh_pagenum && stop_refresh_pagenum();"" onmousedown=""refresh_pagenum && refresh_pagenum(" &(1)& ");"" src=""asset/image/icon_page_next2.gif"" class=""pageicon"" title=""�¾�1ҳ"" /></a>"
	end if%>

	<div class="region page-list">
		<h3 class="title"><%=pagetitle%></h3>
		<div class="content">
			<%if ShowAdvPageList then%>
				<div class="nav backward-nav"><%=str_first_page & str_largeprev_page & str_prev_page & str_first2_page & str_largeprev2_page & str_prev2_page%></div>
				<div class="nav forward-nav"><%=str_next_page & str_largenext_page & str_last_page & str_next2_page & str_largenext2_page & str_last2_page%></div>
			<%end if%>
			<form method="get" action="<%=filename%>">
			<div class="pagenum-list">
				<%for j=start_page to end_page%><a name="pagenum" class="pagenum<%if j=CurPage then Response.Write " pagenum-current"%>" href="<%=filename%>?page=<%=j & param%>"><%=j%></a><%next%>
			</div>
			<div class="goto">(��<%=PagesCount%>ҳ)��ת��ҳ��<input type="text" name="page" class="page" maxlength="10" /> <input type="submit" class="submit" value="GO" /></div>
			<%
			for i=0 to ubound(arr_param)
				if arr_param(i)<>"" then%>
					<%t_name=left(arr_param(i),instr(arr_param(i),"=")-1)
					t_value=right(arr_param(i),len(arr_param(i))-len(t_name)-1)
					if t_name<>"page" then%><input type="hidden" name="<%=t_name%>" value="<%=UrlDecode(t_value)%>" /><%end if%>
				<%end if
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

	if clng(admin_faceid)<>0 and admin_faceurl="" then admin_faceurl=FacePath & admin_faceid  & ".gif"
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
	if Len(admin_email)>0 then%><a class="icon" href="mailto:<%=admin_email%>" title="�������䣺<%=admin_email%>"><img src="asset/image/icon_mail.gif"/></a><%end if
	if Len(admin_qqid)>0 then%><span class="icon" title="����QQ��<%=admin_qqid%>"><img src="asset/image/icon_qq.gif"/></span><%end if
	if Len(admin_msnid)>0 then%><span class="icon"><img src="asset/image/icon_skype.gif" alt="����Skype��<%=admin_msnid%>"/></span><%end if
	if Len(admin_homepage)>0 then%><a class="icon" href="<%=admin_homepage%>" target="_blank" title="������ҳ��<%=admin_homepage%>"><img src="asset/image/icon_homepage.gif"/></a><%end if
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
		<div class="name">ѡ�����ԣ�</div>
		<div class="tools">
			<span class="tool"><input type="checkbox" name="seltodel" class="seltodel checkbox" id="c<%=rs.Fields("id")%>" value="<%=rs("id")%>" /><label for="c<%=rs.Fields("id")%>">(ѡ��)</label></span>
		</div>
	</div>
	<div class="group">
		<div class="name">�ÿ����ԣ�</div>
		<div class="tools">
			<%if clng(guestflag and 16)<>0 then%><a class="tool" href="admin_passaudit.asp<%=param_url%>" title="ͨ�����"<%if PassAuditTip=true then response.Write " onclick=""return confirm('ȷʵҪͨ�������');"""%>><img src="asset/image/icon_pass.gif" />[ͨ�����]</a><%end if%>
			<a class="tool" href="admin_edit.asp<%=param_url%>" title="�༭�ÿ�����"><img src="asset/image/icon_edit.gif" />[�༭����]</a>
			<%if clng(guestflag and 32)<>0 then%><a class="tool" href="admin_pubwhisper.asp<%=param_url%>" title="�������Ļ�"<%if PubWhisperTip=true then response.Write " onclick=""return confirm('ȷʵҪ�������Ļ���');"""%>><img src="asset/image/icon_pub.gif" />[�������Ļ�]</a><%end if%>
			<%if clng(guestflag and 256)=0 then%><a class="tool" href="admin_hidecontact.asp<%=param_url%>" title="���طÿ���ϵ��ʽ"><img src="asset/image/icon_hidecontact.gif" />[������ϵ]</a><%end if%>
			<%if clng(guestflag and 256)<>0 then%><a class="tool" href="admin_unhidecontact.asp<%=param_url%>" title="�����ÿ���ϵ��ʽ"><img src="asset/image/icon_unhidecontact.gif" />[������ϵ]</a><%end if%>
			<%if clng(guestflag and 40)=0 then%><a class="tool" href="admin_hideword.asp<%=param_url%>" title="���طÿ���������"><img src="asset/image/icon_hide.gif" />[��������]</a><%end if%>
			<%if clng(guestflag and 40)=8 then%><a class="tool" href="admin_unhideword.asp<%=param_url%>" title="�����ÿ���������"><img src="asset/image/icon_unhide.gif" />[��������]</a><%end if%>
		</div>
	</div>
	<div class="group">
		<div class="name">�������ܣ�</div>
		<div class="tools">
			<a class="tool" href="admin_reply.asp<%=param_url%>" title="<%if clng(guestflag and 16)<>0 then response.write "ͨ����˲�"%>�ظ�������"<%if clng(guestflag and 16)<>0 and PassAuditTip=true then response.Write " onclick=""return confirm('ȷʵҪͨ�������');"""%>><img src="asset/image/icon_reply.gif" />[<%if clng(guestflag and 16)<>0 then response.write "ͨ����˲�"%>�ظ�����]</a>
			<a class="tool" href="admin_del.asp<%=param_url%>" title="ɾ������(�����ظ�)"<%if DelTip=true then Response.Write " onclick=""return confirm('ȷʵҪɾ��������');"""%>><img src="asset/image/icon_del.gif" />[ɾ������]</a>
			<%if CBool(rs.Fields("replied") AND 1) then %><a class="tool" href="admin_delreply.asp<%=param_url%>" title="ɾ���ظ�"<%if DelReTip=true then Response.Write " onclick=""return confirm('ȷʵҪɾ���ظ���');"""%>><img src="asset/image/icon_delreply.gif" />[ɾ���ظ�]</a><%end if%>
		</div>
	</div>
	<%if rs.Fields("parent_id")<=0 then%>
	<div class="group">
		<div class="name">���Կ��ƣ�</div>
		<div class="tools">
			<%if rs.Fields("parent_id")=0 then%><a class="tool" href="admin_lock2top.asp<%=param_url%>" title="������ʼ����ʾ�ڶ���"<%if Lock2TopTip=true then Response.Write " onclick=""return confirm('ȷʵҪ�ö�������');"""%>><img src="asset/image/icon_toplocked.gif" />[�ö�����]</a><%end if%>
			<%if rs.Fields("parent_id")<0 then%><a class="tool" href="admin_unlock2top.asp<%=param_url%>" title="ȡ�������ö�"<%if Lock2TopTip=true then Response.Write " onclick=""return confirm('ȷʵҪ�ö�������');"""%>><img src="asset/image/icon_untoplocked.gif" />[ȡ���ö�]</a><%end if%>
			<a class="tool" href="admin_bring2top.asp<%=param_url%>" title="��ǰ���Ե���ǰ"<%if Bring2TopTip=true then Response.Write " onclick=""return confirm('ȷʵҪ��ǰ������');"""%>><img src="asset/image/icon_top.gif" />[��ǰ����]</a>
			<%if rs.Fields("parent_id")<0 or rs.Fields("logdate")<>rs.Fields("lastupdated") then%><a class="tool" href="admin_reorder.asp<%=param_url%>" title="ʹ���Իָ���ԭʼ����λ��"<%if ReorderTip=true then Response.Write " onclick=""return confirm('ȷʵҪ��������˳����');"""%>><img src="asset/image/icon_reorder.gif" />[����˳��]</a><%end if%>
			<%if clng(rs.Fields("guestflag") and 512)=0 then%><a class="tool" href="admin_lockreply.asp<%=param_url%>" title="�����ÿͻظ�"><img src="asset/image/icon_lockreply.gif" />[�����ظ�]</a><%end if%>
			<%if clng(rs.Fields("guestflag") and 512)<>0 then%><a class="tool" href="admin_unlockreply.asp<%=param_url%>" title="����ÿͻظ�"><img src="asset/image/icon_reply.gif" />[����ظ�]</a><%end if%>
		</div>
	</div>
	<%end if%>
	<%if clng(guestflag and 952)<>0 OR rs.Fields("parent_id")<0 then%>
	<div class="group">
		<div class="name">����״̬��</div>
		<div class="tools">
			<%if clng(guestflag and 16)<>0 then%><span class="tool"><img src="asset/image/icon_wait2pass.gif" />�ȴ����</span><%end if%>
			<%if clng(guestflag and 32)<>0 then%><span class="tool"><img src="asset/image/icon_whisper.gif" />���Ļ�<%if clng(guestflag and 64)<>0 then response.write ",�Ѽ���"%></span><%end if%>
			<%if rs.Fields("parent_id")<0 then%><span class="tool"><img src="asset/image/icon_toplocked.gif" />�������ö�</span><%end if%>
			<%if clng(guestflag and 512)<>0 and rs.Fields("parent_id")<=0 then%><span class="tool"><img src="asset/image/icon_lockreply.gif" />�ظ�������</span><%end if%>
			<%if clng(guestflag and 256)<>0 then%><span class="tool"><img src="asset/image/icon_hidecontact.gif" />��ϵ������</span><%end if%>
			<%if clng(guestflag and 40)=8 then%><span class="tool"><img src="asset/image/icon_hide.gif" />����������</span><%end if%>
			<%if clng(guestflag and 128)<>0 then%><span class="tool"><img src="asset/image/icon_mail.gif" />�ظ�֪ͨ<%if MailReplyInform=false then response.write ",�ѽ���"%></span><%end if%>
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
	<%if parent_id<0 then%><span class="tool"><img src="asset/image/icon_toplocked.gif"/>(�ö�)</span><%end if%>
	<%if show_reply then%><span class="tool"><a href="<%=url%>"><img src="asset/image/icon_reply.gif"/>[�ظ�]</a></span><%end if%>
</div>
<%end sub

sub showGuestIcons(rs)
	Dim email,qqid,msnid,homepage,ipv4addr,ipv6addr,originalipv4,originalipv6

	email=rs.Fields("email")
	if Len(email)>0 then%><a class="icon" href="mailto:<%=email%>" title="�������䣺<%=email%>"><img src="asset/image/icon_mail.gif"/></a><%end if
	qqid=rs.Fields("qqid")
	if Len(qqid)>0 then%><span class="icon" title="����QQ��<%=qqid%>"><img src="asset/image/icon_qq.gif"/></span><%end if
	msnid=rs.Fields("msnid")
	if Len(msnid)>0 then%><span class="icon" title="����Skype��<%=msnid%>"><img src="asset/image/icon_skype.gif"/></span><%end if
	homepage=rs.Fields("homepage")
	if Len(homepage)>0 then%><a class="icon" href="<%=homepage%>" target="_blank" title="������ҳ��<%=homepage%>"><img src="asset/image/icon_homepage.gif" /></a><%end if

	if left(pagename,5)="admin" then
		ipv4addr=rs.Fields("ipv4addr")
		if AdminShowIPv4>0 and Len(ipv4addr)>0 then
			ipv4addr=GetIPv4WithMask(ipv4addr,AdminShowIPv4)%><span class="icon-entry"><span class="icon" title="IP��<%=ipv4addr%>"><img src="asset/image/icon_ip.gif"/></span></span><%end if
		ipv6addr=rs.Fields("ipv6addr")
		if AdminShowIPv6>0 and Len(ipv6addr)>0 then
			ipv6addr=GetIPv6WithMask(ipv6addr,AdminShowIPv6)%><span class="icon-entry"><span class="icon" title="IP��<%=ipv6addr%>"><img src="asset/image/icon_ip.gif"/></span></span><%end if
		originalipv4=rs.Fields("originalipv4")
		if AdminShowOriginalIPv4>0 and Len(originalipv4)>0 then
			originalipv4=GetIPv4WithMask(originalipv4,AdminShowOriginalIPv4)%><span class="icon-entry"><span class="icon" title="ԭʼIP��<%=originalipv4%>"><img src="asset/image/icon_ip2.gif"/></span></span><%end if
		originalipv6=rs.Fields("originalipv6")
		if AdminShowOriginalIPv6>0 and Len(originalipv6)>0 then
			originalipv6=GetIPv6WithMask(originalipv6,AdminShowOriginalIPv6)%><span class="icon-entry"><span class="icon" title="ԭʼIP��<%=originalipv6%>"><img src="asset/image/icon_ip2.gif"/></span></span><%end if
	else
		ipv4addr=rs.Fields("ipv4addr")
		if ShowIPv4>0 and Len(ipv4addr)>0 then
			ipv4addr=GetIPv4WithMask(ipv4addr,ShowIPv4)%><span class="icon" title="IP��<%=ipv4addr%>"><img src="asset/image/icon_ip.gif"/></span><%end if
		ipv6addr=rs.Fields("ipv6addr")
		if ShowIPv6>0 and Len(ipv6addr)>0 then
			ipv6addr=GetIPv6WithMask(ipv6addr,ShowIPv6)%><span class="icon" title="IP��<%=ipv6addr%>"><img src="asset/image/icon_ip.gif"/></span><%end if
	end if
end sub

sub inneradminreply(byref rs2)%>
<div class="message inner-message admin-message">
	<div class="summary">
		<div class="name">����<%if admin_name<>"" then response.write "(" & admin_name & ")"%>�ظ���</div>
		<div class="date">(<%=rs2("replydate")%>)</div>
		<div class="icons"><%showAdminIcons()%></div>
	</div>
	<%if admin_faceurl<>"" then%><img class="face" src="<%=admin_faceurl%>"/><%end if%>
	<div class="words">
		<%if encrypted=true and pagename<>"showword" and left(pagename,5)<>"admin" then%>
			<span class="inner-hint"><img src="asset/image/icon_key.gif"/>(��ҪԤ���������ܲ鿴...)[<a href="showword.asp?id=<%=rs2("id")%>">���������֤...</a>]</span>
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
	<h2 class="title">[�����ظ���]</h2>
	<div class="info">
		<%if admin_faceurl<>"" then%><img class="face" src="<%=admin_faceurl%>"/><%end if%>
		<%if admin_email<>"" or admin_qqid<>"" or admin_msnid<>"" or admin_homepage<>"" then%>
			<div class="icons"><%showAdminIcons()%></div>
		<%end if%>
		<div class="date"><%=rs2("replydate")%></div>
	</div>
	<div class="detail">
		<div class="words">
			<%if encrypted=true and pagename<>"showword" and left(pagename,5)<>"admin" then%>
				<span class="inner-hint"><img src="asset/image/icon_key.gif"/>(��ҪԤ���������ܲ鿴...)[<a href="showword.asp?id=<%=rs2("id")%>">���������֤...</a>]</span>
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
		<h2 class="title">(���Դ����...)</h2>
	</div>
	<div class="words">
		<span class="inner-hint"><img src="asset/image/icon_wait2pass.gif"/>(���Դ����...)</span>
	</div>
</div>
<%end sub

sub outeraudit(t_rs)%>
<div class="message outer-message guest-message auditing-message">
	<div class="info">
		<%fid=t_rs("faceid")
		if isnumeric(fid) and StatusShowHead=true then
			if fid>=1 and fid<=FaceCount then%>
				<img class="face" src="<%=FacePath & cstr(fid)%>.gif" />
			<%end if
		end if%>

		<div class="date"><%=t_rs("logdate")%></div>
	</div>
	<div class="detail">
		<h2 class="title">(���Դ����...)</h2>
		<%if rs.Fields("parent_id")<=0 and left(pagename,5)<>"admin" then showGuestMessageTools rs.Fields("id"),rs.Fields("parent_id"),StatusWrite and StatusGuestReply and clng(rs.Fields("guestflag") and 512)=0%>
		<div class="words">
			<span class="inner-hint"><img src="asset/image/icon_wait2pass.gif" />(���Դ����...)</span>
		</div>
	</div>
</div>
<%end sub

sub innerword(byref t_rs)%>
<div class="message inner-message guest-message">
	<div class="summary">
		<div class="name"><%=t_rs("name")%>��</div>
		<div class="date">(<%=t_rs("logdate")%>)</div>
			<%if (iswhisper=false and clng(guestflag and 256)=0) or (pagename="showword" and needverify) or left(pagename,5)="admin" then%>
				<div class="icons"><%showGuestIcons(t_rs)%></div>
			<%end if%>
		<h2 class="title"><%if iswhisper=true and pagename<>"showword" and left(pagename,5)<>"admin" then response.write "(�����������Ļ�...)" else response.write t_rs("title")%></h2>
	</div>
	<%if StatusShowHead and t_rs.Fields("faceid")>=1 and t_rs.Fields("faceid")<=FaceCount then%><img class="face" src="<%=FacePath & t_rs.Fields("faceid")%>.gif"/><%end if%>
	<div class="words">
		<%if iswhisper=true and pagename<>"showword" and left(pagename,5)<>"admin" then%>
			<span class="inner-hint"><img src="asset/image/icon_whisper.gif"/>(�����������Ļ�...)</span>
		<%elseif ishidden=true and left(pagename,5)<>"admin" then%>
			<span class="inner-hint"><img src="asset/image/icon_hide.gif"/>(���Ա�����Ա����...)</span>
		<%else
			guest_txt="" & t_rs("article") & ""
			if left(pagename,5)="admin" and AdminViewCode=true then
				guest_txt=server.htmlEncode(guest_txt)
			else
				convertstr guest_txt,guestflag,false
			end if

			if guest_txt<>"" then
				Response.Write guest_txt
			else
				Response.Write "(������)"
			end if
		end if%>
	</div>
</div>
<%end sub

sub outerword(byref rs)%>
<div class="message outer-message guest-message">
	<div class="info">
		<div class="name"><%=rs("name")%></div>

		<%fid=rs("faceid")
		if isnumeric(fid) and StatusShowHead=true then
			if fid>=1 and fid<=FaceCount then%>
				<img class="face" src="<%=FacePath & cstr(fid)%>.gif" />
			<%end if
		end if%>

		<%if (iswhisper=false and clng(guestflag and 256)=0) or (pagename="showword" and needverify) or left(pagename,5)="admin" then%>
			<div class="icons"><%showGuestIcons(rs)%></div>
		<%end if%>

		<div class="date"><%=rs("logdate")%></div>
	</div>
	<div class="detail">
		<h2 class="title"><%if iswhisper=true and pagename<>"showword" and left(pagename,5)<>"admin" then response.write "(�����������Ļ�...)" else response.write rs("title")%></h2>
		<%if rs.Fields("parent_id")<=0 and left(pagename,5)<>"admin" then showGuestMessageTools rs.Fields("id"),rs.Fields("parent_id"),StatusWrite and StatusGuestReply and clng(rs.Fields("guestflag") and 512)=0%>
		<div class="words">
			<%if iswhisper=true and pagename<>"showword" and left(pagename,5)<>"admin" then%>
				<span class="inner-hint"><img src="asset/image/icon_whisper.gif"/>(�����������Ļ�...)</span>
			<%elseif ishidden=true and left(pagename,5)<>"admin" then%>
				<span class="inner-hint"><img src="asset/image/icon_hide.gif"/>(���Ա�����Ա����...)</span>
			<%else
				guest_txt="" & rs("article") & ""
				if left(pagename,5)="admin" and AdminViewCode=true then
					guest_txt=server.htmlEncode(guest_txt)
				else
					convertstr guest_txt,guestflag,false
				end if

				if guest_txt<>"" then
					Response.Write guest_txt
				else
					Response.Write "(������)"
				end if
			end if
			%>
		</div>

		<%
		if CBool(rs.Fields("replied") AND 1) and ReplyInWord=true then inneradminreply(rs)	'��Ƕ

		if (pagename="admin" or pagename="admin_search" or pagename="admin_showword") and ReplyInWord=true then
			Call showAdminMessageTools(rs)
		end if

		if rs.Fields("parent_id")<=0 and clng(rs.Fields("replied") AND 2)<>0 and (encrypted=false or pagename="showword" or left(pagename,5)="admin") and ReplyInWord=true then
			dim hidden_condition
			hidden_condition=""
			if left(pagename,5)<>"admin" then hidden_condition=GetHiddenWordCondition()

			set rs1=Server.CreateObject("ADODB.Recordset")
			rs1.Open Replace(Replace(sql_common2_guestreply,"{0}",rs.Fields("id")),"{1}",hidden_condition),cn,0,1,1
			while not rs1.eof
				guestflag=rs1("guestflag")
				ishidden=(clng(guestflag and 40)=8)
				iswhisper=(clng(guestflag and 32)<>0)
				encrypted=(clng(guestflag and 96)=96)

				if clng(guestflag and 16)<>0 and left(pagename,5)<>"admin" then		'�����
					inneraudit()
				else
					innerword(rs1)
					if rs1.Fields("replied") then inneradminreply(rs1)
				end if

				if pagename="admin" or pagename="admin_search" or pagename="admin_showword" then
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