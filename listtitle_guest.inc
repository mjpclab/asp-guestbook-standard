<table id="titlelist" border="1" cellpadding="2" class="generalwindow grid" style="color:<%=TableContentColor%>; table-layout:fixed;">
<tr class="header"><th style="width:25px;"> </th><th>标题</th><th style="width:100px;">姓名</th><th style="width:120px;">留言时间</th></tr>
<%
'dim guest_txt,guestflag,isauditting,ishidden,iswhisper,encrypted
'dim reply_txt,reply_htmlflag
n_i=0

for i=1 to TitlesPerPage
	if rs.eof=false and Response.IsClientConnected then
		guestflag=rs("guestflag")
		if clng(guestflag and 16)<>0 then isauditting=true else isauditting=false
		if clng(guestflag and 40)=8 then ishidden=true else ishidden=false
		if clng(guestflag and 32)<>0 then iswhisper=true else iswhisper=false
		if clng(guestflag and 96)=96 then encrypted=true else encrypted=false

		Response.Write Replace("<tr style=""background-color:{0}"">","{0}",TableContentBGC)
		
		Response.Write "<td style=""text-align:center;"">"
		n_i=n_i+1
		Response.Write "<a name=""a" &n_i& """ id=""a" &n_i& """></a>"
		if isauditting then
			Response.Write "<img class=""forumicon"" src=""image/icon_wait2pass.gif"" alt=""待审核"" />"
		elseif ishidden then
			Response.Write "<img class=""forumicon"" src=""image/icon_hide.gif"" alt=""被隐藏"" />"
		elseif iswhisper then 
			Response.Write "<img class=""forumicon"" src=""image/icon_whisper.gif"" alt=""悄悄话"" />"
		elseif rs.Fields("parent_id")<0 then
			Response.Write "<img class=""forumicon"" src=""image/icon_toplocked.gif"" alt=""置顶"" />"
		end if
		Response.Write "</td>"
		
		dim title_pattern
		title_pattern="<td nowrap=""nowrap"" style=""overflow:hidden;""><a href=""showword.asp?id={0}"">{1}</a></td>"
		if isauditting then
			Response.Write Replace(Replace(title_pattern,"{0}",rs.Fields("id")),"{1}","(留言待审核...)")
		elseif ishidden then
			Response.Write Replace(Replace(title_pattern,"{0}",rs.Fields("id")),"{1}","(留言被管理员隐藏...)")
		elseif iswhisper then 
			Response.Write Replace(Replace(title_pattern,"{0}",rs.Fields("id")),"{1}","(给版主的悄悄话...)")
		elseif Trim(rs.Fields("title"))="" then
			Response.Write Replace(Replace(title_pattern,"{0}",rs.Fields("id")),"{1}","(无标题)")
		else
			Response.Write Replace(Replace(title_pattern,"{0}",rs.Fields("id")),"{1}",rs.Fields("title"))
		end if

		if isauditting=false then
			Response.Write "<td nowrap=""nowrap"" style=""overflow:hidden;"">" & rs.Fields("name") & "</td>"
		else
			Response.Write "<td nowrap=""nowrap"" style=""overflow:hidden;""></td>"
		end if
		Response.Write "<td nowrap=""nowrap"" style=""overflow:hidden;"">" & rs.Fields("logdate") & "</td>"
		
		Response.Write "</tr>"
		
		rs.MoveNext
	else
		exit for
	end if
next
%>
</table>
<!-- #include file="rowlight.inc" -->