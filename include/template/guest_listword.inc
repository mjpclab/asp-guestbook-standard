<%
dim guest_txt,guestflag,iswhisper,ishidden,encrypted
dim reply_txt,reply_htmlflag
dim n_i
n_i=0

getadmininfo()

while Not rs.EOF
	guestflag=rs("guestflag")
	ishidden=((guestflag AND 40)=8)
	iswhisper=CBool(guestflag AND 32)
	encrypted=((guestflag AND 96)=96)

	n_i=n_i+1
	%>

	<a name="a<%=n_i%>" id="a<%=n_i%>"></a>

	<div class="topic"><%
	if CBool(guestflag and 16) then		'待审核
		outeraudit(rs)
	else
		outerword(rs)
		if CBool(rs.Fields("replied") AND 1) and ReplyInWord=false then outeradminreply(rs)

		if rs.Fields("parent_id")<=0 and CBool(rs.Fields("replied") AND 2) and (encrypted=false or pagename="showword") and ReplyInWord=false then
			set rs1=Server.CreateObject("ADODB.Recordset")
			rs1.open Replace(Replace(sql_listword_guest,"{0}",rs.Fields("id")),"{1}",GetHiddenWordCondition()),cn,0,1,1
			while not rs1.eof
				guestflag=rs1("guestflag")
				ishidden=((guestflag AND 40)=8)
				iswhisper=CBool(guestflag AND 32)
				encrypted=((guestflag AND 96)=96)

				if CBool(guestflag and 16) then		'待审核
					outeraudit(rs1)
				else
					outerword(rs1)
					if CBool(rs1.Fields("replied") AND 1) then outeradminreply(rs1)
				end if
				rs1.movenext
			wend
			rs1.close
			set rs1=nothing
		end if

		if pagename<>"showword" and (rs.Fields("guestflag") and 96)=96 and rs.Fields("parent_id")<=0 then%>
			<div class="outer-hint">悄悄话已预设密码，[<a href="showword.asp?id=<%=rs("id")%>">点击这里验证...</a>]</div>
		<%end if
	end if%>
	</div>

	<%
	rs.MoveNext
wend
%>