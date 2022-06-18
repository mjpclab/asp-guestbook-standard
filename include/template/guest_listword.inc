<%
dim guest_txt,guestflag,iswhisper,ishidden,encrypted
dim reply_txt,reply_htmlflag

getadmininfo()

while Not rs.EOF
	guestflag=rs("guestflag")
	isauditting=CBool(guestflag AND 16)
	ishidden=((guestflag AND 40)=8)
	iswhisper=CBool(guestflag AND 32)
	encrypted=((guestflag AND 96)=96)
	%>

	<div class="topic"><%
	if isauditting then
		outeraudit(rs)
	else
		outerword(rs)
		if CBool(rs.Fields("replied") AND 1) and Not ReplyInWord then outeradminreply(rs)

		if rs.Fields("parent_id")<=0 and CBool(rs.Fields("replied") AND 2) and (encrypted=false or pagename="showword") and Not ReplyInWord then
			set rs1=Server.CreateObject("ADODB.Recordset")
			rs1.open Replace(Replace(sql_listword_guest,"{0}",rs.Fields("id")),"{1}",GetHiddenWordCondition()),cn,0,1,1
			while not rs1.eof
				guestflag=rs1("guestflag")
				isauditting=CBool(guestflag AND 16)
				ishidden=((guestflag AND 40)=8)
				iswhisper=CBool(guestflag AND 32)
				encrypted=((guestflag AND 96)=96)

				if isauditting then
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
	end if%>
	</div>

	<%
	rs.MoveNext
wend
%>