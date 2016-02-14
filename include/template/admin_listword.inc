<%
dim guest_txt,guestflag
dim reply_txt,reply_htmlflag
dim js_status
dim param_url

getadmininfo()

while Not rs.EOF
	guestflag=rs("guestflag")
	%>

	<div class="topic"><%
		outerword(rs)
		if CBool(rs.Fields("replied") AND 1) and Not ReplyInWord then outeradminreply(rs)
		if Not ReplyInWord then
			Call showAdminMessageTools(rs)
		end if

		if rs.Fields("parent_id")<=0 and CBool(rs.Fields("replied") AND 2) and Not ReplyInWord then
			set rs1=Server.CreateObject("ADODB.Recordset")
			rs1.open Replace(sql_listword_admin,"{0}",rs.Fields("id")),cn,0,1,1
			while not rs1.eof
				guestflag=rs1("guestflag")
				ishidden=((guestflag AND 40)=8)
				iswhisper=CBool(guestflag AND 32)
				encrypted=((guestflag AND 96)=96)

				outerword(rs1)
				if CBool(rs1.Fields("replied") AND 1) then outeradminreply(rs1)
				Call showAdminMessageTools(rs1)
				rs1.movenext
			wend
			rs1.close
			set rs1=nothing
		end if%>
	</div>

	<%rs.MoveNext
wend
%>