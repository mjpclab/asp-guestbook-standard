<!-- #include file="common2.asp" -->
<%
dim guest_txt,guestflag
dim reply_txt,reply_htmlflag
dim js_status
dim param_url

getadmininfo()

while rs.eof=false
	guestflag=rs("guestflag")
	%>

	<div class="topic"><%
		outerword(rs)
		if clng(rs.Fields("replied") AND 1)<>0 and ReplyInWord=false then outeradminreply(rs)
		if ReplyInWord=false then
			set p_rs=rs
			%><!-- #include file="admintools.inc" --><%
		end if

		if rs.Fields("parent_id")<=0 and clng(rs.Fields("replied") AND 2)<>0 and ReplyInWord=false then
			set rs1=Server.CreateObject("ADODB.Recordset")
			rs1.open Replace(sql_listword_admin,"{0}",rs.Fields("id")),cn,0,1,1
			while not rs1.eof
				guestflag=rs1("guestflag")
				ishidden=(clng(guestflag and 40)=8)
				iswhisper=(clng(guestflag and 32)<>0)
				encrypted=(clng(guestflag and 96)=96)

				outerword(rs1)
				if clng(rs1.Fields("replied") AND 1)<>0 then outeradminreply(rs1)
				set p_rs=rs1
				%><!-- #include file="admintools.inc" --><%
				rs1.movenext
			wend
			rs1.close
			set rs1=nothing
		end if%>
	</div>

	<%rs.MoveNext
wend
%>