<!-- #include file="common2.asp" -->
<%
dim guest_txt,guestflag
dim reply_txt,reply_htmlflag
dim js_status
dim param_url
js_status=" onmouseover=""if(this.alt&amp;&amp;typeof(window.status).toLowerCase()=='string')window.status=this.alt;"" onmouseout=""if(this.alt&amp;&amp;typeof(window.status).toLowerCase()=='string')window.status='';"""

getadmininfo()

for i=1 to ItemsPerPage
	if rs.eof=false and Response.IsClientConnected then
		guestflag=rs("guestflag")
		%>
			<table border="1" cellpadding="0" cellspacing="0" class="onetopic">
			<%
			outerword(rs)
			if clng(rs.Fields("replied") AND 1)<>0 and ReplyInWord=false then outerreply(rs)
			if ReplyInWord=false then
				set p_rs=rs
				%><!-- #include file="admincontrols.inc" --><%
			end if
			
			if rs.Fields("parent_id")<=0 and clng(rs.Fields("replied") AND 2)<>0 and ReplyInWord=false then
				set rs1=Server.CreateObject("ADODB.Recordset")
				rs1.open Replace(sql_listword_admin,"{0}",rs.Fields("id")),cn,0,1,1
				while not rs1.eof
					guestflag=rs1("guestflag")
					if clng(guestflag and 40)=8 then ishidden=true else ishidden=false
					if clng(guestflag and 32)<>0 then iswhisper=true else iswhisper=false
					if clng(guestflag and 96)=96 then encrypted=true else encrypted=false
					
					outerword(rs1)
					if clng(rs1.Fields("replied") AND 1)<>0 then outerreply(rs1)
					set p_rs=rs1
					%><!-- #include file="admincontrols.inc" --><%
					rs1.movenext
				wend
				rs1.close
				set rs1=nothing
			end if
			%>
			</table>
		<%rs.MoveNext
	else
		exit for
	end if
next
%>