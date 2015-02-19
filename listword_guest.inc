<!-- #include file="common2.asp" -->
<%
dim guest_txt,guestflag,iswhisper,ishidden,encrypted
dim reply_txt,reply_htmlflag
dim js_status,n_i
js_status=" onmouseover=""if(this.alt&amp;&amp;typeof(window.status).toLowerCase()=='string')window.status=this.alt;"" onmouseout=""if(this.alt&amp;&amp;typeof(window.status).toLowerCase()=='string')window.status='';"""
n_i=0

getadmininfo()

for i=1 to ItemsPerPage
	if rs.eof=false and Response.IsClientConnected then
		guestflag=rs("guestflag")
		if clng(guestflag and 40)=8 then ishidden=true else ishidden=false
		if clng(guestflag and 32)<>0 then iswhisper=true else iswhisper=false
		if clng(guestflag and 96)=96 then encrypted=true else encrypted=false

		n_i=n_i+1
		%>
		
		<a name="a<%=n_i%>" id="a<%=n_i%>"></a>
		<table border="1" cellpadding="0" cellspacing="0" class="onetopic">
		<%
		if clng(guestflag and 16)<>0 then		'待审核
			outeraudit()
		else
			outerword(rs)
			if clng(rs.Fields("replied") AND 1)<>0 and ReplyInWord=false then outerreply(rs)
			
			if rs.Fields("parent_id")<=0 and clng(rs.Fields("replied") AND 2)<>0 and (encrypted=false or pagename="showword") and ReplyInWord=false then
				set rs1=Server.CreateObject("ADODB.Recordset")
				rs1.open Replace(Replace(sql_listword_guest,"{0}",rs.Fields("id")),"{1}",GetHiddenWordCondition()),cn,0,1,1
				while not rs1.eof
					guestflag=rs1("guestflag")
					if clng(guestflag and 40)=8 then ishidden=true else ishidden=false
					if clng(guestflag and 32)<>0 then iswhisper=true else iswhisper=false
					if clng(guestflag and 96)=96 then encrypted=true else encrypted=false
					
					if clng(guestflag and 16)<>0 then		'待审核
						outeraudit()
					else
						outerword(rs1)
						if clng(rs1.Fields("replied") AND 1)<>0 then outerreply(rs1)
					end if
					rs1.movenext
				wend
				rs1.close
				set rs1=nothing
			end if
		%>	  
			<%if pagename<>"showword" and clng(rs.Fields("guestflag") and 96)=96 and rs.Fields("parent_id")<=0 then%>
				<tr>
					<td colspan="2" class="vguestpass">
					悄悄话已预设密码，[<a href="showword.asp?id=<%=rs("id")%>">点击这里验证...</a>]
					</td>
				</tr>
			<%end if%>
		<%end if%>
		</table>

		<%
		rs.MoveNext
	else
		exit for
	end if
next
%>