<%set rs2=server.CreateObject("ADODB.Recordset")
rs2.open sql_topbulletin,cn,0,1,1

pub_flag=rs2("declareflag")
pub="" &rs2("declare")& ""
convertstr pub,pub_flag,true

rs2.close
set rs2=nothing
if pub<>"" then
%>
<div class="region">
	<h3 class="title"><img alt="" src="asset/image/icon_bulletin.gif" class="imgicon" />置顶公告：</h3>
	<div class="content"><%=pub%></div>
</div>
<%end if%>