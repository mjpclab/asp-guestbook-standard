<%set rs2=server.CreateObject("ADODB.Recordset")
rs2.open sql_topbulletin,cn,0,1,1

pub_flag=rs2("declareflag")
pub="" &rs2("declare")& ""
convertstr pub,pub_flag,true

rs2.close
set rs2=nothing
if pub<>"" then
%>
<table border="1" cellpadding="2" class="onetopic" style="overflow:hidden;">
	<tr style="height:25px">
		<td class="bulletintitle"><img alt="" src="image/icon_bulletin.gif" class="imgicon" />ÖÃ¶¥¹«¸æ£º</td>
	</tr>
	<tr>
		<td class="wordscontent"><%=pub%></td>
	</tr>
</table>
<%end if%>