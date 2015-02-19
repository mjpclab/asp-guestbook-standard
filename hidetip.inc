<%if isnumeric(Session("guestflag")) and Session("guestflag")<>"" and (Session("guestflag") and 48)<>0 and (HideAudit or HideWhisper) then%>
<table border="1" cellpadding="2" class="onetopic" style="overflow:hidden;">
	<tr style="height:25px">
		<td class="bulletintitle"><img alt="" src="image/icon_tip.gif" class="imgicon" />留言提示：</td>
	</tr>
	<tr>
		<td class="wordscontent">
			您已成功留言，但由于如下原因，暂时还不会显示：<br/>
			<%if HideAudit and (Session("guestflag") and 16)<>0 then%>■　留言正等待审核，在通过之前不会显示。<br/><%end if%>
			<%if HideWhisper and (Session("guestflag") and 32)<>0 then%>■　您写的是悄悄话，在版主回复之前不会显示。<br/><%end if%>
		</td>
	</tr>
</table>
<%
end if
Session("guestflag")=0
%>