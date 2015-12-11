<%if isnumeric(Session("guestflag")) and (Session("guestflag") and 48)<>0 and (HideAudit or HideWhisper) then%>
<div class="region">
	<h3 class="title"><img src="image/icon_tip.gif" />留言提示：</h3>
	<div class="content">
		<p>您已成功留言，但由于如下原因，暂时还不会显示：</p>
		<%if HideAudit and (Session("guestflag") and 16)<>0 then%><p>■　留言正等待审核，在通过之前不会显示。</p><%end if%>
		<%if HideWhisper and (Session("guestflag") and 32)<>0 then%><p>■　您写的是悄悄话，在版主回复之前不会显示。</p><%end if%>
	</div>
</div>
<%
end if
Session("guestflag")=0
%>