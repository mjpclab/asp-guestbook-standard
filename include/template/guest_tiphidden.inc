<%if isnumeric(Session("guestflag")) and (Session("guestflag") and 48)<>0 and (HideAudit or HideWhisper) then%>
<div class="region">
	<h3 class="title"><img src="asset/image/icon_tip.gif" />������ʾ��</h3>
	<div class="content">
		<p>���ѳɹ����ԣ�����������ԭ����ʱ��������ʾ��</p>
		<%if HideAudit and (Session("guestflag") and 16)<>0 then%><p>�����������ȴ���ˣ���ͨ��֮ǰ������ʾ��</p><%end if%>
		<%if HideWhisper and (Session("guestflag") and 32)<>0 then%><p>������д�������Ļ����ڰ����ظ�֮ǰ������ʾ��</p><%end if%>
	</div>
</div>
<%
end if
Session("guestflag")=0
%>