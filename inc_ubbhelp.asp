<table id="ubblist" cellpadding="2" class="grid" style="width:auto; border:solid 1px <%=TableBorderColor%>; border-collapse:collapse; margin:0px auto;">
	<tr style="font-weight:bold;"><th>����</th><th>��ʽ</th><th>״̬</th><th>����</th></tr>
	<tr><td>ͼƬ</td>					<td>[img]ͼƬ��ַ[/img]<br/>[img=��,��]ͼƬ��ַ[/img]</td>			<td><%=getstatus(UbbFlag_image)%></td>		<td>ͼƬ</td></tr>
	<tr><td>������</td>					<td>[url]��Դ��ַ[/url]<br/>[url=��Դ��ַ]��������[/url]</td>		<td><%=getstatus(UbbFlag_url)%></td>		<td>URL��Email</td></tr>
	<tr><td>�����ʼ�����</td>			<td>[email]�ʼ���ַ[/email]</td>										<td><%=getstatus(UbbFlag_url)%></td>		<td>URL��Email</td></tr>
	<tr><td>�Զ�ʶ����ַ</td>			<td>������ַ���Զ���Ϊ������</td>									<td><%=getstatus(UbbFlag_autourl)%></td>	<td>�Զ�ʶ����ַ</td></tr>
	<tr><td>Flash</td>					<td>[flash]Flash��ַ[/flash]<br/>[flash=��,��]Flash��ַ[/flash]</td>	<td><%=getstatus(UbbFlag_player)%></td>		<td>���ſؼ�</td></tr>
	<tr><td>Windows Media Player</td>	<td>[mp]ý���ļ���ַ[/mp]<br/>[mp=��,��]ý���ļ���ַ[/mp]</td>		<td><%=getstatus(UbbFlag_player)%></td>		<td>���ſؼ�</td></tr>
	<tr><td>Real Player</td>			<td>[rm]ý���ļ���ַ[/rm]<br/>[rm=��,��]ý���ļ���ַ[/rm]</td>		<td><%=getstatus(UbbFlag_player)%></td>		<td>���ſؼ�</td></tr>
	<tr><td>���ã�������</td>			<td>[quote]��������[/quote]</td>										<td><%=getstatus(UbbFlag_paragraph)%></td>	<td>������ʽ</td></tr>
	<tr><td>�����б�</td>				<td>[li]�б���[/li]</td>												<td><%=getstatus(UbbFlag_paragraph)%></td>	<td>������ʽ</td></tr>
	<tr><td>ָ������</td>				<td>[font=������]����[/font]</td>										<td><%=getstatus(UbbFlag_fontstyle)%></td>	<td>������ʽ</td></tr>
	<tr><td>������</td>					<td>[b]��������[/b]</td>												<td><%=getstatus(UbbFlag_fontstyle)%></td>	<td>������ʽ</td></tr>
	<tr><td>б����</td>					<td>[i]б������[/i]</td>												<td><%=getstatus(UbbFlag_fontstyle)%></td>	<td>������ʽ</td></tr>
	<tr><td>�»���</td>					<td>[u]���»��ߵ�����[/u]</td>										<td><%=getstatus(UbbFlag_fontstyle)%></td>	<td>������ʽ</td></tr>
	<tr><td>ɾ����</td>					<td>[strike]��ɾ���ߵ�����[/strike]</td>								<td><%=getstatus(UbbFlag_fontstyle)%></td>	<td>������ʽ</td></tr>
	<tr><td>������ɫ</td>				<td>[color=��ɫ]Ӧ����ɫ������[/color]</td>							<td><%=getstatus(UbbFlag_fontcolor)%></td>	<td>������ɫ</td></tr>
	<tr><td>���ֱ���ɫ</td>			<td>[bgcolor=��ɫ]������ɫ������[/bgcolor]</td>						<td><%=getstatus(UbbFlag_fontcolor)%></td>	<td>������ɫ</td></tr>
	<tr><td>ָ�����뷽ʽ</td>			<td>[align=��ʽ(left,center,right)]Ӧ���ڶ��������[/align]</td>		<td><%=getstatus(UbbFlag_alignment)%></td>	<td>���뷽ʽ</td></tr>
	<tr><td>���ж���</td>				<td>[center]Ӧ���ھ��ж��������[/center]</td>						<td><%=getstatus(UbbFlag_alignment)%></td>	<td>���뷽ʽ</td></tr>
	<tr><td>���ع���������</td>		<td>[fly]��������[/fly]</td>											<td><%=getstatus(UbbFlag_movement)%></td>	<td>�ƶ�Ч��</td></tr>
	<tr><td>�������������</td>		<td>[move]��������[/move]</td>											<td><%=getstatus(UbbFlag_movement)%></td>	<td>�ƶ�Ч��</td></tr>
	<tr><td>��������Ч��</td>			<td>[glow=����������,��ɫ,ǿ��]����[/glow]</td>					<td><%=getstatus(UbbFlag_cssfilter)%></td>	<td>�˾�Ч��</td></tr>
	<tr><td>������ӰЧ��</td>			<td>[shadow=������Ӱ���,��ɫ,ǿ��]����[/shadow]</td>				<td><%=getstatus(UbbFlag_cssfilter)%></td>	<td>�˾�Ч��</td></tr>
	<tr><td>����ͼ��</td>				<td>[faceX] ��XΪ�����ţ�</td>										<td><%=getstatus(UbbFlag_face)%></td>		<td>����ͼ��</td></tr>
</table>
<!-- #include file="rowlight.inc" -->