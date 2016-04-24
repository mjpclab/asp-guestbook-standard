<table id="ubblist" class="topic-list">
<thead>
	<tr><th style="width:8em;">功能</th><th>格式</th><th style="width:3em;">状态</th><th style="width:6em;">分组</th></tr>
</thead>
<tbody>
	<tr><td>图片</td>					<td>[img]图片地址[/img]<br/>[img=宽,高]图片地址[/img]</td>			<td><%=getstatus(UbbFlag_image)%></td>		<td>图片</td></tr>
	<tr><td>超链接</td>					<td>[url]资源地址[/url]<br/>[url=资源地址]链接文字[/url]</td>		<td><%=getstatus(UbbFlag_url)%></td>		<td>URL、Email</td></tr>
	<tr><td>电子邮件链接</td>				<td>[email]邮件地址[/email]</td>									<td><%=getstatus(UbbFlag_url)%></td>		<td>URL、Email</td></tr>
	<tr><td>自动识别网址</td>				<td>输入网址后自动变为超链接</td>									<td><%=getstatus(UbbFlag_autourl)%></td>	<td>自动识别网址</td></tr>
	<tr><td>Flash</td>					<td>[flash]Flash地址[/flash]<br/>[flash=宽,高]Flash地址[/flash]</td>	<td><%=getstatus(UbbFlag_player)%></td>		<td>播放控件</td></tr>
	<tr><td>Windows Media Player</td>	<td>[mp]媒体文件地址[/mp]<br/>[mp=宽,高]媒体文件地址[/mp]</td>		<td><%=getstatus(UbbFlag_player)%></td>		<td>播放控件</td></tr>
	<tr><td>HTML5 视频</td>				<td>[video]视频文件地址[/video]<br/>[video=宽,高]视频文件地址[/video]</td>	<td><%=getstatus(UbbFlag_player)%></td>		<td>播放控件</td></tr>
	<tr><td>引用（缩进）</td>				<td>[quote]引用文字[/quote]</td>									<td><%=getstatus(UbbFlag_paragraph)%></td>	<td>段落样式</td></tr>
	<tr><td>无序列表</td>				<td>[li]列表项[/li]</td>											<td><%=getstatus(UbbFlag_paragraph)%></td>	<td>段落样式</td></tr>
	<tr><td>指定字体</td>				<td>[font=字体名]文字[/font]</td>									<td><%=getstatus(UbbFlag_fontstyle)%></td>	<td>字体样式</td></tr>
	<tr><td>粗体字</td>					<td>[b]粗体文字[/b]</td>											<td><%=getstatus(UbbFlag_fontstyle)%></td>	<td>字体样式</td></tr>
	<tr><td>斜体字</td>					<td>[i]斜体文字[/i]</td>											<td><%=getstatus(UbbFlag_fontstyle)%></td>	<td>字体样式</td></tr>
	<tr><td>下划线</td>					<td>[u]加下划线的文字[/u]</td>									<td><%=getstatus(UbbFlag_fontstyle)%></td>	<td>字体样式</td></tr>
	<tr><td>删除线</td>					<td>[strike]加删除线的文字[/strike]</td>							<td><%=getstatus(UbbFlag_fontstyle)%></td>	<td>字体样式</td></tr>
	<tr><td>文字颜色</td>				<td>[color=颜色]应用颜色的文字[/color]</td>						<td><%=getstatus(UbbFlag_fontcolor)%></td>	<td>字体颜色</td></tr>
	<tr><td>文字背景色</td>				<td>[bgcolor=颜色]含背景色的文字[/bgcolor]</td>					<td><%=getstatus(UbbFlag_fontcolor)%></td>	<td>字体颜色</td></tr>
	<tr><td>指定对齐方式</td>				<td>[align=方式(left,center,right)]应用于对齐的文字[/align]</td>	<td><%=getstatus(UbbFlag_alignment)%></td>	<td>对齐方式</td></tr>
	<tr><td>居中对齐</td>				<td>[center]应用于居中对齐的文字[/center]</td>						<td><%=getstatus(UbbFlag_alignment)%></td>	<td>对齐方式</td></tr>
	<tr><td>表情图标</td>				<td>[faceX] （X为表情编号）</td>									<td><%=getstatus(UbbFlag_face)%></td>		<td>表情图标</td></tr>
</tbody>
</table>
