<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
	<title><%=HomeName%> 留言本 留言本参数</title>
	
	<!-- #include file="style.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"管理"%>
	<!-- #include file="admintool.inc" -->

	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")

	CreateConn cn,dbtype
	rs.Open sql_adminconfig_config,cn,,,1
	%>

	<table border="1" bordercolor="<%=TableBorderColor%>" cellpadding="2" class="generalwindow">
		<tr>
			<td class="centertitle">留言本参数设置</td>
		</tr>
		<tr>
			<td style="width:100%; text-align:left; vertical-align:top; color:<%=TableContentColor%>; background-color:<%=TableContentBGC%>;">
			<table border="1" bordercolor="<%=TableBorderColor%>" style="width:100%; text-align:center; border-collapse:collapse;">
			<tr style="cursor:pointer;">
				<td style="width:20%;" onclick="window.location='admin_config.asp?page=1'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin_config.asp?page=1" onmouseover="return true;">[基本配置]</a></td>
				<td style="width:20%;" onclick="window.location='admin_config.asp?page=2'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin_config.asp?page=2" onmouseover="return true;">[邮件通知]</a></td>
				<td style="width:20%;" onclick="window.location='admin_config.asp?page=4'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin_config.asp?page=4" onmouseover="return true;">[界面尺寸]</a></td>
				<td style="width:20%;" onclick="window.location='admin_config.asp?page=8'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin_config.asp?page=8" onmouseover="return true;">[功能设置]</a></td>
				<td style="width:20%;" onclick="window.location='admin_config.asp'" onmouseover="this.firstChild.style.textDecoration='underline'" onmouseout="this.firstChild.style.textDecoration=''"><a href="admin_config.asp" onmouseover="return true;">[全部参数]</a></td>
			</tr>
			</table>

			<br/>
			<form method="post" action="admin_saveconfig.asp" name="configform" onsubmit="return check();">
		
			<%
			showpage=15
			if isnumeric(Request.QueryString("page")) and len(Request.QueryString("page"))<=2 and Request.QueryString("page")<>"" then showpage=clng(Request.QueryString("page"))
		
			if clng(showpage and 1)<> 0 then
			%>
		
				<span style="font-weight:bold">基本配置：</span><br/><br/>
				<%tstatus=rs("status")%>
				留言本状态：　　　　　<input type="radio" name="status1" value="1" id="status11"<%=cked(clng(tstatus and 1)<>0)%> /><label for="status11">开启</label>　　<input type="radio" name="status1" value="0" id="status12"<%=cked(clng(tstatus and 1)=0)%> /><label for="status12">关闭</label> (关闭时"访客留言权限"设置无效)<br/><br/>

				访客留言权限：　　　　<input type="radio" name="status2" value="2" id="status21"<%=cked(clng(tstatus and 2)<>0)%> /><label for="status21">开启</label>　　<input type="radio" name="status2" value="0" id="status22"<%=cked(clng(tstatus and 2)=0)%> /><label for="status22">关闭</label><br/><br/>

				访客搜索留言权限：　　<input type="radio" name="status3" value="4" id="status31"<%=cked(clng(tstatus and 4)<>0)%> /><label for="status31">开启</label>　　<input type="radio" name="status3" value="0" id="status32"<%=cked(clng(tstatus and 4)=0)%> /><label for="status32">关闭</label><br/><br/>

				访客头像功能：　　　　<input type="radio" name="status7" value="8" id="status71"<%=cked(clng(tstatus and 8)<>0)%> /><label for="status71">开启</label>　　<input type="radio" name="status7" value="0" id="status72"<%=cked(clng(tstatus and 8)=0)%> /><label for="status72">关闭</label><br/><br/>

				留言显示前需审核：　　<input type="radio" name="status4" value="16" id="status41"<%=cked(clng(tstatus and 16)<>0)%> /><label for="status41">审核</label>　　<input type="radio" name="status4" value="0" id="status42"<%=cked(clng(tstatus and 16)=0)%> /><label for="status42">不审核</label><br/><br/>

				悄悄话功能：　　　　　<input type="radio" name="status5" value="32" id="status51"<%=cked(clng(tstatus and 32)<>0)%> /><label for="status51">开启</label>　　<input type="radio" name="status5" value="0" id="status52"<%=cked(clng(tstatus and 32)=0)%> /><label for="status52">关闭</label><br/><br/>

				访客为悄悄话加密：　　<input type="radio" name="status6" value="64" id="status61"<%=cked(clng(tstatus and 64)<>0)%> /><label for="status61">允许</label>　　<input type="radio" name="status6" value="0" id="status62"<%=cked(clng(tstatus and 64)=0)%> /><label for="status62">禁止</label><br/><br/>

				允许访客回复：　　　　<input type="radio" name="status8" value="128" id="status81"<%=cked(clng(tstatus and 128)<>0)%> /><label for="status81">允许</label>　　<input type="radio" name="status8" value="0" id="status82"<%=cked(clng(tstatus and 128)=0)%> /><label for="status82">禁止</label><br/><br/>

				所属网站Logo地址：　　<input type="text" size="<%=ConfigTextWidth%>" maxlength="127" name="homelogo" value="<%=rs("homelogo")%>" /><br/><br/>
				所属网站名称：　　　　<input type="text" size="<%=ConfigTextWidth%>" maxlength="30" name="homename" value="<%=rs("homename")%>" /><br/><br/>
				所属网站地址：　　　　<input type="text" size="<%=ConfigTextWidth%>" maxlength="127" name="homeaddr" value="<%=rs("homeaddr")%>" /><br/><br/>

				<%adminlimit=rs("adminhtml")%>
				管理员安全性设置：　　<input type="checkbox" value="1" name="adminviewcode" id="adminviewcode"<%=cked(clng(adminlimit and 8)<>0)%> /><label for="adminviewcode">访客留言显示实际HTML或UBB代码</label><br/><br/>
				管理员默认HTML权限：　<input type="checkbox" value="1" name="adminhtml" id="adminhtml"<%=cked(clng(adminlimit and 1)<>0)%> /><label for="adminhtml">管理员回复、公告默认支持HTML</label><br/>
				　　　　　　　　　　　<input type="checkbox" value="1" name="adminubb" id="adminubb"<%=cked(clng(adminlimit and 2)<>0)%> /><label for="adminubb">管理员回复、公告默认支持UBB</label><br/>
				　　　　　　　　　　　<input type="checkbox" value="1" name="adminertn" id="adminertn"<%=cked(clng(adminlimit and 4)<>0)%> /><label for="adminertn">管理员不支持HTML和UBB时，默认允许回车换行</label><br/><br/>

				<%guestlimit=rs("guesthtml")%>
				访客HTML权限：　　　　<input type="checkbox" value="1" name="guesthtml" id="guesthtml"<%=cked(clng(guestlimit and 1)<>0)%> /><label for="guesthtml">访客留言支持HTML</label><br/>
				　　　　　　　　　　　<input type="checkbox" value="1" name="guestubb" id="guestubb"<%=cked(clng(guestlimit and 2)<>0)%> /><label for="guestubb">访客留言支持UBB</label><br/>
				　　　　　　　　　　　<input type="checkbox" value="1" name="guestertn" id="guestertn"<%=cked(clng(guestlimit and 4)<>0)%> /><label for="guestertn">访客不支持HTML和UBB时，允许回车换行</label><br/><br/>
		
				管理员登录超时：　　　<input type="text" size="4" maxlength="4" name="admintimeout" value="<%=rs("admintimeout")%>" />分 (默认=20)<br/><br/>
				为访客显示IP前：　　　<input type="text" size="4" maxlength="1" name="showip" value="<%=rs("showip")%>" />位 (可选值：0～4)<br/><br/>
				为管理员显示IP前：　　<input type="text" size="4" maxlength="1" name="adminshowip" value="<%=rs("adminshowip")%>" />位 (可选值：0～4)<br/><br/>
				为管理员显示原IP前：　<input type="text" size="4" maxlength="1" name="adminshoworiginalip" value="<%=rs("adminshoworiginalip")%>" />位 (可选值：0～4,使用代理服务器时此项显示原始IP)<br/><br/>
				
				登录验证码长度：　　　<input type="text" size="4" maxlength="2" name="vcodecount" value="<%=clng(rs("vcodecount") and &H0F)%>" />位 (可选值：0～10)<br/><br/>
				留言验证码长度：　　　<input type="text" size="4" maxlength="2" name="writevcodecount" value="<%=clng(rs("vcodecount") and &HF0) \ &H10%>" />位 (可选值：0～10)<br/><br/><br/>
			<%
			end if
		
			if clng(showpage and 2)<> 0 then
			%>
				<%tmailflag=rs("mailflag")%>
				<span style="font-weight:bold">邮件通知：</span><br/><br/>
				新留言到达通知：　　　<input type="checkbox" value="1" name="mailnewinform" id="mailnewinform"<%=cked(clng(tmailflag and 1)<>0)%> /><label for="mailnewinform">启用</label><br/><br/>
				版主回复通知选项：　　<input type="checkbox" value="1" name="mailreplyinform" id="mailreplyinform"<%=cked(clng(tmailflag and 2)<>0)%> /><label for="mailreplyinform">开启</label><br/><br/>
				邮件发送组件：　　　　<input type="radio" value="0" name="mailcomponent" id="mailcomponent0"<%=cked(clng(tmailflag and 4)=0)%> /><label for="mailcomponent0">JMail</label>　<input type="radio" value="1" name="mailcomponent" id="mailcomponent1"<%=cked(clng(tmailflag and 4)<>0)%> /><label for="mailcomponent1">CDO</label><br/><br/>
				新留言通知接收地址：　<input type="text" size="<%=ConfigTextWidth%>" maxlength="48" name="mailreceive" value="<%=rs("mailreceive")%>" /><br/><br/>
				发件人地址：　　　　　<input type="text" size="<%=ConfigTextWidth%>" maxlength="48" name="mailfrom" value="<%=rs("mailfrom")%>" /><br/><br/>
				发件人SMTP服务器地址：<input type="text" size="<%=ConfigTextWidth%>" maxlength="48" name="mailsmtpserver" value="<%=rs("mailsmtpserver")%>" /><br/><br/>
				登录用户名(如需要)：　<input type="text" size="<%=ConfigTextWidth%>" maxlength="48" name="mailuserid" value="<%=rs("mailuserid")%>" /><br/><br/>
				登录密码(如需要)：　　<input type="password" size="<%=ConfigTextWidth%>" maxlength="48" name="mailuserpass" value="<%=rs("mailuserpass")%>" /><br/><br/>
				邮件紧急程度：　　　　<input type="text" size="4" maxlength="1" name="maillevel" value="<%=rs("maillevel")%>" /> (1最快～5最慢)<br/><br/><br/>
			<%end if
		
			if clng(showpage and 4)<> 0 then
			%>
		
				<span style="font-weight:bold">界面尺寸：</span><br/><br/>
				字体名：　　　　　　　<input type="text" size="10" maxlength="30" name="cssfontfamily" value="<%=rs("cssfontfamily")%>" /><br/><br/>
				字体大小：　　　　　　<input type="text" size="10" maxlength="30" name="cssfontsize" value="<%=rs("cssfontsize")%>" /><br/><br/>
				文字行间距：　　　　　<input type="text" size="10" maxlength="30" name="csslineheight" value="<%=rs("csslineheight")%>" /><br/><br/>			
				留言本总宽度：　　　　<input type="text" size="10" maxlength="5" name="tablewidth" value="<%=rs("tablewidth")%>" /> (默认=630,可用百分比)<br/><br/>
				窗格区块间距：　　　　<input type="text" size="10" maxlength="3" name="windowspace" value="<%=rs("windowspace")%>" /> (默认=20,单位:象素)<br/><br/>
				留言本左窗格宽度：　　<input type="text" size="10" maxlength="5" name="tableleftwidth" value="<%=rs("tableleftwidth")%>" /> (默认=150,可用百分比)<br/><br/>
				“写留言”文本框宽度：<input type="text" size="10" maxlength="3" name="leavetextwidth" value="<%=rs("leavetextwidth")%>" /> (默认=20,单位:字母宽度)<br/><br/>
				“验证码”文本框宽度：<input type="text" size="10" maxlength="3" name="leavevcodewidth" value="<%=rs("leavevcodewidth")%>" /> (默认=10,单位:字母宽度)<br/><br/>
				“留言内容”文本宽度：<input type="text" size="10" maxlength="3" name="leavecontentwidth" value="<%=rs("leavecontentwidth")%>" /> (默认=50,单位:字母宽度)<br/><br/>
				“留言内容”文本高度：<input type="text" size="10" maxlength="3" name="leavecontentheight" value="<%=rs("leavecontentheight")%>" /> (默认=7,单位:字母高度)<br/><br/>
				UBB 工具栏宽度：　　　<input type="text" size="10" maxlength="4" name="ubbtoolwidth" value="<%=rs("ubbtoolwidth")%>" /> (默认=320,单位:象素)<br/><br/>
				UBB 工具栏高度：　　　<input type="text" size="10" maxlength="4" name="ubbtoolheight" value="<%=rs("ubbtoolheight")%>" /> (默认=100,单位:象素)<br/><br/>
				搜索框宽度：　　　　　<input type="text" size="10" maxlength="3" name="searchtextwidth" value="<%=rs("searchtextwidth")%>" /> (默认=20,单位:字母宽度)<br/><br/>
				“高级删除”中文本宽：<input type="text" size="10" maxlength="3" name="advdeltextwidth" value="<%=rs("advdeltextwidth")%>" /> (默认=20,单位:字母宽度)<br/><br/>
				“修改资料”中文本宽：<input type="text" size="10" maxlength="3" name="setinfotextwidth" value="<%=rs("setinfotextwidth")%>" /> (默认=40,单位:字母宽度,包括密码框)<br/><br/>
				“参数”中文本宽度：　<input type="text" size="10" maxlength="3" name="configtextwidth" value="<%=rs("configtextwidth")%>" /> (默认=75,单位:字母宽度,即"基本配置"页的文本)<br/><br/>
				“内容过滤”中文本宽：<input type="text" size="10" maxlength="3" name="filtertextwidth" value="<%=rs("filtertextwidth")%>" /> (默认=62,单位:字母宽度)<br/><br/>
				回复、公告编辑框宽度：<input type="text" size="10" maxlength="3" name="replytextwidth" value="<%=rs("replytextwidth")%>" /> (默认=98,单位:字母宽度)<br/><br/>
				回复、公告编辑框高度：<input type="text" size="10" maxlength="3" name="replytextheight" value="<%=rs("replytextheight")%>" /> (默认=10,单位:字母高度)<br/><br/>

				每页显示的留言数：　　<input type="text" size="10" maxlength="5" name="itemsperpage" value="<%=rs("itemsperpage")%>" /> (默认=5)<br/><br/>
				每页显示的标题数：　　<input type="text" size="10" maxlength="5" name="titlesperpage" value="<%=rs("titlesperpage")%>" /> (默认=20)<br/><br/>
				头像每行显示的数目：　<input type="text" size="10" maxlength="3" name="picturesperrow" value="<%=rs("picturesperrow")%>" /> (默认=5)<br/><br/>
				少量载入的头像数：　　<input type="text" size="10" maxlength="3" name="frequentfacecount" value="<%=rs("frequentfacecount")%>" /> (默认=15,单击"更多头像"显示全部)<br/><br/><br/>
		
			<%
			end if
		
			if clng(showpage and 8)<> 0 then
			%>	
		
				<span style="font-weight:bold">功能设置：</span><br/><br/>
				<%tvisualflag=rs("visualflag")%>
				默认版面模式：　　　　<input type="radio" name="displaymode" value="1" id="displaymode1"<%=cked(clng(tvisualflag and 1024)<>0)%> /><label for="displaymode1">论坛列表样式</label>　　<input type="radio" name="displaymode" value="0" id="displaymode2"<%=cked(clng(tvisualflag and 1024)=0)%> /><label for="displaymode2">留言本样式</label><br/><br/>
				回复内容显示位置：　　<input type="radio" name="replyinword" value="1" id="replyinword1"<%=cked(clng(tvisualflag and 1)<>0)%> /><label for="replyinword1">内嵌于访客留言</label>　　<input type="radio" name="replyinword" value="0" id="replyinword2"<%=cked(clng(tvisualflag and 1)=0)%> /><label for="replyinword2">显示在访客留言下方</label><br/><br/>
				隐藏内容被隐藏的留言：<input type="radio" name="hidehidden" value="1" id="hidehidden1"<%=cked(clng(tvisualflag and 128)<>0)%> /><label for="hidehidden1" >隐藏</label>　　<input type="radio" name="hidehidden" value="0" id="hidehidden2"<%=cked(clng(tvisualflag and 128)=0)%> /><label for="hidehidden2" >显示</label><br/><br/>
				隐藏待审核留言：　　　<input type="radio" name="hideaudit" value="1" id="hideaudit1"<%=cked(clng(tvisualflag and 256)<>0)%> /><label for="hideaudit1" >隐藏</label>　　<input type="radio" name="hideaudit" value="0" id="hideaudit2"<%=cked(clng(tvisualflag and 256)=0)%> /><label for="hideaudit2" >显示</label><br/><br/>
				隐藏版主未回复悄悄话：<input type="radio" name="hidewhisper" value="1" id="hidewhisper1"<%=cked(clng(tvisualflag and 512)<>0)%> /><label for="hidewhisper1" >隐藏</label>　　<input type="radio" name="hidewhisper" value="0" id="hidewhisper2"<%=cked(clng(tvisualflag and 512)=0)%> /><label for="hidewhisper2" >显示</label><br/><br/>
				Logo显示模式：　　　　<input type="radio" name="logobannermode" value="1" id="logobannermode1"<%=cked(clng(tvisualflag and 2048)<>0)%> /><label for="logobannermode1" >横幅模式</label>　　<input type="radio" name="logobannermode" value="0" id="logobannermode2"<%=cked(clng(tvisualflag and 2048)=0)%> /><label for="logobannermode2" >图片模式</label><br/><br/>
				分页窗口显示位置：　　<input type="radio" name="showpagelist" value="3" id="showpagelist3"<%=cked(clng(tvisualflag and 12)=12)%> /><label for="showpagelist3">上下方</label>　<input type="radio" name="showpagelist" value="1" id="showpagelist1"<%=cked(clng(tvisualflag and 12)=4)%> /><label for="showpagelist1">上方</label>　　<input type="radio" name="showpagelist" value="2" id="showpagelist2"<%=cked(clng(tvisualflag and 12)=8)%> /><label for="showpagelist2">下方</label><br/><br/>
				访客搜索窗口显示位置：<input type="radio" name="showsearchbox" value="3" id="showsearchbox3"<%=cked(clng(tvisualflag and 48)=48)%> /><label for="showsearchbox3">上下方</label>　<input type="radio" name="showsearchbox" value="1" id="showsearchbox1"<%=cked(clng(tvisualflag and 48)=16)%> /><label for="showsearchbox1">上方</label>　　<input type="radio" name="showsearchbox" value="2" id="showsearchbox2"<%=cked(clng(tvisualflag and 48)=32)%> /><label for="showsearchbox2">下方</label><br/><br/>
				分页列表模式：　　　　<input type="radio" name="advpagelist" value="1" id="advpagelist1"<%=cked(clng(tvisualflag and 64)<>0)%> /><label for="advpagelist1">区段式</label>　<input type="radio" name="advpagelist" value="0" id="advpagelist2"<%=cked(clng(tvisualflag and 64)=0)%> /><label for="advpagelist2">平面式</label><br/><br/>
				
				区段式分页项数：　　　<input type="text" size="5" maxlength="3" name="advpagelistcount" value="<%=rs("advpagelistcount")%>" /> (默认=10)<br/><br/>
				
				UBB 工具栏：　　　　　<input type="radio" name="showubbtool" value="1" id="showubbtool1"<%=cked(clng(tvisualflag and 2)<>0)%> /><label for="showubbtool1">显示</label>　　<input type="radio" name="showubbtool" value="0" id="showubbtool2"<%=cked(clng(tvisualflag and 2)=0)%> /><label for="showubbtool2">隐藏</label><br/><br/>
				UBB开关(须启用UBB)：　<input type="checkbox" name="ubbflag_image" id="ubbflag_image" value="1"<%=cked(UbbFlag_image)%> /><label for="ubbflag_image">图片</label>　　　
									<input type="checkbox" name="ubbflag_url" id="ubbflag_url" value="1"<%=cked(UbbFlag_url)%> /><label for="ubbflag_url">URL、Email</label>
									<input type="checkbox" name="ubbflag_autourl" id="ubbflag_autourl" value="1"<%=cked(UbbFlag_autourl)%> /><label for="ubbflag_autourl">自动识别网址</label><br/>

				　　　　　　　　　　　<input type="checkbox" name="ubbflag_player" id="ubbflag_player" value="1"<%=cked(UbbFlag_player)%> /><label for="ubbflag_player">播放控件</label>　
									<input type="checkbox" name="ubbflag_paragraph" id="ubbflag_paragraph" value="1"<%=cked(UbbFlag_paragraph)%> /><label for="ubbflag_paragraph">段落样式</label>　
									<input type="checkbox" name="ubbflag_fontstyle" id="ubbflag_fontstyle" value="1"<%=cked(UbbFlag_fontstyle)%> /><label for="ubbflag_fontstyle">字体样式</label><br/>

				　　　　　　　　　　　<input type="checkbox" name="ubbflag_fontcolor" id="ubbflag_fontcolor" value="1"<%=cked(UbbFlag_fontcolor)%> /><label for="ubbflag_fontcolor">字体颜色</label>　
									<input type="checkbox" name="ubbflag_alignment" id="ubbflag_alignment" value="1"<%=cked(UbbFlag_alignment)%> /><label for="ubbflag_alignment">对齐方式</label>　
									<input type="checkbox" name="ubbflag_movement" id="ubbflag_movement" value="1"<%=cked(UbbFlag_movement)%> /><label for="ubbflag_movement">移动效果(<span title="启用移动效果将无法通过W3C XHTML 1.0 Traditional 认证">反对使用</span>)</label><br/>

				　　　　　　　　　　　<input type="checkbox" name="ubbflag_cssfilter" id="ubbflag_cssfilter" value="1"<%=cked(UbbFlag_cssfilter)%> /><label for="ubbflag_cssfilter">滤镜效果</label>　
									<input type="checkbox" name="ubbflag_face" id="ubbflag_face" value="1"<%=cked(UbbFlag_face)%> /><label for="ubbflag_face">表情图标</label><br/><br/>

				<%talign=rs("tablealign")%>
				留言本对齐方式：　　　<input type="radio" name="tablealign" value="left" id="align1"<%=cked(talign<>"center" and talign<>"right")%> /><label for="align1">左对齐</label>　<input type="radio" name="tablealign" value="center" id="align2"<%=cked(talign="center")%> /><label for="align2">居中</label>　　<input type="radio" name="tablealign" value="right" id="align3"<%=cked(talign="right")%> /><label for="align3">右对齐</label><br/><br/>
				<%tpagecontrol=rs("pagecontrol")%>
				留言本边框线：　　　　<input type="radio" name="showborder" value="1" id="showborder1"<%=cked(clng(tpagecontrol and 1)<>0)%> /><label for="showborder1">显示</label>　　<input type="radio" name="showborder" value="0" id="showborder2"<%=cked(clng(tpagecontrol and 1)=0)%> /><label for="showborder2">隐藏</label><br/><br/>
				留言本总标题：　　　　<input type="radio" name="showtitle" value="1" id="showtitle1"<%=cked(clng(tpagecontrol and 2)<>0)%> /><label for="showtitle1">显示</label>　　<input type="radio" name="showtitle" value="0" id="showtitle2"<%=cked(clng(tpagecontrol and 2)=0)%> /><label for="showtitle2">隐藏</label><br/><br/>
				网页右键菜单：　　　　<input type="radio" name="showcontext" value="1" id="showcontext1"<%=cked(clng(tpagecontrol and 4)<>0)%> /><label for="showcontext1">启用</label>　　<input type="radio" name="showcontext" value="0" id="showcontext2"<%=cked(clng(tpagecontrol and 4)=0)%> /><label for="showcontext2">禁用</label><br/><br/>
				选择网页内容：　　　　<input type="radio" name="selectcontent" value="1" id="selectcontent1"<%=cked(clng(tpagecontrol and 8)<>0)%> /><label for="selectcontent1">允许</label>　　<input type="radio" name="selectcontent" value="0" id="selectcontent2"<%=cked(clng(tpagecontrol and 8)=0)%> /><label for="selectcontent2">禁止</label><br/><br/>
				复制网页(Ctrl-C)：　　<input type="radio" name="copycontent" value="1" id="copycontent1"<%=cked(clng(tpagecontrol and 16)<>0)%> /><label for="copycontent1">允许</label>　　<input type="radio" name="copycontent" value="0" id="copycontent2"<%=cked(clng(tpagecontrol and 16)=0)%> /><label for="copycontent2">禁止</label><br/><br/>
				被IFrame内嵌：　　　　<input type="radio" name="beframed" value="1" id="beframed1"<%=cked(clng(tpagecontrol and 32)<>0)%> /><label for="beframed1">允许</label>　　<input type="radio" name="beframed" value="0" id="beframed2"<%=cked(clng(tpagecontrol and 32)=0)%> /><label for="beframed2">禁止</label><br/><br/>
				留言字数限制：　　　　<input type="text" size="10" maxlength="10" name="wordslimit" value="<%=rs("wordslimit")%>" /> (0=不限,回车换行占2字,按UBB编码前统计)<br/><br/>
				
				<%tdelconfirm=rs("delconfirm")%>
				留言通过审核前提示：　<input type="radio" name="passaudittip" value="1" id="passaudittip1"<%=cked(clng(tdelconfirm and 16)<>0)%> /><label for="passaudittip1">提示</label>　　<input type="radio" name="passaudittip" value="0" id="passaudittip2"<%=cked(clng(tdelconfirm and 16)=0)%> /><label for="passaudittip2">不提示</label><br/><br/>
				选定留言通过审核提示：<input type="radio" name="passseltip" value="1" id="passseltip1"<%=cked(clng(tdelconfirm and 32)<>0)%> /><label for="passseltip1">提示</label>　　<input type="radio" name="passseltip" value="0" id="passseltip2"<%=cked(clng(tdelconfirm and 32)=0)%> /><label for="passseltip2">不提示</label><br/><br/>
				删除留言时提示：　　　<input type="radio" name="deltip" value="1" id="deltip1"<%=cked(clng(tdelconfirm and 1)<>0)%> /><label for="deltip1">提示</label>　　<input type="radio" name="deltip" value="0" id="deltip2"<%=cked(clng(tdelconfirm and 1)=0)%> /><label for="deltip2">不提示</label><br/><br/>
				删除回复时提示：　　　<input type="radio" name="delretip" value="1" id="delretip1"<%=cked(clng(tdelconfirm and 2)<>0)%> /><label for="delretip1">提示</label>　　<input type="radio" name="delretip" value="0" id="delretip2"<%=cked(clng(tdelconfirm and 2)=0)%> /><label for="delretip2">不提示</label><br/><br/>
				删除选定留言时提示：　<input type="radio" name="delseltip" value="1" id="delseltip1"<%=cked(clng(tdelconfirm and 4)<>0)%> /><label for="delseltip1">提示</label>　　<input type="radio" name="delseltip" value="0" id="delseltip2"<%=cked(clng(tdelconfirm and 4)=0)%> /><label for="delseltip2">不提示</label><br/><br/>
				执行高级删除时提示：　<input type="radio" name="deladvtip" value="1" id="deladvtip1"<%=cked(clng(tdelconfirm and 8)<>0)%> /><label for="deladvtip1">提示</label>　　<input type="radio" name="deladvtip" value="0" id="deladvtip2"<%=cked(clng(tdelconfirm and 8)=0)%> /><label for="deladvtip2">不提示</label><br/><br/>
				公开悄悄话时提示：　　<input type="radio" name="pubwhispertip" value="1" id="pubwhispertip1"<%=cked(clng(tdelconfirm and 64)<>0)%> /><label for="pubwhispertip1">提示</label>　　<input type="radio" name="pubwhispertip" value="0" id="pubwhispertip2"<%=cked(clng(tdelconfirm and 64)=0)%> /><label for="pubwhispertip2">不提示</label><br/><br/>
				置顶留言时提示：　　　<input type="radio" name="lock2toptip" value="1" id="lock2toptip1"<%=cked(clng(tdelconfirm and 256)<>0)%> /><label for="lock2toptip1">提示</label>　　<input type="radio" name="lock2toptip" value="0" id="lock2toptip2"<%=cked(clng(tdelconfirm and 256)=0)%> /><label for="lock2toptip2">不提示</label><br/><br/>
				提前留言时提示：　　　<input type="radio" name="bring2toptip" value="1" id="bring2toptip1"<%=cked(clng(tdelconfirm and 128)<>0)%> /><label for="bring2toptip1">提示</label>　　<input type="radio" name="bring2toptip" value="0" id="bring2toptip2"<%=cked(clng(tdelconfirm and 128)=0)%> /><label for="bring2toptip2">不提示</label><br/><br/>
				重置留言顺序时提示：　<input type="radio" name="reordertip" value="1" id="reordertip1"<%=cked(clng(tdelconfirm and 512)<>0)%> /><label for="reordertip1">提示</label>　　<input type="radio" name="reordertip" value="0" id="reordertip2"<%=cked(clng(tdelconfirm and 512)=0)%> /><label for="reordertip2">不提示</label><br/><br/>

				留言本配色方案：　　　<select name="style">
				<%
				stylename=rs("stylename")
				rs.Close
				rs.Open sql_adminconfig_style,cn,,,1

				dim onestyle
				while rs.EOF=false
					onestyle=rs("stylename")
					Response.Write "<option value=" &chr(34)& onestyle &chr(34)
					if onestyle=stylename or stylename="" then Response.Write " selected=""selected"""
					Response.Write ">" &onestyle& "</option>"
					rs.MoveNext
				wend
				%>
				</select><br/><br/><br/>
			<%end if%>
			<input type="hidden" name="page" value="<%=showpage%>" />
			<p style="text-align:center;"><input value="更新数据" type="submit" name="submit1" /></p>
			</form>
			<br/>
		</td>
	</tr>
</table>
<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>

</div>

<script type="text/javascript" defer="defer">
//<![CDATA[

function check()
{
	var tv,showpage=<%=showpage%>;
	if ((showpage & 1) != 0)
	{
		if (isNaN(tv=Number(document.configform.admintimeout.value)))
			{alert('“管理员登录超时”必须为数字。');document.configform.admintimeout.select();return false;}
		else if (tv<1 || tv>1440)
			{alert('“管理员登录超时”必须在1～1440的范围内。');document.configform.admintimeout.select();return false;}

		if (isNaN(tv=Number(document.configform.showip.value)))
			{alert('“为访客显示IP”必须为数字。');document.configform.showip.select();return false;}
		else if (tv<0 || tv>4 || document.configform.showip.value=='')
			{alert('“为访客显示IP”必须在0～4的范围内。');document.configform.showip.select();return false;}

		if (isNaN(tv=Number(document.configform.adminshowip.value)))
			{alert('“为管理员显示IP”必须为数字。');document.configform.adminshowip.select();return false;}
		else if (tv<0 || tv>4 || document.configform.adminshowip.value=='')
			{alert('“为管理员显示IP”必须在0～4的范围内。');document.configform.adminshowip.select();return false;}

		if (isNaN(tv=Number(document.configform.adminshoworiginalip.value)))
			{alert('“为管理员显示原IP”必须为数字');document.configform.adminshoworiginalip.select();return false;}
		else if (tv<0 || tv>4 || document.configform.adminshoworiginalip.value=='')
			{alert('“为管理员显示原IP”必须在0～4的范围内。');document.configform.adminshoworiginalip.select();return false;}

		if (isNaN(tv=Number(document.configform.vcodecount.value)))
			{alert('“登录验证码长度”必须为数字。');document.configform.vcodecount.select();return false;}
		else if (tv<0 || tv>10)
			{alert('“登录验证码长度”必须在0～10的范围内。');document.configform.vcodecount.select();return false;}

		if (isNaN(tv=Number(document.configform.writevcodecount.value)))
			{alert('“留言验证码长度”必须为数字。');document.configform.writevcodecount.select();return false;}
		else if (tv<0 || tv>10)
			{alert('“留言验证码长度”必须在0～10的范围内。');document.configform.writevcodecount.select();return false;}
	}
	if ((showpage & 2) != 0)
	{
		if (isNaN(tv=Number(document.configform.maillevel.value)))
			{alert('“邮件紧急程度”必须为数字。');document.configform.maillevel.select();return false;}
		else if (tv<1 || tv>5)
			{alert('“邮件紧急程度”必须在1～5的范围内。');document.configform.maillevel.select();return false;}	
	}
	if ((showpage & 4) != 0)
	{
		if (isNaN(tv=Number(document.configform.tablewidth.value)))
			{if (/^\d+%$/.test(document.configform.tablewidth.value)==false) {alert('“留言本总宽度”必须为正数或正百分比。');document.configform.tablewidth.select();return false;}}
		else if (tv<1)
			{alert('“留言本总宽度”必须大于零。');document.configform.tablewidth.select();return false;}

		if (isNaN(tv=Number(document.configform.tableleftwidth.value)))
			{if (/^\d+%$/.test(document.configform.tableleftwidth.value)==false) {alert('“留言本左窗格宽度”必须为正数或正百分比。');document.configform.tableleftwidth.select();return false;}}
		else if (tv<1)
			{alert('“留言本左窗格宽度”必须大于零。');document.configform.tableleftwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.windowspace.value)))
			{alert('“窗口区块间距”必须为数字。');document.configform.windowspace.select();return false;}
		else if (tv<1 || tv>255)
			{alert('“窗口区块间距”必须在1～255的范围内。');document.configform.windowspace.select();return false;}

		if (isNaN(tv=Number(document.configform.leavetextwidth.value)))
			{alert('“‘写留言’文本框宽度”必须为数字。');document.configform.leavetextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('“‘写留言’文本框宽度”必须在1～255的范围内。');document.configform.leavetextwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.leavevcodewidth.value)))
			{alert('“‘验证码’文本框宽度”必须为数字。');document.configform.leavevcodewidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('“‘验证码’文本框宽度”必须在1～255的范围内。');document.configform.leavevcodewidth.select();return false;}

		if (isNaN(tv=Number(document.configform.leavecontentwidth.value)))
			{alert('“‘留言内容’文本宽度”必须为数字。');document.configform.leavecontentwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('“‘留言内容’文本宽度”必须在1～255的范围内。');document.configform.leavecontentwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.leavecontentheight.value)))
			{alert('“‘留言内容’文本高度”必须为数字。');document.configform.leavecontentheight.select();return false;}
		else if (tv<1 || tv>255)
			{alert('“‘留言内容’文本高度”必须在1～255的范围内。');document.configform.leavecontentheight.select();return false;}

		if (isNaN(tv=Number(document.configform.ubbtoolwidth.value)))
			{alert('“‘写留言’UBB工具宽”必须为数字。');document.configform.ubbtoolwidth.select();return false;}
		else if (tv<1 || tv>9999)
			{alert('“‘写留言’UBB工具宽”必须在1～9999的范围内。');document.configform.ubbtoolwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.ubbtoolheight.value)))
			{alert('“‘写留言’UBB工具高”必须为数字。');document.configform.ubbtoolheight.select();return false;}
		else if (tv<1 || tv>9999)
			{alert('“‘写留言’UBB工具高”必须在1～9999的范围内。');document.configform.ubbtoolheight.select();return false;}

		if (isNaN(tv=Number(document.configform.searchtextwidth.value)))
			{alert('“搜索框宽度”必须为数字。');document.configform.searchtextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('“搜索框宽度”必须在1～255的范围内。');document.configform.searchtextwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.advdeltextwidth.value)))
			{alert('“‘高级删除’中文本宽”必须为数字。');document.configform.advdeltextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('“‘高级删除’中文本宽”必须在1～255的范围内。');document.configform.advdeltextwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.setinfotextwidth.value)))
			{alert('“‘修改资料’中文本宽”必须为数字。');document.configform.setinfotextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('“‘修改资料’中文本宽”必须在1～255的范围内。');document.configform.setinfotextwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.configtextwidth.value)))
			{alert('“‘参数’中文本宽度”必须为数字。');document.configform.configtextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('“‘参数’中文本宽度”必须在1～255的范围内。');document.configform.configtextwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.filtertextwidth.value)))
			{alert('“‘内容过滤’中文本宽”必须为数字。');document.configform.filtertextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('“‘内容过滤’中文本宽”必须在1～255的范围内。');document.configform.filtertextwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.replytextwidth.value)))
			{alert('“回复、公告编辑框宽度”必须为数字。');document.configform.replytextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('“回复、公告编辑框宽度”必须在1～255的范围内。');document.configform.replytextwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.replytextheight.value)))
			{alert('“回复、公告编辑框高度”必须为数字。');document.configform.replytextheight.select();return false;}
		else if (tv<1 || tv>255)
			{alert('“回复、公告编辑框高度”必须在1～255的范围内。');document.configform.replytextheight.select();return false;}

		if (isNaN(tv=Number(document.configform.itemsperpage.value)))
			{alert('“每页显示的留言数”必须为数字。');document.configform.itemsperpage.select();return false;}
		else if (tv<1 || tv>32767)
			{alert('“每页显示的留言数”必须在1～32767的范围内。');document.configform.itemsperpage.select();return false;}

		if (isNaN(tv=Number(document.configform.titlesperpage.value)))
			{alert('“每页显示的标题数”必须为数字。');document.configform.titlesperpage.select();return false;}
		else if (tv<1 || tv>32767)
			{alert('“每页显示的标题数”必须在1～32767的范围内。');document.configform.titlesperpage.select();return false;}

		if (isNaN(tv=Number(document.configform.picturesperrow.value)))
			{alert('“头像每行显示的数目”必须为数字。');document.configform.picturesperrow.select();return false;}
		else if (tv<1 || tv>255)
			{alert('“头像每行显示的数目”必须在1～255的范围内。');document.configform.picturesperrow.select();return false;}

		if (isNaN(tv=Number(document.configform.frequentfacecount.value)))
			{alert('“少量载入的头像数”必须为数字。');document.configform.frequentfacecount.select();return false;}
		else if (tv<0 || tv>255)
			{alert('“少量载入的头像数”必须在0～255的范围内。');document.configform.frequentfacecount.select();return false;}
	}
	if ((showpage & 8) != 0)
	{
		if (isNaN(tv=Number(document.configform.advpagelistcount.value)))
			{alert('“区段式分页项数”必须为数字。');document.configform.advpagelistcount.select();return false;}
		else if (tv<1 || tv>255 || document.configform.wordslimit.value=='')
			{alert('“区段式分页项数”必须在1～255的范围内。');document.configform.advpagelistcount.select();return false;}

		if (isNaN(tv=Number(document.configform.wordslimit.value)))
			{alert('“留言字数限制”必须为数字。');document.configform.wordslimit.select();return false;}
		else if (tv<0 || tv>2147483647 || document.configform.wordslimit.value=='')
			{alert('“留言字数限制”必须在0～2147483647的范围内。');document.configform.wordslimit.select();return false;}
	}
	document.configform.submit1.disabled=true;
	return true;
}

//]]>
</script>

<!-- #include file="bottom.asp" -->
</body>
</html>