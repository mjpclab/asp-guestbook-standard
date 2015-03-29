<!-- #include file="common.asp" -->
<%
dim lcn,lrs
set lcn=server.CreateObject("ADODB.Connection")
set lrs=server.CreateObject("ADODB.Recordset")
CreateConn lcn,dbtype
lrs.Open sql_loadconfig_config,lcn,0,1,1

status=lrs("status")
if clng(status and 1)<>0 then	'留言本开启
	StatusOpen=true
else
	StatusOpen=false
end if
if clng(status and 2)<>0 then	'留言权限开启
	StatusWrite=true
else
	StatusWrite=false
end if
if clng(status and 4)<>0 then	'搜索权限开启
	StatusSearch=true
else
	StatusSearch=false
end if
if clng(status and 8)<>0 then	'访客头像显示开关
	StatusShowHead=true
else
	StatusShowHead=false
end if
if clng(status and 16)<>0 then	'留言需要审核
	StatusNeedAudit=true
else
	StatusNeedAudit=false
end if
if clng(status and 32)<>0 then	'允许悄悄话
	StatusWhisper=true
else
	StatusWhisper=false
end if
if clng(status and 64)<>0 then	'允许访客加密悄悄话回复
	StatusEncryptWhisper=true
else
	StatusEncryptWhisper=false
end if
if clng(status and 128)<>0 then	'允许访客回复
	StatusGuestReply=true
else
	StatusGuestReply=false
end if
if clng(status and 256)<>0 then	'开启统计
	StatusStatistics=true
else
	StatusStatistics=false
end if

IPConStatus=lrs("ipconstatus")	'IP屏蔽策略，低4位用于IPv4，高4位用于IPv6
IPv4ConStatus=IPConStatus mod 16
if IPv4ConStatus<0 or IPv4ConStatus>2 then IPv4ConStatus=0
IPv6ConStatus=IPConStatus \ 16
if IPv6ConStatus<0 or IPv6ConStatus>2 then IPv6ConStatus=0

HomeLogo=lrs("homelogo")			'网站Logo地址
HomeName=lrs("homename")			'网站名称
HomeAddr=lrs("homeaddr")	'网站地址



'========HTML权限设定========
AdminHTMLSupport=false
AdminUBBSupport=false
AdminAllowNewLine=false
AdminViewCode=false

HTMLSupport=false
UBBSupport=false
AllowNewLine=false

dim adminlimit,guestlimit
adminlimit=lrs("adminhtml")
if clng(adminlimit and 1) <>0 then AdminHTMLSupport=true		'管理员回复、公告默认是否支持HTML True:是 False:否
if clng(adminlimit and 2) <>0 then AdminUBBSupport=true		'管理员回复、公告默认是否支持UBB标记
if clng(adminlimit and 4) <>0 then AdminAllowNewLine=true	'管理员不支持HTML和UBB时，默认是否允许回车换行
if clng(adminlimit and 8) <>0 then AdminViewCode=true		'为管理员显示实际HTML代码

guestlimit=lrs("guesthtml")
if clng(guestlimit and 1) <>0 then HTMLSupport=true			'访客留言是否支持HTML
if clng(guestlimit and 2) <>0 then UBBSupport=true			'访客留言是否支持UBB标记
if clng(guestlimit and 4) <>0 then AllowNewLine=true			'访客不支持HTML和UBB时，是否允许回车换行


'========安全性设置========
AdminTimeOut=lrs("admintimeout")		'管理员登录超时(分)
ShowIP=lrs("showip")			'留言者IP显示，低4位用于IPv4，高4位用于IPv6
ShowIPv4=ShowIP mod 16
ShowIPv6=ShowIP \ 16
AdminShowIP=lrs("adminshowip")	'为管理员显示IP位数，低4位用于IPv4，高4位用于IPv6
AdminShowIPv4=AdminShowIP mod 16
AdminShowIPv6=AdminShowIP \ 16
AdminShowOriginalIP=lrs("adminshoworiginalip")	'为管理员显示原始IP位数，低4位用于IPv4，高4位用于IPv6
AdminShowOriginalIPv4=AdminShowOriginalIP mod 16
AdminShowOriginalIPv6=AdminShowOriginalIP \ 16

VcodeCount=clng(lrs("vcodecount") and &H0F)		'登录验证码长度
WriteVcodeCount=clng(lrs("vcodecount") and &HF0) \ &H10		'留言验证码长度

'========邮件设置========
MailFlag=lrs("mailflag")
if clng(Mailflag and 1)<>0 then MailNewInform=true else MailNewInform=false		'新留言通知
if clng(Mailflag and 2)<>0 then MailReplyInform=true else MailReplyInform=false		'回复通知
if clng(Mailflag and 4)<>0 then MailComponent="cdo" else MailComponent="jmail"
MailReceive=lrs("mailreceive")			'新留言通知接收地址
MailFrom=lrs("mailfrom")				'发件人地址
MailSmtpServer=lrs("mailsmtpserver")	'发件人SMTP服务器地址
MailUserId=lrs("mailuserid")			'登录用户名
MailUserPass=lrs("mailuserpass")		'登录密码
MailLevel=lrs("maillevel")				'紧急程度

'========界面设置========
CssFontFamily=lrs("cssfontfamily")
CssFontSize=lrs("cssfontsize")
CssLineHeight=lrs("csslineheight")

VisualFlag=lrs("visualflag")
if clng(VisualFlag and 1)<>0 then ReplyInWord=true else ReplyInWord=false					'回复内嵌于留言
if clng(VisualFlag and 2)<>0 then ShowUbbTool=true else ShowUbbTool=false					'显示UBB工具栏
if clng(VisualFlag and 4)<>0 then ShowTopPageList=true else ShowTopPageList=false			'上方显示分页
if clng(VisualFlag and 8)<>0 then ShowBottomPageList=true else ShowBottomPageList=false		'下方显示分页
if ShowTopPageList=false and ShowBottomPageList=false then ShowBottomPageList=true
if clng(VisualFlag and 16)<>0 then ShowTopSearchBox=true else ShowTopSearchBox=false		'上方显示搜索
if clng(VisualFlag and 32)<>0 then ShowBottomSearchBox=true else ShowBottomSearchBox=false	'下方显示搜索
if ShowTopSearchBox=false and ShowBottomSearchBox=false then ShowBottomSearchBox=true
if clng(VisualFlag and 64)<>0 then ShowAdvPageList=true else ShowAdvPageList=false			'区段式分页
if clng(VisualFlag and 128)<>0 then HideHidden=true else HideHidden=false					'隐藏内容被隐藏的留言
if clng(VisualFlag and 256)<>0 then HideAudit=true else HideAudit=false						'隐藏待审核留言
if clng(VisualFlag and 512)<>0 then HideWhisper=true else HideWhisper=false					'隐藏版主未回复悄悄话
if clng(VisualFlag and 1024)<>0 then DisplayMode="forum" else DisplayMode="book"			'默认版面模式
if clng(VisualFlag and 2048)<>0 then LogoBannerMode=true else LogoBannerMode=false			'Logo显示模式

AdvPageListCount=lrs("advpagelistcount")	'区段式分页项数

UbbFlag=lrs("ubbflag")
if clng(UbbFlag and 1)<>0 then UbbFlag_image=true else UbbFlag_image=false
if clng(UbbFlag and 2)<>0 then UbbFlag_url=true else UbbFlag_url=false
if clng(UbbFlag and 4)<>0 then UbbFlag_autourl=true else UbbFlag_autourl=false
if clng(UbbFlag and 8)<>0 then UbbFlag_player=true else UbbFlag_player=false
if clng(UbbFlag and 16)<>0 then UbbFlag_paragraph=true else UbbFlag_paragraph=false
if clng(UbbFlag and 32)<>0 then UbbFlag_fontstyle=true else UbbFlag_fontstyle=false
if clng(UbbFlag and 64)<>0 then UbbFlag_fontcolor=true else UbbFlag_fontcolor=false
if clng(UbbFlag and 128)<>0 then UbbFlag_alignment=true else UbbFlag_alignment=false
'if clng(UbbFlag and 256)<>0 then UbbFlag_movement=true else UbbFlag_movement=false
'if clng(UbbFlag and 512)<>0 then UbbFlag_cssfilter=true else UbbFlag_cssfilter=false
if clng(UbbFlag and 1024)<>0 then UbbFlag_face=true else UbbFlag_face=false

TableAlign=lrs("tablealign")		'表格对齐方式 left:左对齐 center:居中 right:右对齐
TableWidth=lrs("tablewidth")		'表格宽度，可用百分比
TableLeftWidth=lrs("tableleftwidth")	'表格左方宽度，可用百分比
WindowSpace=lrs("windowspace")			'窗格区块间距

LeaveTextWidth=lrs("leavetextwidth")				'“写留言”文本宽度
LeaveVcodeWidth=lrs("leavevcodewidth")				'“验证码”文本框宽度
LeaveContentWidth=lrs("leavecontentwidth")			'“留言内容”文本宽度
LeaveContentHeight=lrs("leavecontentheight")		'“留言内容”文本高度
UbbToolWidth=lrs("ubbtoolwidth")					'“写留言”UBB工具宽
UbbToolHeight=lrs("ubbtoolheight")					'“写留言”UBB工具高
SearchTextWidth=lrs("searchtextwidth")				'“留言搜索”文本宽度
AdvDelTextWidth=lrs("advdeltextwidth")				'“高级删除”中文本宽
SetInfoTextWidth=lrs("setinfotextwidth")			'“修改资料”中文本宽
ConfigTextWidth=lrs("configtextwidth")				'“参数”中文本宽度
FilterTextWidth=lrs("filtertextwidth")				'“内容过滤”中文本宽
ReplyTextWidth=lrs("replytextwidth")				'回复、公告编辑框宽度
ReplyTextHeight=lrs("replytextheight")				'回复、公告编辑框高度

ItemsPerPage=lrs("itemsperpage")		'每页显示的留言数
TitlesPerPage=lrs("titlesperpage")		'每页显示的标题数(论坛模式)
PicturesPerRow=lrs("picturesperrow")		'每行显示的头像数
FrequentFaceCount=lrs("frequentfacecount")	'少量载入的头像数
if FrequentFaceCount>FaceCount then FrequentFaceCount=FaceCount

PageControl=clng(lrs("pagecontrol"))
ShowBorder=false
ShowTitle=false
ShowContext=false
SelectContent=false
CopyContent=false
BeFramed=false
if clng(PageControl and 1)<> 0 then ShowBorder=true			'显示边框
if clng(PageControl and 2)<> 0 then ShowTitle=true			'显示标题栏
if clng(PageControl and 4)<> 0 then ShowContext=true		'显示右键菜单
if clng(PageControl and 8)<> 0 then SelectContent=true		'允许选择
if clng(PageControl and 16)<> 0 then CopyContent=true		'允许复制
if clng(PageControl and 32)<> 0 then BeFramed=true			'允许frame
if ShowBorder=true then
	TableBorderWidth=1
else
	TableBorderWidth=0
end if

WordsLimit=abs(lrs("wordslimit"))		'最大留言数

DelTip=false
DelReTip=false
DelSelTip=false
DelAdvTip=false
PassAuditTip=false
PassSelTip=false
PubWhisperTip=false
Bring2TopTip=false
Lock2TopTip=false
ReorderTip=false
DelConfirm=lrs("delconfirm")
if clng(DelConfirm and 1)<>0 then DelTip=true
if clng(DelConfirm and 2)<>0 then DelReTip=true
if clng(DelConfirm and 4)<>0 then DelSelTip=true
if clng(DelConfirm and 8)<>0 then DelAdvTip=true
if clng(DelConfirm and 16)<>0 then PassAuditTip=true
if clng(DelConfirm and 32)<>0 then PassSelTip=true
if clng(DelConfirm and 64)<>0 then PubWhisperTip=true
if clng(DelConfirm and 128)<>0 then Bring2TopTip=true
if clng(DelConfirm and 256)<>0 then Lock2TopTip=true
if clng(DelConfirm and 512)<>0 then ReorderTip=true

styleid=lrs("styleid")

lrs.Close
lrs.Open Replace(sql_loadconfig_style,"{0}",styleid),lcn,0,1,1
if lrs.EOF=true then
	lrs.Close
	lrs.Open sql_loadconfig_top1style ,lcn
end if

PageBackColor=lrs("pagebackcolor")			'网页背景色
PageBackImage=lrs("pagebackimage")			'网页背景图片，使用相对路径

TableBorderColor=lrs("tablebordercolor")		'表格边框颜色
TableBorderColorInactive=lrs("tablebordercolorinactive")		'表格边框颜色（非活动）
TableBorderWidth=TableBorderWidth*lrs("tableborderwidth")		'表格外框粗细
TableBorderPadding=lrs("tableborderpadding")					'表格外框边距
TableBGC=lrs("tablebgc")						'表格主体背景色
TablePic=lrs("tablepic")						'表格主题背景图片

TitleColor=lrs("titlecolor")		'留言本标题文字颜色
TitleBGC=lrs("titlebgc")			'留言本标题背景色
TitlePic=lrs("titlepic")			'留言本标题背景图片

TableBulletinTitleColor=lrs("tablebulletintitlecolor")	'表格－公告标题文字颜色
TableBulletinTitleBGC=lrs("tablebulletintitlebgc")		'表格－公告标题文字颜色
TableBulletinTitlePic=lrs("tablebulletintitlepic")		'表格－公告标题文字颜色

TableTitleColor=lrs("tabletitlecolor")		'表格－留言标题文字颜色
TableTitleBGC=lrs("tabletitlebgc")			'表格－留言/公告标题背景色
TableTitlePic=lrs("tabletitlepic")			'表格－留言/公告标题背景图片

TableContentColor=lrs("tablecontentcolor")		'表格－文章内容文字颜色
TableContentBGC=lrs("tablecontentbgc")			'表格－文章内容背景色

TableReplyColor=lrs("tablereplycolor")		'表格－版主回复标题文字颜色
TableReplyBGC=lrs("tablereplybgc")			'表格－版主回复标题背景色
TableReplyPic=lrs("tablereplypic")			'表格－版主回复标题背景图片

TableReplyContentColor=lrs("tablereplycontentcolor")	'表格－版主回复内容文字颜色

TableGuestInfoColor=lrs("tableguestinfocolor")	'表格－留言者信息框(左方)文字颜色
TableGuestInfoBGC=lrs("tableguestinfobgc")		'表格－留言者信息框背景色
TableGuestInfoPic=lrs("tableguestinfopic")		'表格－留言者信息框背景图片

BlockBGC=lrs("blockbgc")					'内嵌区块背景色

FormColor=lrs("formcolor")
FormBGC=lrs("formbgc")

LinkNormal=lrs("linknormal")		'超链接－正常颜色
LinkHover=lrs("linkhover")			'超链接－鼠标悬停颜色
LinkActive=lrs("linkactive")		'超链接－活动颜色
LinkVisited=lrs("linkvisited")		'超链接－已访问颜色

PageNumColor_Curr=lrs("pagenumcolor_curr")			'当前页号颜色
PageNumColor_Normal=lrs("pagenumcolor_normal")		'一般页号颜色

Additional_Css=lrs("additional_css")	'附加CSS

lrs.Close
lrs.Open sql_loadconfig_floodconfig,lcn,0,1,1
flood_minwait=abs(lrs.Fields("minwait"))
flood_searchrange=abs(lrs.Fields("searchrange"))
flood_searchflag=lrs.Fields("searchflag")
if clng(flood_searchflag and 1)<>0 then flood_sfnewword=true else flood_sfnewword=false
if clng(flood_searchflag and 2)<>0 then flood_sfnewreply=true else flood_sfnewreply=false
if clng(flood_searchflag and 16)<>0 then flood_include=true else flood_include=false
if clng(flood_searchflag and 32)<>0 then flood_equal=true else flood_equal=false
if clng(flood_searchflag and 256)<>0 then flood_sititle=true else flood_sititle=false
if clng(flood_searchflag and 512)<>0 then flood_sicontent=true else flood_sicontent=false


lrs.Close : lcn.Close : set lrs=nothing : set lcn=nothing
%>