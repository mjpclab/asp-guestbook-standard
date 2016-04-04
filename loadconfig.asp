<!-- #include file="include/sql/loadconfig.asp" -->
<%
dim lcn,lrs
set lcn=server.CreateObject("ADODB.Connection")
set lrs=server.CreateObject("ADODB.Recordset")
Call CreateConn(lcn)
lrs.Open sql_loadconfig_config,lcn,0,1,1

status=lrs("status")
StatusOpen=CBool(status and 1)	'���Ա�����
StatusWrite=CBool(status and 2)	'����Ȩ�޿���
StatusSearch=CBool(status and 4)	'����Ȩ�޿���
StatusShowHead=CBool(status and 8)	'�ÿ�ͷ����ʾ����
StatusNeedAudit=CBool(status and 16)	'������Ҫ���
StatusWhisper=CBool(status and 32)	'�������Ļ�
StatusEncryptWhisper=CBool(status and 64)	'����ÿͼ������Ļ��ظ�
StatusGuestReply=CBool(status and 128)	'����ÿͻظ�
StatusStatistics=CBool(status and 256)	'����ͳ��

IPConStatus=lrs("ipconstatus")	'IP���β��ԣ���4λ����IPv4����4λ����IPv6
IPv4ConStatus=IPConStatus mod 16
if IPv4ConStatus<0 or IPv4ConStatus>2 then IPv4ConStatus=0
IPv6ConStatus=IPConStatus \ 16
if IPv6ConStatus<0 or IPv6ConStatus>2 then IPv6ConStatus=0

HomeLogo=lrs("homelogo")			'��վLogo��ַ
HomeName=lrs("homename")			'��վ����
HomeAddr=lrs("homeaddr")	'��վ��ַ



'========HTMLȨ���趨========
dim adminlimit,guestlimit
adminlimit=lrs("adminhtml")
AdminHTMLSupport=CBool(adminlimit and 1)	'����Ա�ظ�������Ĭ���Ƿ�֧��HTML True:�� False:��
AdminUBBSupport=CBool(adminlimit and 2)		'����Ա�ظ�������Ĭ���Ƿ�֧��UBB���
AdminAllowNewLine=CBool(adminlimit and 4)	'����Ա��֧��HTML��UBBʱ��Ĭ���Ƿ�����س�����
AdminViewCode=CBool(adminlimit and 8)		'Ϊ����Ա��ʾʵ��HTML����

guestlimit=lrs("guesthtml")
HTMLSupport=CBool(guestlimit and 1)			'�ÿ������Ƿ�֧��HTML
UBBSupport=CBool(guestlimit and 2)			'�ÿ������Ƿ�֧��UBB���
AllowNewLine=CBool(guestlimit and 4)		'�ÿͲ�֧��HTML��UBBʱ���Ƿ�����س�����


'========��ȫ������========
AdminTimeOut=lrs("admintimeout")		'����Ա��¼��ʱ(��)
ShowIP=lrs("showip")			'������IP��ʾ����4λ����IPv4����4λ����IPv6
ShowIPv4=ShowIP mod 16
ShowIPv6=ShowIP \ 16
AdminShowIP=lrs("adminshowip")	'Ϊ����Ա��ʾIPλ������4λ����IPv4����4λ����IPv6
AdminShowIPv4=AdminShowIP mod 16
AdminShowIPv6=AdminShowIP \ 16
AdminShowOriginalIP=lrs("adminshoworiginalip")	'Ϊ����Ա��ʾԭʼIPλ������4λ����IPv4����4λ����IPv6
AdminShowOriginalIPv4=AdminShowOriginalIP mod 16
AdminShowOriginalIPv6=AdminShowOriginalIP \ 16

VcodeCount=lrs("vcodecount") mod 16		'��¼��֤�볤��
WriteVcodeCount=lrs("vcodecount") \ 16	'������֤�볤��

'========�ʼ�����========
MailFlag=lrs("mailflag")
MailNewInform=CBool(Mailflag and 1)		'������֪ͨ
MailReplyInform=CBool(Mailflag and 2)		'�ظ�֪ͨ
if CBool(Mailflag and 4) then MailComponent="cdo" else MailComponent="jmail"
MailReceive=lrs("mailreceive")			'������֪ͨ���յ�ַ
MailFrom=lrs("mailfrom")				'�����˵�ַ
MailSmtpServer=lrs("mailsmtpserver")	'������SMTP��������ַ
MailUserId=lrs("mailuserid")			'��¼�û���
MailUserPass=lrs("mailuserpass")		'��¼����
MailLevel=lrs("maillevel")				'�����̶�

'========��������========
CssFontFamily=lrs("cssfontfamily")
CssFontSize=lrs("cssfontsize")
CssLineHeight=lrs("csslineheight")

ServerTimezoneOffset=lrs("servertimezoneoffset")
DisplayTimezoneOffset=lrs("displaytimezoneoffset")

VisualFlag=lrs("visualflag")
ReplyInWord=CBool(VisualFlag and 1)					'�ظ���Ƕ������
ShowUbbTool=CBool(VisualFlag and 2)					'��ʾUBB������
ShowTopPageList=CBool(VisualFlag and 4)			'�Ϸ���ʾ��ҳ
ShowBottomPageList=CBool(VisualFlag and 8)		'�·���ʾ��ҳ
if Not ShowTopPageList and Not ShowBottomPageList then ShowBottomPageList=true
ShowTopSearchBox=CBool(VisualFlag and 16)		'�Ϸ���ʾ����
ShowBottomSearchBox=CBool(VisualFlag and 32)	'�·���ʾ����
if Not ShowTopSearchBox and Not ShowBottomSearchBox then ShowBottomSearchBox=true
ShowAdvPageList=CBool(VisualFlag and 64)			'����ʽ��ҳ
HideHidden=CBool(VisualFlag and 128)					'�������ݱ����ص�����
HideAudit=CBool(VisualFlag and 256)						'���ش��������
HideWhisper=CBool(VisualFlag and 512)					'���ذ���δ�ظ����Ļ�
if CBool(VisualFlag and 1024) then DisplayMode="forum" else DisplayMode="book"			'Ĭ�ϰ���ģʽ
LogoBannerMode=CBool(VisualFlag and 2048)			'Logo��ʾģʽ

AdvPageListCount=lrs("advpagelistcount")	'����ʽ��ҳ����

UbbFlag=lrs("ubbflag")
UbbFlag_image=CBool(UbbFlag and 1)
UbbFlag_url=CBool(UbbFlag and 2)
UbbFlag_autourl=CBool(UbbFlag and 4)
UbbFlag_player=CBool(UbbFlag and 8)
UbbFlag_paragraph=CBool(UbbFlag and 16)
UbbFlag_fontstyle=CBool(UbbFlag and 32)
UbbFlag_fontcolor=CBool(UbbFlag and 64)
UbbFlag_alignment=CBool(UbbFlag and 128)
'UbbFlag_movement=CBool(UbbFlag and 256)
'UbbFlag_cssfilter=CBool(UbbFlag and 512)
UbbFlag_face=CBool(UbbFlag and 1024)

TableWidth=lrs("tablewidth")		'����ȣ����ðٷֱ�
TableLeftWidth=lrs("tableleftwidth")	'����󷽿�ȣ����ðٷֱ�
WindowSpace=lrs("windowspace")			'����������

LeaveContentHeight=lrs("leavecontentheight")		'���������ݡ��ı��߶�
SearchTextWidth=lrs("searchtextwidth")				'�������������ı����
ReplyTextHeight=lrs("replytextheight")				'�ظ�������༭��߶�

ItemsPerPage=lrs("itemsperpage")		'ÿҳ��ʾ��������
TitlesPerPage=lrs("titlesperpage")		'ÿҳ��ʾ�ı�����(��̳ģʽ)
PicturesPerRow=lrs("picturesperrow")		'ÿ����ʾ��ͷ����
FrequentFaceCount=lrs("frequentfacecount")	'���������ͷ����
if FrequentFaceCount>FaceCount then FrequentFaceCount=FaceCount

PageControl=clng(lrs("pagecontrol"))
'ShowBorder=CBool(PageControl and 1)		'��ʾ�߿�
ShowTitle=CBool(PageControl and 2)			'��ʾ������
ShowContext=CBool(PageControl and 4)		'��ʾ�Ҽ��˵�
SelectContent=CBool(PageControl and 8)		'����ѡ��
CopyContent=CBool(PageControl and 16)		'������
BeFramed=CBool(PageControl and 32)			'����frame

WordsLimit=abs(lrs("wordslimit"))		'���������

DelConfirm=lrs("delconfirm")
DelTip=CBool(DelConfirm and 1)
DelReTip=CBool(DelConfirm and 2)
DelSelTip=CBool(DelConfirm and 4)
DelAdvTip=CBool(DelConfirm and 8)
PassAuditTip=CBool(DelConfirm and 16)
PassSelTip=CBool(DelConfirm and 32)
PubWhisperTip=CBool(DelConfirm and 64)
Bring2TopTip=CBool(DelConfirm and 128)
Lock2TopTip=CBool(DelConfirm and 256)
ReorderTip=CBool(DelConfirm and 512)

styleid=lrs("styleid")

lrs.Close
lrs.Open Replace(sql_loadconfig_style,"{0}",styleid),lcn,0,1,1
if lrs.EOF then
	lrs.Close
	lrs.Open sql_loadconfig_top1style,lcn,0,1,1
end if

ThemePath=lrs("themepath")

lrs.Close
lrs.Open sql_loadconfig_floodconfig,lcn,0,1,1
flood_minwait=abs(lrs.Fields("minwait"))
flood_searchrange=abs(lrs.Fields("searchrange"))
flood_searchflag=lrs.Fields("searchflag")
flood_sfnewword=CBool(flood_searchflag and 1)
flood_sfnewreply=CBool(flood_searchflag and 2)
flood_include=CBool(flood_searchflag and 16)
flood_equal=CBool(flood_searchflag and 32)
flood_sititle=CBool(flood_searchflag and 256)
flood_sicontent=CBool(flood_searchflag and 512)


lrs.Close : lcn.Close : set lrs=nothing : set lcn=nothing
%>