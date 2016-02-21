<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_saveconfig.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<!-- #include file="tips.asp" -->
<%
sub errbox(byval errmsg)
	Call TipsPage(errmsg,"admin_config.asp?page=" &Request.Form("page"))
	Response.End
end sub

function isnum(byval pdnum)
if isnumeric(pdnum)=true then
	isnum=true
	exit function
end if

if right(pdnum,1)="%" then
	pdnum=left(pdnum,len(pdnum)-1)
	if isnumeric(pdnum) then
		isnum=true
		exit function
	end if
end if

isnum=false
end function

Response.Expires=-1

if Not IsNumeric(request.Form("admintimeout")) then
	errbox "“管理员登录超时”必须为数字。"
elseif Not IsNumeric(request.Form("showipv4")) then
	errbox "“为访客显示IPv4”必须为数字。"
elseif Not IsNumeric(request.Form("showipv6")) then
	errbox "“为访客显示IPv6”必须为数字。"
elseif Not IsNumeric(request.Form("adminshowipv4")) then
	errbox "“为管理员显示IPv4”必须为数字。"
elseif Not IsNumeric(request.Form("adminshowipv6")) then
	errbox "“为管理员显示IPv6”必须为数字。"
elseif Not IsNumeric(request.Form("adminshoworiginalipv4")) then
	errbox "“为管理员显示原IPv4”必须为数字。"
elseif Not IsNumeric(request.Form("adminshoworiginalipv6")) then
	errbox "“为管理员显示原IPv6”必须为数字。"
elseif Not IsNumeric(request.Form("vcodecount")) then
	errbox "“登录验证码长度”必须为数字。"
elseif Not IsNumeric(request.Form("writevcodecount")) then
	errbox "“留言验证码长度”必须为数字。"
elseif Not IsNumeric(request.Form("maillevel")) then
	errbox "“邮件紧急程度”必须为数字。"
elseif isnum(request.Form("tablewidth"))=false then
	errbox "“留言本最大宽度”必须为数字或百分比。"
elseif isnum(request.Form("tableleftwidth"))=false then
	errbox "“留言本左窗格宽度”必须为数字或百分比。"
elseif Not IsNumeric(request.Form("windowspace")) then
	errbox "“窗口区块间距”必须为数字。"
elseif Not IsNumeric(request.Form("leavecontentheight")) then
	errbox "“‘留言内容’文本高度”必须为数字。"
elseif Not IsNumeric(request.Form("searchtextwidth")) then
	errbox "“搜索框宽度”必须为数字。"
elseif Not IsNumeric(request.Form("replytextheight")) then
	errbox "“回复、公告编辑框高度”必须为数字。"
elseif Not IsNumeric(request.Form("itemsperpage")) then
	errbox "“每页显示的留言数”必须为数字。"
elseif Not IsNumeric(request.Form("titlesperpage")) then
	errbox "“每页显示的标题数”必须为数字。"
elseif Not IsNumeric(request.Form("picturesperrow")) then
	errbox "“头像每行显示的数目”必须为数字。"
elseif Not IsNumeric(request.Form("frequentfacecount")) then
	errbox "“少量载入的头像数”必须为数字。"
elseif Not IsNumeric(request.Form("advpagelistcount")) then
	errbox "“区段式分页项数”必须为数字。"
elseif Not IsNumeric(request.Form("wordslimit")) then
	errbox "“留言字数限制”必须为数字。"
else
	tstatus=0
	if Request.Form("status1")="1" then tstatus=tstatus OR 1
	if Request.Form("status2")="1" then tstatus=tstatus OR 2
	if Request.Form("status3")="1" then tstatus=tstatus OR 4
	if Request.Form("status4")="1" then tstatus=tstatus OR 8
	if Request.Form("status5")="1" then tstatus=tstatus OR 16
	if Request.Form("status6")="1" then tstatus=tstatus OR 32
	if Request.Form("status7")="1" then tstatus=tstatus OR 64
	if Request.Form("status8")="1" then tstatus=tstatus OR 128
	if Request.Form("status9")="1" then tstatus=tstatus OR 256

	thomelogo=Trim(Request.Form("homelogo"))
	if thomelogo<>"" then
		thomelogo=textfilter(thomelogo,true)
	end if

	thomename=server.htmlEncode(Request.Form("homename"))

	thomeaddr=Trim(Request.Form("homeaddr"))
	if thomeaddr<>"" then
		thomeaddr=textfilter(thomeaddr,true)
	end if

	tadminhtml=0
	if Request.Form("adminhtml")="1" then tadminhtml=tadminhtml OR 1
	if Request.Form("adminubb")="1" then tadminhtml=tadminhtml OR 2
	if Request.Form("adminertn")="1" then tadminhtml=tadminhtml OR 4
	if Request.Form("adminviewcode")="1" then tadminhtml=tadminhtml OR 8

	tguesthtml=0
	if Request.Form("guesthtml")="1" then tguesthtml=tguesthtml OR 1
	if Request.Form("guestubb")="1" then tguesthtml=tguesthtml OR 2
	if Request.Form("guestertn")="1" then tguesthtml=tguesthtml OR 4

	tadmintimeout=Request.Form("admintimeout")
	if len(cstr(tadmintimeout))>4 or Not IsNumeric(tadmintimeout) then tadmintimeout=1440
	if clng(tadmintimeout)>1440 then tadmintimeout=1440
	if clng(tadmintimeout)<1 then tadmintimeout="20"

	tshowip=0
	tshowipv4=Request.Form("showipv4")
	if tshowipv4="" or Not IsNumeric(tshowipv4) then tshowipv4=2 else tshowipv4=clng(tshowipv4)
	if tshowipv4<0 or tshowipv4>4 then tshowipv4=2
	tshowipv6=Request.Form("showipv6")
	if tshowipv6="" or Not IsNumeric(tshowipv6) then tshowipv6=2 else tshowipv6=clng(tshowipv6)
	if tshowipv6<0 or tshowipv6>8 then tshowipv6=2
	tshowip=tshowipv6*16 + tshowipv4

	tadminshowip=0
	tadminshowipv4=Request.Form("adminshowipv4")
	if tadminshowipv4="" or Not IsNumeric(tadminshowipv4) then tadminshowipv4=2 else tadminshowipv4=clng(tadminshowipv4)
	if tadminshowipv4<0 or tadminshowipv4>4 then tadminshowipv4=2
	tadminshowipv6=Request.Form("adminshowipv6")
	if tadminshowipv6="" or Not IsNumeric(tadminshowipv6) then tadminshowipv6=2 else tadminshowipv6=clng(tadminshowipv6)
	if tadminshowipv6<0 or tadminshowipv6>8 then tadminshowipv6=2
	tadminshowip=tadminshowipv6*16 + tadminshowipv4

	tadminshoworiginalip=0
	tadminshoworiginalipv4=Request.Form("adminshoworiginalipv4")
	if tadminshoworiginalipv4="" or Not IsNumeric(tadminshoworiginalipv4) then tadminshoworiginalipv4=2 else tadminshoworiginalipv4=clng(tadminshoworiginalipv4)
	if tadminshoworiginalipv4<0 or tadminshoworiginalipv4>4 then tadminshoworiginalipv4=2
	tadminshoworiginalipv6=Request.Form("adminshoworiginalipv6")
	if tadminshoworiginalipv6="" or Not IsNumeric(tadminshoworiginalipv6) then tadminshoworiginalipv6=2 else tadminshoworiginalipv6=clng(tadminshoworiginalipv6)
	if tadminshoworiginalipv6<0 or tadminshoworiginalipv6>8 then tadminshoworiginalipv6=2
	tadminshoworiginalip=tadminshoworiginalipv6*16 + tadminshoworiginalipv4

	tvcodecount=Request.Form("vcodecount")
	if len(cstr(tvcodecount))>2 then tvcodecount=4
	if clng(tvcodecount)>10 then tvcodecount=4
	if clng(tvcodecount)<0 then tvcodecount=4
	tvcodecount=clng(tvcodecount)

	twritevcodecount=Request.Form("writevcodecount")
	if len(cstr(twritevcodecount))>2 then twritevcodecount=4
	if clng(twritevcodecount)>10 then twritevcodecount=4
	if clng(twritevcodecount)<0 then twritevcodecount=4
	twritevcodecount=clng(twritevcodecount) * &H10


	tmailflag=0
	if Request.Form("mailnewinform")="1" then tmailflag=tmailflag OR 1
	if Request.Form("mailreplyinform")="1" then tmailflag=tmailflag OR 2
	if Request.Form("mailcomponent")="1" then tmailflag=tmailflag OR 4

	tmailreceive=Request.Form("mailreceive")
	if len(tmailreceive)>48 then tmailreceive=left(tmailreceive,48)

	tmailfrom=Request.Form("mailfrom")
	if len(tmailfrom)>48 then tmailfrom=left(tmailfrom,48)

	tmailsmtpserver=Request.Form("mailsmtpserver")
	if len(tmailsmtpserver)>48 then tmailsmtpserver=left(tmailsmtpserver,48)

	tmailuserid=Request.Form("mailuserid")
	if len(tmailuserid)>48 then tmailuserid=left(tmailuserid,48)

	tmailuserpass=Request.Form("mailuserpass")
	if len(tmailuserpass)>48 then tmailuserpass=left(tmailuserpass,48)

	tmaillevel=Request.Form("maillevel")
	if len(cstr(tmaillevel))>1 then maillevel=3
	if clng(tmaillevel)<1 or clng(tmaillevel)>5 then tmaillevel=3


	tcssfontfamily=Request.Form("cssfontfamily")
	if len(tcssfontfamily)>48 then tcssfontfamily=left(tcssfontfamily,48)

	tcssfontsize=Request.Form("cssfontsize")
	if len(tcssfontsize)>8 then tcssfontsize=left(tcssfontsize,8)

	tcsslineheight=Request.Form("csslineheight")
	if len(tcsslineheight)>8 then tcsslineheight=left(tcsslineheight,8)

	ttablewidth=Request.Form("tablewidth")
	if len(cstr(ttablewidth))>5 then ttablewidth="630"
	if isnumeric(ttablewidth) then
		if clng(ttablewidth)<1 then ttablewidth="630"
		ttablewidth=cstr(round(clng(ttablewidth)))
	else
		if left(ttablewidth,1)="-" then ttablewidth=right(ttablewidth,len(ttablewidth)-1)
	end if

	ttableleftwidth=Request.Form("tableleftwidth")
	if len(cstr(ttableleftwidth))>5 then ttableleftwidth="150"
	if isnumeric(ttableleftwidth) then
		if clng(ttableleftwidth)<1 then ttableleftwidth="150"
		ttableleftwidth=cstr(round(clng(ttableleftwidth)))
	else
		if left(ttableleftwidth,1)="-" then ttableleftwidth=right(ttableleftwidth,len(ttableleftwidth)-1)
	end if

	twindowspace=Request.Form("windowspace")
	if len(cstr(twindowspace))>3 then twindowspace=20
	if clng(twindowspace)>255 then twindowspace=20
	if clng(twindowspace)<1 then twindowspace=20

	tleavecontentheight=Request.Form("leavecontentheight")
	if len(cstr(tleavecontentheight))>3 then tleavecontentheight=7
	if clng(tleavecontentheight)>255 then tleavecontentheight=255
	if clng(tleavecontentheight)<1 then tleavecontentheight=7

	tsearchtextwidth=Request.Form("searchtextwidth")
	if len(cstr(tsearchtextwidth))>3 then tsearchtextwidth=20
	if clng(tsearchtextwidth)>255 then tsearchtextwidth=255
	if clng(tsearchtextwidth)<1 then tsearchtextwidth=20

	treplytextheight=Request.Form("replytextheight")
	if len(cstr(treplytextheight))>3 then treplytextheight=62
	if clng(treplytextheight)>255 then treplytextheight=255
	if clng(treplytextheight)<1 then treplytextheight=62

	titemsperpage=Request.Form("itemsperpage")
	if len(cstr(titemsperpage))>5 then titemsperpage=5
	if clng(titemsperpage)>32767 then titemsperpage=32767
	if clng(titemsperpage)<1 then titemsperpage=5

	ttitlesperpage=Request.Form("titlesperpage")
	if len(cstr(ttitlesperpage))>5 then ttitlesperpage=5
	if clng(ttitlesperpage)>32767 then ttitlesperpage=32767
	if clng(ttitlesperpage)<1 then ttitlesperpage=5

	tpicturesperrow=Request.Form("picturesperrow")
	if len(cstr(tpicturesperrow))>3 then tpicturesperrow=6
	if clng(tpicturesperrow)>255 then tpicturesperrow=255
	if clng(tpicturesperrow)<1 then tpicturesperrow=6

	tfrequentfacecount=Request.Form("frequentfacecount")
	if len(cstr(tfrequentfacecount))>3 then tfrequentfacecount=14
	if tfrequentfacecount>255 then tfrequentfacecount=255
	if clng(tfrequentfacecount)<0 then tfrequentfacecount=14
	if clng(tfrequentfacecount)>clng(FaceCount) then tfrequentfacecount=FaceCount

	tstyleid=Request.Form("style")
	if Not IsNumeric(tstyleid) then
		tstyleid=0
	else
		tstyleid=clng(tstyleid)
	end if


	tvisualflag=0
	if Request.Form("replyinword")="1" then tvisualflag=tvisualflag OR 1
	if Request.Form("showubbtool")="1" then tvisualflag=tvisualflag OR 2
	select case Request.Form("showpagelist") : case "1","2","3"
		tvisualflag=tvisualflag OR clng(Request.Form("showpagelist"))*4				'左移2位后累加
	end select
	select case Request.Form("showsearchbox") : case "1","2","3"
		tvisualflag=tvisualflag OR clng(Request.Form("showsearchbox"))*16				'左移4位后累加
	end select
	if Request.Form("advpagelist")="1" then tvisualflag=tvisualflag OR 64
	if Request.Form("hidehidden")="1" then tvisualflag=tvisualflag OR 128
	if Request.Form("hideaudit")="1" then tvisualflag=tvisualflag OR 256
	if Request.Form("hidewhisper")="1" then tvisualflag=tvisualflag OR 512
	if Request.Form("displaymode")="1" then tvisualflag=tvisualflag OR 1024
	if Request.Form("logobannermode")="1" then tvisualflag=tvisualflag OR 2048

	tadvpagelistcount=Request.Form("advpagelistcount")
	if len(cstr(tadvpagelistcount))>3 then
		tadvpagelistcount=10
	else
		tadvpagelistcount=clng(tadvpagelistcount)
	end if
	if tadvpagelistcount>255 then tadvpagelistcount=255
	if tadvpagelistcount<1 then tadvpagelistcount=1

	tubbflag=0
	if Request.Form("ubbflag_image")="1" then tubbflag=tubbflag OR 1
	if Request.Form("ubbflag_url")="1" then tubbflag=tubbflag OR 2
	if Request.Form("ubbflag_autourl")="1" then tubbflag=tubbflag OR 4
	if Request.Form("ubbflag_player")="1" then tubbflag=tubbflag OR 8
	if Request.Form("ubbflag_paragraph")="1" then tubbflag=tubbflag OR 16
	if Request.Form("ubbflag_fontstyle")="1" then tubbflag=tubbflag OR 32
	if Request.Form("ubbflag_fontcolor")="1" then tubbflag=tubbflag OR 64
	if Request.Form("ubbflag_alignment")="1" then tubbflag=tubbflag OR 128
	'if Request.Form("ubbflag_movement")="1" then tubbflag=tubbflag OR 256
	'if Request.Form("ubbflag_cssfilter")="1" then tubbflag=tubbflag OR 512
	if Request.Form("ubbflag_face")="1" then tubbflag=tubbflag OR 1024

	tpagecontrol=0
	'if Request.Form("showborder")="1" then tpagecontrol=tpagecontrol OR 1
	if Request.Form("showtitle")="1" then tpagecontrol=tpagecontrol OR 2
	if Request.Form("showcontext")="1" then tpagecontrol=tpagecontrol OR 4
	if Request.Form("selectcontent")="1" then tpagecontrol=tpagecontrol OR 8
	if Request.Form("copycontent")="1" then tpagecontrol=tpagecontrol OR 16
	if Request.Form("beframed")="1" then tpagecontrol=tpagecontrol OR 32

	twordslimit=Request.Form("wordslimit")
	if len(cstr(twordslimit))>10 then twordslimit=0
	if twordslimit>2147483647 then twordslimit=2147483647
	twordslimit=abs(twordslimit)

	tdelconfirm=0
	if Request.Form("deltip")="1" then tdelconfirm=tdelconfirm OR 1
	if Request.Form("delretip")="1" then tdelconfirm=tdelconfirm OR 2
	if Request.Form("delseltip")="1" then tdelconfirm=tdelconfirm OR 4
	if Request.Form("deladvtip")="1" then tdelconfirm=tdelconfirm OR 8
	if Request.Form("passaudittip")="1" then tdelconfirm=tdelconfirm OR 16
	if Request.Form("passseltip")="1" then tdelconfirm=tdelconfirm OR 32
	if Request.Form("pubwhispertip")="1" then tdelconfirm=tdelconfirm OR 64
	if Request.Form("bring2toptip")="1" then tdelconfirm=tdelconfirm OR 128
	if Request.Form("lock2toptip")="1" then tdelconfirm=tdelconfirm OR 256
	if Request.Form("reordertip")="1" then tdelconfirm=tdelconfirm OR 512

	set cn1=server.CreateObject("ADODB.Connection")
	set rs1=server.CreateObject("ADODB.Recordset")

	Call CreateConn(cn1)
	rs1.open sql_adminsaveconfig,cn1,0,3

	rs1("status")=tstatus
	rs1("homelogo")=thomelogo
	rs1("homename")=thomename
	rs1("homeaddr")=thomeaddr
	rs1("adminhtml")=tadminhtml
	rs1("guesthtml")=tguesthtml
	rs1("admintimeout")=tadmintimeout
	rs1("showip")=tshowip
	rs1("adminshowip")=tadminshowip
	rs1("adminshoworiginalip")=tadminshoworiginalip
	rs1("vcodecount")=tvcodecount + twritevcodecount

	rs1("mailflag")=tmailflag
	rs1("mailreceive")=tmailreceive
	rs1("mailfrom")=tmailfrom
	rs1("mailsmtpserver")=tmailsmtpserver
	rs1("mailuserid")=tmailuserid
	rs1("mailuserpass")=tmailuserpass
	rs1("maillevel")=tmaillevel

	rs1("cssfontfamily")=tcssfontfamily
	rs1("cssfontsize")=tcssfontsize
	rs1("csslineheight")=tcsslineheight
	rs1("tablewidth")=ttablewidth
	rs1("tableleftwidth")=ttableleftwidth
	rs1("windowspace")=twindowspace
	rs1("leavecontentheight")=tleavecontentheight
	rs1("searchtextwidth")=tsearchtextwidth
	rs1("replytextheight")=treplytextheight
	rs1("itemsperpage")=titemsperpage
	rs1("titlesperpage")=ttitlesperpage
	rs1("picturesperrow")=tpicturesperrow
	rs1("frequentfacecount")=tfrequentfacecount
	rs1("styleid")=tstyleid

	rs1("visualflag")=tvisualflag
	rs1("advpagelistcount")=tadvpagelistcount
	rs1("ubbflag")=tubbflag
	rs1("pagecontrol")=tpagecontrol
	rs1("wordslimit")=twordslimit
	rs1("delconfirm")=tdelconfirm

	rs1.Update
	rs1.Close : cn1.close : set rs1=nothing : set cn1=nothing
end if

Response.Redirect "admin_config.asp?tabIndex=" & Request.Form("tabIndex")
%>