<!-- #include file="config.asp" -->
<link rel="stylesheet" type="text/css" href="css/style.css"/>
<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
<!-- #include file="css/style.asp" -->
<!-- #include file="css/adminstyle.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
sub errbox(byval errmsg)
	Call MessagePage(errmsg,"admin_config.asp?page=" &Request.Form("page"))
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

showpage=15
if isnumeric(Request.Form("page")) and len(Request.Form("page"))<=2 and Request.Form("page")<>"" then showpage=clng(Request.Form("page"))

if isnumeric(request.Form("admintimeout"))=false and clng(showpage and 1)<> 0 then
	errbox "“管理员登录超时”必须为数字。"
elseif isnumeric(request.Form("showipv4"))=false and clng(showpage and 1)<> 0 then
	errbox "“为访客显示IPv4”必须为数字。"
elseif isnumeric(request.Form("showipv6"))=false and clng(showpage and 1)<> 0 then
	errbox "“为访客显示IPv6”必须为数字。"
elseif isnumeric(request.Form("adminshowipv4"))=false and clng(showpage and 1)<> 0 then
	errbox "“为管理员显示IPv4”必须为数字。"
elseif isnumeric(request.Form("adminshowipv6"))=false and clng(showpage and 1)<> 0 then
	errbox "“为管理员显示IPv6”必须为数字。"
elseif isnumeric(request.Form("adminshoworiginalipv4"))=false and clng(showpage and 1)<> 0 then
	errbox "“为管理员显示原IPv4”必须为数字。"
elseif isnumeric(request.Form("adminshoworiginalipv6"))=false and clng(showpage and 1)<> 0 then
	errbox "“为管理员显示原IPv6”必须为数字。"
elseif isnumeric(request.Form("vcodecount"))=false and clng(showpage and 1)<> 0 then
	errbox "“登录验证码长度”必须为数字。"
elseif isnumeric(request.Form("writevcodecount"))=false and clng(showpage and 1)<> 0 then
	errbox "“留言验证码长度”必须为数字。"
elseif isnumeric(request.Form("maillevel"))=false and clng(showpage and 2)<> 0 then
	errbox "“邮件紧急程度”必须为数字。"
elseif isnum(request.Form("tablewidth"))=false and clng(showpage and 4)<> 0 then
	errbox "“留言本最大宽度”必须为数字或百分比。"
elseif isnum(request.Form("tableleftwidth"))=false and clng(showpage and 4)<> 0 then
	errbox "“留言本左窗格宽度”必须为数字或百分比。"
elseif isnumeric(request.Form("windowspace"))=false and clng(showpage and 4)<> 0 then
	errbox "“窗口区块间距”必须为数字。"
elseif isnumeric(request.Form("leavecontentheight"))=false and clng(showpage and 4)<> 0 then
	errbox "“‘留言内容’文本高度”必须为数字。"
elseif isnumeric(request.Form("searchtextwidth"))=false and clng(showpage and 4)<> 0 then
	errbox "“搜索框宽度”必须为数字。"
elseif isnumeric(request.Form("replytextheight"))=false and clng(showpage and 4)<> 0 then
	errbox "“回复、公告编辑框高度”必须为数字。"
elseif isnumeric(request.Form("itemsperpage"))=false and clng(showpage and 4)<> 0 then
	errbox "“每页显示的留言数”必须为数字。"
elseif isnumeric(request.Form("titlesperpage"))=false and clng(showpage and 4)<> 0 then
	errbox "“每页显示的标题数”必须为数字。"
elseif isnumeric(request.Form("picturesperrow"))=false and clng(showpage and 4)<> 0 then
	errbox "“头像每行显示的数目”必须为数字。"
elseif isnumeric(request.Form("frequentfacecount"))=false and clng(showpage and 4)<> 0 then
	errbox "“少量载入的头像数”必须为数字。"
elseif isnumeric(request.Form("advpagelistcount"))=false and clng(showpage and 8)<> 0 then
	errbox "“区段式分页项数”必须为数字。"
elseif isnumeric(request.Form("wordslimit"))=false and clng(showpage and 8)<> 0 then
	errbox "“留言字数限制”必须为数字。"
else
	if clng(showpage and 1)<> 0 then
		tstatus1=Request.Form("status1")
		if tstatus1<>"0" and tstatus1<>"1" then tstatus1=1
				
		tstatus2=Request.Form("status2")
		if tstatus2<>"0" and tstatus2<>"2" then tstatus2=2
				
		tstatus3=Request.Form("status3")
		if tstatus3<>"0" and tstatus3<>"4" then tstatus3=4
				
		tstatus4=Request.Form("status4")
		if tstatus4<>"0" and tstatus4<>"16" then tstatus4=16
				
		tstatus5=Request.Form("status5")
		if tstatus5<>"0" and tstatus5<>"32" then tstatus5=32
				
		tstatus6=Request.Form("status6")
		if tstatus6<>"0" and tstatus6<>"64" then tstatus6=64
				
		tstatus7=Request.Form("status7")
		if tstatus7<>"0" and tstatus7<>"8" then tstatus6=8
				
		tstatus8=Request.Form("status8")
		if tstatus8<>"0" and tstatus8<>"128" then tstatus8=128

		tstatus9=Request.Form("status9")
		if tstatus9<>"0" and tstatus9<>"256" then tstatus9=256

		tstatus=clng(tstatus1)+clng(tstatus2)+clng(tstatus3)+clng(tstatus4)+clng(tstatus5)+clng(tstatus6)+clng(tstatus7)+clng(tstatus8)+clng(tstatus9)
				
		thomelogo=server.htmlEncode(Request.Form("homelogo"))
		thomelogo=textfilter(replace(thomelogo," ",""),true)
		if lcase(left(thomelogo,4))="www." then thomelogo="http://" & thomelogo

		thomename=server.htmlEncode(Request.Form("homename"))
				
		thomeaddr=server.htmlEncode(Request.Form("homeaddr"))
		thomeaddr=textfilter(replace(thomeaddr," ",""),true)
		if lcase(left(thomeaddr,4))="www." then thomeaddr="http://" & thomeaddr
				
		tadminhtml=0
		if Request.Form("adminhtml")="1" then tadminhtml=tadminhtml+1
		if Request.Form("adminubb")="1" then tadminhtml=tadminhtml+2
		if Request.Form("adminertn")="1" then tadminhtml=tadminhtml+4
		if Request.Form("adminviewcode")="1" then tadminhtml=tadminhtml+8

		tguesthtml=0
		if Request.Form("guesthtml")="1" then tguesthtml=tguesthtml+1
		if Request.Form("guestubb")="1" then tguesthtml=tguesthtml+2
		if Request.Form("guestertn")="1" then tguesthtml=tguesthtml+4

		tadmintimeout=Request.Form("admintimeout")
		if len(cstr(tadmintimeout))>4 or isnumeric(tadmintimeout)=false then tadmintimeout=1440
		if clng(tadmintimeout)>1440 then tadmintimeout=1440
		if clng(tadmintimeout)<1 then tadmintimeout="20"

		tshowip=0
		tshowipv4=Request.Form("showipv4")
		if Len(tshowipv4)=0 or IsNumeric(tshowipv4)=false then tshowipv4=2 else tshowipv4=clng(tshowipv4)
		if tshowipv4<0 or tshowipv4>4 then tshowipv4=2
		tshowipv6=Request.Form("showipv6")
		if Len(tshowipv6)=0 or IsNumeric(tshowipv6)=false then tshowipv6=2 else tshowipv6=clng(tshowipv6)
		if tshowipv6<0 or tshowipv6>8 then tshowipv6=2
		tshowip=tshowipv6*16 + tshowipv4

		tadminshowip=0
		tadminshowipv4=Request.Form("adminshowipv4")
		if Len(tadminshowipv4)=0 or IsNumeric(tadminshowipv4)=false then tadminshowipv4=2 else tadminshowipv4=clng(tadminshowipv4)
		if tadminshowipv4<0 or tadminshowipv4>4 then tadminshowipv4=2
		tadminshowipv6=Request.Form("adminshowipv6")
		if Len(tadminshowipv6)=0 or IsNumeric(tadminshowipv6)=false then tadminshowipv6=2 else tadminshowipv6=clng(tadminshowipv6)
		if tadminshowipv6<0 or tadminshowipv6>8 then tadminshowipv6=2
		tadminshowip=tadminshowipv6*16 + tadminshowipv4

		tadminshoworiginalip=0
		tadminshoworiginalipv4=Request.Form("adminshoworiginalipv4")
		if Len(tadminshoworiginalipv4)=0 or IsNumeric(tadminshoworiginalipv4)=false then tadminshoworiginalipv4=2 else tadminshoworiginalipv4=clng(tadminshoworiginalipv4)
		if tadminshoworiginalipv4<0 or tadminshoworiginalipv4>4 then tadminshoworiginalipv4=2
		tadminshoworiginalipv6=Request.Form("adminshoworiginalipv6")
		if Len(tadminshoworiginalipv6)=0 or IsNumeric(tadminshoworiginalipv6)=false then tadminshoworiginalipv6=2 else tadminshoworiginalipv6=clng(tadminshoworiginalipv6)
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

	end if
	
	if clng(showpage and 2)<> 0 then
		tmailflag=0
		if Request.Form("mailnewinform")="1" then tmailflag=tmailflag+1
		if Request.Form("mailreplyinform")="1" then tmailflag=tmailflag+2
		if Request.Form("mailcomponent")="1" then tmailflag=tmailflag+4
		
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
	end if
	
	if clng(showpage and 4)<> 0 then
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

	end if
	
	if clng(showpage and 8)<> 0	then
		tvisualflag=0
		if Request.Form("replyinword")="1" then tvisualflag=tvisualflag+1
		if Request.Form("showubbtool")="1" then tvisualflag=tvisualflag+2
		select case Request.Form("showpagelist") : case "1","2","3"
			tvisualflag=tvisualflag+clng(Request.Form("showpagelist"))*4				'左移2位后累加
		end select
		select case Request.Form("showsearchbox") : case "1","2","3"
			tvisualflag=tvisualflag+clng(Request.Form("showsearchbox"))*16				'左移4位后累加
		end select
		if Request.Form("advpagelist")="1" then tvisualflag=tvisualflag+64
		if Request.Form("hidehidden")="1" then tvisualflag=tvisualflag+128
		if Request.Form("hideaudit")="1" then tvisualflag=tvisualflag+256
		if Request.Form("hidewhisper")="1" then tvisualflag=tvisualflag+512
		if Request.Form("displaymode")="1" then tvisualflag=tvisualflag+1024
		if Request.Form("logobannermode")="1" then tvisualflag=tvisualflag+2048

		tadvpagelistcount=Request.Form("advpagelistcount")
		if len(cstr(tadvpagelistcount))>3 then
			tadvpagelistcount=10
		else
			tadvpagelistcount=clng(tadvpagelistcount)
		end if
		if tadvpagelistcount>255 then tadvpagelistcount=255
		if tadvpagelistcount<1 then tadvpagelistcount=1

		tubbflag=0
		if Request.Form("ubbflag_image")="1" then tubbflag=tubbflag+1
		if Request.Form("ubbflag_url")="1" then tubbflag=tubbflag+2
		if Request.Form("ubbflag_autourl")="1" then tubbflag=tubbflag+4
		if Request.Form("ubbflag_player")="1" then tubbflag=tubbflag+8
		if Request.Form("ubbflag_paragraph")="1" then tubbflag=tubbflag+16
		if Request.Form("ubbflag_fontstyle")="1" then tubbflag=tubbflag+32
		if Request.Form("ubbflag_fontcolor")="1" then tubbflag=tubbflag+64
		if Request.Form("ubbflag_alignment")="1" then tubbflag=tubbflag+128
		'if Request.Form("ubbflag_movement")="1" then tubbflag=tubbflag+256
		'if Request.Form("ubbflag_cssfilter")="1" then tubbflag=tubbflag+512
		if Request.Form("ubbflag_face")="1" then tubbflag=tubbflag+1024
		
		ttablealign=Request.Form("tablealign")
		if ttablealign<>"left" and ttablealign<>"center" and ttablealign<>"right" then ttablealign="left"
				
		tpagecontrol=0			
		if Request.Form("showborder")="1" then tpagecontrol=tpagecontrol+1
		if Request.Form("showtitle")="1" then tpagecontrol=tpagecontrol+2
		if Request.Form("showcontext")="1" then tpagecontrol=tpagecontrol+4
		if Request.Form("selectcontent")="1" then tpagecontrol=tpagecontrol+8
		if Request.Form("copycontent")="1" then tpagecontrol=tpagecontrol+16
		if Request.Form("beframed")="1" then tpagecontrol=tpagecontrol+32

		twordslimit=Request.Form("wordslimit")
		if len(cstr(twordslimit))>10 then twordslimit=0
		if twordslimit>2147483647 then twordslimit=2147483647
		twordslimit=abs(twordslimit)
		
		tdelconfirm=0
		if Request.Form("deltip")="1" then tdelconfirm=tdelconfirm+1
		if Request.Form("delretip")="1" then tdelconfirm=tdelconfirm+2
		if Request.Form("delseltip")="1" then tdelconfirm=tdelconfirm+4
		if Request.Form("deladvtip")="1" then tdelconfirm=tdelconfirm+8
		if Request.Form("passaudittip")="1" then tdelconfirm=tdelconfirm+16
		if Request.Form("passseltip")="1" then tdelconfirm=tdelconfirm+32
		if Request.Form("pubwhispertip")="1" then tdelconfirm=tdelconfirm+64
		if Request.Form("bring2toptip")="1" then tdelconfirm=tdelconfirm+128
		if Request.Form("lock2toptip")="1" then tdelconfirm=tdelconfirm+256
		if Request.Form("reordertip")="1" then tdelconfirm=tdelconfirm+512

		tstyleid=Request.Form("style")
		if isnumeric(tstyleid)=false then tstyleid=0
		tstyleid=clng(tstyleid)
	end if

	set cn1=server.CreateObject("ADODB.Connection")
	set rs1=server.CreateObject("ADODB.Recordset")

	CreateConn cn1,dbtype
	rs1.open sql_adminsaveconfig,cn1,0,3

	if clng(showpage and 1)<> 0	then
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
	end if
	if clng(showpage and 2)<> 0	then
		rs1("mailflag")=tmailflag
		rs1("mailreceive")=tmailreceive
		rs1("mailfrom")=tmailfrom
		rs1("mailsmtpserver")=tmailsmtpserver
		rs1("mailuserid")=tmailuserid
		rs1("mailuserpass")=tmailuserpass
		rs1("maillevel")=tmaillevel
	end if
	if clng(showpage and 4)<> 0	then
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
	end if
	if clng(showpage and 8)<> 0	then
		rs1("visualflag")=tvisualflag
		rs1("advpagelistcount")=tadvpagelistcount
		rs1("ubbflag")=tubbflag
		rs1("tablealign")=ttablealign
		rs1("pagecontrol")=tpagecontrol
		rs1("wordslimit")=twordslimit
		rs1("delconfirm")=tdelconfirm
		rs1("styleid")=tstyleid
	end if
	rs1.Update
	
	rs1.Close : cn1.close : set rs1=nothing : set cn1=nothing
end if

Response.Redirect "admin_config.asp?page=" & showpage
%>