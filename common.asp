<!-- #include file="ubbcode.asp" -->
<!-- #include file="sql.asp" -->
<%
'===================
function CreateConn(byref tconn,byval tDBType)
	'On Error Resume Next

	select case tDBType
	case 1
		tconn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & server.Mappath(dbfile) & """;Jet OLEDB:Database Password=""" & dbfilepassword & """;"
		tconn.Open
	case 2
		tconn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & server.Mappath(dbfile) & """;Jet OLEDB:Database Password=""" & dbfilepassword & """;"
		tconn.Open
	case 3
		tconn.ConnectionString="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""" & server.Mappath(dbfile) & """;Jet OLEDB:Database Password=""" & dbfilepassword & """;"
		tconn.Open
	case 10
		dim sql_port
		if dbport<>"" then sql_port="," & dbport else sql_port=""
		tconn.ConnectionString= _
			" Provider=SQLOLEDB.1" & _
			";Data Source=""" & dbserver & sql_port & """" & _
			";User Id=""" & dbuserid & """" & _
			";Password=""" & dbpassword & """" & _
			";Application Name=Shen Guest Book Standard - " & InstanceName & _
			""
		if dbcatalog<>"" then tconn.ConnectionString=tconn.ConnectionString & ";Initial Catalog=""" & dbcatalog & """"
		tconn.Open
	case else
		Err.Raise 513,,"无效的数据库类型号(dbtype)"
	end select
	
	'If Err.number<>0 Then
		'Response.Write "连接数据库出错，请检查配置文件。错误描述：<br/><br/>" & Chr(13) & Chr(10)
		'Response.Write Err.Description		'显示错误消息
		'Response.End
	'End If
end function
'===================
function fullsize(byref istr,byref isize)
	Dim strLength
	strLength=Len(istr)
	if strLength<isize then
		fullsize=string(isize-strLength,"0") & istr
	else
		fullsize=istr
	end if
end function
'===================
function IsIPv4Range(byref ipStr)
	if Instr(ipStr,".")>0 then
		IsIPv4Range=true
	elseif Len(ipStr)>0 and IsNumeric(ipStr) then
		IsIPv4Range=true
	else
		IsIPv4Range=false
	end if
end function
'===================
function IsIPv6Range(byref ipStr)
	if Instr(ipStr,":")>0 then
		IsIPv6Range=true
	elseif Len(ipStr)>0 and IsNumeric("&H" & ipStr) then
		IsIPv6Range=true
	else
		IsIPv6Range=false
	end if
end function
'===================
function IsIPv4(byref ipStr)
	IsIPv4=(Instr(ipStr,".")>0)
end function
'===================
function IsIPv6(byref ipStr)
	IsIPv6=(Instr(ipStr,":")>0)
end function
'===================
function IPv4toHex(byref ipStr,byval toMax)
	Dim items,i,hexStr,value,defaultFill
	hexStr=""
	items=Split(ipStr,".")
	ReDim Preserve items(3)

	if toMax then
		defaultFill="FF"
	else
		defaultFill="00"
	end if

	for i=0 to 3
		value=items(i)
		if isnumeric(Cstr(value)) then
			value=clng(value) mod 256
			hexStr=hexStr & fullsize(hex(value),2)
		else
			hexStr=hexStr & defaultFill
		end if
	next

	IPv4toHex=UCase(hexStr)
end function
'===================
function IPv6toHex(byref ipStr,byval toMax)
	ipStr=expandIPv6(ipStr,toMax)
	IPv6toHex=Replace(ipStr,":","")
end function
'===================
function expandIPv6(byval ipStr,byval toMax)
	Dim hexStr,items,itemCount,lackedItemCount,i,j,value,defaultFill
	hexStr=""

	if Left(ipStr,1)=":" then ipStr=Mid(ipStr,2)
	items=Split(ipStr,":")
	itemCount=UBound(items)+1
	if itemCount>8 then itemCount=8
	lackedItemCount=9-itemCount     'optimize for 8-(itemCount-1)
	Redim Preserve items(7)

	if toMax then
		defaultFill="FFFF"
	else
		defaultFill="0000"
	end if

	Dim NeedReplaceX
	NeedReplaceX=false

	for i=0 to 7
		value=items(i)
		if Len(value)>0 then
			value=Left(value,4)
			if IsNumeric("&H" & value) then
				if NeedReplaceX then
					hexStr=Replace(hexStr,"XXXX","0000")
					NeedReplaceX=false
				end if
				hexStr=hexStr & fullsize(value,4) & ":"
			else
				hexStr=hexStr & "XXXX:"
				NeedReplaceX=true
			end if
		else
			for j=1 to lackedItemCount
				hexStr=hexStr & "XXXX:"
				NeedReplaceX=true
			next
		end if
	next

	hexStr=UCase(Left(hexStr,39))
	if NeedReplaceX then
		hexStr=Replace(hexStr,"XXXX",defaultFill)
		NeedReplaceX=false
	end if

	expandIPv6=hexStr
end function
'===================
function compactIPv6(byref ipStr)
	Dim items,maxZeroItems,maxZeroIndex,value,foundZero,zeroItems,zeroIndex,i
	items=Split(ipStr,":")
	maxZeroIndex=-1
	maxZeroItems=0

	foundZero=false
	zeroItems=0
	for i=0 to UBound(items)
		value=items(i)
		if value="0000" then
			if not foundZero then
				zeroIndex=i
			end if
			foundZero=true
			zeroItems=zeroItems+1
			if zeroItems>maxZeroItems then
				maxZeroIndex=zeroIndex
				maxZeroItems=zeroItems
			end if
		else
			foundZero=false
			zeroItems=0
		end if
	next

	Dim result
	result=""
	if maxZeroIndex>-1 then
		Dim compactIndexMin, compactIndexMax
		compactIndexMin=maxZeroIndex
		compactIndexMax=maxZeroIndex+maxZeroItems-1
		for i=0 to UBound(items)
			if i<compactIndexMin or i>compactIndexMax then
				result=result & items(i) & ":"
			else
				if i=0 then
					result=result & ":"
				elseif i=Ubound(items) then
					result=result & ":"
				end if
				if i=compactIndexMax then
					result=result & ":"
				end if
			end if
		next
	else
		for i=0 to UBound(items)
			result=result & items(i) & ":"
		next
	end if

	if Len(result)>0 then
		result=":" & Left(result,Len(result)-1)
	end if
	result=Replace(result,":0",":")
	result=Replace(result,":0",":")
	result=Replace(result,":0",":")
	result=Mid(result,2)

	compactIPv6=UCase(result)
end function
'===================
function IsBannedIPv4(byref ipStr)
	dim cnx,rsx,iphex

	if IPv4ConStatus = 0 then
		IsBannedIPv4=false
	else
		iphex=IPv4ToHex(ipStr,false)
		set cnx=server.CreateObject("ADODB.Connection")
		set rsx=server.CreateObject("ADODB.Recordset")
		CreateConn cnx,dbtype
		rsx.Open Replace(Replace(sql_common_isbanipv4,"{0}",IPv4ConStatus),"{1}",iphex),cnx,,1

		if IPv4ConStatus=1 then
			IsBannedIPv4=not rsx.EOF
		elseif IPv4ConStatus=2 then
			IsBannedIPv4=rsx.EOF
		end if

		rsx.Close
		cnx.Close
		set rsx=nothing
		set cnx=nothing
	end if
end function
'===================
function IsBannedIPv6(byref ipStr)
	dim cnx,rsx,iphex

	if IPv6ConStatus = 0 then
		IsBannedIPv6=false
	else
		iphex=IPv6ToHex(ipStr,false)
		set cnx=server.CreateObject("ADODB.Connection")
		set rsx=server.CreateObject("ADODB.Recordset")
		CreateConn cnx,dbtype
		rsx.Open Replace(Replace(sql_common_isbanipv6,"{0}",IPv6ConStatus),"{1}",iphex),cnx,,1

		if IPv6ConStatus=1 then
			IsBannedIPv6=not rsx.EOF
		elseif IPv6ConStatus=2 then
			IsBannedIPv6=rsx.EOF
		end if

		rsx.Close
		cnx.Close
		set rsx=nothing
		set cnx=nothing
	end if
end function
'===================
function IsBannedIP(byref ipStr)
	if IsIPv4(ipStr) then
		IsBannedIP=IsBannedIPv4(ipStr)
	elseif IsIPv6(ipStr) then
		IsBannedIP=IsBannedIPv6(ipStr)
	else
		IsBannedIP=false
	end if
end function
'===================
function checkIsBannedIP
	if IsBannedIP(Request.ServerVariables("REMOTE_ADDR")) then
		checkIsBannedIP=true
	elseif IsBannedIP(Request.ServerVariables("HTTP_X_FORWARDED_FOR")) then
		checkIsBannedIP=true
	else
		checkIsBannedIP=false
	end if
end function
'===================
function hexToIPv4(byref hexip)
	dim i,strip
	strip=""
	for i=0 to 3
		if i>0 then strip=strip & "."
		strip=strip & cstr(cint("&H" & mid(hexip,i*2+1,2)))
	next

	hexToIPv4=strip
end function
'===================
function hexToIPv6(byref hexip)
	dim i,strip
	strip=""
	for i=0 to 7
		if i>0 then strip=strip & ":"
		strip=strip & mid(hexip,i*4+1,4)
	next

	hexToIPv6=compactIPv6(strip)
end function
'=====================
function addstat(byref fieldname)
set cna=server.CreateObject("ADODB.Connection")
set rsa=server.CreateObject("ADODB.Recordset")

CreateConn cna,dbtype
rsa.open Replace(sql_common_getstat,"{0}",fieldname),cna,0,1,1

if rsa.EOF then
	cna.Execute Replace(sql_common_initstat,"{0}",DateTimeStr(Now())),,1
elseif isdate(rsa.Fields("startdate"))=false then
	cna.Execute Replace(sql_common_updatetime,"{0}",DateTimeStr(Now())),,1
end if
rsa.Close
cna.Execute Replace(sql_common_addstat,"{0}",fieldname),,1

cna.Close
set rsa=nothing
set cna=nothing

end function
'================================
function bodylimit()
dim limitstr
limitstr=""

if ShowContext=false then limitstr=limitstr & " oncontextmenu=""return false;"""
if SelectContent=false then limitstr=limitstr & " onselectstart=""return false;"""
if CopyContent=false then limitstr=limitstr & " oncopy=""return false;"""

bodylimit=limitstr

end function
'==================================
function framecheck()
if BeFramed=false then
	framecheck="if (window.top!=window) window.top.location.href=window.location.href;"
else
	framecheck=""
end if
end function
'==================================
function textfilter(byref str,byval filterScript)
if str="" then
	textfilter=""
else
	dim re,strContent
	set re=new RegExp
	re.IgnoreCase=true
	strContent=str

	if filterScript=true then
		on error resume next
		
		re.pattern="(?:javascript|vbscript)\s*:"
		if err.number=0 then strContent=re.replace(strContent,"")

		re.pattern="(style=.*?):expression"
		if err.number=0 then strContent=re.replace(strContent,"$1")

		re.pattern="(<.*?style\s*=.*?)behavior:(.*?>)"
		if err.number=0 then strContent=re.replace(strContent,"$1$2")

		on error goto 0
	end if

	strContent=replace(strContent,chr(13)&chr(10)," ")
	textfilter=strContent
end if
end function
'==================================
function convertstr(byref str,byval htmlflag,byval allubbflags)
dim tHTML,tUBB,tNewline
tHTML=false
tUBB=false
tNewline=false

if cint(htmlflag and 1)<>0 then tHTML=true
if cint(htmlflag and 2)<>0 then tUBB=true
if cint(htmlflag and 4)<>0 then tNewline=true

if tHTML=false then
	str=server.HTMLEncode(str)
	str=replace(str,chr(9),"        ")
	'str=replace(str," ","&nbsp;")
end if
if tUBB=true then str=ubbcode(str,allubbflags)
if tHTML=false and tUBB=false then
	if tNewline=true then
		str=replace(str,chr(13)&chr(10),"<br/>")
	else
		str=replace(str,chr(13)&chr(10)," ")
	end if
end if
end function
'==================================
function getvcode(byval codelen)
dim rvalue
randomize

rvalue=cstr(fix(abs(rnd*(10^codelen))))
if len(rvalue)<codelen then rvalue=string(codelen-len(rvalue),"0") & rvalue

getvcode=rvalue
end function
'==================================
function GetIPv4(byref strIP, byval itemCount)
	if Len(strIP)=0 then
		GetIPv4=""
	elseif itemCount>=4 then
		GetIPv4=strIP
	else
		Dim sepPos,i,buffer
		sepPos=0
		for i=1 to itemCount
			sepPos=Instr(sepPos+1,strIP,".")
		next
		if sepPos>0 then
			buffer=Mid(strIP,1,sepPos-1)
		end if
		for i=1 to 4-itemCount
			buffer=buffer & ".*"
		next
		GetIPv4=buffer
	end if
end function
'==================================
function GetIPv6(byref strIP, byval itemCount)
	if Len(strIP)=0 then
		GetIPv6=""
	elseif itemCount>=8 then
		GetIPv6=strIP
	else
		Dim sepPos,i,buffer
		sepPos=0

		for i=1 to itemCount
			sepPos=Instr(sepPos+1,strIP,":")
		next
		if sepPos>0 then
			buffer=Mid(strIP,1,sepPos-1)
		end if
		for i=1 to 8-itemCount
			buffer=buffer & ":*"
		next
		GetIPv6=buffer
	end if
end function
'==================================
sub get_divided_page(byref cn,byref rs,byval pk,byval countsql,byval sql,byval keyword,byval request_page,byval ItemsEachPage,byref ItemsCount,byref PagesCount,byref CurrentItemsCount,byref ipage)
if clng(ItemsEachPage)>0 then 
	ItemsEachPage=clng(ItemsEachPage)
else
	ItemsEachPage=10
end if

rs.Open countsql,cn,0,1,1
ItemsCount=rs(0)
rs.Close

if ItemsCount mod ItemsEachPage <>0 then
	PagesCount=ItemsCount \ ItemsEachPage +1
else
	PagesCount=ItemsCount \ ItemsEachPage
end if

if isnumeric(request_page) and request_page<>"" then
	if request_page>2147483647 then
		ipage=PagesCount
	elseif clng(request_page)>PagesCount then
		ipage=PagesCount
	elseif clng(request_page)>0 then
		ipage=clng(request_page)
	else
		ipage=1
	end if
else
	ipage=1
end if

if ipage=PagesCount then
	CurrentItemsCount=ItemsCount - ItemsEachPage*(ipage -1)
else
	CurrentItemsCount=ItemsEachPage
end if

if ItemsCount>0 then
	sql_from=mid(sql,instr(1,sql,"FROM",1))
	
	sql_innerfields=replace(replace(keyword," INC","")," DEC","")
	if instr(sql_innerfields,pk & ",")=0 AND instr(sql_innerfields,"," & pk)=0 AND sql_innerfields<>pk then sql_innerfields=sql_innerfields & "," & pk
	
	where_and=""
	if instr(1,sql,"where",1)>0 then
		where_and=" AND "
	else
		where_and=" WHERE"
	end if
	
	rs.Open Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sql_common_dividepage _
		,"{0}",CurrentItemsCount) _
		,"{1}",ItemsEachPage*(ipage-1)+CurrentItemsCount) _
		,"{2}",sql_innerfields) _
		,"{3}",sql_from) _
		,"{4}",replace(replace(keyword,"INC","ASC"),"DEC","DESC")) _
		,"{5}",replace(replace(keyword,"INC","DESC"),"DEC","ASC")) _
		,"{6}",sql) _
		,"{7}",where_and) _
		,"{8}",pk) _
	,cn,0,1,1
end if
end sub
'==================================
sub show_page_list(byval CurPage,byval PagesCount,byval filename,byval pagetitle,byval param)

if param<>"" then
	if left(param,1)<>"&" then param="&" & param
end if

arr_param=split(param,"&")

if ShowAdvPageList then
	txt_align="center"
	if CurPage-Int(AdvPageListCount/2)<1 then start_page=1 else start_page=CurPage-Int(AdvPageListCount/2)
	if start_page+AdvPageListCount-1>PagesCount then end_page=PagesCount else end_page=start_page+AdvPageListCount-1
	if end_page-start_page+1<AdvPageListCount and start_page>1 then
		if end_page-AdvPageListCount+1>1 then
			start_page=end_page-AdvPageListCount+1
		else
			start_page=1
		end if
	end if
else
	txt_align="left"
	start_page=1
	end_page=PagesCount
end if

'计算各分页控件目标页号
if ShowAdvPageList then
	first_page_no=1
	last_page_no=PagesCount
	if CurPage+1>PagesCount then next_page_no=PagesCount else next_page_no=CurPage+1
	if CurPage-1<1 then prev_page_no=1 else prev_page_no=CurPage-1
	if CurPage+AdvPageListCount>PagesCount then largenext_page_no=PagesCount else largenext_page_no=CurPage+AdvPageListCount
	if CurPage-AdvPageListCount<1 then largeprev_page_no=1 else largeprev_page_no=CurPage-AdvPageListCount
	
	str_first_page=		"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &first_page_no& param & """><img src=""image/icon_page_first.gif"" title=""第" &first_page_no& "页"" /></a>"
	str_largeprev_page=	"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &largeprev_page_no& param & """><img src=""image/icon_page_largeprev.gif"" title=""第" &largeprev_page_no& "页"" /></a>"
	str_prev_page=		"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &prev_page_no& param & """><img src=""image/icon_page_prev.gif"" title=""第" &prev_page_no& "页"" /></a>"
	
	str_last_page=		"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &last_page_no& param & """><img src=""image/icon_page_last.gif"" title=""第" &last_page_no& "页"" /></a>"
	str_largenext_page=	"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &largenext_page_no& param & """><img src=""image/icon_page_largenext.gif"" title=""第" &largenext_page_no& "页"" /></a>"
	str_next_page=		"<a class=""page-control"" name=""page-control"" href=""" &filename& "?page=" &next_page_no& param & """><img src=""image/icon_page_next.gif"" title=""第" &next_page_no& "页"" /></a>"
	
	str_first2_page=	"<a class=""js-page-control"" name=""js-page-control""><img onmouseup=""if(ghandle)clearTimeout(ghandle);"" onmousedown=""if(ghandle)clearTimeout(ghandle); refresh_pagenum(" &(-PagesCount+1)& ");"" src=""image/icon_page_first2.gif"" title=""卷至首页"" /></a>"
	str_largeprev2_page="<a class=""js-page-control"" name=""js-page-control""><img onmouseup=""if(ghandle)clearTimeout(ghandle);"" onmousedown=""if(ghandle)clearTimeout(ghandle); refresh_pagenum(" &(-AdvPageListCount)& ");"" src=""image/icon_page_largeprev2.gif"" class=""pageicon"" title=""上卷" &AdvPageListCount& "页"" /></a>"
	str_prev2_page=		"<a class=""js-page-control"" name=""js-page-control""><img onmouseup=""if(ghandle)clearTimeout(ghandle);"" onmousedown=""if(ghandle)clearTimeout(ghandle); refresh_pagenum(" &(-1)& ");"" src=""image/icon_page_prev2.gif"" class=""pageicon"" title=""上卷1页"" /></a>"

	str_last2_page=		"<a class=""js-page-control"" name=""js-page-control""><img onmouseup=""if(ghandle)clearTimeout(ghandle);"" onmousedown=""if(ghandle)clearTimeout(ghandle); refresh_pagenum(" &(PagesCount-1)& ");"" src=""image/icon_page_last2.gif"" class=""pageicon"" title=""卷至末页"" /></a>"
	str_largenext2_page="<a class=""js-page-control"" name=""js-page-control""><img onmouseup=""if(ghandle)clearTimeout(ghandle);"" onmousedown=""if(ghandle)clearTimeout(ghandle); refresh_pagenum(" &(AdvPageListCount)& ");"" src=""image/icon_page_largenext2.gif"" class=""pageicon"" title=""下卷" &AdvPageListCount& "页"" /></a>"
	str_next2_page=		"<a class=""js-page-control"" name=""js-page-control""><img onmouseup=""if(ghandle)clearTimeout(ghandle);"" onmousedown=""if(ghandle)clearTimeout(ghandle); refresh_pagenum(" &(1)& ");"" src=""image/icon_page_next2.gif"" class=""pageicon"" title=""下卷1页"" /></a>"
end if%>

<div class="region page-list">
	<h3 class="title"><%=pagetitle%></h3>
	<div class="content">
		<%if ShowAdvPageList then%>
			<div class="nav backward-nav"><%=str_first_page & str_largeprev_page & str_prev_page & str_first2_page & str_largeprev2_page & str_prev2_page%></div>
			<div class="nav forward-nav"><%=str_next_page & str_largenext_page & str_last_page & str_next2_page & str_largenext2_page & str_last2_page%></div>
		<%end if%>
		<form method="get" action="<%=filename%>">
		<div class="pagenum-list">
			<%for j=start_page to end_page%><a name="pagenum" class="pagenum<%if j=CurPage then Response.Write " pagenum-current"%>" href="<%=filename%>?page=<%=j & param%>"><%=j%></a><%next%>
		</div>
		<div class="goto">(共<%=PagesCount%>页)　转到页数<input type="text" name="page" class="page" maxlength="10" /> <input type="submit" class="submit" value="GO" /></div>
		<%
		for i=0 to ubound(arr_param)
			if arr_param(i)<>"" then%>
				<%t_name=left(arr_param(i),instr(arr_param(i),"=")-1)
				t_value=right(arr_param(i),len(arr_param(i))-len(t_name)-1)
				if t_name<>"page" then%><input type="hidden" name="<%=t_name%>" value="<%=UrlDecode(t_value)%>" /><%end if%>
			<%end if
		next
		%>
		</form>
	</div>
</div>
<!-- #include file="pagecontrols.inc" -->
<%end sub
'==================================
sub show_book_title(byval layercount, byval param)%>
<div class="header">
	<%if HomeLogo<>"" then%><img class="logo<%if LogoBannerMode then Response.write " banner"%>" src="<%=HomeLogo%>" alt="<%=HomeName%>"/><%end if%>
	<div class="breadcrumb">
		<%if HomeName<>"" then
			if HomeAddr<>"" then%>
				<a class="name" href="<%=HomeAddr%>"><%=HomeName%></a>
			<%else%>
				<span class="name"><%=HomeName%></span>
			<%end if
		end if%>

		<%if layercount=2 then%>
			<%if HomeName<>"" then%><span class="separator">&gt;&gt;</span><%end if%>
			<span class="guestbook">留言本</span>
		<%elseif layercount=3 then%>
			<%if HomeName<>"" then%><span class="separator">&gt;&gt;</span><%end if%>
			<a class="guestbook" href="index.asp">留言本</a>
			<span class="separator">&gt;&gt;</span>
			<span class="page"><%=param%></span>
		<%end if%>
	</div>
</div>
<%end sub
'==================================
function cked(exp)
	if exp=true then
		cked=" checked=""checked"""
	else
		cked=""
	end if
end function
'==================================
function dised(exp)
	if exp=true then
		dised=" disabled=""disabled"""
	else
		dised=""
	end if
end function
'==================================
function seled(exp)
	if exp=true then
		seled=" selected=""selected"""
	else
		seled=""
	end if
end function
'==================================
'isvalidhex: prepare for UrlDecode()
function isvalidhex(str)
	isvalidhex=true
	str=ucase(str)
	if len(str)<>3 then isvalidhex=false:exit function
	if left(str,1)<>"%" then isvalidhex=false:exit function
	c=mid(str,2,1)
	if not (((c>="0") and (c<="9")) or ((c>="A") and (c<="Z"))) then isvalidhex=false:exit function
	c=mid(str,3,1)
	if not (((c>="0") and (c<="9")) or ((c>="A") and (c<="Z"))) then isvalidhex=false:exit function
end function

function UrlDecode(enStr)
	dim deStr
	dim c,i,v
	deStr=""
	for i=1 to len(enStr)
		c=Mid(enStr,i,1)
		if c="%" then
			v=eval("&h"+Mid(enStr,i+1,2))
			if v<128 then
				deStr=deStr&chr(v)
				i=i+2
			else
				if isvalidhex(mid(enstr,i,3)) then
					if isvalidhex(mid(enstr,i+3,3)) then
						v=eval("&h"+Mid(enStr,i+1,2)+Mid(enStr,i+4,2))
						deStr=deStr&chr(v)
						i=i+5
					else
						v=eval("&h"+Mid(enStr,i+1,2)+cstr(hex(asc(Mid(enStr,i+3,1)))))
						deStr=deStr&chr(v)
						i=i+3 
					end if 
				else 
					destr=destr&c
				end if
			end if
		else
			if c="+" then
				deStr=deStr&" "
			else
				deStr=deStr&c
			end if
		end if
	next
	UrlDecode=deStr
end function
'==================================
Function HtmlDecode(sText)
    sText = Replace(sText, "&quot;", Chr(34))
    sText = Replace(sText, "&lt;"  , Chr(60))
    sText = Replace(sText, "&gt;"  , Chr(62))
    sText = Replace(sText, "&amp;" , Chr(38))
    sText = Replace(sText, "&nbsp;", Chr(32))
    HtmlDecode = sText
End Function
'==================================
function GetHiddenWordCondition()
	dim t_condition
	t_condition=""
	if HideHidden then t_condition=t_condition & sql_condition_hidehidden
	if HideAudit then t_condition=t_condition & sql_condition_hideaudit
	if HideWhisper then t_condition=t_condition & sql_condition_hidewhisper
	
	GetHiddenWordCondition=t_condition
end function
'==================================
function SetTimelessCookie(byval CookieName,byval CookieValue)
	Response.Cookies(CookieName)=CookieValue
	Response.Cookies(CookieName).Expires=CDate("2038-1-18 0:0:0")
	Response.Cookies(CookieName).Path="/"
end function
'==================================
function FormOrCookie(byval strname)
	if not isempty(Request.Form(strname)) then
		FormOrCookie=Request.Form(strname)
	elseif not isempty(Request.Cookies(strname)) then
		FormOrCookie=Request.Cookies(strname)
	else
		FormOrCookie=""
	end if
end function
'==================================
function FormOrSession(byval strname)
	if not isempty(Request.Form(strname)) then
		FormOrSession=Request.Form(strname)
	elseif not isempty(Session(strname)) then
		FormOrSession=Session(strname)
	else
		FormOrSession=""
	end if
end function
'==================================
function GetDisplayMode(FieldName)
	if isempty(Request.Cookies(FieldName)) or (Request.Cookies(FieldName)<>"book" and Request.Cookies(FieldName)<>"forum") then
		GetDisplayMode=DisplayMode
	else
		GetDisplayMode=Request.Cookies(FieldName)
	end if
end function

function GuestDisplayMode
	GuestDisplayMode=GetDisplayMode(InstanceName & "_DisplayMode")
end function

function AdminDisplayMode
	AdminDisplayMode=GetDisplayMode(InstanceName & "_AdminDisplayMode")
end function
'==================================
function GetRequests
	dim rvalue
	rvalue=""
	if Request.QueryString<>"" then
		for each item in Request.QueryString
			if item<>"mode" and item<>"modeflag" and item<>"rpage" then
				if Request.QueryString(item)<>"" then rvalue=rvalue & "&" & item & "=" & Server.URLEncode(Request.QueryString(item))
			end if
		next
	elseif Request.Form<>"" then
		for each item in Request.Form
			if item<>"mode" and item<>"modeflag" and item<>"rpage" then
				if Request.Form(item)<>"" then rvalue=rvalue & "&" & item & "=" & Server.URLEncode(Request.Form(item))
			end if
		next
	end if
	GetRequests=Server.HTMLEncode(rvalue)
end function
'==================================
function FilterGuestLike(byref str)
	dim r_str
	r_str=str
	
	r_str=replace(r_str,"'","''")
	r_str=replace(r_str,"[","[[]")
	r_str=replace(r_str,"%","[%]")
	r_str=replace(r_str,"_","[_]")
	r_str=replace(r_str,chr(34),"")
	
	FilterGuestLike=r_str
end function
'==================================
function FilterAdminLike(byref str)
	dim r_str
	r_str=str
	
	r_str=replace(r_str,"'","''")
	r_str=replace(r_str,"[","[[]")
	r_str=replace(r_str,chr(34),"")
	
	FilterAdminLike=r_str
end function
'==================================
function FilterStr(byref str,byval delquote)
	dim r_str
	r_str=str
	
	if delquote then r_str=replace(r_str,"'","")
	r_str=FilterGuestLike(r_str)
	
	FilterStr=r_str
end function
'==================================
function FilterSqlStr(byref str)
	dim r_str
	r_str=str
	
	r_str=replace(r_str,"'","''")
	r_str=replace(r_str,"[","[[]")
	r_str=replace(r_str,"%","[%]")
	r_str=replace(r_str,"_","[_]")
	'r_str=replace(r_str,chr(34),"")	'the difference between FilterGuestLike()
	
	FilterSqlStr=r_str
end function
'==================================
function FilterQuote(byref str)
	FilterQuote=Replace(str,"'","''")
end function
'==================================
function FilterSql(byref str)
	dim r_str
	r_str=str
	
	r_str=replace(r_str,"'","")
	r_str=replace(r_str,";","")
	r_str=replace(r_str," ","")
	r_str=replace(r_str,"[","")
	r_str=replace(r_str,"--","")
	r_str=replace(r_str,"/*","")
	r_str=replace(r_str,"#","")
	r_str=replace(r_str,"@","")
	r_str=replace(r_str,"/","")
	r_str=replace(r_str,"\","")
	r_str=replace(r_str,"?","")
	r_str=replace(r_str,"&","")
	r_str=replace(r_str,"=","")
	
	FilterSql=r_str
end function
'==================================
function FilterKeyword(byref str)
	dim r_str
	r_str=str

	r_str=replace(r_str,"'","")
	r_str=replace(r_str,"""","")
	r_str=replace(r_str,",","")
	r_str=replace(r_str,".","")
	r_str=replace(r_str,":","")
	r_str=replace(r_str,";","")
	r_str=replace(r_str,"`","")
	r_str=replace(r_str,"~","")
	r_str=replace(r_str," ","")
	r_str=replace(r_str,"(","")
	r_str=replace(r_str,")","")
	r_str=replace(r_str,"[","")
	r_str=replace(r_str,"]","")
	r_str=replace(r_str,"{","")
	r_str=replace(r_str,"}","")
	r_str=replace(r_str,"+","")
	r_str=replace(r_str,"^","")
	r_str=replace(r_str,"+","")
	r_str=replace(r_str,"-","")
	r_str=replace(r_str,"*","")
	r_str=replace(r_str,"#","")
	r_str=replace(r_str,"@","")
	r_str=replace(r_str,"$","")
	r_str=replace(r_str,"%","")
	r_str=replace(r_str,"/","")
	r_str=replace(r_str,"\","")
	r_str=replace(r_str,"?","")
	r_str=replace(r_str,"&","")
	r_str=replace(r_str,"=","")

	FilterKeyword=r_str
end function
'==================================
sub MessagePage(strMessage,backPage)
	%>
	<!-- #include file="inc_dtd.asp" -->
	<html>
	<head>
		<!-- #include file="inc_metatag.asp" -->
		<title><%=HomeName%> 留言本</title>
		<!-- #include file="inc_stylesheet.asp" -->
	</head>
	<body>
		<%=strMessage%><br />
		<a href="<%=backPage%>">[返回]</a>
		<script type="text/javascript" defer="defer" async="async">
			alert('<%=strMessage%>');window.location.replace('<%=backPage%>');
		</script>
	</body>
	</html>
	<%
end sub
'==================================
Function DateTimeStr(theTime)
	dim t
	t=CDate(theTime)
	DateTimeStr = Year(t) & "-" & Month(t) & "-" & Day(t) & " " & Hour(t) & ":" & Minute(t) & ":" & Second(t)
End Function
'==================================
Function CssBgImg(FileName)
	if FileName<>"" then
		CssBgImg = "background-image:url(" & FileName & ");"
	else
		CssBgImg = "background-image:none;"
	end if
End Function
'==================================
Function CssBgColor(ColorName)
	if ColorName<>"" then
		CssBgColor = "background-color:" & ColorName & ";"
	else
		CssBgColor = "background-color:transparent;"
	end if
End Function
'==================================
Function CssOptionalSize(property,size)
	if(property<>"") then
		if(isnumeric(size)) then
			CssOptionalSize=property & ":" & size & "px;"
		else
			CssOptionalSize=property & ":" & size & ";"
		end if
	else
		CssOptionalSize=""
	end if
End Function
%>