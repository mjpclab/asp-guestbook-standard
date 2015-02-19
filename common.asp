<!-- #include file="ubbcode.asp" -->
<!-- #include file="sql.asp" -->
<%
'===================
function CreateConn(byref tconn,byval tDBType)
	'On Error Resume Next

	select case tDBType
	case 1
		tconn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & server.Mappath(dbfile) & """"
		tconn.Open
	case 2
		tconn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & server.Mappath(dbfile) & """"
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
	if len(istr)<isize then
		fullsize=string(isize-len(istr),"0") & istr
	else
		fullsize=istr
	end if
end function
'===================
function iptohex(byref strip)
	dim hexip,iparr,vi
	iparr=split(strip,".")
	hexip=""

	for vi=0 to ubound(iparr)
		if len(iparr(vi))<=3 then
			if isnumeric(iparr(vi))=true then
				if cint(iparr(vi))<=255 then hexip=hexip & fullsize(hex(iparr(vi)),2)
			end if
		end if
	next
	iptohex=hexip
end function
'=====================
function isbanip(byval ipaddr)
if IPConStatus=0 or ipaddr="" or lcase(ipaddr)="unknown" then
	isbanip=false
else
	ipaddr=iptohex(ipaddr)
	dim cnx,rsx
	set cnx=server.CreateObject("ADODB.Connection")
	set rsx=server.CreateObject("ADODB.Recordset")

	CreateConn cnx,dbtype
	rsx.Open Replace(Replace(sql_common_isbanip,"{0}",IPConStatus),"{1}",ipaddr),cnx
	if IPConStatus=1 then
		if rsx.EOF=false then
			isbanip=true
		else
			isbanip=false
		end if
	elseif IPConStatus=2 then
		if rsx.EOF=false then
			isbanip=false
		else
			isbanip=true
		end if
	else
		isbanip=false
	end if	

	rsx.Close
	cnx.Close
	set rsx=nothing
	set cnx=nothing
end if
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
function getip(byval sip ,byval PosCount)
	if ("" & sip & "")="" then
		getip=""
	elseif PosCount>=4 then
		getip=sip
	else
		dim t_str,t_char,t_count
		t_str=""
		t_count=0
		for ii=1 to len("" & sip & "")
			if t_count=PosCount then exit for
			t_char=mid(sip,ii,1)
			if isnumeric(t_char) then
				t_str=t_str & t_char
			else
				t_str=t_str & "."
				t_count=t_count+1
			end if
		next

		if right(t_str,1)="." then t_str=left(t_str,len(t_str)-1)
		for ii=1 to 4-PosCount
			t_str=t_str & ".*"
		next
		getip=t_str
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
sub show_page_list(byval CurPage,byval PagesCount,byval filename,byval pagetitle,byval titlealign,byval param)

if param<>"" then
	if left(param,1)<>"&" then param="&" & param
end if

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
	
	str_first_page=		"<a		 name=""page_control"" href=""" &filename& "?page=" &first_page_no& param & """ style=""float:left;""><img src=""image/icon_page_first.gif"" class=""pageicon"" title=""第" &first_page_no& "页"" /></a>"
	str_largeprev_page=	"<a		 name=""page_control"" href=""" &filename& "?page=" &largeprev_page_no& param & """ style=""float:left;""><img src=""image/icon_page_largeprev.gif"" class=""pageicon"" title=""第" &largeprev_page_no& "页"" /></a>"
	str_prev_page=		"<a		 name=""page_control"" href=""" &filename& "?page=" &prev_page_no& param & """ style=""float:left;""><img src=""image/icon_page_prev.gif"" class=""pageicon"" title=""第" &prev_page_no& "页"" /></a>"
	
	str_last_page=		"<a		 name=""page_control"" href=""" &filename& "?page=" &last_page_no& param & """ style=""float:right;""><img src=""image/icon_page_last.gif"" class=""pageicon"" title=""第" &last_page_no& "页"" /></a>"
	str_largenext_page=	"<a		 name=""page_control"" href=""" &filename& "?page=" &largenext_page_no& param & """ style=""float:right;""><img src=""image/icon_page_largenext.gif"" class=""pageicon"" title=""第" &largenext_page_no& "页"" /></a>"
	str_next_page=		"<a		 name=""page_control"" href=""" &filename& "?page=" &next_page_no& param & """ style=""float:right;""><img src=""image/icon_page_next.gif"" class=""pageicon"" title=""第" &next_page_no& "页"" /></a>"
	
	str_first2_page=	"<img	 name=""js_page_control"" onmouseup=""if(ghandle)clearTimeout(ghandle);"" onmousedown=""if(ghandle)clearTimeout(ghandle); refresh_pagenum(" &(-PagesCount+1)& ");"" src=""image/icon_page_first2.gif"" class=""pageicon"" style=""float:left; visibility:hidden; cursor:pointer;"" title=""卷至首页"" />"
	str_largeprev2_page="<img	 name=""js_page_control"" onmouseup=""if(ghandle)clearTimeout(ghandle);"" onmousedown=""if(ghandle)clearTimeout(ghandle); refresh_pagenum(" &(-AdvPageListCount)& ");"" src=""image/icon_page_largeprev2.gif"" class=""pageicon"" style=""float:left; visibility:hidden; cursor:pointer;"" title=""上卷" &AdvPageListCount& "页"" />"
	str_prev2_page=		"<img	 name=""js_page_control"" onmouseup=""if(ghandle)clearTimeout(ghandle);"" onmousedown=""if(ghandle)clearTimeout(ghandle); refresh_pagenum(" &(-1)& ");"" src=""image/icon_page_prev2.gif"" class=""pageicon"" style=""float:left; visibility:hidden; cursor:pointer;"" title=""上卷1页"" />"

	str_last2_page=		"<img	 name=""js_page_control"" onmouseup=""if(ghandle)clearTimeout(ghandle);"" onmousedown=""if(ghandle)clearTimeout(ghandle); refresh_pagenum(" &(PagesCount-1)& ");"" src=""image/icon_page_last2.gif"" class=""pageicon"" style=""float:right; visibility:hidden; cursor:pointer;"" title=""卷至末页"" />"
	str_largenext2_page="<img	 name=""js_page_control"" onmouseup=""if(ghandle)clearTimeout(ghandle);"" onmousedown=""if(ghandle)clearTimeout(ghandle); refresh_pagenum(" &(AdvPageListCount)& ");"" src=""image/icon_page_largenext2.gif"" class=""pageicon"" style=""float:right; visibility:hidden; cursor:pointer;"" title=""下卷" &AdvPageListCount& "页"" />"
	str_next2_page=		"<img	 name=""js_page_control"" onmouseup=""if(ghandle)clearTimeout(ghandle);"" onmousedown=""if(ghandle)clearTimeout(ghandle); refresh_pagenum(" &(1)& ");"" src=""image/icon_page_next2.gif"" class=""pageicon"" style=""float:right; visibility:hidden; cursor:pointer;"" title=""下卷1页"" />"
end if

Response.Write "<table cellpadding=""2"" class=""generalwindow"">" & _
	"<tr>" & _
		"<td class=""centertitle"" style=""text-align:" &titlealign& """>" &pagetitle& "</td>" & _
	"</tr>" & _
	"<tr style=""height:40px;"">" & _
		"<td class=""wordscontent"" style=""text-align:" &txt_align& """>"
	
		if ShowAdvPageList then		'左分页导航器
			Response.Write "<div style=""float:left; border:0px; padding:0px; margin:0px;"">"
			Response.Write str_first_page & str_largeprev_page & str_prev_page
			Response.Write str_first2_page & str_largeprev2_page & str_prev2_page
			Response.Write "</div>"
		end if

		if ShowAdvPageList then		'右分页导航器
			Response.Write "<div style=""float:right; border:0px; padding:0px; margin:0px;"">"
			Response.Write str_last_page & str_largenext_page & str_next_page
			Response.Write str_last2_page & str_largenext2_page & str_next2_page
			Response.Write "</div>"
		end if
	

		Response.Write "<form method=""get"" action=""" &filename& """ style=""margin:10px; text-align:" &txt_align& """>"
		for j=start_page to end_page
			if j=CurPage then
				Response.Write "<a name=""pagenum"" class=""pagenum_curr"" href=""" &filename& "?page=" & cstr(j) & param & """>[" &cstr(j)& "]</a> "
			else
				Response.Write "<a name=""pagenum"" class=""pagenum_normal"" href=""" &filename& "?page=" & cstr(j) & param & """>[" &cstr(j)& "]</a> "
			end if
		next
		
		Response.Write "<br/>(共" & PagesCount & "页)　转到<input type=""text"" name=""page"" size=""5"" maxlength=""10"" />页"
		arr_param=split(param,"&")
		for i=0 to ubound(arr_param)
			if arr_param(i)<>"" then
				t_name=left(arr_param(i),instr(arr_param(i),"=")-1)
				t_value=right(arr_param(i),len(arr_param(i))-len(t_name)-1)
				if t_name<>"page" then Response.Write "<input type=""hidden"" name=""" &t_name& """ value=""" &UrlDecode(t_value)& """ />"
			end if
		next
		Response.Write "　<input type=""submit"" value=""GO"" />"
		Response.Write "</form>"
	
Response.Write "</td></tr></table>"
%><!-- #include file="pagecontrols.inc" --><%
end sub
'==================================
sub show_book_title(byval layercount, byval param)
Response.Write "" & _
"<table cellpadding=""0"" style=""width:100%; overflow:hidden; border-width:0px; border-collapse:collapse; background-color:" &TitleBGC& "; background-image:url(" &TitlePic& ");"">" & _
	"<tr style=""height:60px;"">" & _
		"<td style=""width:100%;"" class=""booktitle"">"

if HomeLogo<>"" then
	Response.Write "<img style=""border-width:0px;"
	if LogoBannerMode then Response.Write "display:block;"
	Response.Write """ src=""" &HomeLogo& """ alt=""" &HomeName& """ />"
end if
if HomeName<>"" then
	if HomeAddr<>"" then
		Response.Write "<a href=""" &HomeAddr& """ class=""booktitle"">" &HomeName& "</a>"
	else
		Response.Write HomeName
	end if
end if

if layercount=2 then
	if HomeName<>"" then Response.Write " &gt;&gt; "
	Response.Write "留言本"
elseif layercount=3 then
	if HomeName<>"" then Response.Write " &gt;&gt; "
	Response.Write "<a href=""index.asp"" class=""booktitle"">留言本</a>"
	
	Response.Write " &gt;&gt; " & param
end if

Response.Write "</td></tr></table>"
end sub
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
sub MessagePage(strMessage,backPage)
	%>
	<!-- #include file="inc_dtd.asp" -->
	<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
		<title><%=HomeName%> 留言本</title>
		<!-- #include file="style.asp" -->
	</head>
	<body>
		<%=strMessage%><br />
		<a href="<%=backPage%>">[返回]</a>
		<script type="text/javascript" defer="defer">
			//<![CDATA[
			alert('<%=strMessage%>');window.location.replace('<%=backPage%>');
			//]]>
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
%>