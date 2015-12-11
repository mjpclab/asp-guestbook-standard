<%
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

Function HtmlDecode(sText)
    sText = Replace(sText, "&quot;", """")
    sText = Replace(sText, "&lt;"  , "<")
    sText = Replace(sText, "&gt;"  , ">")
    sText = Replace(sText, "&amp;" , "&")
    sText = Replace(sText, "&nbsp;", " ")
    HtmlDecode = sText
End Function

function SetTimelessCookie(byval CookieName,byval CookieValue)
	Response.Cookies(CookieName)=CookieValue
	Response.Cookies(CookieName).Expires=CDate("2038-1-18 0:0:0")
	Response.Cookies(CookieName).Path="/"
end function

function FormOrCookie(byval strname)
	if not isempty(Request.Form(strname)) then
		FormOrCookie=Request.Form(strname)
	elseif not isempty(Request.Cookies(strname)) then
		FormOrCookie=Request.Cookies(strname)
	else
		FormOrCookie=""
	end if
end function

function FormOrSession(byval strname)
	if not isempty(Request.Form(strname)) then
		FormOrSession=Request.Form(strname)
	elseif not isempty(Session(strname)) then
		FormOrSession=Session(strname)
	else
		FormOrSession=""
	end if
end function

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

function getvcode(byval codelen)
	dim rvalue
	randomize

	rvalue=cstr(fix(abs(rnd*(10^codelen))))
	if len(rvalue)<codelen then rvalue=string(codelen-len(rvalue),"0") & rvalue

	getvcode=rvalue
end function

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

function GetRequests
	dim rvalue
	rvalue=""
	if Request.QueryString<>"" then
		for each item in Request.QueryString
			if item<>"mode" and item<>"modeflag" and item<>"rpage" then
				if Request.QueryString(item)<>"" then rvalue=rvalue & "&" & item & "=" & Server.URLEncode(Request.QueryString(item))
			end if
		next
	end if
	if Request.Form<>"" then
		for each item in Request.Form
			if item<>"mode" and item<>"modeflag" and item<>"rpage" then
				if Request.Form(item)<>"" then rvalue=rvalue & "&" & item & "=" & Server.URLEncode(Request.Form(item))
			end if
		next
	end if
	GetRequests=Server.HTMLEncode(rvalue)
end function

Function DateTimeStr(theTime)
	dim t
	t=CDate(theTime)
	DateTimeStr = Year(t) & "-" & Month(t) & "-" & Day(t) & " " & Hour(t) & ":" & Minute(t) & ":" & Second(t)
End Function

function geturlpath()
	dim host,url,buffer,port,httpHost
	host="http://"
	httpHost=Request.ServerVariables("HTTP_HOST")
	if Len(httpHost)>0 then
		host=host & httpHost
	else
		buffer=Request.ServerVariables("SERVER_NAME")
		if IsIPv6(buffer) then
			buffer = "[" & buffer & "]"
		end if
		host=host & buffer

		port=Request.ServerVariables("SERVER_PORT")
		if Len(port)>0 and port<>"80" then
			host=host & ":" & port
		end if
	end if
	if Request.ServerVariables("PATH_INFO")<>"" then
		host=host & Request.ServerVariables("PATH_INFO")
	else
		host=host & Request.ServerVariables("URL")
	end if
	url=left(host,InStrRev(host,"/"))

	geturlpath=url
end function
%>