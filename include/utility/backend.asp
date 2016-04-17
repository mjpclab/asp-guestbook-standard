<%
Dim charCr, charLf, charCrLf
charCr=Chr(13)
charLf=Chr(10)
charCrLf=charCr & charLf

Function HtmlEncode(byref text)
	if Instr(text,"&")>0 then
		text=Replace(text,"&","&amp;")
	end if
	if Instr(text,"<")>0 then
		text=Replace(text,"<","&lt;")
	end if
	if Instr(text,">")>0 then
		text=Replace(text,">","&gt;")
	end if
	if Instr(text,"""")>0 then
		text=Replace(text,"""","&quot;")
	end if
	if Instr(text,"'")>0 then
		text=Replace(text,"'","&#x27;")
	end if

	HtmlEncode=text
End Function

Function HtmlNewLineEncode(byref text)
	if Instr(text,charCrLf)>0 then
		text=Replace(text,charCrLf,"&#xd;&#xa;")
	end if
	if Instr(text,charCr)>0 then
		text=Replace(text,charCr,"&#xd;")
	end if
	if Instr(text,charLf)>0 then
		text=Replace(text,charLf,"&#xa;")
	end if

	HtmlNewLineEncode=text
End Function

Function HtmlDecode(byref text)
	if Instr(text,"&quot;")>0 then
		text = Replace(text, "&quot;", """")
	end if
	if Instr(text,"&apos;")>0 then
		text = Replace(text, "&apos;", "'")
	end if
	if Instr(text,"&lt;")>0 then
		text = Replace(text, "&lt;"  , "<")
	end if
	if Instr(text,"&gt;")>0 then
		text = Replace(text, "&gt;"  , ">")
	end if
	if Instr(text,"&amp;")>0 then
		text = Replace(text, "&amp;" , "&")
	end if
	if Instr(text,"&nbsp;")>0 then
		text = Replace(text, "&nbsp;", " ")
	end if

	HtmlDecode = text
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

	Call CreateConn(cna)
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
	GetRequests=HtmlEncode(rvalue)
end function

Function DateTimeStr(theTime)
	dim t
	t=CDate(theTime)
	DateTimeStr = Year(t) & "-" & Month(t) & "-" & Day(t) & " " & Hour(t) & ":" & Minute(t) & ":" & Second(t)
End Function

Function ServerTimeToUTC(time)
	ServerTimeToUTC=DateAdd("n",-ServerTimezoneOffset,time)
End Function

Function UTCToDisplayTime(time)
	UTCToDisplayTime=DateAdd("n",DisplayTimezoneOffset,time)
End Function

function geturlpath()
	dim host,url,buffer,port,httpHost
	if LCase(Request.ServerVariables("HTTPS"))="on" then
		host="https://"
	else
		host="http://"
	end if
	httpHost=Request.ServerVariables("HTTP_HOST")
	if httpHost<>"" then
		host=host & httpHost
	else
		buffer=Request.ServerVariables("SERVER_NAME")
		if IsIPv6(buffer) then
			buffer = "[" & buffer & "]"
		end if
		host=host & buffer

		port=Request.ServerVariables("SERVER_PORT")
		if port<>"" and port<>"80" then
			host=host & ":" & port
		end if
	end if
	host=host & Request.ServerVariables("URL")
	url=left(host,InStrRev(host,"/"))

	geturlpath=url
end function
%>