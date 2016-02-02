<%
Function UrlDecode(encoded)
	Dim startPos, nextPos, decoded, encodedLength
	Dim num1, str1, code1, num2, str2, code2, behindLeadingChar

	Const escape = "%"
	startPos = 1
	decoded = ""
	encodedLength = Len(encoded)

	Do While True
		nextPos = InStr(startPos, encoded, escape)
		If CBool(nextPos) Then
			decoded = decoded & Mid(encoded, startPos, nextPos - startPos)
			str1 = Mid(encoded, nextPos + 1, 2)
			code1 = "&H" & str1
			If Len(str1) > 0 And IsNumeric(code1) Then
				num1 = CLng(code1)
				If num1 < 128 Then
					decoded = decoded & Chr(num1)
					startPos = nextPos + 3
				Else
					behindLeadingChar = Mid(encoded, nextPos + 3, 1)
					If behindLeadingChar = escape Then
						str2 = Mid(encoded, nextPos + 4, 2)
						code2 = "&H" & str2
						If Len(str2) > 0 And IsNumeric(code2) Then
							decoded = decoded & Chr(CLng("&H" & str1 & str2))
							startPos = nextPos + 6
						Else
							decoded = decoded & Chr(CLng("&H" & str1 & "25"))   '25 is the hex of escape
							startPos = nextPos + 4
						End If
					ElseIf Len(behindLeadingChar) > 0 Then
						num2 = Asc(behindLeadingChar)
						If num2 < 16 Then
							str2 = "0" & Hex(num2)
						Else
							str2 = Hex(num2)
						End If
						decoded = decoded & Chr(CLng("&H" & str1 & str2))
						startPos = nextPos + 4
					End If
				End If
			Else
				decoded = decoded & escape
				startPos = nextPos + 1
			End If
		Else
			decoded = decoded & Mid(encoded, startPos)
			Exit Do
		End If
	Loop

	UrlDecode = decoded
End Function

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
	GetRequests=Server.HTMLEncode(rvalue)
end function

Function DateTimeStr(theTime)
	dim t
	t=CDate(theTime)
	DateTimeStr = Year(t) & "-" & Month(t) & "-" & Day(t) & " " & Hour(t) & ":" & Minute(t) & ":" & Second(t)
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
	if Request.ServerVariables("PATH_INFO")<>"" then
		host=host & Request.ServerVariables("PATH_INFO")
	else
		host=host & Request.ServerVariables("URL")
	end if
	url=left(host,InStrRev(host,"/"))

	geturlpath=url
end function
%>