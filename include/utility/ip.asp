<%
function fullsize(byref istr,byref isize)
	Dim strLength
	strLength=Len(istr)
	if strLength<isize then
		fullsize=string(isize-strLength,"0") & istr
	else
		fullsize=istr
	end if
end function

function IsIPv4Range(byref ipStr)
	if Instr(ipStr,".")>0 then
		IsIPv4Range=true
	elseif ipStr<>"" and IsNumeric(ipStr) then
		IsIPv4Range=true
	else
		IsIPv4Range=false
	end if
end function

function IsIPv6Range(byref ipStr)
	if Instr(ipStr,":")>0 then
		IsIPv6Range=true
	elseif ipStr<>"" and IsNumeric("&H" & ipStr) then
		IsIPv6Range=true
	else
		IsIPv6Range=false
	end if
end function

function IsIPv4(byref ipStr)
	IsIPv4=(Instr(ipStr,".")>0)
end function

function IsIPv6(byref ipStr)
	IsIPv6=(Instr(ipStr,":")>0)
end function

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

function IPv6toHex(byref ipStr,byval toMax)
	ipStr=expandIPv6(ipStr,toMax)
	IPv6toHex=Replace(ipStr,":","")
end function

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
		if value<>"" then
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

	if result<>"" then
		result=":" & Left(result,Len(result)-1)
	end if
	result=Replace(result,":0",":")
	result=Replace(result,":0",":")
	result=Replace(result,":0",":")
	result=Mid(result,2)

	compactIPv6=UCase(result)
end function

function IsBannedIPv4(byref ipStr)
	dim cnx,rsx,iphex

	if IPv4ConStatus = 0 then
		IsBannedIPv4=false
	else
		iphex=IPv4ToHex(ipStr,false)
		set cnx=server.CreateObject("ADODB.Connection")
		set rsx=server.CreateObject("ADODB.Recordset")
		Call CreateConn(cnx)
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

function IsBannedIPv6(byref ipStr)
	dim cnx,rsx,iphex

	if IPv6ConStatus = 0 then
		IsBannedIPv6=false
	else
		iphex=IPv6ToHex(ipStr,false)
		set cnx=server.CreateObject("ADODB.Connection")
		set rsx=server.CreateObject("ADODB.Recordset")
		Call CreateConn(cnx)
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

function IsBannedIP(byref ipStr)
	if IsIPv4(ipStr) then
		IsBannedIP=IsBannedIPv4(ipStr)
	elseif IsIPv6(ipStr) then
		IsBannedIP=IsBannedIPv6(ipStr)
	else
		IsBannedIP=false
	end if
end function

function checkIsBannedIP
	if IsBannedIP(Request.ServerVariables("REMOTE_ADDR")) then
		checkIsBannedIP=true
	elseif IsBannedIP(Request.ServerVariables("HTTP_X_FORWARDED_FOR")) then
		checkIsBannedIP=true
	else
		checkIsBannedIP=false
	end if
end function

function hexToIPv4(byref hexip)
	dim i,strip
	strip=""
	for i=0 to 3
		if i>0 then strip=strip & "."
		strip=strip & cstr(cint("&H" & mid(hexip,i*2+1,2)))
	next

	hexToIPv4=strip
end function

function hexToIPv6(byref hexip)
	dim i,strip
	strip=""
	for i=0 to 7
		if i>0 then strip=strip & ":"
		strip=strip & mid(hexip,i*4+1,4)
	next

	hexToIPv6=compactIPv6(strip)
end function

function GetIPv4WithMask(byref strIP, byval itemCount)
	if strIP="" then
		GetIPv4WithMask=""
	elseif itemCount>=4 then
		GetIPv4WithMask=strIP
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
		GetIPv4WithMask=buffer
	end if
end function

function GetIPv6WithMask(byref strIP, byval itemCount)
	if strIP="" then
		GetIPv6WithMask=""
	elseif itemCount>=8 then
		GetIPv6WithMask=strIP
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
		GetIPv6WithMask=buffer
	end if
end function
%>