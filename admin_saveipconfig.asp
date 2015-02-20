<!-- #include file="config.asp" -->
<link rel="stylesheet" type="text/css" href="style.css"/>
<!-- #include file="style.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1

dim ids,iids
if Request.Form("savediplist1")<>"" then
ids=split(Request.Form("savediplist1"),",")
	for each iids in ids
		if isnumeric(iids)=false or iids="" then
			Response.Redirect "admin.asp"
			Response.End
		end if
	next
end if
set ids=nothing
if Request.Form("savediplist2")<>"" then
ids=split(Request.Form("savediplist2"),",")
	for each iids in ids
		if isnumeric(iids)=false or iids="" then
			Response.Redirect "admin.asp"
			Response.End
		end if
	next
end if


sub splitip(byref sourceip,byref dest1,byref dest2)
	dim si,ini

	for si=0 to ubound(sourceip)
		ini=instr(1,sourceip(si),"-")
		if ini>1 and ini<len(sourceip(si)) then
			dest1(si)=iptohex(left(sourceip(si),ini-1))
			dest2(si)=iptohex(mid(sourceip(si),ini+1))
		end if
	next
end sub


'Get and Process Parameters

tip1=split(Request.Form("txt1"),chr(13)&chr(10))
redim tstartip1(ubound(tip1)),tendip1(ubound(tip1))
splitip tip1,tstartip1,tendip1

tip2=split(Request.Form("txt2"),chr(13)&chr(10))
redim tstartip2(ubound(tip2)),tendip2(ubound(tip2))
splitip tip2,tstartip2,tendip2


tipconstatus=Request.Form("ipconstatus")
if tipconstatus<>"0" and tipconstatus<>"1" and tipconstatus<>"2" then tipconstatus=0


'Write to DB
set cn=server.CreateObject("ADODB.Connection")
CreateConn cn,dbtype

if Request.Form("savediplist1")<>"" then cn.Execute Replace(sql_adminsaveipconfig_delete1,"{0}",replace(Request.Form("savediplist1"),"'","")),,1
if Request.Form("savediplist2")<>"" then cn.Execute Replace(sql_adminsaveipconfig_delete2,"{0}",replace(Request.Form("savediplist2"),"'","")),,1

dim i
for i = 0 to ubound(tstartip1)
	if len(tstartip1(i))=8 and len(tendip1(i))=8 and tstartip1(i)<=tendip1(i) then cn.Execute Replace(Replace(sql_adminsaveipconfig_insert1,"{0}",tstartip1(i)),"{1}",tendip1(i)),,1
next
for i = 0 to ubound(tstartip2)
	if len(tstartip2(i))=8 and len(tendip2(i))=8 and tstartip2(i)<=tendip2(i) then cn.Execute Replace(Replace(sql_adminsaveipconfig_insert2,"{0}",tstartip2(i)),"{1}",tendip2(i)),,1
next

cn.Execute sql_adminsaveipconfig_update &cstr(tipconstatus),,1
cn.Close : set cn=nothing
		
Response.Redirect "admin_ipconfig.asp"
%>