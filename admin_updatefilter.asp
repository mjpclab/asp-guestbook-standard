<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_updatefilter.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
Response.Expires=-1

tfindexp=Request.Form("findexp")
if tfindexp<>"" then
	Dim tfiltermode
	tfiltermode=0

	Dim tsearchmode
	tsearchmode=Request.Form("searchmode")
	if tsearchmode="256" then
		tfiltermode=tfiltermode+256
	elseif tsearchmode="512" then
		tfiltermode=tfiltermode+512
	elseif tsearchmode="1024" then
		tfiltermode=tfiltermode+1024
	elseif tsearchmode="2048" then
		tfiltermode=tfiltermode+2048
	end if
	if Request.Form("matchcase")="8192" then tfiltermode=tfiltermode+8192

	arr_findrange=split(Request.Form("findrange"),",")
	for fi=0 to ubound(arr_findrange)
		if len(arr_findrange(fi))<=5 then
			if isnumeric(arr_findrange(fi)) then
				if cint(arr_findrange(fi))<=64 then
					tfiltermode=tfiltermode+cint(arr_findrange(fi))
				end if
			end if
		end if
	next

	Dim tfiltermethod,treplacestr
	tfiltermethod=Request.Form("filtermethod")
	if tfiltermethod="0" then
		treplacestr=Request.Form("replacetxt")
	else
		treplacestr=""
		if tfiltermethod="4096" then
			tfiltermode=tfiltermode+4096
		elseif tfiltermethod="16384" then
			tfiltermode=tfiltermode+16384
		end if
	end if

	Dim tmemo
	tmemo=Request.Form("memo")
	if len(tmemo)>25 then
		tmemo=left(tmemo,25)
	end if

	Dim tfilterid
	tfilterid=Request.Form("filterid")
	if isnumeric(tfilterid) and tfilterid<>"" then
		tfilterid=clng(tfilterid)

		set cn=server.CreateObject("ADODB.Connection")
		set rs=server.CreateObject("ADODB.Recordset")
		Call CreateConn(cn)

		rs.Open sql_adminupdatefilter &tfilterid,cn,1,3,1
		if not rs.EOF then
			rs("regexp")=tfindexp
			rs("filtermode")=tfiltermode
			rs("replacestr")=treplacestr
			rs("memo")=tmemo
			rs.Update
		end if

		rs.Close : cn.Close : set rs=nothing : set cn=nothing
	end if
end if

Response.Redirect "admin_filter.asp"
%>