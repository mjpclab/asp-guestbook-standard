<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_savefloodconfig.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
if Not IsEmpty(Request.Form) then
	dim flood_minwait,flood_searchrange,flood_searchflag

	if isnumeric(Request.Form("minwait")) and Request.Form("minwait")<>"" then
		if Request.Form("minwait")>2147483647 then
			flood_minwait=2147483647
		else
			flood_minwait=abs(clng(Request.Form("minwait")))
		end if
	else
		flood_minwait=0
	end if

	if isnumeric(Request.Form("searchrange")) and Request.Form("searchrange")<>"" then
		if Request.Form("searchrange")>2147483647 then
			flood_searchrange=2147483647
		else
			flood_searchrange=abs(clng(Request.Form("searchrange")))
		end if
	else
		flood_searchrange=0
	end if

	flood_searchflag=0
	if Request.Form("flag_newword")="1" then flood_searchflag=flood_searchflag+1
	if Request.Form("flag_newreply")="1" then flood_searchflag=flood_searchflag+2
	if Request.Form("flag_include_equal")="1" then
		flood_searchflag=flood_searchflag+16
	elseif Request.Form("flag_include_equal")="2" then
		flood_searchflag=flood_searchflag+32
	end if
	if Request.Form("flag_title")="1" then flood_searchflag=flood_searchflag+256
	if Request.Form("flag_content")="1" then flood_searchflag=flood_searchflag+512

	dim cn1
	set cn1=server.CreateObject("ADODB.Connection")
	Call CreateConn(cn1)
	cn1.Execute Replace(Replace(Replace(sql_adminsavefloodconfig,"{0}",flood_minwait),"{1}",flood_searchrange),"{2}",flood_searchflag),,1

	cn1.Close : set cn1=nothing
end if
Response.Redirect "admin_floodconfig.asp"
%>