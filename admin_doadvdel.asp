<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_doadvdel.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<!-- #include file="tips.asp" -->
<%
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
Call CreateConn(cn)

dim tparam
tparam=Request.Form("iparam")
select case Request.Form("option")
case "1"
	if isdate(tparam) then
		tparam=DateTimeStr(tparam)
		cn.Execute Replace(sql_admindoadvdel_beforedate_main,"{0}",tparam),,1
		Call TipsPage("删除操作完成。","admin_advdel.asp")
	else
		Call TipsPage("您输入的日期有误，请检查。","admin_advdel.asp")
	end if
case "2"
	if isdate(tparam) then
		tparam=DateTimeStr(tparam)
		cn.Execute Replace(sql_admindoadvdel_afterdate_main,"{0}",tparam),,1
		Call TipsPage("删除操作完成。","admin_advdel.asp")
	else
		Call TipsPage("您输入的日期有误，请检查。","admin_advdel.asp")
	end if
case "3"
	if isnumeric(tparam) then
		if clng(tparam)>0 then
			'cn.Execute "delete reply.* from main,reply where main.id=reply.articleid and main.id<=(select max(id) from (select top " &tparam& " id from main order by id ASC))",,1
			cn.Execute Replace(sql_admindoadvdel_firstn_main,"{0}",tparam),,1
			Call TipsPage("删除操作完成。","admin_advdel.asp")
		else
			Call TipsPage("请输正确的数值。","admin_advdel.asp")
		end if
	else
		Call TipsPage("请输正确的数值。","admin_advdel.asp")
	end if
case "4"
	if isnumeric(tparam) then
		if clng(tparam)>0 then
			'cn.Execute "delete reply.* from main,reply where main.id=reply.articleid and main.id>=(select min(id) from (select top " &tparam& " id from main order by id DESC))",,1
			cn.Execute Replace(sql_admindoadvdel_lastn_main,"{0}",tparam),,1
			Call TipsPage("删除操作完成。","admin_advdel.asp")
		else
			Call TipsPage("请输正确的数值。","admin_advdel.asp")
		end if
	else
		Call TipsPage("请输正确的数值。","admin_advdel.asp")
	end if
case "5"
	tparam=replace(tparam,"'","''")
	tparam=replace(tparam,"[","[\[]")
	while left(tparam,1)="%" or left(tparam,1)="_"
		tparam=right(tparam,len(tparam)-1)
	wend
	while right(tparam,1)="%" or right(tparam,1)="_"
		tparam=left(tparam,len(tparam)-1)
	wend
	if tparam<>"" then
		'cn.Execute "delete reply.* from main,reply where main.id=reply.articleid and main.name like ('%" &tparam& "%')",,1
		cn.Execute Replace(sql_admindoadvdel_name_main,"{0}",tparam),,1
		Call TipsPage("删除操作完成。","admin_advdel.asp")
	else
		Call TipsPage("不能输入空字符串或全部为通配符。","admin_advdel.asp")
	end if
case "6"
	tparam=replace(tparam,"'","''")
	tparam=replace(tparam,"[","[\[]")
	while left(tparam,1)="%" or left(tparam,1)="_"
		tparam=right(tparam,len(tparam)-1)
	wend
	while right(tparam,1)="%" or right(tparam,1)="_"
		tparam=left(tparam,len(tparam)-1)
	wend
	if tparam<>"" then
		'cn.Execute "delete reply.* from main,reply where main.id=reply.articleid and main.title like ('%" &tparam& "%')",,1
		cn.Execute Replace(sql_admindoadvdel_title_main,"{0}",tparam),,1
		Call TipsPage("删除操作完成。","admin_advdel.asp")
	else
		Call TipsPage("不能输入空字符串或全部为通配符。","admin_advdel.asp")
	end if
case "7"
	tparam=replace(tparam,"'","''")
	tparam=replace(tparam,"[","[\[]")
	while left(tparam,1)="%" or left(tparam,1)="_"
		tparam=right(tparam,len(tparam)-1)
	wend
	while right(tparam,1)="%" or right(tparam,1)="_"
		tparam=left(tparam,len(tparam)-1)
	wend
	if tparam<>"" then
		'cn.Execute "delete reply.* from main,reply where main.id=reply.articleid and main.article like ('%" &tparam& "%')",,1
		cn.Execute Replace(sql_admindoadvdel_article_main,"{0}",tparam),,1
		Call TipsPage("删除操作完成。","admin_advdel.asp")
	else
		Call TipsPage("不能输入空字符串或全部为通配符。","admin_advdel.asp")
	end if
case "8"
	'cn.Execute "delete from reply",,1
	cn.Execute Replace(sql_admindoadvdel_main,"{0}",tparam),,1
	Call TipsPage("删除操作完成。","admin_advdel.asp")
case else
	Response.Redirect "admin_advdel.asp"
	cn.Close : set cn=nothing
	Response.End
end select

cn.Execute sql_admindoadvdel_clearfragment_main,,1
cn.Execute sql_admindoadvdel_clearfragment_reply,,1
cn.Execute sql_admindoadvdel_adjustguestreply_flag,,1
cn.Close : set cn=nothing
%>