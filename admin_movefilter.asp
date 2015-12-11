<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_movefilter.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
Response.Expires=-1

if isnumeric(Request.Form("filterid")) and Request.Form("filterid")<>"" and (Request.Form("movedirection")="up" or Request.Form("movedirection")="down" or Request.Form("movedirection")="top" or Request.Form("movedirection")="bottom") then
	 filterid1=clng(Request.Form("filterid"))

	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	CreateConn cn,dbtype
	
	rs.Open sql_adminmovefilter_sort1 &filterid1,cn,,,1
	if rs.EOF=false then
		qid1=rs(0)
		rs.Close
	
		if Request.Form("movedirection")="up" then
			rs.Open sql_adminmovefilter_sort2_up &qid1,cn,,,1
		elseif Request.Form("movedirection")="down" then
			rs.Open sql_adminmovefilter_sort2_down &qid1,cn,,,1
		elseif Request.Form("movedirection")="top" then
			rs.Open sql_adminmovefilter_sort2_top,cn,,,1
		elseif Request.Form("movedirection")="bottom" then
			rs.Open sql_adminmovefilter_sort2_bottom,cn,,,1
		end if
		if rs.EOF=false then
			qid2=rs(0)
			rs.Close
			
			if qid2<>"" and qid1<>qid2 then
				rs.Open sql_adminmovefilter_id2 &qid2,cn,,,1
				filterid2=rs(0)
				rs.Close

				cn.BeginTrans
				
				select case Request.Form("movedirection")
				case "up","down"
					cn.Execute sql_adminmovefilter_update_updown1 &filterid1,,1
					cn.Execute Replace(Replace(sql_adminmovefilter_update_updown2,"{0}",qid1),"{1}",qid2),,1
					cn.Execute Replace(sql_adminmovefilter_update_updown3,"{0}",qid2),,1
				case "top"
					cn.Execute sql_adminmovefilter_update_top1 &qid1,,1
					cn.Execute Replace(Replace(sql_adminmovefilter_update_top2,"{0}",qid2),"{1}",filterid1),,1
				case "bottom"
					cn.Execute sql_adminmovefilter_update_bottom1 &qid1,,1
					cn.Execute Replace(Replace(sql_adminmovefilter_update_bottom2,"{0}",qid2),"{1}",filterid1),,1
				end select
									
				cn.CommitTrans
			end if
		else
			rs.Close
		end if
	else
		rs.Close
	end if
	cn.Close : set rs=nothing : set cn=nothing
end if

Response.Redirect "admin_filter.asp"
%>