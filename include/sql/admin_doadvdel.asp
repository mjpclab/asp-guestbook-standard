<%
if IsAccess then
	sql_admindoadvdel_beforedate_main=	"DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated<=#{0}#"
	sql_admindoadvdel_afterdate_main=	"DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>=#{0}#"
elseif IsSqlServer then
	sql_admindoadvdel_beforedate_main=	"DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated<='{0}'"
	sql_admindoadvdel_afterdate_main=	"DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>='{0}'"
end if
sql_admindoadvdel_firstn_main=			"DELETE FROM " &table_main& " WHERE id IN (SELECT top {0} id FROM " &table_main& " WHERE parent_id<=0 ORDER BY lastupdated ASC)"
sql_admindoadvdel_lastn_main=			"DELETE FROM " &table_main& " WHERE id IN (SELECT top {0} id FROM " &table_main& " WHERE parent_id<=0 ORDER BY lastupdated DESC)"
if IsAccess then
	sql_admindoadvdel_name_main=			"DELETE FROM " &table_main& " WHERE name LIKE '%{0}%'"
	sql_admindoadvdel_title_main=			"DELETE FROM " &table_main& " WHERE title LIKE '%{0}%'"
	sql_admindoadvdel_article_main=			"DELETE FROM " &table_main& " WHERE article LIKE '%{0}%'"
elseif IsSqlServer then
	sql_admindoadvdel_name_main=			"DELETE FROM " &table_main& " WHERE name LIKE N'%{0}%'"
	sql_admindoadvdel_title_main=			"DELETE FROM " &table_main& " WHERE title LIKE N'%{0}%'"
	sql_admindoadvdel_article_main=			"DELETE FROM " &table_main& " WHERE article LIKE N'%{0}%'"
end if
sql_admindoadvdel_main=					"DELETE FROM " &table_main
sql_admindoadvdel_clearfragment_main=	"DELETE FROM " &table_main& " WHERE parent_id>0 AND parent_id NOT IN (SELECT DISTINCT id FROM " &table_main& " WHERE parent_id<=0)"
sql_admindoadvdel_clearfragment_reply=	"DELETE FROM " &table_reply& " WHERE articleid NOT IN(SELECT id FROM " &table_main& ")"
if IsAccess then
	sql_admindoadvdel_adjustguestreply_flag="UPDATE " &table_main& " SET replied=replied-((replied MOD 4)\2)*2 WHERE parent_id<=0 AND id NOT IN(" & _
												"SELECT DISTINCT parent_id FROM " &table_main & _
											")"
elseif IsSqlServer then
	sql_admindoadvdel_adjustguestreply_flag="UPDATE " &table_main& " SET replied=replied-(replied&2) WHERE parent_id<=0 AND id NOT IN(" & _
												"SELECT DISTINCT parent_id FROM " &table_main & _
											")"
end if
%>