<%
if IsAccess then
	sql_noguestreply_flag=	"UPDATE " &table_main& " SET replied=replied-((replied MOD 4)\2)*2 WHERE parent_id<=0 AND id NOT IN({0}) AND id IN(" & _
										"SELECT parent_id FROM " &table_main& " WHERE id IN({0})" & _
									") AND NOT EXISTS (" & _
										"SELECT 1 FROM " &table_main& " AS temp WHERE temp.parent_id=" &table_main& ".id AND temp.id NOT IN({0})" & _
									")"
elseif IsSqlServer then
	sql_noguestreply_flag=	"UPDATE " &table_main& " SET replied=replied-(replied&2) WHERE parent_id<=0 AND id NOT IN({0}) AND id IN(" & _
										"SELECT parent_id FROM " &table_main& " WHERE id IN({0})" & _
									") AND NOT EXISTS (" & _
										"SELECT 1 FROM " &table_main& " AS temp WHERE temp.parent_id=" &table_main& ".id AND temp.id NOT IN({0})" & _
									")"
end if
%>