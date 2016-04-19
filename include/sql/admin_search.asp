<%
if IsAccess then
	sql_adminsearch_condition_audit=	"((guestflag mod 32)\16)="
	sql_adminsearch_condition_reply=	"id IN (SELECT articleid FROM " &table_reply& " WHERE reinfo LIKE '%{0}%')"
	sql_adminsearch_condition_else=		"{0} LIKE '%{1}%'"
	sql_adminsearch_count_inner=		"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_adminsearch_count=				"SELECT Count(*) FROM ({0})"
	sql_adminsearch_full_inner=			"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_adminsearch_full=				"SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id IN({0})"
elseif IsSqlServer then
	sql_adminsearch_condition_audit=	"guestflag % 32 / 16="
	sql_adminsearch_condition_reply=	"id IN (SELECT articleid FROM " &table_reply& " WHERE reinfo LIKE N'%{0}%')"
	sql_adminsearch_condition_else=		"{0} LIKE N'%{1}%'"
	sql_adminsearch_count_inner=		"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_adminsearch_count=				"SELECT Count(*) FROM ({0}) AS __temp1"
	sql_adminsearch_full_inner=			"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_adminsearch_full=				"SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id IN({0})"
end if
%>