<%
if IsAccess then
	sql_search_condition_reply=		"((guestflag mod 32)\16)=0 AND (((guestflag mod 64)\32)=0 OR ((guestflag mod 128)\64)=0) AND id IN (SELECT articleid FROM " &table_reply& " WHERE reinfo LIKE '%{0}%')"
	sql_search_condition_name=		"((guestflag mod 32)\16)=0 AND name LIKE '%{0}%'"
	sql_search_condition_title=		"((guestflag mod 32)\16)=0 AND ((guestflag mod 64)\32)=0 AND title LIKE '%{0}%'"
	sql_search_condition_article=	"((guestflag mod 16)\8)=0 AND ((guestflag mod 32)\16)=0 AND ((guestflag mod 64)\32)=0 AND article LIKE '%{0}%'"
	sql_search_count_inner=			"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_search_count=				"SELECT Count(*) FROM ({0})"
	sql_search_full_inner=			"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_search_full=				"SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id IN({0})"
elseif IsSqlServer then
	sql_search_condition_reply=		"guestflag & 16=0 AND (guestflag & 32=0 OR guestflag & 64=0) AND id IN (SELECT articleid FROM " &table_reply& " WHERE reinfo LIKE '%{0}%')"
	sql_search_condition_name=		"guestflag & 16=0 AND name LIKE '%{0}%'"
	sql_search_condition_title=		"guestflag & 48=0 AND title LIKE '%{0}%'"
	sql_search_condition_article=	"guestflag & 56=0 AND article LIKE '%{0}%'"
	sql_search_count_inner=			"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_search_count=				"SELECT Count(*) FROM ({0}) AS __temp1"
	sql_search_full_inner=			"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_search_full=				"SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id IN({0})"
end if
%>