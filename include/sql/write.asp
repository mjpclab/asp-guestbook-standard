<%
sql_write_filter=					"SELECT regexp,filtermode,replacestr FROM " &table_filterconfig& " ORDER BY filtersort ASC"
sql_write_flood_ids=				"SELECT DISTINCT id FROM " &table_main& " WHERE id="
sql_write_flood_idnull=				"SELECT DISTINCT id FROM " &table_main& " WHERE id IS NULL"
sql_write_flood_head=				"SELECT TOP 1 id FROM " &table_main& " WHERE ("
sql_write_flood_titlelike=			" title LIKE '%{0}%'"
sql_write_flood_titleequal=			" title='{0}'"
if IsAccess then
	sql_write_flood_articlelike=	" article LIKE '%{0}%'"
	sql_write_flood_articleequal=	" article LIKE '{0}'"
elseif IsSqlServer then
	sql_write_flood_articlelike=	" article LIKE CAST('%{0}%' AS NTEXT)"
	sql_write_flood_articleequal=	" article LIKE CAST('{0}' AS NTEXT)"
end if
sql_write_flood_wordstail=			") AND id IN (SELECT TOP {0} id FROM " &table_main& " WHERE parent_id<=0 ORDER BY id DESC)"
sql_write_flood_replytail=			") AND id IN (SELECT TOP {0} id FROM " &table_main& " WHERE id={1} OR parent_id={1} ORDER BY id DESC)"
sql_write_idnull=					"SELECT * FROM " &table_main& " WHERE id IS NULL"
if IsAccess then
	sql_write_verify_repliable=		"SELECT id FROM " &table_main& " WHERE parent_id<=0 AND (guestflag mod 1024)\512=0 AND id="
	sql_write_updateparentflag=		"UPDATE " &table_main& " SET replied=(replied MOD 2) + 2 WHERE id="
elseif IsSqlServer then
	sql_write_verify_repliable=		"SELECT id FROM " &table_main& " WHERE parent_id<=0 AND guestflag & 512=0 AND id="
	sql_write_updateparentflag=		"UPDATE " &table_main& " SET replied=replied|2 WHERE id="
end if
sql_write_updatelastupdated=		"UPDATE " &table_main& " SET lastupdated='{0}' WHERE id={1}"
%>