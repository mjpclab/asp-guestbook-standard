<%
IsAccess=false
IsSqlServer=false

if dbtype=1 or dbtype=2 then
	IsAccess=true
    table_config			="[" & prefix & "config]"
    table_filterconfig		="[" & prefix & "filterconfig]"
    table_floodconfig		="[" & prefix & "floodconfig]"
    table_ipconfig			="[" & prefix & "ipconfig]"
    table_main				="[" & prefix & "main]"
    table_reply				="[" & prefix & "reply]"
    table_stats				="[" & prefix & "stats]"
    table_stats_clientinfo	="[" & prefix & "stats_clientinfo]"
    table_style				="[" & prefix & "style]"
    table_supervisor		="[" & prefix & "supervisor]"
elseif dbtype=10 then
	IsSqlServer=true
	if dbschema<>"" then
		schema="[" &dbschema& "]."
	else
		schema=""
	end if
    table_config			=schema & "[" & prefix & "config]"
    table_filterconfig		=schema & "[" & prefix & "filterconfig]"
    table_floodconfig		=schema & "[" & prefix & "floodconfig]"
    table_ipconfig			=schema & "[" & prefix & "ipconfig]"
    table_main				=schema & "[" & prefix & "main]"
    table_reply				=schema & "[" & prefix & "reply]"
    table_stats				=schema & "[" & prefix & "stats]"
    table_stats_clientinfo	=schema & "[" & prefix & "stats_clientinfo]"
    table_style				=schema & "[" & prefix & "style]"
    table_supervisor		=schema & "[" & prefix & "supervisor]"
end if

'PKs
sql_pk_main="id"
sql_pksearch_main="root_id"

'GLOBALs
if IsAccess then
	sql_global_noguestreply_flag=	"UPDATE " &table_main& " SET replied=replied-((replied MOD 4)\2)*2 WHERE parent_id<=0 AND id NOT IN({0}) AND id IN(" & _
										"SELECT parent_id FROM " &table_main& " WHERE id IN({0})" & _
									") AND NOT EXISTS (" & _
										"SELECT 1 FROM " &table_main& " AS temp WHERE temp.parent_id=" &table_main& ".id AND temp.id NOT IN({0})" & _
									")"
elseif IsSqlServer then
	sql_global_noguestreply_flag=	"UPDATE " &table_main& " SET replied=replied-(replied&2) WHERE parent_id<=0 AND id NOT IN({0}) AND id IN(" & _
										"SELECT parent_id FROM " &table_main& " WHERE id IN({0})" & _
									") AND NOT EXISTS (" & _
										"SELECT 1 FROM " &table_main& " AS temp WHERE temp.parent_id=" &table_main& ".id AND temp.id NOT IN({0})" & _
									")"
end if

'common
sql_common_isbanip=		"SELECT TOP 1 1 FROM " &table_ipconfig& " WHERE ipconstatus={0} And '{1}'>=startip And '{1}'<=endip"
sql_common_getstat=		"SELECT TOP 1 startdate,[{0}] FROM " &table_stats
sql_common_initstat=	"INSERT INTO " &table_stats& "(startdate) VALUES('{0}')"
sql_common_updatetime=	"UPDATE " &table_stats& " SET startdate='{0}'"
sql_common_addstat=		"UPDATE " &table_stats& " SET [{0}]=[{0}]+1"
if IsAccess then
	'sql_common_dividepage=	"SELECT TOP {0} * FROM (" & _
	'							"SELECT * FROM (" & _
	'								"SELECT TOP {0} * FROM (" & _
	'									"SELECT TOP {1} {2} ORDER BY {3}" & _
	'								") ORDER BY {4}" & _
	'							") ORDER BY {3}" & _
	'						")"
	sql_common_dividepage=	"{6} {7} {8} IN(" & _
								"SELECT TOP {0} {8} FROM (" & _
									"SELECT {8} FROM (" & _
										"SELECT TOP {0} {2} FROM (" & _
											"SELECT TOP {1} {2} {3} ORDER BY {4}" & _
										") ORDER BY {5}" & _
									") ORDER BY {4}" & _
								")" & _
							") ORDER BY {4}"
elseif IsSqlServer then
	'sql_common_dividepage=	"SELECT * FROM (" & _
	'							"SELECT TOP {0} * FROM (" & _
	'								"SELECT TOP {1} {2} ORDER BY {3}" & _
	'							") AS __temp1 ORDER BY {4}" & _
	'						") AS __temp2 ORDER BY {3}"
	sql_common_dividepage=	"{6} {7} {8} IN(" & _
								"SELECT TOP {0} {8} FROM (" & _
									"SELECT TOP {1} {2} {3} ORDER BY {4}" & _
								") AS __temp1 ORDER BY {5}" & _
							") ORDER BY {4}"
end if

'common2
sql_common2_getadmininfo=	"SELECT name,faceid,faceurl,email,qqid,msnid,homepage FROM " &table_supervisor
'sql_common2_reply=			"SELECT replydate,htmlflag,reinfo FROM " &table_reply& " WHERE articleid="
sql_common2_guestreply=		"SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE parent_id={0}{1} ORDER BY id ASC"
sql_common2_replyinform=	"SELECT parent_id,email,guestflag,logdate,article FROM " &table_main& " WHERE id="

'loadconfig
sql_loadconfig_config=		"SELECT TOP 1 * FROM " & table_config
sql_loadconfig_style=		"SELECT * FROM " &table_style& " WHERE stylename='{0}'"
sql_loadconfig_top1style=	"SELECT Top 1 * FROM " &table_style
sql_loadconfig_floodconfig=	"SELECT TOP 1 * FROM " &table_floodconfig

'index/search/listword_guest/common2-outerword() condition
if IsAccess then
	sql_condition_hidehidden=	" AND (((guestflag mod 16)\8)=0 OR replied<>0)"
	sql_condition_hideaudit=	" AND (((guestflag mod 32)\16)=0 OR replied<>0)"
	sql_condition_hidewhisper=	" AND (((guestflag mod 64)\32)=0 OR replied<>0)"
elseif IsSqlServer then
	sql_condition_hidehidden=	" AND (guestflag & 8=0 OR replied<>0)"
	sql_condition_hideaudit=	" AND (guestflag & 16=0 OR replied<>0)"
	sql_condition_hidewhisper=	" AND (guestflag & 32=0 OR replied<>0)"
end if

'index
sql_index_words_count="SELECT COUNT(1) FROM " &table_main& " WHERE parent_id<=0"
sql_index_words_query="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE parent_id<=0"

'search
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

'write
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

'topbulletin
sql_topbulletin="SELECT declareflag,[declare] FROM " &table_supervisor

'showword
sql_showword_count=	"SELECT COUNT(1) FROM " &table_main
sql_showword=		"SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id="

'tlist
'sql_tlist_itemsperpage=	"SELECT TOP 1 itemsperpage FROM " &table_config
sql_tlist_maxtime=		"SELECT TOP 1 lastupdated FROM " &table_main& " WHERE parent_id<=0 {0} ORDER BY lastupdated DESC"
if IsAccess then
	sql_tlist_mintime=	"SELECT TOP 1 lastupdated FROM (SELECT TOP {0} parent_id,lastupdated FROM " &table_main& " WHERE parent_id<=0 AND (guestflag mod 32)\16=0 AND (guestflag mod 64)\32=0 {1} ORDER BY parent_id ASC,lastupdated DESC) ORDER BY lastupdated ASC"
	sql_tlist=			"SELECT guestflag,title FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>=#{0}# AND lastupdated<=#{1}# {2} ORDER BY parent_id ASC,lastupdated DESC"
elseif IsSqlServer then
	sql_tlist_mintime=	"SELECT TOP 1 lastupdated FROM (SELECT TOP {0} parent_id,lastupdated FROM " &table_main& " WHERE parent_id<=0 AND guestflag & 48=0 {1} ORDER BY parent_id ASC,lastupdated DESC) AS __temp1 ORDER BY lastupdated ASC"
	sql_tlist=			"SELECT guestflag,title FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>='{0}' AND lastupdated<='{1}' {2} ORDER BY parent_id ASC,lastupdated DESC"
end if

'saveclientinfo
sql_saveclientinfo="INSERT INTO " &table_stats_clientinfo& "(os,browser,screenwidth,screenheight,timesect,sourceaddr,fullsource) VALUES('{0}','{1}',{2},{3},'{4}','{5}','{6}')"

'listword_guest
sql_listword_guest="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE parent_id={0}{1} ORDER BY id ASC"

'listword_admin
sql_listword_admin="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE parent_id={0} ORDER BY id ASC"

'admin
sql_admin_words_count="SELECT COUNT(1) FROM " &table_main& " WHERE parent_id<=0"
sql_admin_words_query="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE parent_id<=0"

'admin_showword
sql_admin_showword="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id="

'admin_search
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
	sql_adminsearch_condition_reply=	"id IN (SELECT articleid FROM " &table_reply& " WHERE reinfo LIKE '%{0}%')"
	sql_adminsearch_condition_else=		"{0} LIKE '%{1}%'"
	sql_adminsearch_count_inner=		"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_adminsearch_count=				"SELECT Count(*) FROM ({0}) AS __temp1"
	sql_adminsearch_full_inner=			"SELECT DISTINCT root_id FROM " &table_main& " WHERE "
	sql_adminsearch_full=				"SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id IN({0})"
end if

'admin_edit
sql_adminedit="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id="

'admin_saveedit
sql_adminsaveedit_open="SELECT guestflag,title,article FROM " &table_main& " WHERE id="

'admin_reply
sql_adminreply_reply="SELECT htmlflag,reinfo FROM " &table_reply& " WHERE articleid="
sql_adminreply_words="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id="


'admin_delreply
sql_admindelreply_delete="DELETE FROM " &table_reply& " WHERE articleid="
if IsAccess then
	sql_admindelreply_update="UPDATE " &table_main& " SET replied=replied-(replied mod 2) WHERE id="
elseif IsSqlServer then
	sql_admindelreply_update="UPDATE " &table_main& " SET replied=replied-(replied & 1) WHERE id="
end if

'admin_savereply
sql_adminsavereply_main=			"SELECT guestflag,replied FROM " &table_main& " WHERE id="
sql_adminsavereply_reply=			"SELECT articleid FROM " &table_reply& " WHERE articleid="
sql_adminsavereply_insert=			"INSERT INTO " &table_reply& "(articleid,replydate,htmlflag,reinfo) VALUES('{0}','{1}',{2},'{3}')"
sql_adminsavereply_update=			"UPDATE " &table_reply& " SET replydate='{0}',htmlflag={1},reinfo='{2}' WHERE articleid={3}"
if IsAccess then
	sql_adminsavereply_lock2top=	"UPDATE " &table_main& " SET parent_id=-1 WHERE parent_id=0 AND id=(SELECT root_id FROM " &table_main& " WHERE id={0})"
	sql_adminsavereply_bring2top=	"UPDATE " &table_main& " SET lastupdated='{0}' WHERE parent_id<=0 AND id=(SELECT root_id FROM " &table_main& " WHERE id={1})"
elseif IsSqlServer then
	sql_adminsavereply_lock2top=	"UPDATE " &table_main& " SET parent_id=-1 WHERE parent_id=0 AND id=(SELECT root_id FROM " &table_main& " WHERE id={0})"
	sql_adminsavereply_bring2top=	"UPDATE " &table_main& " SET lastupdated='{0}' WHERE parent_id<=0 AND id=(SELECT root_id FROM " &table_main& " WHERE id={1})"
end if

'admin_[action]
if IsAccess then
	sql_adminhidecontact=	"UPDATE " &table_main& " SET guestflag=guestflag+256 WHERE (guestflag mod 512)\256=0 AND id="
	sql_adminunhidecontact=	"UPDATE " &table_main& " SET guestflag=guestflag-256 WHERE (guestflag mod 512)\256<>0 AND id="
	sql_adminhideword=		"UPDATE " &table_main& " SET guestflag=guestflag+8 WHERE (guestflag mod 16)\8=0 AND id="
	sql_adminunhideword=	"UPDATE " &table_main& " SET guestflag=guestflag-8 WHERE (guestflag mod 16)\8<>0 AND id="
	sql_adminlockreply=		"UPDATE " &table_main& " SET guestflag=guestflag+512 WHERE (guestflag mod 1024)\512=0 AND id="
	sql_adminunlockreply=	"UPDATE " &table_main& " SET guestflag=guestflag-512 WHERE (guestflag mod 1024)\512<>0 AND id="
	sql_adminpassaudit=		"UPDATE " &table_main& " SET guestflag=guestflag-16 WHERE (guestflag mod 32)\16<>0 AND id="
	sql_adminpubwhisper=	"UPDATE " &table_main& " SET guestflag=guestflag-32 WHERE (guestflag mod 64)\32<>0 AND id="
elseif IsSqlServer then
	sql_adminhidecontact=	"UPDATE " &table_main& " SET guestflag=guestflag+256 WHERE guestflag & 256=0 AND id="
	sql_adminunhidecontact=	"UPDATE " &table_main& " SET guestflag=guestflag-256 WHERE guestflag & 256<>0 AND id="
	sql_adminhideword=		"UPDATE " &table_main& " SET guestflag=guestflag+8 WHERE guestflag & 8=0 AND id="
	sql_adminunhideword=	"UPDATE " &table_main& " SET guestflag=guestflag-8 WHERE guestflag & 8<>0 AND id="
	sql_adminlockreply=		"UPDATE " &table_main& " SET guestflag=guestflag+512 WHERE guestflag & 512=0 AND id="
	sql_adminunlockreply=	"UPDATE " &table_main& " SET guestflag=guestflag-512 WHERE guestflag & 512<>0 AND id="
	sql_adminpassaudit=		"UPDATE " &table_main& " SET guestflag=guestflag-16 WHERE guestflag & 16<>0 AND id="
	sql_adminpubwhisper=	"UPDATE " &table_main& " SET guestflag=guestflag-32 WHERE guestflag & 32<>0 AND id="
end if
sql_admindel_reply=			"DELETE FROM " &table_reply& " WHERE articleid IN(SELECT id FROM " &table_main& " WHERE parent_id={0} OR id={0})"
sql_admindel_main=			"DELETE FROM " &table_main& " WHERE parent_id={0} OR id={0}"
sql_adminlock2top=			"UPDATE " &table_main& " SET parent_id=-1 WHERE parent_id=0 AND id="
sql_adminunlock2top=		"UPDATE " &table_main& " SET parent_id=0 WHERE parent_id<0 AND id="
sql_adminbring2top=			"UPDATE " &table_main& " SET lastupdated='{0}' WHERE parent_id<=0 AND id={1}"
sql_adminreorder=			"UPDATE " &table_main& " SET parent_id=0,lastupdated=logdate WHERE parent_id<=0 AND id={0}"

'admin_mdel
sql_adminmdel_reply="DELETE FROM " &table_reply& " WHERE articleid IN(SELECT id FROM " &table_main& " WHERE parent_id IN({0}) OR id IN({0}))"
sql_adminmdel_main=	"DELETE FROM " &table_main& " WHERE parent_id IN({0}) OR id IN({0})"

'admin_mpass
if IsAccess then
	sql_adminmpass="UPDATE " &table_main& " SET guestflag=guestflag-16 WHERE ((guestflag mod 32)\16)<>0 AND id IN ({0})"
elseif IsSqlServer then
	sql_adminmpass="UPDATE " &table_main& " SET guestflag=guestflag-16 WHERE guestflag & 16<>0 AND id IN ({0})"
end if

'admin_doadvdel
if IsAccess then
	sql_admindoadvdel_beforedate_main=	"DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated<=#{0}#"
	sql_admindoadvdel_afterdate_main=	"DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>=#{0}#"
elseif IsSqlServer then
	sql_admindoadvdel_beforedate_main=	"DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated<='{0}'"
	sql_admindoadvdel_afterdate_main=	"DELETE FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>='{0}'"
end if
sql_admindoadvdel_firstn_main=			"DELETE FROM " &table_main& " WHERE id IN (SELECT top {0} id FROM " &table_main& " WHERE parent_id<=0 ORDER BY lastupdated ASC)"
sql_admindoadvdel_lastn_main=			"DELETE FROM " &table_main& " WHERE id IN (SELECT top {0} id FROM " &table_main& " WHERE parent_id<=0 ORDER BY lastupdated DESC)"	
sql_admindoadvdel_name_main=			"DELETE FROM " &table_main& " WHERE name LIKE '%{0}%'"
sql_admindoadvdel_title_main=			"DELETE FROM " &table_main& " WHERE title LIKE '%{0}%'"
sql_admindoadvdel_article_main=			"DELETE FROM " &table_main& " WHERE article LIKE '%{0}%'"
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

'admin_setbulletin
sql_adminsetbulletin="SELECT declareflag,[declare] FROM " &table_supervisor

'admin_savebulletin
sql_adminsavebulletin="UPDATE " &table_supervisor& " SET declareflag={0},[declare]='{1}'"

'admin_config
sql_adminconfig_config=	"SELECT TOP 1 * FROM " &table_config
sql_adminconfig_style=	"SELECT stylename FROM " &table_style

'admin_saveconfig
sql_adminsaveconfig="SELECT TOP 1 * FROM " &table_config

'admin_ipconfig
sql_adminipconfig_status=	"SELECT ipconstatus FROM " &table_config
sql_adminipconfig_status1=	"SELECT listid,startip,endip FROM " &table_ipconfig& " WHERE ipconstatus=1"
sql_adminipconfig_status2=	"SELECT listid,startip,endip FROM " &table_ipconfig& " WHERE ipconstatus=2"

'admin_saveipconfig
sql_adminsaveipconfig_delete1=	"DELETE FROM " &table_ipconfig& " WHERE listid IN ({0})"
sql_adminsaveipconfig_delete2=	"DELETE FROM " &table_ipconfig& " WHERE listid IN ({0})"
sql_adminsaveipconfig_insert1=	"INSERT INTO " &table_ipconfig& "(ipconstatus,startip,endip) VALUES (1,'{0}','{1}')"
sql_adminsaveipconfig_insert2=	"INSERT INTO " &table_ipconfig& "(ipconstatus,startip,endip) VALUES (2,'{0}','{1}')"
sql_adminsaveipconfig_update=	"UPDATE " &table_config& " SET ipconstatus="

'admin_filter
sql_adminfilter="SELECT * FROM " &table_filterconfig& " ORDER BY filtersort ASC"

'admin_appendfilter
sql_adminfilter_null=	"SELECT * FROM " &table_filterconfig& " WHERE filterid IS NULL"
sql_adminfilter_max=	"SELECT MAX(filterid) FROM " &table_filterconfig
sql_adminfilter_update=	"UPDATE " &table_filterconfig& " SET filtersort={0} WHERE filterid={0}"

'admin_updatefilter
sql_adminupdatefilter="SELECT regexp,filtermode,replacestr,[memo] FROM " &table_filterconfig& " WHERE filterid="

'admin_delfilter
sql_admindelfilter="DELETE FROM " &table_filterconfig& " WHERE filterid="

'admin_movefilter
sql_adminmovefilter_sort1=			"SELECT filtersort FROM " &table_filterconfig& " WHERE filterid="
sql_adminmovefilter_sort2_up=		"SELECT MAX(filtersort) FROM " &table_filterconfig& " WHERE filtersort<"
sql_adminmovefilter_sort2_down=		"SELECT MIN(filtersort) FROM " &table_filterconfig& " WHERE filtersort>"
sql_adminmovefilter_sort2_top=		"SELECT MIN(filtersort) FROM " &table_filterconfig
sql_adminmovefilter_sort2_bottom=	"SELECT MAX(filtersort) FROM " &table_filterconfig
sql_adminmovefilter_id2=			"SELECT filterid FROM " &table_filterconfig& " WHERE filtersort="
sql_adminmovefilter_update_updown1=	"UPDATE " &table_filterconfig& " SET filtersort=-1 WHERE filterid="
sql_adminmovefilter_update_updown2=	"UPDATE " &table_filterconfig& " SET filtersort={0} WHERE filtersort={1}"
sql_adminmovefilter_update_updown3=	"UPDATE " &table_filterconfig& " SET filtersort={0} WHERE filtersort=-1"
sql_adminmovefilter_update_top1=	"UPDATE " &table_filterconfig& " SET filtersort=filtersort+1 WHERE filtersort<"
sql_adminmovefilter_update_top2=	"UPDATE " &table_filterconfig& " SET filtersort={0} WHERE filterid={1}"
sql_adminmovefilter_update_bottom1=	"UPDATE " &table_filterconfig& " SET filtersort=filtersort-1 WHERE filtersort>"
sql_adminmovefilter_update_bottom2=	"UPDATE " &table_filterconfig& " SET filtersort={0} WHERE filterid={1}"

'admin_floodconfig
sql_adminfloodconfig="SELECT TOP 1 * FROM " &table_floodconfig

'admin_savefloodconfig
sql_adminsavefloodconfig="UPDATE " &table_floodconfig& " SET minwait={0},searchrange={1},searchflag={2}"

'admin_stats
sql_adminstats_startdate=			"SELECT startdate FROM " &table_stats
sql_adminstats_insert=				"INSERT INTO " &table_stats& "(startdate) VALUES('{0}')"
sql_adminstats=						"SELECT TOP 1 * FROM " &table_stats
sql_adminstats_client_count=		"SELECT Count(*) FROM " &table_stats_clientinfo
if IsAccess then
	sql_adminstats_client_os=		"SELECT TOP 10 * FROM (SELECT os,Count(os) FROM " &table_stats_clientinfo& " GROUP BY os ORDER BY Count(os) DESC)"
	sql_adminstats_client_browser=	"SELECT TOP 10 * FROM (SELECT browser,Count(browser) FROM " &table_stats_clientinfo& " GROUP BY browser ORDER BY Count(browser) DESC)"
	sql_adminstats_client_screen=	"SELECT TOP 10 * FROM (SELECT screenwh,Count(screenwh) FROM (SELECT screenwidth &'*'& screenheight AS screenwh FROM " &table_stats_clientinfo& ") GROUP BY screenwh ORDER BY Count(screenwh) DESC)"
	sql_adminstats_client_timesect=	"SELECT hsect,Count(hsect) FROM (SELECT Hour(timesect) AS hsect FROM " &table_stats_clientinfo& ") GROUP BY hsect ORDER BY hsect ASC"
	sql_adminstats_client_week=		"SELECT weekno,Count(weekno) FROM (SELECT Weekday(timesect,1) AS weekno FROM " &table_stats_clientinfo& ") GROUP BY weekno ORDER BY weekno ASC"
	sql_adminstats_client_30day=	"SELECT * FROM (SELECT TOP 30 * FROM (SELECT datesect,Count(datesect) FROM (SELECT Year(timesect) & '-' & Month(timesect) & '-' & Day(timesect) AS datesect FROM " &table_stats_clientinfo& ") GROUP BY datesect ORDER BY CDate(datesect) DESC)) ORDER BY CDate(datesect) ASC"
	sql_adminstats_client_source=	"SELECT TOP 10 * FROM (SELECT sourceaddr,Count(sourceaddr) FROM " &table_stats_clientinfo& " GROUP BY sourceaddr ORDER BY Count(sourceaddr) DESC)"
elseif IsSqlServer then
	sql_adminstats_client_os=		"SELECT TOP 10 * FROM (SELECT TOP 100 PERCENT os,Count(os) AS count_os FROM " &table_stats_clientinfo& " GROUP BY os ORDER BY count_os DESC) AS __temp1"
	sql_adminstats_client_browser=	"SELECT TOP 10 * FROM (SELECT TOP 100 PERCENT browser,Count(browser) AS count_browser FROM " &table_stats_clientinfo& " GROUP BY browser ORDER BY count_browser DESC) AS __temp1"
	sql_adminstats_client_screen=	"SELECT TOP 10 * FROM (SELECT TOP 100 PERCENT screenwh,Count(screenwh) AS count_screenwh FROM (SELECT CAST(screenwidth AS varchar) +'*'+ CAST(screenheight AS varchar) AS screenwh FROM " &table_stats_clientinfo& ") AS __temp1 GROUP BY screenwh ORDER BY count_screenwh DESC) AS __temp2"
	sql_adminstats_client_timesect=	"SELECT hsect,Count(hsect) AS count_hsect FROM (SELECT DATEPART(hh,timesect) AS hsect FROM " &table_stats_clientinfo& ") AS __temp1 GROUP BY hsect ORDER BY hsect ASC"
	sql_adminstats_client_week=		"SET DATEFIRST 7; SELECT weekno,Count(weekno) FROM (SELECT DATEPART(dw,timesect) AS weekno FROM " &table_stats_clientinfo& ") AS __temp1 GROUP BY weekno ORDER BY weekno ASC"
	sql_adminstats_client_30day=	"SELECT * FROM (SELECT TOP 30 * FROM (SELECT TOP 100 PERCENT datesect,Count(datesect) AS count_datesect FROM (SELECT CAST(Year(timesect) AS varchar) + '-' + CAST(Month(timesect) AS varchar) + '-' + CAST(Day(timesect) AS varchar) AS datesect FROM " &table_stats_clientinfo& ") AS __temp1 GROUP BY datesect ORDER BY CAST(datesect AS datetime) DESC) AS __temp2) AS __temp3 ORDER BY CAST(datesect AS datetime) ASC"
	sql_adminstats_client_source=	"SELECT TOP 10 * FROM (SELECT TOP 100 PERCENT sourceaddr,Count(sourceaddr) AS count_sourceaddr FROM " &table_stats_clientinfo& " GROUP BY sourceaddr ORDER BY count_sourceaddr DESC) AS __temp1"
end if

'admin_clearstats
sql_adminclearstats_startdate=	"UPDATE " &table_stats& " SET startdate='{0}',[view]=0,[search]=0,[leaveword]=0,[written]=0,[filtered]=0,[banned]=0,[login]=0,[loginfailed]=0"
sql_adminclearstats_client=		"DELETE FROM " &table_stats_clientinfo

'admin_setinfo
sql_adminsetinfo="SELECT TOP 1 name,faceid,faceurl,email,qqid,msnid,homepage FROM " &table_supervisor

'admin_saveinfo
sql_adminsaveinfo="SELECT TOP 1 name,faceid,faceurl,email,qqid,msnid,homepage FROM " &table_supervisor

'admin_savepass
sql_adminsavepass_select="SELECT TOP 1 adminpass FROM " &table_supervisor
sql_adminsavepass_update="UPDATE " &table_supervisor& " SET adminpass='{0}'"

'login_verify
sql_loginverify="SELECT TOP 1 adminpass FROM " &table_supervisor

'admin_verify
sql_adminverify="SELECT TOP 1 adminpass FROM " &table_supervisor

'compact
if IsSqlServer then
	sql_compact_dblog="DECLARE @curr_dbname NVARCHAR(255); SET @curr_dbname=DB_NAME(0); BACKUP LOG @curr_dbname WITH NO_LOG"
	sql_compact_dbfile="DECLARE @curr_dbname NVARCHAR(255); SET @curr_dbname=DB_NAME(0); DBCC SHRINKDATABASE(@curr_dbname,8)"
end if
%>