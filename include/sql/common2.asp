<%
sql_common2_getadmininfo=	"SELECT name,faceid,faceurl,email,qqid,msnid,homepage FROM " &table_supervisor
'sql_common2_reply=			"SELECT replydate,htmlflag,reinfo FROM " &table_reply& " WHERE articleid="
sql_common2_guestreply=		"SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE parent_id={0}{1} ORDER BY id ASC"
sql_common2_replyinform=	"SELECT parent_id,email,guestflag,logdate,article FROM " &table_main& " WHERE id="
%>