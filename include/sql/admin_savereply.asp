<%
sql_adminsavereply_main=			"SELECT guestflag,replied FROM " &table_main& " WHERE id="
sql_adminsavereply_reply=			"SELECT articleid FROM " &table_reply& " WHERE articleid="
if IsAccess then
	sql_adminsavereply_insert=		"INSERT INTO " &table_reply& "(articleid,replydate,htmlflag,reinfo) VALUES('{0}','{1}',{2},'{3}')"
	sql_adminsavereply_update=		"UPDATE " &table_reply& " SET replydate='{0}',htmlflag={1},reinfo='{2}' WHERE articleid={3}"
elseif IsSqlServer then
	sql_adminsavereply_insert=		"INSERT INTO " &table_reply& "(articleid,replydate,htmlflag,reinfo) VALUES('{0}','{1}',{2},N'{3}')"
	sql_adminsavereply_update=		"UPDATE " &table_reply& " SET replydate='{0}',htmlflag={1},reinfo=N'{2}' WHERE articleid={3}"
end if
sql_adminsavereply_lock2top=    	"UPDATE " &table_main& " SET parent_id=-1 WHERE parent_id=0 AND id=(SELECT root_id FROM " &table_main& " WHERE id={0})"
sql_adminsavereply_bring2top=	    "UPDATE " &table_main& " SET lastupdated='{0}' WHERE parent_id<=0 AND id=(SELECT root_id FROM " &table_main& " WHERE id={1})"
%>