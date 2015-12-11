<%
sql_adminreply_reply="SELECT htmlflag,reinfo FROM " &table_reply& " WHERE articleid="
sql_adminreply_words="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE id="
%>