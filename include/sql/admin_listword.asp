<%
sql_listword_admin="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE parent_id={0} ORDER BY id ASC"
%>