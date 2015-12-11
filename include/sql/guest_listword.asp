<%
sql_listword_guest="SELECT * FROM " &table_main& " LEFT JOIN " &table_reply& " ON " &table_main& ".id=" &table_reply& ".articleid WHERE parent_id={0}{1} ORDER BY id ASC"
%>