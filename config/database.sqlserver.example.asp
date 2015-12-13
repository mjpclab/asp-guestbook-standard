<%
const dbtype=10		'数据库类型, 10:MSSQL  11:SQLNCLI  12:SQLNCLI10  13:SQLNCLI11

const dbserver="127.0.0.1"		'数据库服务器地址，可为ip，主机名或命名管道
const dbport=""		'数据库服务器端口号，默认可留空
const dbuserid="user"		'用户名
const dbpassword="pass"		'密码
const dbcatalog="dbname"    '数据库名
const dbschema="dbo"		'架构/所有者，如不清楚用途则设为"dbo"
const dbprefix=""		'表名前缀，多个程序共用一个数据库时，通过表名前缀防止冲突
%>