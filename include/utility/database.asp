<%
function GetAccessConnStr(dbprovider)
	GetAccessConnStr="Provider=" & dbprovider & ";Data Source=""" & server.Mappath(dbfile) & """;Jet OLEDB:Database Password=""" & Replace(dbfilepassword,"""","""""") & """;"
end function

function GetSqlServerConnStr(dbprovider)
	dim connstr, sql_port
	if dbport<>"" then sql_port="," & dbport else sql_port=""
	connstr="Provider=" & dbprovider & _
		";Data Source=""" & dbserver & sql_port & """" & _
		";User Id=""" & dbuserid & """" & _
		";Password=""" & Replace(dbpassword,"""","""""") & """" & _
		";Application Name=Shen Guest Book Standard - " & InstanceName
	if dbcatalog<>"" then connstr=connstr & ";Initial Catalog=""" & dbcatalog & """"
	GetSqlServerConnStr=connstr
end function

function CreateConn(byref tconn)
	select case dbtype
	case 1
		tconn.ConnectionString=GetAccessConnStr("Microsoft.Jet.OLEDB.4.0")
		tconn.Open
	case 2
		tconn.ConnectionString=GetAccessConnStr("Microsoft.Jet.OLEDB.4.0")
		tconn.Open
	case 3
		tconn.ConnectionString=GetAccessConnStr("Microsoft.ACE.OLEDB.12.0")
		tconn.Open
	case 10
		tconn.ConnectionString=GetSqlServerConnStr("SQLOLEDB.1")
		tconn.Open
	case 11
		tconn.ConnectionString=GetSqlServerConnStr("SQLNCLI")
		tconn.Open
	case 12
		tconn.ConnectionString=GetSqlServerConnStr("SQLNCLI10")
		tconn.Open
	case 13
		tconn.ConnectionString=GetSqlServerConnStr("SQLNCLI11")
		tconn.Open
	case else
		Err.Raise 513,,"无效的数据库类型号(dbtype)"
	end select
end function
%>