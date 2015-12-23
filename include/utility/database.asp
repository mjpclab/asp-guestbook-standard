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

Dim dbconnstr
function GetConnStr()
	if IsEmpty(dbconnstr) then
		select case dbtype
		case 1
			dbconnstr=GetAccessConnStr("Microsoft.Jet.OLEDB.4.0")
		case 2
			dbconnstr=GetAccessConnStr("Microsoft.Jet.OLEDB.4.0")
		case 3
			dbconnstr=GetAccessConnStr("Microsoft.ACE.OLEDB.12.0")
		case 10
			dbconnstr=GetSqlServerConnStr("SQLOLEDB.1")
		case 11
			dbconnstr=GetSqlServerConnStr("SQLNCLI")
		case 12
			dbconnstr=GetSqlServerConnStr("SQLNCLI10")
		case 13
			dbconnstr=GetSqlServerConnStr("SQLNCLI11")
		case else
			Err.Raise 513,,"无效的数据库类型号(dbtype)"
		end select
	end if
	GetConnStr=dbconnstr
end function

function CreateConn(byref tconn)
	tconn.ConnectionString=GetConnStr()
	tconn.Open
end function
%>