<%
function CreateConn(byref tconn,byval tDBType)
	'On Error Resume Next

	select case tDBType
	case 1
		tconn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & server.Mappath(dbfile) & """;Jet OLEDB:Database Password=""" & dbfilepassword & """;"
		tconn.Open
	case 2
		tconn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & server.Mappath(dbfile) & """;Jet OLEDB:Database Password=""" & dbfilepassword & """;"
		tconn.Open
	case 3
		tconn.ConnectionString="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""" & server.Mappath(dbfile) & """;Jet OLEDB:Database Password=""" & dbfilepassword & """;"
		tconn.Open
	case 10
		dim sql_port
		if dbport<>"" then sql_port="," & dbport else sql_port=""
		tconn.ConnectionString= _
			" Provider=SQLOLEDB.1" & _
			";Data Source=""" & dbserver & sql_port & """" & _
			";User Id=""" & dbuserid & """" & _
			";Password=""" & dbpassword & """" & _
			";Application Name=Shen Guest Book Standard - " & InstanceName & _
			""
		if dbcatalog<>"" then tconn.ConnectionString=tconn.ConnectionString & ";Initial Catalog=""" & dbcatalog & """"
		tconn.Open
	case else
		Err.Raise 513,,"无效的数据库类型号(dbtype)"
	end select

	'If Err.number<>0 Then
		'Response.Write "连接数据库出错，请检查配置文件。错误描述：<br/><br/>" & Chr(13) & Chr(10)
		'Response.Write Err.Description		'显示错误消息
		'Response.End
	'End If
end function
%>