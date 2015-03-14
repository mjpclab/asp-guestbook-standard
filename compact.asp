<!-- #include file="config.asp" -->
<link rel="stylesheet" type="text/css" href="css/style.css"/>
<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
<!-- #include file="css/style.asp" -->
<!-- #include file="css/adminstyle.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
sub fsosupported()
	dim tfso
	on error resume next
	err.number=0
	set tfso=Server.CreateObject("Scripting.FileSystemObject")
	set tfso=nothing
	if err.number<>0 then
		Call MessagePage("服务器不支持FSO(File System Object)，操作无法继续。","index.asp")
		response.end
	end if
	on error goto 0
end sub

Response.Expires=-1
if dbtype=1 or dbtype=2 then
	Call fsosupported
	Dim fso, Engine
	r_dbfile=server.mappath(dbfile)
	r_dbfile_tmp=server.mappath(dbfile) & ".tmp"
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(r_dbfile_tmp) Then fso.DeleteFile (r_dbfile_tmp)

	If fso.FileExists(r_dbfile) Then
		'set ofile=fso.GetFile(r_dbfile)
		'oldsize=ofile.size
		'set ofile=nothing
		
		Set Engine = Server.CreateObject("JRO.JetEngine")
		if dbtype=1 then	'Access97
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & r_dbfile & """","Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & r_dbfile_tmp & """;Jet OLEDB:Engine Type=4"
		elseif dbtype=2 then	'Access2000
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & r_dbfile & """","Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & r_dbfile_tmp & """;Jet OLEDB:Engine Type=5"
		end if
		fso.CopyFile r_dbfile_tmp,r_dbfile
		fso.DeleteFile (r_dbfile_tmp)
		
		'set ofile=fso.GetFile(r_dbfile)
		'newsize=ofile.size
		'set ofile=nothing	
		
	End If
	Set fso = nothing
	Set Engine = nothing
elseif dbtype=10 then
	set cn=Server.CreateObject("ADODB.Connection")
	CreateConn cn,dbtype
	cn.Execute sql_compact_dblog,,1
	cn.Execute sql_compact_dbfile,,1
	
	cn.Close
	set cn=nothing
end if

Call MessagePage("数据库压缩完成。","admin.asp")
%>