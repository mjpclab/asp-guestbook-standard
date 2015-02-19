<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
set cn1=server.CreateObject("ADODB.Connection")
set rs1=server.CreateObject("ADODB.Recordset")
CreateConn cn1,dbtype
rs1.open sql_adminsaveinfo,cn1,0,3,1
		
tname=server.htmlEncode(Request.Form("aname"))

tfaceid=Request.Form("afaceid")
if len(cstr(tfaceid))>3 or Request.Form("afaceurl")<>"" then
	tfaceid=0
elseif isnumeric(tfaceid)=false or tfaceid="" then
	tfaceid=0
elseif clng(tfaceid)<0 or clng(tfaceid)>255 then
	tfaceid=0
else
	tfaceid=cbyte(tfaceid)
end if

tfaceurl=replace(Server.HTMLEncode(Request.Form("afaceurl"))," ","")
if lcase(left(tfaceurl,7))<>"http://" and lcase(left(tfaceurl,6))<>"ftp://" and tfaceurl<>"" then tfaceurl="http://" & tfaceurl

temail=replace(server.htmlEncode(Request.Form("aemail"))," ","")
tqqid=replace(server.htmlEncode(Request.Form("aqqid"))," ","")
tmsnid=replace(server.htmlEncode(Request.Form("amsnid"))," ","")

thomepage=replace(server.htmlEncode(Request.Form("ahomepage"))," ","")
if lcase(left(thomepage,7))<>"http://" and lcase(left(thomepage,6))<>"ftp://" and thomepage<>"" then thomepage="http://" & thomepage

rs1("name")=textfilter(tname,false)
rs1("faceid")=tfaceid
rs1("faceurl")=textfilter(tfaceurl,true)
rs1("email")=textfilter(temail,true)
rs1("qqid")=textfilter(tqqid,false)
rs1("msnid")=textfilter(tmsnid,false)
rs1("homepage")=textfilter(thomepage,true)
rs1.Update

rs1.Close : cn1.close : set rs1=nothing : set cn1=nothing
		
Response.Redirect "admin.asp"
%>