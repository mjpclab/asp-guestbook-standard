<%@ CodePage="936" Language="VBScript" %>
<%
const dbtype=2		'数据库类型,1:Access97  2:Access2000-2003  3:Access 2007-2013  10:MSSQL
const prefix=""		'表名前缀，多个程序共用一个数据库时，通过表名前缀防止冲突

'<Access数据库参数>
const dbfile="database/data.mdb.db"		'数据库文件位置，使用相对路径
'</Access数据库参数>

'<SQL Server数据库参数>
const dbserver=""
const dbport=""
const dbuserid=""
const dbpassword=""
const dbcatalog=""
const dbschema="dbo"
'</SQL Server数据库参数>

const FacePath="face/"		'头像文件所在目录，可使用相对路径
const FaceCount=51			'头像总数
const SmallFacePath="smallface/"		'表情图标所在目录，可使用相对路径
const SmallFaceCount=77					'表情总数

%>




<!-- #include file="loadconfig.asp" -->
<%'此部分代码供内部处理之用,请勿修改
const InstanceName=""

Session.CodePage=936
Response.Buffer=True
Response.ContentType="text/html"
Response.CharSet="gbk"


if right(FacePath,1)<>"/" then FacePath=FacePath & "/"
if FrequentFaceCount>FaceCount then FrequentFaceCount=FaceCount

if right(SmallFacePath,1)<>"/" then SmallFacePath=SmallFacePath & "/"
%>