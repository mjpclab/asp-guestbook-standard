<%@ CodePage="936" Language="VBScript" %>
<%
const dbtype=2		'���ݿ�����,1:Access97  2:Access2000/2002/2003  10:MSSQL
const prefix=""		'����ǰ׺�����������һ�����ݿ�ʱ��ͨ������ǰ׺��ֹ��ͻ

'<Access���ݿ����>
const dbfile="database/data.asp"		'���ݿ��ļ�λ�ã�ʹ�����·��
'</Access���ݿ����>

'<SQL Server���ݿ����>
const dbserver=""
const dbport=""
const dbuserid=""
const dbpassword=""
const dbcatalog=""
const dbschema="dbo"
'</SQL Server���ݿ����>

const FacePath="face/"		'ͷ���ļ�����Ŀ¼����ʹ�����·��
const FaceCount=51			'ͷ������
const SmallFacePath="smallface/"		'����ͼ������Ŀ¼����ʹ�����·��
const SmallFaceCount=77					'��������

%>




<!-- #include file="loadconfig.asp" -->
<%'�˲��ִ��빩�ڲ�����֮��,�����޸�
const InstanceName=""

Session.CodePage=936
Response.Buffer=True
'if InStr(Request.ServerVariables("HTTP_ACCEPT"),"text/xml")>0 then
'	Response.ContentType="text/xml"
'elseif IsEmpty(Request.ServerVariables("HTTP_ACCEPT")) then
'	Response.ContentType="text/xml"
'elseif Trim(Request.ServerVariables("HTTP_ACCEPT"))="" then
'	Response.ContentType="text/xml"
'else
	Response.ContentType="text/html"
'end if
Response.CharSet="gb2312"


if right(FacePath,1)<>"/" then FacePath=FacePath & "/"
if FrequentFaceCount>FaceCount then FrequentFaceCount=FaceCount

if right(SmallFacePath,1)<>"/" then SmallFacePath=SmallFacePath & "/"
%>