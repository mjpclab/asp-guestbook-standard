<%@ CodePage="936" Language="VBScript" %>
<%
const dbtype=2		'���ݿ�����,1:Access97  2:Access2000-2003  3:Access 2007-2013  10:MSSQL
const prefix=""		'����ǰ׺�����������һ�����ݿ�ʱ��ͨ������ǰ׺��ֹ��ͻ

'<Access���ݿ����>
const dbfile="database/data.mdb.db"		'���ݿ��ļ�λ�ã�ʹ�����·��
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
Response.ContentType="text/html"
Response.CharSet="gbk"


if right(FacePath,1)<>"/" then FacePath=FacePath & "/"
if FrequentFaceCount>FaceCount then FrequentFaceCount=FaceCount

if right(SmallFacePath,1)<>"/" then SmallFacePath=SmallFacePath & "/"
%>