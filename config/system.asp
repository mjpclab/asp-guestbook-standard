<%
const FacePath="asset/face/"		'ͷ���ļ�����Ŀ¼����ʹ�����·��
const FaceCount=51			'ͷ������
const SmallFacePath="asset/smallface/"		'����ͼ������Ŀ¼����ʹ�����·��
const SmallFaceCount=77					'��������




'�˲��ִ��빩�ڲ�����֮��,�����޸�
const InstanceName=""

Response.Buffer=True
Response.ContentType="text/html; Charset=gbk"

if right(FacePath,1)<>"/" then FacePath=FacePath & "/"
if right(SmallFacePath,1)<>"/" then SmallFacePath=SmallFacePath & "/"
%>