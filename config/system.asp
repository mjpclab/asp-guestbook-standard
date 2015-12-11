<%
const FacePath="face/"		'头像文件所在目录，可使用相对路径
const FaceCount=51			'头像总数
const SmallFacePath="smallface/"		'表情图标所在目录，可使用相对路径
const SmallFaceCount=77					'表情总数




'此部分代码供内部处理之用,请勿修改
const InstanceName=""

Session.CodePage=936
Response.Buffer=True
Response.ContentType="text/html"
Response.CharSet="gbk"


if right(FacePath,1)<>"/" then FacePath=FacePath & "/"
if right(SmallFacePath,1)<>"/" then SmallFacePath=SmallFacePath & "/"
%>