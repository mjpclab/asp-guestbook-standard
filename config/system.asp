<%
const FacePath="asset/face/"		'头像文件所在目录，可使用相对路径
const FaceCount=51			'头像总数
const SmallFacePath="asset/smallface/"		'表情图标所在目录，可使用相对路径
const SmallFaceCount=77					'表情总数




'此部分代码供内部处理之用,请勿修改
const InstanceName=""

Response.Buffer=True
Response.ContentType="text/html; Charset=gbk"

if right(FacePath,1)<>"/" then FacePath=FacePath & "/"
if right(SmallFacePath,1)<>"/" then SmallFacePath=SmallFacePath & "/"
%>