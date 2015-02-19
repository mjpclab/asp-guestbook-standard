<!-- #include file="config.asp" -->
<%
if session("gotclientinfo")<>true then
	Dim cn
	Dim os,browser,screenwidth,screenheight,sourceaddr,fullsource

	'Get Parameters
	os=Request.QueryString("sys")
	browser=Request.QueryString("brow")
	screenwidth=Request.QueryString("sw")
	screenheight=Request.QueryString("sh")
	sourceaddr=Request.QueryString("src")
	fullsource=Request.QueryString("fsrc")

	if not (os="" and browser="" and screenwidth="" and screenheight="" and sourceaddr="" and fullsource="") then
		'Verify Parameters
		if len(os)>25 then os=left(os,25)
		if os="" then os="未知"

		if len(browser)>16 then browser=left(browser,16)
		if browser="" then browser="未知"

		if isnumeric(screenwidth)=false or len(cstr(screenwidth))>9 then screenwidth=0 else screenwidth=abs(clng(screenwidth))
		if isnumeric(screenheight)=false or len(cstr(screenheight))>9 then screenheight=0 else screenheight=abs(clng(screenheight))
		if screenwidth=0 or screenheight=0 then
			screenwidth=0
			screenheight=0
		end if

		if len(sourceaddr)>50 then sourceaddr=left(sourceaddr,50)
		if sourceaddr="" then sourceaddr="收藏夹或手动输入"

		if len(fullsource)>255 then fullsource=left(fullsource,255)
		if fullsource="" then fullsource="收藏夹或手动输入"

		'Filter Parameters
		os=replace(replace(os,"'","''"),"[","[[]")
		browser=replace(replace(browser,"'","''"),"[","[[]")
		sourceaddr=replace(replace(sourceaddr,"'","''"),"[","[[]")
		fullsource=replace(replace(fullsource,"'","''"),"[","[[]")

		'Save Parameters
		set cn=server.CreateObject("ADODB.Connection")
		CreateConn cn,dbtype
		cn.Execute Replace(Replace(Replace(Replace(Replace(Replace(Replace(sql_saveclientinfo,"{0}",os),"{1}",browser),"{2}",screenwidth),"{3}",screenheight),"{4}",now()),"{5}",sourceaddr),"{6}",fullsource),,1
		cn.Close : set cn=nothing
		
		'Got Complete
		session("gotclientinfo")=true
	end if
end if
%>