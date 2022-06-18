<%
function FilterGuestLike(byref str)
	dim r_str
	r_str=str
	
	r_str=replace(r_str,"'","''")
	r_str=replace(r_str,"[","[[]")
	r_str=replace(r_str,"%","[%]")
	r_str=replace(r_str,"_","[_]")
	r_str=replace(r_str,"""","")
	
	FilterGuestLike=r_str
end function

function FilterAdminLike(byref str)
	dim r_str
	r_str=str
	
	r_str=replace(r_str,"'","''")
	r_str=replace(r_str,"[","[[]")
	r_str=replace(r_str,"""","")
	
	FilterAdminLike=r_str
end function

function FilterStr(byref str,byval delquote)
	dim r_str
	r_str=str
	
	if delquote then r_str=replace(r_str,"'","")
	r_str=FilterGuestLike(r_str)
	
	FilterStr=r_str
end function

function FilterSqlStr(byref str)
	dim r_str
	r_str=str
	
	r_str=replace(r_str,"'","''")
	r_str=replace(r_str,"[","[[]")
	r_str=replace(r_str,"%","[%]")
	r_str=replace(r_str,"_","[_]")
	'r_str=replace(r_str,"""","")	'the difference between FilterGuestLike()
	
	FilterSqlStr=r_str
end function

function FilterQuote(byref str)
	FilterQuote=Replace(str,"'","''")
end function

function FilterSql(byref str)
	dim r_str
	r_str=str
	
	r_str=replace(r_str,"'","")
	r_str=replace(r_str,";","")
	r_str=replace(r_str," ","")
	r_str=replace(r_str,"[","")
	r_str=replace(r_str,"--","")
	r_str=replace(r_str,"/*","")
	r_str=replace(r_str,"#","")
	r_str=replace(r_str,"@","")
	r_str=replace(r_str,"/","")
	r_str=replace(r_str,"\","")
	r_str=replace(r_str,"?","")
	r_str=replace(r_str,"&","")
	r_str=replace(r_str,"=","")
	
	FilterSql=r_str
end function

function FilterKeyword(byref str)
	dim r_str
	r_str=str

	r_str=replace(r_str,"'","")
	r_str=replace(r_str,"""","")
	r_str=replace(r_str,",","")
	r_str=replace(r_str,".","")
	r_str=replace(r_str,":","")
	r_str=replace(r_str,";","")
	r_str=replace(r_str,"`","")
	r_str=replace(r_str,"~","")
	r_str=replace(r_str," ","")
	r_str=replace(r_str,"(","")
	r_str=replace(r_str,")","")
	r_str=replace(r_str,"[","")
	r_str=replace(r_str,"]","")
	r_str=replace(r_str,"{","")
	r_str=replace(r_str,"}","")
	r_str=replace(r_str,"^","")
	r_str=replace(r_str,"+","")
	r_str=replace(r_str,"-","")
	r_str=replace(r_str,"*","")
	r_str=replace(r_str,"#","")
	r_str=replace(r_str,"@","")
	r_str=replace(r_str,"$","")
	r_str=replace(r_str,"%","")
	r_str=replace(r_str,"/","")
	r_str=replace(r_str,"\","")
	r_str=replace(r_str,"?","")
	r_str=replace(r_str,"&","")
	r_str=replace(r_str,"=","")

	FilterKeyword=r_str
end function
%>