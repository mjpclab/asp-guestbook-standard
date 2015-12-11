<%
Dim CurrentDisplayMode
CurrentDisplayMode=AdminDisplayMode()
if CurrentDisplayMode="book" then
	if Request("type")<>"" and Request("searchtxt")<>"" then
		Response.Redirect "admin_search.asp?page=" & Request("page") & "&type=" & server.URLEncode(Request("type")) & "&searchtxt=" & server.URLEncode(Request("searchtxt"))
	else
		Response.Redirect "admin.asp?page=" & Request("page")
	end if
elseif CurrentDisplayMode="forum" then
	Response.Redirect "admin_showword.asp?id=" & Request("rootid") & "&page=" & Request("page") & "&type=" & server.URLEncode(Request("type")) & "&searchtxt=" & server.URLEncode(Request("searchtxt"))
end if
%>