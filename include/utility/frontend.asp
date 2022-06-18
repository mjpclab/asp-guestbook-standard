<%
function bodylimit()
	dim limitstr
	limitstr=""

	if Not ShowContext then limitstr=limitstr & " oncontextmenu=""return false;"""
	if Not SelectContent then limitstr=limitstr & " onselectstart=""return false;"""
	if Not CopyContent then limitstr=limitstr & " oncopy=""return false;"""

	bodylimit=limitstr
end function

function framecheck()
	if Not BeFramed then
		framecheck="if (window.top!=window) window.top.location.href=window.location.href;"
	else
		framecheck=""
	end if
end function

function cked(exp)
	if exp then
		cked=" checked=""checked"""
	else
		cked=""
	end if
end function

function dised(exp)
	if exp then
		dised=" disabled=""disabled"""
	else
		dised=""
	end if
end function

function seled(exp)
	if exp then
		seled=" selected=""selected"""
	else
		seled=""
	end if
end function

Function CssOptionalSize(property,size)
	if(property<>"") then
		if(isnumeric(size)) then
			CssOptionalSize=property & ":" & size & "px;"
		else
			CssOptionalSize=property & ":" & size & ";"
		end if
	else
		CssOptionalSize=""
	end if
End Function

Function CssOptionalSeparator(value,separator)
	Dim result
	result = Trim(value)
	if result <> "" then
		result = result & separator
	end if

	CssOptionalSeparator = result
End Function
%>