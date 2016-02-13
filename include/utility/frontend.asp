<%
function bodylimit()
	dim limitstr
	limitstr=""

	if ShowContext=false then limitstr=limitstr & " oncontextmenu=""return false;"""
	if SelectContent=false then limitstr=limitstr & " onselectstart=""return false;"""
	if CopyContent=false then limitstr=limitstr & " oncopy=""return false;"""

	bodylimit=limitstr
end function

function framecheck()
	if BeFramed=false then
		framecheck="if (window.top!=window) window.top.location.href=window.location.href;"
	else
		framecheck=""
	end if
end function

function cked(exp)
	if exp=true then
		cked=" checked=""checked"""
	else
		cked=""
	end if
end function

function dised(exp)
	if exp=true then
		dised=" disabled=""disabled"""
	else
		dised=""
	end if
end function

function seled(exp)
	if exp=true then
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