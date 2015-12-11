<%
sub show_book_title(byval layercount, byval param)%>
<div class="header">
	<%if HomeLogo<>"" then%><img class="logo<%if LogoBannerMode then Response.write " banner"%>" src="<%=HomeLogo%>" alt="<%=HomeName%>"/><%end if%>
	<div class="breadcrumb">
		<%if HomeName<>"" then
			if HomeAddr<>"" then%>
				<a class="name" href="<%=HomeAddr%>"><%=HomeName%></a>
			<%else%>
				<span class="name"><%=HomeName%></span>
			<%end if
		end if%>

		<%if layercount=2 then%>
			<%if HomeName<>"" then%><span class="separator">&gt;&gt;</span><%end if%>
			<span class="guestbook">¡Ù—‘±æ</span>
		<%elseif layercount=3 then%>
			<%if HomeName<>"" then%><span class="separator">&gt;&gt;</span><%end if%>
			<a class="guestbook" href="index.asp">¡Ù—‘±æ</a>
			<span class="separator">&gt;&gt;</span>
			<span class="page"><%=param%></span>
		<%end if%>
	</div>
</div>
<%end sub
%>