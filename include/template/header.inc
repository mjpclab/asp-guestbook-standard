<div class="header">
	<%if BannerUrl<>"" then%>
		<%if HomeAddr<>"" then%>
			<div class="banner"><a href="<%=HomeAddr%>"><img src="<%=BannerUrl%>" alt="<%=HomeName%>"/></a></div>
		<%else%>
			<div class="banner"><img src="<%=BannerUrl%>" alt="<%=HomeName%>"/></div>
		<%end if%>
	<%end if%>

	<div class="header-container">
	<%if LogoUrl<>"" then%>
		<%if HomeAddr<>"" then%>
			<div class="logo"><a href="<%=HomeAddr%>"><img src="<%=LogoUrl%>" alt="<%=HomeName%>"/></a></div>
		<%else%>
			<div class="logo"><img src="<%=LogoUrl%>" alt="<%=HomeName%>"/></div>
		<%end if%>
	<%end if%>

	<%if Not IsEmpty(BreadcrumbItems) then%>
	<div class="breadcrumb">
		<% Dim itemIndex, url, text, renderedCount
		renderedCount=0
		for itemIndex = 0 to UBound(BreadcrumbItems)
			url=BreadcrumbItems(itemIndex,0)
			text=BreadcrumbItems(itemIndex,1)
			if text<>"" then
				if renderedCount>0 then%><span class="separator">&raquo;</span><%end if
				renderedCount=renderedCount+1
				if url<>"" then%>
					<a href="<%=url%>"><%=text%></a>
				<%else%>
					<span><%=text%></span>
				<%end if
			end if
		next%>
	</div>
	<%end if%>
	</div>
</div>