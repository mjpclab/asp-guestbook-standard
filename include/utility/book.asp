<%
Dim BannerUrl,LogoUrl,BreadcrumbItems()
Sub InitHeaderData(PageName)
	if LogoBannerMode then
		BannerUrl = HomeLogo
	else
		LogoUrl = HomeLogo
	end if

	if PageName="" then
		Redim BreadcrumbItems(1,1)
	else
		Redim BreadcrumbItems(2,1)
	end if

	BreadcrumbItems(0,0)=HomeAddr
	BreadcrumbItems(0,1)=HomeName
	BreadcrumbItems(1,0)="index.asp"
	BreadcrumbItems(1,1)="留言本"
	if PageName<>"" then
		BreadcrumbItems(2,0)=""
		BreadcrumbItems(2,1)=PageName
	end if
End Sub
%>