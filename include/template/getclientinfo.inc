<%if StatusStatistics and session("gotclientinfo")<>true then%>
<script type="text/javascript" src="asset/js/stats.js"></script>
<script type="text/javascript">
	var url = location.href.substring(0, location.href.lastIndexOf('/') + 1);
	url += 'saveclientinfo.asp?sys=' + encodeURIComponent(getsysname()) +
		'&brow=' + encodeURIComponent(getbrowsername()) +
		'&sw=' + encodeURIComponent(getscreenwidth()) +
		'&sh=' + encodeURIComponent(getscreenheight()) +
		'&src=' + encodeURIComponent(getsourceaddr()) +
		'&fsrc=' + encodeURIComponent(document.referrer);
	RequestScript(url);
</script>
<%end if%>