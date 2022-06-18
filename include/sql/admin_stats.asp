<%
sql_adminstats_startdate=			"SELECT startdate FROM " &table_stats
sql_adminstats_insert=				"INSERT INTO " &table_stats& "(startdate) VALUES('{0}')"
sql_adminstats=						"SELECT TOP 1 * FROM " &table_stats
sql_adminstats_client_count=		"SELECT Count(*) FROM " &table_stats_clientinfo
if IsAccess then
	sql_adminstats_client_os=		"SELECT TOP 30 * FROM (SELECT os,Count(os) AS count_os FROM " &table_stats_clientinfo& " GROUP BY os ORDER BY Count(os) DESC) ORDER BY count_os DESC"
	sql_adminstats_client_browser=	"SELECT TOP 30 * FROM (SELECT browser,Count(browser) AS count_browser FROM " &table_stats_clientinfo& " GROUP BY browser ORDER BY Count(browser) DESC) ORDER BY count_browser DESC"
	sql_adminstats_client_screen=	"SELECT TOP 30 * FROM (SELECT screenwh,Count(screenwh) AS count_screenwh FROM (SELECT screenwidth &'*'& screenheight AS screenwh FROM " &table_stats_clientinfo& ") GROUP BY screenwh ORDER BY Count(screenwh) DESC) ORDER BY count_screenwh DESC"
	sql_adminstats_client_timesect=	"SELECT hsect,Count(hsect) FROM (SELECT Hour(timesect) AS hsect FROM " &table_stats_clientinfo& ") GROUP BY hsect ORDER BY hsect ASC"
	sql_adminstats_client_week=		"SELECT weekno,Count(weekno) FROM (SELECT Weekday(timesect,1) AS weekno FROM " &table_stats_clientinfo& ") GROUP BY weekno ORDER BY weekno ASC"
	sql_adminstats_client_30day=	"SELECT * FROM (SELECT TOP 30 * FROM (SELECT datesect,Count(datesect) FROM (SELECT Year(timesect) & '-' & Month(timesect) & '-' & Day(timesect) AS datesect FROM " &table_stats_clientinfo& ") GROUP BY datesect ORDER BY CDate(datesect) DESC)) ORDER BY CDate(datesect) ASC"
	sql_adminstats_client_source=	"SELECT TOP 30 * FROM (SELECT sourceaddr,Count(sourceaddr) AS count_sourceaddr FROM " &table_stats_clientinfo& " GROUP BY sourceaddr ORDER BY Count(sourceaddr) DESC) ORDER BY count_sourceaddr DESC"
elseif IsSqlServer then
	sql_adminstats_client_os=		"SELECT TOP 30 * FROM (SELECT TOP 100 PERCENT os,Count(os) AS count_os FROM " &table_stats_clientinfo& " GROUP BY os ORDER BY count_os DESC) AS __temp1 ORDER BY count_os DESC"
	sql_adminstats_client_browser=	"SELECT TOP 30 * FROM (SELECT TOP 100 PERCENT browser,Count(browser) AS count_browser FROM " &table_stats_clientinfo& " GROUP BY browser ORDER BY count_browser DESC) AS __temp1 ORDER BY count_browser DESC"
	sql_adminstats_client_screen=	"SELECT TOP 30 * FROM (SELECT TOP 100 PERCENT screenwh,Count(screenwh) AS count_screenwh FROM (SELECT CAST(screenwidth AS varchar) +'*'+ CAST(screenheight AS varchar) AS screenwh FROM " &table_stats_clientinfo& ") AS __temp1 GROUP BY screenwh ORDER BY count_screenwh DESC) AS __temp2 ORDER BY count_screenwh DESC"
	sql_adminstats_client_timesect=	"SELECT hsect,Count(hsect) AS count_hsect FROM (SELECT DATEPART(hh,timesect) AS hsect FROM " &table_stats_clientinfo& ") AS __temp1 GROUP BY hsect ORDER BY hsect ASC"
	sql_adminstats_client_week=		"SET DATEFIRST 7; SELECT weekno,Count(weekno) FROM (SELECT DATEPART(dw,timesect) AS weekno FROM " &table_stats_clientinfo& ") AS __temp1 GROUP BY weekno ORDER BY weekno ASC"
	sql_adminstats_client_30day=	"SELECT * FROM (SELECT TOP 30 * FROM (SELECT TOP 100 PERCENT datesect,Count(datesect) AS count_datesect FROM (SELECT CAST(Year(timesect) AS varchar) + '-' + CAST(Month(timesect) AS varchar) + '-' + CAST(Day(timesect) AS varchar) AS datesect FROM " &table_stats_clientinfo& ") AS __temp1 GROUP BY datesect ORDER BY CAST(datesect AS datetime) DESC) AS __temp2 ORDER BY CAST(datesect AS datetime) DESC) AS __temp3 ORDER BY CAST(datesect AS datetime) ASC"
	sql_adminstats_client_source=	"SELECT TOP 30 * FROM (SELECT TOP 100 PERCENT sourceaddr,Count(sourceaddr) AS count_sourceaddr FROM " &table_stats_clientinfo& " GROUP BY sourceaddr ORDER BY count_sourceaddr DESC) AS __temp1 ORDER BY count_sourceaddr DESC"
end if
%>