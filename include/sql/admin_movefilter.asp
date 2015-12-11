<%
sql_adminmovefilter_sort1=			"SELECT filtersort FROM " &table_filterconfig& " WHERE filterid="
sql_adminmovefilter_sort2_up=		"SELECT MAX(filtersort) FROM " &table_filterconfig& " WHERE filtersort<"
sql_adminmovefilter_sort2_down=		"SELECT MIN(filtersort) FROM " &table_filterconfig& " WHERE filtersort>"
sql_adminmovefilter_sort2_top=		"SELECT MIN(filtersort) FROM " &table_filterconfig
sql_adminmovefilter_sort2_bottom=	"SELECT MAX(filtersort) FROM " &table_filterconfig
sql_adminmovefilter_id2=			"SELECT filterid FROM " &table_filterconfig& " WHERE filtersort="
sql_adminmovefilter_update_updown1=	"UPDATE " &table_filterconfig& " SET filtersort=-1 WHERE filterid="
sql_adminmovefilter_update_updown2=	"UPDATE " &table_filterconfig& " SET filtersort={0} WHERE filtersort={1}"
sql_adminmovefilter_update_updown3=	"UPDATE " &table_filterconfig& " SET filtersort={0} WHERE filtersort=-1"
sql_adminmovefilter_update_top1=	"UPDATE " &table_filterconfig& " SET filtersort=filtersort+1 WHERE filtersort<"
sql_adminmovefilter_update_top2=	"UPDATE " &table_filterconfig& " SET filtersort={0} WHERE filterid={1}"
sql_adminmovefilter_update_bottom1=	"UPDATE " &table_filterconfig& " SET filtersort=filtersort-1 WHERE filtersort>"
sql_adminmovefilter_update_bottom2=	"UPDATE " &table_filterconfig& " SET filtersort={0} WHERE filterid={1}"
%>