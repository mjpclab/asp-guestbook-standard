<%
if IsAccess then
	sql_adminhidecontact=	"UPDATE " &table_main& " SET guestflag=guestflag+256 WHERE (guestflag mod 512)\256=0 AND id="
	sql_adminunhidecontact=	"UPDATE " &table_main& " SET guestflag=guestflag-256 WHERE (guestflag mod 512)\256<>0 AND id="
	sql_adminhideword=		"UPDATE " &table_main& " SET guestflag=guestflag+8 WHERE (guestflag mod 16)\8=0 AND id="
	sql_adminunhideword=	"UPDATE " &table_main& " SET guestflag=guestflag-8 WHERE (guestflag mod 16)\8<>0 AND id="
	sql_adminlockreply=		"UPDATE " &table_main& " SET guestflag=guestflag+512 WHERE (guestflag mod 1024)\512=0 AND id="
	sql_adminunlockreply=	"UPDATE " &table_main& " SET guestflag=guestflag-512 WHERE (guestflag mod 1024)\512<>0 AND id="
	sql_adminpassaudit=		"UPDATE " &table_main& " SET guestflag=guestflag-16 WHERE (guestflag mod 32)\16<>0 AND id="
	sql_adminpubwhisper=	"UPDATE " &table_main& " SET guestflag=guestflag-32 WHERE (guestflag mod 64)\32<>0 AND id="
elseif IsSqlServer then
	sql_adminhidecontact=	"UPDATE " &table_main& " SET guestflag=guestflag+256 WHERE guestflag & 256=0 AND id="
	sql_adminunhidecontact=	"UPDATE " &table_main& " SET guestflag=guestflag-256 WHERE guestflag & 256<>0 AND id="
	sql_adminhideword=		"UPDATE " &table_main& " SET guestflag=guestflag+8 WHERE guestflag & 8=0 AND id="
	sql_adminunhideword=	"UPDATE " &table_main& " SET guestflag=guestflag-8 WHERE guestflag & 8<>0 AND id="
	sql_adminlockreply=		"UPDATE " &table_main& " SET guestflag=guestflag+512 WHERE guestflag & 512=0 AND id="
	sql_adminunlockreply=	"UPDATE " &table_main& " SET guestflag=guestflag-512 WHERE guestflag & 512<>0 AND id="
	sql_adminpassaudit=		"UPDATE " &table_main& " SET guestflag=guestflag-16 WHERE guestflag & 16<>0 AND id="
	sql_adminpubwhisper=	"UPDATE " &table_main& " SET guestflag=guestflag-32 WHERE guestflag & 32<>0 AND id="
end if
sql_adminlock2top=			"UPDATE " &table_main& " SET parent_id=-1 WHERE parent_id=0 AND id="
sql_adminunlock2top=		"UPDATE " &table_main& " SET parent_id=0 WHERE parent_id<0 AND id="
sql_adminbring2top=			"UPDATE " &table_main& " SET lastupdated='{0}' WHERE parent_id<=0 AND id={1}"
sql_adminreorder=			"UPDATE " &table_main& " SET parent_id=0,lastupdated=logdate WHERE parent_id<=0 AND id={0}"
%>