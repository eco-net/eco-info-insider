<!--#include virtual="/shared/sqlcheck.asp"-->
<%
'**** Calculating sql-statement

DIM theWhere, searchDescr
theWhere="0=0"


if request("cat")<>"" then
	theWhere=theWhere & " AND i.cat_id=" & request("cat")
end if
if request("size")<>"" then
	theWhere=theWhere & " AND i.size_id=" & request("size")
end if
if request("sizes")<>"" then
	if request("sizes")="1" then
	theWhere=theWhere & " AND i.branding=1 "
	end if
	if request("sizes")="2" then
	theWhere=theWhere & " AND i.splash=1 "
	end if
	if request("sizes")="4" then
	theWhere=theWhere & " AND i.rightcol=1 "
	end if
	if request("sizes")="5" then
	theWhere=theWhere & " AND i.threecol=1 "
	end if
	if request("sizes")="6" then
	theWhere=theWhere & " AND i.twocol=1 "
	end if
	if request("sizes")="7" then
	theWhere=theWhere & " AND i.tema=1 "
	end if
end if
if request("img_id")<>"" then
	theWhere=theWhere & " AND i.id=" & request("img_id")
end if 
if request("searchtext")<>"" then
	theWhere=theWhere & " AND (i.filename LIKE '%" & request("searchtext") &_
			"%' OR i.subtext LIKE '%" & request("searchtext") &_
			"%' OR i.source LIKE '%" & request("searchtext") & "%')"
end if
%>
