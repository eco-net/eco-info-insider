<!--#include virtual="/connections/ecoinfo.asp" -->
<% 
Dim day
set rsDay = Server.CreateObject("ADODB.Recordset")
rsDay.ActiveConnection = MM_ecoinfo_STRING
rsDay.Source = "SELECT day  FROM eisys_insideruser_count"
rsDay.CursorType = 3
rsDay.CursorLocation = 2
rsDay.LockType = 2
rsDay.Open()
day=rsDay("day")
rsDay.Close()
%>

<%

Function reg(page)
	Application.Lock
	remote=Request.ServerVariables("REMOTE_ADDR")
	if remote<>"80.160.41.182" then
		reguser(page)
		addToAllDays(page)
	else
		regeconet(page)
		addToAllDays("e" & page)
	end if
	checkDay()
	Application.Unlock
end Function

Function regeconet(page)
	if not Session("userid")=Session.SessionID then
		Application("eusers")=Application("eusers") + 1
		addToAllDays("eusers")
	end if
	Session("userid")=Session.SessionID
	
	if page<>"" then
		if page="homeindex" then
			Application("ehomeindex")=Application("ehomeindex") + 1
		elseif page="dgslist" then
			Application("edgslist")=Application("edgslist") + 1
		elseif page="dgblist" then
			Application("edgblist")=Application("edgblist") + 1
		elseif page="oklist" then
			Application("eoklist")=Application("eoklist") + 1
		elseif page="kurlist" then
			Application("ekurlist")=Application("ekurlist") + 1
		elseif page="uddlist" then
			Application("euddlist")=Application("euddlist") + 1
		elseif page="artlist" then
			Application("eartlist")=Application("eartlist") + 1
		end if
	end if
end Function

Function reguser(page)
	if not Session("userid")=Session.SessionID then
		Application("users")=Application("users") + 1
		addToAllDays("users")
	end if
	Session("userid")=Session.SessionID
	if page<>"" then
		if page="homeindex" then
			Application("homeindex")=Application("homeindex") + 1
		elseif page="dgslist" then
			Application("dgslist")=Application("dgslist") + 1
		elseif page="dgblist" then
			Application("dgblist")=Application("dgblist") + 1
		elseif page="oklist" then
			Application("oklist")=Application("oklist") + 1
		elseif page="kurlist" then
			Application("kurlist")=Application("kurlist") + 1
		elseif page="uddlist" then
			Application("uddlist")=Application("uddlist") + 1
		elseif page="artlist" then
			Application("artlist")=Application("artlist") + 1
		end if
	end if
end Function

Function addToAllDays(att)
	set rsAll = Server.CreateObject("ADODB.Recordset")
	rsAll.ActiveConnection = MM_ecoinfo_STRING
	rsAll.Source = "SELECT * FROM eisys_insideruser_count"
	rsAll.CursorType = 0
	rsAll.CursorLocation = 2
	rsAll.LockType = 2
	rsAll.Open()
	on error resume next
	if rsAll(att)=0 then
		c=1
	else
		c=rsAll(att) + 1
	end if
	if err then
		att=replace(att,"index","list")
		if rsAll(att)=0 then
			c=1
		else
			c=rsAll(att) + 1
		end if
	end if
	rsAll(att)=c
	rsAll.Update
	rsAll.Close
end Function

Function checkDay()
	if day<>DatePart("d",Date) then
		saveDay()
	end if
end Function

Function saveDay()
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.ActiveConnection = MM_ecoinfo_STRING
	rs.Source = "SELECT *  FROM eisys_insideruser_count"
	rs.CursorType = 0
	rs.CursorLocation = 2
	rs.LockType = 2
	rs.Open()
	rs("day")=DatePart("d",Date)
	rs.Update
	rs.Close()
	initApp()
end Function

Function initApp()
	Application("homeindex")=0
	Application("ehomeindex")=0
	Application("dgslist")=0
	Application("edgslist")=0
	Application("dgblist")=0
	Application("edgblist")=0
	Application("oklist")=0
	Application("eoklist")=0
	Application("kurlist")=0
	Application("ekurlist")=0
	Application("uddlist")=0
	Application("euddlist")=0
	Application("artlist")=0
	Application("eartlist")=0
	Application("users")=0
	Application("eusers")=0
end Function
%>