<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/shared/common.asp" -->

<%
Dim calid,theRS,orgid,status,theID
status=2
if (request.cookies("eiorgname") = "&Oslash;ko-net medarbejder") then 
status=1
end if
calid=request("calid")

SET theRS = Server.CreateObject("ADODB.recordset")
	theRS.ActiveConnection=MM_ecoinfo_STRING
	
	'theOrg
	theRS.Source="SELECT orgid FROM eical_r_org WHERE calid='" & calid & "'"
	
	theRS.CursorType = 0
	theRS.CursorLocation = 2
	theRS.LockType = 3
	
	theRS.Open()
	orgid=theRS.Fields.Item("orgid").Value
	
	theRS.Close()


	'theCal
	theRS.Source="SELECT * FROM eical_maindata WHERE id='" & calid & "'"
	theRS.Open()
	
	theTitle=replace(theRS("title"),"'","''") 
	
	theTitle = theTitle & " KOPI"
	theTitle = replace(theTitle,"'","''") 
		
	'1. Inserting copy of cal
	theSQL="INSERT INTO eical_maindata(" &_
	"title,subtitle,organizers,shortdescription,description,startdate,enddate,fulladdress,postnr,emailaddress,www,phone,validated,keywords,starttime,endtime,place,address,status" &_
	") VALUES (" &_
	"'" & theTitle & "'" & "," &_
	"'" & replace(theRS("subtitle"),"'","''") & "'" & "," &_
	"'" & replace(theRS("organizers"),"'","''") & "'" & "," &_
	"'" & replace(theRS("shortdescription"),"'","''") & "'" & "," &_
	"'" & replace(theRS("description"),"'","''") & "'" & "," &_
	FormatDateDB(theRS("startdate")) & "," &_
	FormatDateDB(theRS("enddate")) & "," &_
	"'" & replace(theRS("fulladdress"),"'","''") & "'" & "," &_
	theRS("postnr") & "," &_
	"'" & theRS("emailaddress") & "'" & "," &_
	"'" & theRS("www") & "'" & "," &_
	"'" & theRS("phone") & "'" & "," &_
	"0" & "," &_
	"'" & theRS("keywords") & "'" & ","  &_
	"'" & theRS("starttime") & "'" & ","  &_
	"'" & theRS("endtime") & "'" & ","  &_
	"'" & theRS("place") & "'" & ","  &_
	"'" & theRS("address") & "'" & "," &_
	status &_
	")"

	theID=InsertRec(theSQL,"eical_maindata","id","title='" & theTitle & "'")
	theRS.Close()

	'2. Inserting Categories.
	theRS.Source="SELECT * FROM eical_r_cat WHERE calid='" & calid & "'"
	theRS.Open()
	theSQL=""
	WHILE NOT theRS.EOF
		theSQL=theSQL & "INSERT INTO eical_r_cat (calid,catid) VALUES (" & theID & "," & theRS("catid") & ")"
	theRS.MoveNext
	WEND
	execCommand theSQL
	theRS.Close()
	
	'3. Inserting Subjects
	theRS.Source="SELECT * FROM eical_r_sbj WHERE calid='" & calid & "'"
	theRS.Open()
	theSQL=""
	WHILE NOT theRS.EOF
		theSQL=theSQL & "INSERT INTO eical_r_sbj (calid,sbjid) VALUES (" & theID & "," & theRS("sbjid") & ");" & chr(13)
	theRS.MoveNext
	WEND
	execCommand theSQL
	

	'4. Linking to Org
	'theRS.Source="SELECT * FROM eical_r_org WHERE calid='" & calid & "'"
	'theRS.Open()
	theSQL=""
	'WHILE NOT theRS.EOF
		theSQL=theSQL & "INSERT INTO eical_r_org (orgid,calid) VALUES (" & orgid & "," & theID & ")"
	'theRS.MoveNext
	'WEND
	execCommand theSQL
	

	'5. Updating include-files
	if request("orgid")<>"" then
	'	UpdateIncludes
	end if

response.redirect "list.asp?valid=0"
%>
