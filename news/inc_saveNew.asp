<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/shared/common.asp" -->
<%
SUB DoCreate


Dim validate,bu,econet,newwin,img

	if request("img_id")<>"" then
	img=CInt(request("img_id"))
	else
	img=0
	end if
	if request("img_id2")<>"" then
	img2=CInt(request("img_id2"))
	else
	img2=0
	end if
	if request("validated")<>"" then
	validate=1 	'valideret
	else 
	validate=0
	end if
	if request("bu")<>"" then
	bu=1 ' vises på bu
	else
	bu=0
	end if
	if request("econet")<>"" then
	econet=1 'vises på econet
	else
	econet=0
	end if
	if request("vindue")=0 then
	newwin=0
	else
	newwin=1
	end if
	if request("miljoinfo")<>"" then
	miljoinfo=1 'Miljø Info tekst vises
	else
	miljoinfo=0
	end if
	'idag=Date
	

	
	cat=CInt(request("cat"))
	
	'1. Inserting news
	theSQL="INSERT INTO bu_Artikel(" &_
	"Header,Approved,ShortDescription,FileName,FileName2,FileName3,link4,Content,img_id,img_text,ArtikelCategory_ID,linktext,linktext2,linktext3,linktext4,maker,Tema_ID,img_id2,img_text2,imagefilename1,imagefilename2,miljoinfo" &_
	") VALUES (" &_
	"'" & toHTML(request("title")) & "'" & "," & validate & "," &_
	"'" & toHTML(request("shortdescription")) & "'" & "," &_
	"'" & request("link") & "'" & "," &_
	"'" & request("link2") & "'" & "," &_
	"'" & request("link3") & "'" & "," &_
	"'" & request("link4") & "'" & "," &_
	"'" & toHTML(request("EditorValue")) & "'" & "," &_
	img & "," &_
	"'" & request("img_text") & "'," &_
	cat & "," &_
	"'" & request("linktext") & "'" & "," &_
	"'" & request("linktext2") & "'" & "," &_
	"'" & request("linktext3") & "'" & "," &_
	"'" & request("linktext4") & "'" & "," &_
	"'" & Session("username") & "'" & "," &_
	request("tema") & "," &_
	img2 &_
	",'" & request("img_text2") & "','" &_
	request("imagefilename1") & "','" &_
	request("imagefilename2") & "'," &_
	miljoinfo & ")"	
	'response.write theSQL
	'response.end
	theID=InsertRec(theSQL,"bu_Artikel","Artikel_ID","")
	
	'2. Inserting Categories.
	theSQL=""
	
	theSQL= "INSERT INTO bu_Artikel_site (artikel_id,newwin,busite,econetsite) VALUES (" & theID & "," & newwin & "," & bu & "," & econet & ");" & chr(13)

	execCommand theSQL
	
END SUB
%>