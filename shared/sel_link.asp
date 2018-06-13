<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/ecoinfo.asp" -->

<%
IF request("step")=2 THEN
	IF request("linktype")=1 THEN 'Organisation
		IF LEN(request("theid"))>0 THEN
			theSQL="SELECT id,name FROM eiorg_maindata WHERE id=" & request("theid")
		ELSE
			theSQL="SELECT id,name FROM eiorg_maindata WHERE name LIKE '%" & request("thetext") & "%' OR firstname LIKE '%" & request("thetext") & "%' OR lastname LIKE '%"  & request("thetext") & "%'"
		END IF
	ELSEIF request("linktype")=2 THEN 'Arrangement
		IF LEN(request("theid"))>0 THEN
			theSQL="SELECT id,title FROM eical_maindata WHERE id=" & request("theid")
		ELSE
			theSQL="SELECT id,title FROM eical_maindata WHERE title LIKE '%" & request("thetext") & "%' OR shortdescription LIKE '%" & request("thetext") & "%' OR description LIKE '%"  & request("thetext") & "%'"
		END IF
	ELSEIF request("linktype")=22 THEN 'Kursus
		IF LEN(request("theid"))>0 THEN
			theSQL="SELECT id,title FROM eicourse_maindata WHERE id=" & request("theid")
		ELSE
			theSQL="SELECT id,title FROM eicourse_maindata WHERE title LIKE '%" & request("thetext") & "%' OR shortdescription LIKE '%" & request("thetext") & "%' OR description LIKE '%"  & request("thetext") & "%'"
		END IF
	ELSEIF request("linktype")=3 THEN 'Publikation
		IF LEN(request("theid"))>0 THEN
			theSQL="SELECT id,title FROM eilib_maindata WHERE id=" & request("theid")
		ELSE
			theSQL="SELECT id,title FROM eilib_maindata WHERE title LIKE '%" & request("thetext") & "%' OR shortdescription LIKE '%" & request("thetext") & "%' OR description LIKE '%"  & request("thetext") & "%'"
		END IF
	ELSEIF request("linktype")=4 THEN 'Opslag
		IF LEN(request("theid"))>0 THEN
			theSQL="SELECT id,title FROM eiopsl_maindata WHERE id=" & request("theid")
		ELSE
			theSQL="SELECT id,title FROM eiopsl_maindata WHERE title LIKE '%" & request("thetext") & "%' OR kort_beskrivelse LIKE '%" & request("thetext") & "%' OR lang_beskrivelse LIKE '%"  & request("thetext") & "%'"
		END IF
	ELSEIF request("linktype")=5 THEN 'Nyhed
		IF LEN(request("theid"))>0 THEN
			theSQL="SELECT id,name FROM eical_maindata WHERE id=" & request("theid")
		ELSE
			theSQL="SELECT id,name FROM eical_maindata WHERE title LIKE '%" & request("thetext") & "%' OR shortdescription LIKE '%" & request("thetext") & "%' OR description LIKE '%"  & request("thetext") & "%'"
		END IF

	ELSEIF request("linktype")=6 THEN 'Kategori i De Grønne Sider
		theSQL="SELECT id,name FROM eiorg_cat_maindata ORDER BY 2"

	ELSEIF request("linktype")=7 THEN 'Kategori i Øko-kalender
		theSQL="SELECT id,name FROM eical_cat_maindata ORDER BY 2"

	ELSEIF request("linktype")=8 THEN 'Kategori i Det grønne Bibliotek
		theSQL="SELECT id,name FROM eilib_cat_maindata ORDER BY 2"

	ELSEIF request("linktype")=9 THEN 'Kategori i Opslagtavlen
		theSQL="SELECT id,name FROM eiopsl_cat_maindata ORDER BY 2"

	ELSEIF request("linktype")=10 THEN 'Kategori i Nyheder
		theSQL="SELECT id,name FROM einews_cat_maindata ORDER BY 2"

	ELSEIF request("linktype")>10 THEN 'Emneord
		theSQL="SELECT id,name FROM eisbj_maindata ORDER BY 2"

	END IF
	
	IF theSQL<>"" THEN theOptions=MakeOptions(theSQL)
END IF
%>

<html><!-- #BeginTemplate "/Templates/noheader.dwt" --><!-- DW6 -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Ecoinfo</title>
<script language="JavaScript">
function setLink(theform) {
	if(theform.linktype.value==1){theform.linktext.value='/dgs/detail.asp?id=' + theform.theid.options[theform.theid.selectedIndex].value;theform.linkdescr.value='Organisationen ' + theform.theid.options[theform.theid.selectedIndex].text + '.'}
	else if(theform.linktype.value==2){theform.linktext.value='/ok/detail.asp?id=' + theform.theid.options[theform.theid.selectedIndex].value;theform.linkdescr.value='Arrangementet ' + theform.theid.options[theform.theid.selectedIndex].text + '.'}
	else if(theform.linktype.value==22){theform.linktext.value='/kurser/detail.asp?id=' + theform.theid.options[theform.theid.selectedIndex].value;theform.linkdescr.value='Kurset ' + theform.theid.options[theform.theid.selectedIndex].text + '.'}
	else if(theform.linktype.value==3){theform.linktext.value='/dgb/detail.asp?id=' + theform.theid.options[theform.theid.selectedIndex].value;theform.linkdescr.value='Publikationen ' + theform.theid.options[theform.theid.selectedIndex].text + '.'}
	else if(theform.linktype.value==4){theform.linktext.value='/opsl/detail.asp?id=' + theform.theid.options[theform.theid.selectedIndex].value;theform.linkdescr.value='Opslaget ' + theform.theid.options[theform.theid.selectedIndex].text + '.'}
	else if(theform.linktype.value==5){theform.linktext.value='/news/detail.asp?id=' + theform.theid.options[theform.theid.selectedIndex].value;theform.linkdescr.value='Nyheden ' + theform.theid.options[theform.theid.selectedIndex].text + '.'}
	else if(theform.linktype.value==6){theform.linktext.value='/dgs/list.asp?dgscat=' + theform.theid.options[theform.theid.selectedIndex].value + '&dgscatname=' + theform.theid.options[theform.theid.selectedIndex].text;theform.linkdescr.value='Kategorien ' + theform.theid.options[theform.theid.selectedIndex].text + ' i De Grønne Sider.'}
	else if(theform.linktype.value==7){theform.linktext.value='/ok/list.asp?okcat=' + theform.theid.options[theform.theid.selectedIndex].value + '&okcatname=' + theform.theid.options[theform.theid.selectedIndex].text;theform.linkdescr.value='Kategorien ' + theform.theid.options[theform.theid.selectedIndex].text + ' i Øko-kalenderen.'}
	else if(theform.linktype.value==8){theform.linktext.value='/dgb/list.asp?dgbcat=' + theform.theid.options[theform.theid.selectedIndex].value + '&dgbcatname=' + theform.theid.options[theform.theid.selectedIndex].text;theform.linkdescr.value='Kategorien ' + theform.theid.options[theform.theid.selectedIndex].text + ' i Det Grønne Bibliotek.'}
	else if(theform.linktype.value==9){theform.linktext.value='/opsl/list.asp?opslcat=' + theform.theid.options[theform.theid.selectedIndex].value + '&opslcatname=' + theform.theid.options[theform.theid.selectedIndex].text;theform.linkdescr.value='Kategorien ' + theform.theid.options[theform.theid.selectedIndex].text + ' i Opslagstavlen.'}
	else if(theform.linktype.value==10){theform.linktext.value='/news/list.asp?newscat=' + theform.theid.options[theform.theid.selectedIndex].value + '&newscatname=' + theform.theid.options[theform.theid.selectedIndex].text;theform.linkdescr.value='Kategorien ' + theform.theid.options[theform.theid.selectedIndex].text + ' i Nyheder.'}
	else if(theform.linktype.value==11){theform.linktext.value='/dgs/list.asp?dgssbj=' + theform.theid.options[theform.theid.selectedIndex].value + '&dgssbjname=' + theform.theid.options[theform.theid.selectedIndex].text;theform.linkdescr.value='Emneordet ' + theform.theid.options[theform.theid.selectedIndex].text + ' i De Grønne Sider.'}
	else if(theform.linktype.value==12){theform.linktext.value='/ok/list.asp?oksbj=' + theform.theid.options[theform.theid.selectedIndex].value + '&oksbjname=' + theform.theid.options[theform.theid.selectedIndex].text;theform.linkdescr.value='Emneordet ' + theform.theid.options[theform.theid.selectedIndex].text + ' i Øko-kalenderen.'}
	else if(theform.linktype.value==13){theform.linktext.value='/dgb/list.asp?dgbsbj=' + theform.theid.options[theform.theid.selectedIndex].value + '&dgbsbjname=' + theform.theid.options[theform.theid.selectedIndex].text;theform.linkdescr.value='Emneordet ' + theform.theid.options[theform.theid.selectedIndex].text + ' i Det Grønne Bibliotek.'}
}
</script>
<!-- #EndEditable --> 
<script src="/shared/common.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="7" marginwidth="0" marginheight="7">
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" name="Pagetable">
<tr> 
<td background="/shared/graphics/layout/dots_vert.gif" width="1" valign="top"><img src="/shared/graphics/layout/cover_dots.gif" width="1" height="18"></td>
<td width="100%" valign="top"> 
<!-- START HEADER -->
<table width="100%" border="0" cellspacing="0" cellpadding="0" name="Header">
<tr> 
<td valign="top" rowspan="3" width="180" heigth="33"><img src="/shared/graphics/logo.gif" width="180" height="33"></td>
<td height="17">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td align="left"><img src="/shared/graphics/logo2.gif" width="21" height="16"></td>
</tr>
</table>
</td>
</tr>
<tr>
<td valign="top" width="99%" background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr>
<td height="16"><br></td>
</tr>
</table>
<!-- END HEADER -->

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td width="100%" valign="top" name="maincontent">
<table border="0" cellspacing="0" cellpadding="10" width="100%" name="Contentarea">
<tr><td valign="top">
<!-- #BeginEditable "maincontent" -->
<form name="form1" method="post" action="">
<%IF LEN(request("step"))=0 OR request("step")=1 THEN%>
<span class="formLabeltext">1.V&aelig;lg hvilken datatype der skal linkes til</span> 
<br>
<br>
<select name="linktype" class="formInputobjectLong" onChange="this.form.linktypename.value=this.form.linktype.options[this.form.linktype.selectedIndex].text;this.form.theid.focus();">
<option value="0">&lt;V&aelig;lg&gt;</option>
<option value="1">Organisation</option>
<option value="2">Arrangement</option>
<option value="22">Kursus</option>
<option value="3">Publikation</option>
<option value="4">Opslag</option>
<option value="5">Nyhed</option>
<option value="6">Kategori i De grønne sider</option>
<option value="7">Kategori i Øko-kalenderen</option>
<option value="8">Kategori i Det Grønne Bibliotek</option>
<option value="9">Kategori i Opslagstavlen</option>
<option value="10">Kategori i Nyheder</option>
<option value="11">Emneord i De grønne sider</option>
<option value="12">Emneord i Øko-kalenderen</option>
<option value="13">Emneord i Det Grønne Bibliotek</option>
</select>
<br>
<input type="hidden" name="linktypename">
<br>
<span class="formLabeltext">2. Indtast idnr. på det ønskede kort ELLER s&oslash;getekst</span> <br>
<br>
Med id nummer:<br>
<input type="text" name="theid" class="formInputobject">
<br>
Indeholder teksten:<br>
<input type="text" name="thetext" class="formInputobject">
<br>
<br>
<input type="button" name="Button" value="N&aelig;ste &gt;&gt;" class="formSubmitbutton" onClick="this.form.step.value='2';this.form.submit();">
<%ELSEIF request("step")=2 THEN%>
<span class="formLabeltext">3. V&aelig;lg &oslash;nsket <%=request("linktypename")%> i listen</span> 
<br>
<br>
<select name="theid" class="formInputobjectLong" onChange="setLink(this.form)">
<%=theOptions%>
</select>
<br>
<input type="hidden" name="linktext">
<input type="hidden" name="linkdescr">

<br>
<input type="button" name="Button" value="&lt;&lt;Tilbage" onClick="this.form.step.value='1';this.form.submit();" class="formbutton">
<input type="button" name="Button" value="OK" class="formSubmitbutton" onClick="alert(this.form.linkdescr.value);opener.document.forms[0].link.value=this.form.linktext.value;opener.document.forms[0].linkdescr.value=this.form.linkdescr.value;window.close();">
<input type="hidden" name="linktype" value="<%=request("linktype")%>">
<input type="hidden" name="linktypename" value="<%=request("linktypename")%>">
<%END IF%>
<input type="hidden" name="step">
</form>
<br>
<!-- #EndEditable -->
</td></tr>
</table>
</td>
</tr>
</table>
</td>
<td background="/shared/graphics/layout/dots_vert.gif" width="1" valign="top"><img src="/shared/graphics/layout/cover_dots.gif" width="1" height="18"></td>
</tr>
<tr height="1"> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1" colspan="3"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
</table>
</body>
<!-- #EndTemplate --></html>
<!--#include virtual="/shared/log.asp"-->
<%
FUNCTION MakeOptions(theSQL)
	DIM theRS, Format_Sel, fso, ts,theData,rowCount, colCount,i,ii, thisRow, theFile
	modelText="<option value=""#0#"">#1#</option>"
	SET theRS= Server.CreateObject("ADODB.Recordset")
	theRS.ActiveConnection = MM_ecoinfo_STRING
	theRS.Source = theSQL
	theRS.Open()
	theData=theRS.GetRows
	theRS.Close()
	theRS=""
	colCount=uBound(theData,2)
	rowCount=uBound(theData,1)
	theFile="<option value=""0"" selected>&lt;V&aelig;lg&gt;</option>"
	FOR i=0 TO colCount
		thisRow=ModelText
		FOR ii=0 TO rowCount
			thisRow=replace(thisRow,"#" & ii & "#",trim(theData(ii,i)) & "")
		NEXT
		theFile=theFile & thisrow
	NEXT
	MakeOptions=theFile
END FUNCTION

%>