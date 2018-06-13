<%IF LEN(request.cookies("eiuserid"))>0 AND rsPageData.Fields.Item("validated").value=0 THEN%>
<input type="checkbox" name="validated" value="1">
<span class="formLabeltext">Godkend</span>
<%END IF%>
