
<%
dim strPathInfo, strPhysicalPath,bytes
Dim objFSO, objFile, objFileItem, objFolder, objFolderContents, Folder, Folders
	Set objFSO = CreateObject("Scripting.FileSystemObject")
Function foldersize(thefolder)
	thebytes=0
	strPhysicalPath = Server.MapPath(thefolder)
	set objFolder = objFSO.GetFolder(strPhysicalPath)
	thebytes=objFolder.size
	
	mb=Int(thebytes/1048576)
	m_b= thebytes -(1048576*mb)
	kb=Int(m_b/1024)
	b=m_b -(1024*kb)
	
	fs=mb & "MB " & kb & "kB " & b & "bytes"
	
	foldersize=fs
end function
Function folderbytes(thefolder)
	thebytes=0
	strPhysicalPath = Server.MapPath(thefolder)
	set objFolder = objFSO.GetFolder(strPhysicalPath)
	folderbytes=objFolder.size
	
end function
Function megabytes(thefolder)
	thebytes=0
	strPhysicalPath = Server.MapPath(thefolder)
	set objFolder = objFSO.GetFolder(strPhysicalPath)
	thebytes=objFolder.size
	megabytes=Int(thebytes/1048576)
end function
Function toMega(b)
	thebytes=b
	
	mb=Int(thebytes/1048576)
	m_b= thebytes -(1048576*mb)
	kb=Int(m_b/1024)
	b=m_b -(1024*kb)
	
	fs=mb & "MB " & kb & "kB " & b & "bytes"
	
	toMega=fs
end function
%>
