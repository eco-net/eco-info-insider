<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%

Dim w2col,w3col,wTema,wRightcol,wThumbnail,width,height,path,theID,filename
w2col=540
w3col=360
wTema=258
wRightcol=160
wThumbnail=75
Dim twocol,threecol,tema,rightcol,thumbnail
twocol=0
threecol=0
tema=0
rightcol=0
thumbnail=0
Set Upload = Server.CreateObject("Persits.Upload.1")

' Upload files
Upload.OverwriteFiles = False ' Generate unique names
'Upload.SetMaxSize 100000,true ' Truncate files above 1MB
Count=Upload.SaveVirtual("upload/") 
Set File = Upload.Files("file")
filename=File.ExtractFileName

if File.ImageType ="JPG" then

width = File.ImageWidth
height = File.ImageHeight
path=Server.MapPath("upload/") & "/" & filename
response.write "Der er oprettet disse billeder i billedbasen: <br>"
Set Image = Server.CreateObject("AspImage.Image")
	if width > w3col then
		if width >= w2col then 
			makeImage w2col,"2col" ' make size = w2col
		else 
			makeImage width,"2col" 'make size max w2col 
		end if
		twocol=1
	end if
	if width > wTema then 
		if width >= w3col then 
			makeImage w3col,"3col" ' make size = w3col
		else
			makeImage width,"3col" ' make size max w3col
		end if
		threecol=1
	end if
	if width > wRightcol then 
		if width >= wTema then 
			makeImage wTema,"tema" ' make size = wTema
		else
			makeImage width,"tema" ' make size max wTema
		end if
		tema=1
	end if
	if width > wThumbnail then 
		if width >= wRightcol then 
			makeImage wRightcol,"rightcol" ' make size = wRightcol
		else
			makeImage width,"rightcol" ' make size max wRightcol
		end if
		rightcol=1
	end if
	if width >=wThumbnail then 
			makeImage wThumbnail,"thumbnail" ' make size = wThumbnail
		thumbnail=1
	end if
	writeDB
File.Delete  
Image.LoadImage(Server.MapPath("thumbnail/") & "/" & filename)
Image.Filename=Server.MapPath("upload/")  & "/" & filename
Image.SaveImage
else 'not .jpg
response.write("Virker kun på .jpg filer<br>")
end if '.jpg
Function makeImage(theSize,sizeName)
	Image.LoadImage(path)
	x= CInt(theSize/width * height)
	if sizeName="thumbnail" then
		if x > 75 then
			x= 75
			theSize=CInt(75/height*width)
		end if
	end if
	Image.Resize theSize,x
	Image.Filename=Server.MapPath(sizeName & "/")  & "/" & filename
	Image.SaveImage

response.write "<img src='" & sizeName  & "/" & filename & "'><br>"
response.write sizeName  & " - " & filename & "<br>"
End Function

Function writeDB()
		Set rs = Server.CreateObject("adodb.recordset")
		rs.Open "images", MM_ecoinfo_STRING, 2, 3
		rs.AddNew
		rs("filename") = filename
		rs("subtext") = Upload.Form("subtext")
		rs("source") = Upload.Form("source")
		rs("cat_id") = CInt(Upload.Form("cat"))
		rs("twocol")= twocol
		rs("threecol")= threecol
		rs("tema")= tema
		rs("rightcol")= rightcol
		rs("thumbnail")= thumbnail
		rs("branding")= 0
		rs("splash")= 0
		rs.Update
		rs.Close
		rs.Open "images", MM_ecoinfo_STRING, 2, 3
		rs.MoveLast
		theID=rs("id")
		rs.Close
		response.write "BilledID: " & theID & "<br>"
end Function



%>