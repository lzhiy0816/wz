<!-- #include file="setup.asp" -->
<!--#include FILE="inc/upfile.asp"-->

<%
if Request.Cookies("username")=empty then
error2("您还未登录社区")
end if

if Request("menu")="up" then

ftime=""&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&""

On Error Resume Next

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set upl = Server.CreateObject("SoftArtisans.FileUp")
If 0 = Err Then

select case ""&upl.ContentType&""
case "application/octet-stream"
error2("未知文件格式")
case "image/gif"
types="gif"
case "image/pjpeg"
types="jpg"
case "image/bmp"
types="bmp"
case "image/x-png"
types="png"
case "text/plain"
types="txt"
case "audio/mid"
types="mid"
case "application/msword"
types="doc"
case "application/x-zip-compressed"
types="zip"
case "application/x-shockwave-flash"
types="swf"
end select

filename="images/upfile/"&ftime&"."&types&""

if upl.TotalBytes > int(MaxFile) then error2("文件大小不得超过 "&int(MaxFile/1024)&" KB\n当前的文件大小为 "&int(upl.TotalBytes/1024)&" KB")


if types="gif" or types="jpg" or types="png" or types="bmp" then
img="[img]"&filename&"[/img]"

elseif types="swf" then
img="[flash]"&filename&"[/flash]"

elseif types="txt" or types="mid" or types="doc" or types="zip" then
img="[img]images/affix.gif[/img][url="&filename&"]相关附件[/url]"
else
error2("对不起，本服务器只支持GIF、JPG、BMP、PNG、TXT、MID、DOC、ZIP、SWF格式的文件\n不支持 "&upl.ContentType&" 格式的文件")
end if


upl.SaveAs Server.mappath(""&filename&"")

sql="insert into upfile(username,filename,types,len) values ('"&Request.Cookies("username")&"','"&ftime&"','"&types&"','"&upl.TotalBytes&"')"
conn.Execute(SQL)

set upl=nothing

else
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

set FileUP=new Upload_file
FileUP.GetDate(-1)
formPath="images/upfile/"
set file=FileUP.file("file")
filename=formPath&ftime&"."&file.FileExt

if file.filesize > int(MaxFile) then error2("文件大小不得超过 "&int(MaxFile/1024)&" KB\n当前的文件大小为 "&int(file.filesize/1024)&" KB")


types=LCase(file.FileExt)

if types="gif" or types="jpg" or types="bmp" or types="png" then
img="[img]"&filename&"[/img]"

elseif types="swf" then
img="[flash]"&filename&"[/flash]"

elseif types="txt" or types="mid" or types="doc" or types="zip" or types="rar" or types="torrent" then
img="[img]images/affix.gif[/img][url="&filename&"]相关附件[/url]"
else
error2("对不起，本服务器只支持GIF、JPG、BMP、PNG、TXT、MID、DOC、ZIP、SWF、RAR、torrent格式的文件\n不支持 "&types&" 格式的文件")
end if

file.SaveToFile Server.mappath(filename)

sql="insert into upfile(username,filename,types,len) values ('"&Request.Cookies("username")&"','"&ftime&"','"&file.FileExt&"','"&file.filesize&"')"
conn.Execute(SQL)

set FileUP=nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If

Response.Cookies("filename")=""&Request.Cookies("filename")&"|"&ftime&""
response.write "<script>document.oncontextmenu = new Function('return false')</script><link href=images/skins/"&Request.Cookies("skins")&"/bbs.css rel=stylesheet><body topmargin=0  rightmargin=0  leftmargin=0><SCRIPT>parent.form.content.value+='\n"&img&"'</SCRIPT>"
response.write "<table cellpadding=0 cellspacing=0 width=100%  height=100% class=a4><tr><td><a target=_blank href="&filename&">"&cluburl&"/"&filename&"</a></td></tr></table>"
responseend

else

%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<script>if(top==self)document.location='';</script>
<meta http-equiv=Content-Type content=text/html;charset=gb2312>
<link href=images/skins/<%=Request.Cookies("skins")%>/bbs.css rel=stylesheet>
<body topmargin=0  rightmargin=0  leftmargin=0>
<form enctype=multipart/form-data method=post action=upfile.asp?menu=up>
<table cellpadding=0 cellspacing=0 width=100% class=a4>
<tr><td>　</td>
<td><input type=file style=FONT-SIZE:9pt name=file size=60> <input style=FONT-SIZE:9pt type="submit" value=" 上 传 " name=Submit></td></tr></table>
<%

end if
%>