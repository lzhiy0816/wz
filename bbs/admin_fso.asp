<!-- #include file="setup.asp" -->
<%
if adminpassword<>session("pass") then response.redirect "admin.asp?menu=login"
log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")


%>
<META http-equiv=Content-Type content=text/html;charset=gb2312>
<link href=images/skins/<%=Request.Cookies("skins")%>/bbs.css rel=stylesheet>
<br><center>
<p></p>

<%
select case Request("menu")

case "files"
files

case "delfiles"
set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
for each ho in request.form("files")
MyFileObject.DeleteFile""&Server.MapPath("./images/upfile/"&ho&"")&""
next
error2("�Ѿ��ɹ�ɾ����ѡ���ļ���")


case "bak"
bak

case "bakbf"
set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
MyFileObject.CopyFile ""&Server.MapPath(Request("yl"))&"",""&Server.MapPath(Request("bf"))&""
error2("���ݳɹ���")

case "bakhf"
set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
MyFileObject.CopyFile ""&Server.MapPath(Request("bf"))&"",""&Server.MapPath(Request("yl"))&""
error2("�ָ��ɹ���")

case "css"
css


case "cssok"
Set fs=Server.CreateObject("Scripting.FileSystemObject")
File=Server.MapPath("images/skins/bbs.css")
Set txtf=fs.OpenTextFile(File,2,true)
txtf.Write Request("css")
error2("���³ɹ���")

case "statroom"
statroom

end select

sub bak
%>


<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>�������ݿ�</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
    
<form method="post" action="?menu=bakbf">
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="90%">
<tr>
<td width="30%">���ݿ�·���� </td>
<td width="70%">
<input size="30" value="<%=""&datapath&""&datafile&""%>" name="yl"></td>
</tr>
<tr>
<td width="30%">���ݵ����ݿ�·����</td>
<td width="70%"><input size="30" value="<%=replace(""&datapath&""&datafile&"",".mdb","bak.mdb")%>" name="bf"></td>
</tr>
<tr>
<td width="100%" align="center" colspan="2">
<input type="submit" value=" �� �� " name="Submit1"><br></td>
</tr>
</table>
</td></tr></table></form>
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>�ָ����ݿ�</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
<form method="post" action="?menu=bakhf">
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="90%">
<tr>
<td width="30%">���ݵ����ݿ�·���� </td>
<td width="70%">
<input size="30" value="<%=replace(""&datapath&""&datafile&"",".mdb","bak.mdb")%>" name="bf"></td>
</tr>
<tr>
<td width="30%">���ݿ�·����</td>
<td width="70%"><input size="30" value="<%=""&datapath&""&datafile&""%>" name="yl"></td>
</tr>
<tr>
<td width="100%" align="center" colspan="2">
<input type="submit" value=" �� �� " name="Submit1"><br></td>
</tr>
</table></td></tr></table>
</form>

<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>ѹ�����ݿ�</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>

<form action=compact.asp>
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="90%">
<tr>

<td width="70%">


<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="90%">
<tr>

<td width="30%">���ݿ�·���� </td>
<td width="70%">
<input size="30" value="<%=""&datapath&""&datafile&""%>" name="dbpath"></td>
</tr>
<tr>
<td width="30%">���ݿ��ʽ��</td>
<td width="70%"><input type="radio" value="True" name="boolIs97" id=boolIs97><label for=boolIs97>Access 97</label>��<input type="radio" value="" name="boolIs97" checked id=boolIs97_1><label for=boolIs97_1>Access 2000��2002��2003</label></td>
</tr>
<tr>
<td width="100%" align="center" colspan="2">
<input type="submit" value=" ѹ �� " name="Submit"></td>
</tr>



</table>
</td>
</table>
</td></tr></table>
</form>

<%
end sub

sub css
%>

<br>
<form method="post" action="?menu=cssok">

<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>�༭Ĭ��CSSģ��</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
    
������������༭ģ���ļ������������CSS���벻Ҫ�Ķ���<br>
<textarea name="css" rows="20" cols="60">
<%
Set fs=Server.CreateObject("Scripting.FileSystemObject")
File=Server.MapPath("images/skins/bbs.css")
Set txtf=fs.OpenTextFile(File)
Response.Write txtf.ReadAll
%></textarea><br>
<input type="submit" value="���������������¡�����">

</td></tr></table>
</form>
<%
end sub


sub statroom
set fso=server.createobject("scripting.filesystemobject")
upfacedir=server.mappath("./images/upface")
set d=fso.getfolder(upfacedir)
upfacesize=d.size

upphotodir=server.mappath("./images/upphoto")
set d=fso.getfolder(upphotodir)
upphotosize=d.size

upfiledir=server.mappath("./images/upfile")
set d=fso.getfolder(upfiledir)
upfilesize=d.size


datadir=server.mappath("./"&datapath&"")
set d=fso.getfolder(datadir)
datasize=d.size


toldir=server.mappath(".")
set d=fso.getfolder(toldir)
tolsize=d.size

%>
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>ͳ��ռ�ÿռ�</td>
  </tr>
   <tr height=25>
    <td class=a3 width="29%">
    

�ϴ�ͷ��ռ�ÿռ�<br>




</td>
    <td class=a3 width="69%">
    

<IMG height=10 src="images/bar1.gif" width=<%=Int(upfacesize/1024/1024/5)%>> <%=Int(upfacesize/1024/1024)%> MB</td></tr>
   <tr height=25>
    <td class=a3 width="29%">
    

�ϴ���Ƭռ�ÿռ�</td>
    <td class=a3 width="69%">
    

<IMG height=10 src="images/bar1.gif" width=<%=Int(upphotosize/1024/1024/5)%>> <%=Int(upphotosize/1024/1024)%> MB</td></tr>
   <tr height=25>
    <td class=a3 width="29%">
    

�ϴ�����ռ�ÿռ�</td>
    <td class=a3 width="69%">
    

<IMG height=10 src="images/bar1.gif" width=<%=Int(upfilesize/1024/1024/5)%>> <%=Int(upfilesize/1024/1024)%> MB</td></tr>


   <tr height=25>
    <td class=a3 width="29%">

ACCESS���ݿ�ռ�ÿռ�</td>
    <td class=a3 width="69%">
<IMG height=10 src="images/bar1.gif" width=<%=Int(datasize/1024/1024/5)%>> <%=Int(datasize/1024/1024)%> MB</td></tr>


   <tr height=25>
    <td class=a3 width="29%">
    

BBSxpĿ¼�ܹ�ռ�ÿռ�</td>
    <td class=a3 width="69%">
    

<IMG height=10 src="images/bar1.gif" width=<%=Int(tolsize/1024/1024/5)%>> <%=Int(tolsize/1024/1024)%> MB</td></tr></table>


<%
end sub


sub files

thisdir=server.mappath("./images/upfile")
Set fs=Server.CreateObject("Scripting.FileSystemObject")
Set fdir=fs.GetFolder(thisdir)




pagesetup=10 '�趨ÿҳ����ʾ����
count=fdir.Files.count

If Count/pagesetup > (Count\pagesetup) then
TotalPage=(Count\pagesetup)+1
else TotalPage=(Count\pagesetup)
End If
if Request.QueryString("ToPage")<>"" then PageCount = cint(Request.QueryString("ToPage"))
if PageCount <=0 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage


%>
<script>
function CheckAll(form){for (var i=0;i<form.elements.length;i++){var e = form.elements[i];if (e.name != 'chkall')e.checked = form.chkall.checked;}}
</script>

���� <font color="FF0000"><b><%=fdir.Files.count%></b></font> ���ļ�
��<form method=post name=form action="?menu=delfiles">
<div align="center">
<table cellspacing="1" cellpadding="2" border="0" class="a2" width=97% >
<tr class=a1 ><td width=50 align="center">

<input type=checkbox name=chkall onclick=CheckAll(this.form) value="ON"></td>
	<td width=208 align=center>����</td>
  <td width=128 align=center>
��С</td><td width=177 align="center">����</td><td align="center">
�޸�ʱ��</td></tr>
<%
pagesize=20
page=request.querystring("page")
if page="" or not isnumeric(page) then
	page=1
else
	page=int(page)
end if
filenum=fdir.Files.count
pagenum=int(filenum/pagesize)
if filenum mod pagesize>0 then
	pagenum=pagenum+1
end if
if page> pagenum then
	page=1
end if
i=0



For each thing in fdir.Files

	i=i+1
	if i>(page-1)*pagesize and i<=page*pagesize then

response.write "<tr class=a3><td align=center><input type='CheckBox' value='"&thing.name&"' name=files></td><td align=center><a target=_blank href=images/upfile/"&thing.Name&">"&thing.Name&"</a></td><td align=center>" & cstr(thing.size) & " byte</td><td align=center>" & thing.type & "</td><td align=center>" & cstr(thing.datelastmodified) & "</td></tr>"


	elseif i>page*pagesize then
		exit for
	end if


Next

%></table>


</div>


<table width=97% cellpadding=2 cellspacing=2>
<tr><td>
<INPUT type=submit value=" ɾ �� " name=Submit>

</td>

	<td align="right">

<%
if page>1 then
	response.write "<a href=?menu=files&page=1>��ҳ</a>&nbsp;&nbsp;<a href=?menu=files&page="& page-1 &">��һҳ</a>&nbsp;&nbsp;"
else
	response.write "��ҳ&nbsp;&nbsp;��һҳ&nbsp;&nbsp;"
end if
if page<i/pagesize then
	response.write "<a href=?menu=files&page="& page+1 &">��һҳ</a>&nbsp;&nbsp;<a href=?menu=files&page="& pagenum &">βҳ</a>"
else
	response.write "��һҳ&nbsp;&nbsp;βҳ"
end if
%>
</td>

</tr></table>
</form>


<%
end sub



htmlend
%>