<!-- #include file="setup.asp" -->
<%
if Request.Cookies("username")=empty then error("<li>����δ<a href=login.asp>��¼</a>����")

id=int(Request("id"))

sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn

if Request.Cookies("userpass")<>rs("userpass") then error("<li>�������")

faction=rs("faction")
experience=rs("experience")
money=rs("money")
rs.close


if Request("menu")="factionadd" then
if faction<>"" then error("<li>���Ѿ����� "&faction&" �ˣ�<li>�����˳� "&faction&" ���ܼ����°��ɣ�")

factionname=Conn.Execute("Select factionname From faction where id="&id&"")(0)
conn.execute("update [user] set faction='"&factionname&"' where username='"&Request.Cookies("username")&"'")

message=message&"<li>����ɹ�<li><a href=faction.asp>������������</a><li><a href=main.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="factionout" then
if faction=empty then error("<li>��Ŀǰû�м����κΰ��ɣ�")
conn.execute("update [user] set faction='' where username='"&Request.Cookies("username")&"'")
message=message&"<li>�˳����ɳɹ�<li><a href=faction.asp>������������</a><li><a href=main.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="factiondel" then

sql="select * from faction where id="&id&""
rs.Open sql,Conn,1,3

if rs("buildman")<>""&Request.Cookies("username")&"" then error("<li>�����Ǹð�İ����޷���ɢ�ð�")

conn.execute("update [user] set faction='' where faction='"&rs("factionname")&"'")
rs.close
conn.execute("delete from [faction] where id="&id&"")

message=message&"<li>��ɢ���ɳɹ�<li><a href=faction.asp>������������</a><li><a href=main.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="look" then
sql="select * from faction where factionname='"&server.htmlencode(Request("faction"))&"'"
rs.Open sql,Conn
%>
<HTML><META http-equiv=Content-Type content="text/html; charset=gb2312">
<link href=images/skins/<%=Request.Cookies("skins")%>/bbs.css rel=stylesheet>
<title><%=rs("factionname")%></title>
<body onload=resizeTo(650,400)>
<table height=269 cellspacing=0 cellpadding=0 width=576 border=0 align="center">
<tbody>
<tr>
<td width="19"><img height=51 alt=""
src="images/plus/index_01.gif" width=19></td>
<td width=557 background=images/plus/index_02.gif
height=51><img height=26 src="images/plus/niu01.gif"
width=26><img height=26
src="images/plus/hyuan2.gif" width=95 border=0></td>
<td width="10"><img height=51 src="images/plus/index_03.gif" width=30></td>
</tr>
<tr>
<td width=19 background=images/plus/index_04.gif
height=199></td>
<td width=557 height=199 bgcolor="FFFFFF">
<table width="82%" border="0" align="center" cellspacing="1" cellpadding="2" bgcolor="CCCCCC" height="150">
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>���ɼ��:</b></font></div>
</td>
<td width="82%"><%=rs("factionname")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>��������:</b></font></div>
</td>
<td width="82%"><%=rs("allname")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>������ּ:</b></font></div>
</td>
<td width="82%"><%=rs("tenet")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>����ʱ��:</b></font></div>
</td>
<td width="82%"><%=rs("addtime")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>��������:</b></font></div>
</td>
<td width="82%"><%=rs("buildman")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>���л�Ա:</b></font></div>
</td>
<td width="82%">
<%
sql="select username from [user] where faction='"&rs("factionname")&"'"
rs.close
rs.Open sql,Conn,1
count=rs.recordcount
%><%=count%>��</td>
</tr>

<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>��Ա����:</b></font></div>
</td>
<td width="82%">
<%
Do While Not RS.EOF 
response.write ""&rs("username")&"&nbsp;"
RS.MoveNext
loop
RS.Close
%>
</td>
</tr>

</table>
</td>
<td height=199 width="10">
<img height=250 alt=""
src="images/plus/index_06.gif" width=30></td>
</tr>
<tr>
<td width="19"><img height=31 src="images/plus/index_07.gif" width=19></td>
<td width=557 background=images/plus/index_08.gif
height=31></td>
<td width="10"><img height=31 src="images/plus/index_09.gif"
width=30></td>
</tr>
</tbody>
</table>
<%

responseend
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="addok" then

factionname=server.htmlencode(Request("factionname"))
allname=server.htmlencode(Request("allname"))
tenet=server.htmlencode(Request("tenet"))


if faction<>empty then
message=message&"<li>���Ѿ������˰��ɣ�"
end if

if experience< 5000 then
message=message&"<li>���ľ���ֵС�� 5000 ��"
end if
if money< 5000 then
message=message&"<li>���Ľ������ 5000 ��"
end if

if factionname="" then
message=message&"<li>���ɼ��û����д"
end if
if allname="" then
message=message&"<li>��������û����д"
end if
if tenet="" then
message=message&"<li>������ּû����д"
end if

if message<>"" then
error(""&message&"")
end if



rs.Open "faction",conn,1,3
rs.addnew
rs("factionname")=factionname
rs("allname")=allname
rs("tenet")=tenet
rs("buildman")=Request.Cookies("username")
rs.update
rs.close

conn.execute("update [user] set faction='"&factionname&"',[money]=[money]-5000 where username='"&Request.Cookies("username")&"'")


message=message&"<li>�������ɳɹ�<li><a href=faction.asp>������������</a><li><a href=main.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="xiuok" then

factionname=server.htmlencode(Request("factionname"))
allname=server.htmlencode(Request("allname"))
tenet=server.htmlencode(Request("tenet"))


if factionname="" then
message=message&"<li>���ɼ��û����д"
end if
if allname="" then
message=message&"<li>��������û����д"
end if
if tenet="" then
message=message&"<li>������ּû����д"
end if

if message<>"" then
error(""&message&"")
end if

sql="select * from faction where id="&id&""
rs.Open sql,Conn,1,3
if rs("buildman")<>""&Request.Cookies("username")&"" then
error("<li>�����Ǹð�İ����޷��޸���Ϣ")
end if
oldfactionname=rs("factionname")
rs("factionname")=factionname
rs("allname")=allname
rs("tenet")=tenet
rs.update
rs.close
conn.execute("update [user] set faction='"&factionname&"' where faction='"&oldfactionname&"'")
message=message&"<li>�޸İ��ɳɹ�<li><a href=faction.asp>������������</a><li><a href=main.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
end if


top
%>

<title>��������</title>
<script>function checkclick(msg){if(confirm(msg)){event.returnValue=true;}else{event.returnValue=false;}}</script>
<br>
<center>


<%

if Request("menu")="add" then



%>


<table cellSpacing=0 cellPadding=0 border=0>
<form method=post name=form action=faction.asp?menu=addok>
  <tr><td width="100%">
<img border="0" src="images/plus/top.gif"></td></tr>
</table>
<TABLE cellSpacing=0 cellPadding=0 border=0 width="751">
<TBODY>
<TR>
<TD vAlign=top width=192>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
<TBODY>
<TR>
<TD><IMG height=115 src="images/plus/left01.gif"
width=192></TD></TR>
<TR>
<TD><IMG height=113 src="images/plus/left02.gif"
width=192></TD></TR>
<TR>
<TD><IMG height=121 src="images/plus/left03.gif"
width=192></TD></TR></TBODY></TABLE></TD>
<TD vAlign=top width=559>
<TABLE cellSpacing=0 cellPadding=0 width="19%" border=0>
<TBODY>
<TR>
<TD width="40%"><IMG height=115 src="images/plus/right01.gif"></TD>
</TR></TBODY></TABLE>
<br>
<table height=74 cellspacing=0 cellpadding=0 width=320 border=0 bgcolor="FFFFFF">
<tbody>
<tr>
<td><img height=51 src="images/plus/index_01.gif" width=19></td>
<td width=271 background=images/plus/index_02.gif height=51><img height=26 src="images/plus/niu03.gif" width=26><img height=26 src="images/plus/hyuan.gif" width=95 border=0></td>
<td><img height=51 src="images/plus/index_03.gif" width=30></td>
</tr>
<tr>
<td width=19 background=images/plus/index_04.gif
height=199></td>
<td width=351 height=199>
<center>
<table cellspacing=0 cellpadding=0 width=442 border=0 align="center" bgcolor="FFFFFF">
<tbody>
<tr>
<td width=187>
<div align="right"><b><font color="0033CC">�������ɼ�ƣ� </font></b></div>
</td>
<td width=339>
<input maxlength=7 name=factionname size="10"> ���7���ַ�</td>
</tr>
<tr>
<td width=187>
<div align="right"><b><font color="0033CC">�����������ƣ� </font></b></div>
</td>
<td width=339>
<input size=30 name=allname>
</td>
</tr>
<tr>
<td width=187 height=15>
<div align="right"><b><font color="0033CC">����������ּ�� </font></b></div>
</td>
<td width=339 height=15>
<input size=40 name=tenet>
</td>
</tr>
<tr>
<td width=526 colspan=2 height=15>
<div align=center>
<input type=submit value=" �� �� " name=Submit3>
<input type=reset value=" �� �� " name=Submit22>
</div>
</td>
</tr>
</tbody>
</table>
</center>
<ol>
�������ɵı�Ҫ������
<li>���ľ���ֵ���� 5000 ����
<li>��Ҫ�۳������� 5000 �����Ϊ���ɻ��� </li>
</ol>
</td>
<td height=199 background="images/plus/index_06.gif">��</td>
</tr>
<tr>
<td><img height=31 src="images/plus/index_07.gif" width=19></td>
<td width=271 background=images/plus/index_08.gif height=31></td>
<td><img height=31 src="images/plus/index_09.gif" width=30></td>
</tr>
</tbody>
</table>
</TD>
</TR></TBODY></form></TABLE>





<%
elseif Request("menu")="xiu" then
sql="select * from faction where id="&id&""
rs.Open sql,Conn

%>

<form method=post action=faction.asp?menu=xiuok&id=<%=rs("id")%>>
<table cellpadding="2" cellspacing="1" width="70%" border="0" class=a2>

<tr>
<td colspan="2" height="25" align="center" class=a1>���������趨</td>
</tr>


<tr class=a3>
<td>�������ɼ�ƣ� </td>
<td>
<input size="20" maxlength=7 name="factionname" value="<%=rs("factionname")%>"> 
���7���ַ�</td>
</tr>
<tr class=a3>
<td>�����������ƣ� </td>
<td><input size="30" name="allname" value="<%=rs("allname")%>"> </td>
</tr>
<tr class=a3>
<td>����������ּ�� </td>
<td><input size="60" name="tenet" value="<%=rs("tenet")%>"> </td>
</tr>
<tr class=a3>
<td colSpan="2">
<div align="center">
<input type="submit" value=" �� �� " name="Submit">
<input type="reset" value=" �� �� " name="Submit2">
</div>
</td>
</tr>
</table>
</form>
<%
else
%>
<p>&nbsp;<a href="?menu=add"><img src="images/plus/niu05.gif" width="26" height="26" border="0"><img src="images/plus/cj.gif" width="95" height="26" border="0"></a>��������<a href="?menu=factionout"><img src="images/plus/niu05.gif" width="26" height="26" border="0"><img src="images/plus/tc.gif" width="95" height="26" border="0"></a>
<br>
</p>

<SCRIPT>valigntop()</SCRIPT>
<table border="0" cellpadding="5" cellspacing="1" class=a2 width=97%>
<tr>
<td width="15%" align="center" height="25" class=a1>�ɱ�</td>
<td width="40%" align="center" height="25" class=a1>��ּ</td>
<td width="15%" align="center" height="25" class=a1>��ʼ��</td>
<td width="10%" align="center" height="25" class=a1>����</td>
<td width="20%" align="center" height="25" class=a1>��������</td>
</tr>


<%


sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn
faction=rs("faction")
rs.close

sql="select * from faction order by addtime Desc"
rs.Open sql,Conn,1


pagesetup=20 '�趨ÿҳ����ʾ����
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '��ҳ��
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '��ת��ָ��ҳ��

i=0
Do While Not RS.EOF and i<pagesetup
i=i+1





%>

<tr>
<td width="10%" align="center" height="25" bgcolor="FFFFFF"> <a style=cursor:hand onclick="javascript:open('faction.asp?menu=look&faction=<%=rs("factionname")%>','','toolbar=no')"><%=rs("factionname")%></a>��</td>
<td width="50%" align="center" height="25" bgcolor="FFFFFF"><%=rs("tenet")%>��</td>
<td width="10%" align="center" height="25" bgcolor="FFFFFF"><a href=Profile.asp?username=<%=rs("buildman")%>><%=rs("buildman")%></a>��</td>
<td width="10%" align="center" height="25" bgcolor="FFFFFF"><%
if faction=rs("factionname") then
response.write "<a onclick=checkclick('��ȷ��Ҫ�˳��ð��ɣ�') href=?menu=factionout><img alt='�˳��˰�' src=images/plus/factionout.gif border=0></a>"
else
response.write "<a onclick=checkclick('��ȷ��Ҫ����ð��ɣ�') href=?menu=factionadd&id="&rs("id")&"><img alt='����˰�' src=images/plus/factionadd.gif border=0></a>"
end if
%>��</td>
<td width="20%" align="center" height="25" bgcolor="FFFFFF"><a href="?menu=xiu&id=<%=rs("id")%>"><img height="19" alt="�޸�����" src="images/plus/factionxiu.gif" border="0"></a><a onclick=checkclick('��ȷ��Ҫ��ɢ�ð��ɣ�') href="?menu=factiondel&id=<%=rs("id")%>"><img height="19" alt="��ɢ�˰�" src="images/plus/factiondel.gif" border="0"></a></td>
</tr>

<%
RS.MoveNext
loop
RS.Close
%>

</table>
<SCRIPT>valignbottom()</SCRIPT>
Page��[
<script>
PageCount=<%=TotalPage%> //��ҳ��
topage=<%=PageCount%>   //��ǰͣ��ҳ

for (var i=1; i <= PageCount; i++) {
if (i <= topage+3 && i >= topage-3 || i==1 || i==PageCount){
if (i > topage+4 || i < topage-2 && i!=1 && i!=2 ){document.write(" ... ");}
if (topage==i){document.write(" "+ i +" ");}
else{
document.write("<a href=?topage="+i+">"+ i +"</a> ");
}
}
}
</script>
 ]<br>

<%
end if



%>

<br>
<INPUT onclick=history.back(-1) type=button value=" << �� �� ">



<%
htmlend
%>