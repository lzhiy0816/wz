<!-- #include file="setup.asp" --><%

if Request.Cookies("username")=empty then error("<li>����δ<a href=login.asp>��¼</a>����")


id=int(Request("id"))


select case Request("menu")
case "add"

if Request.ServerVariables("request_method")="POST" then
url=Request("url")
else
url=Request.ServerVariables("http_referer")
end if

conn.execute("insert into favorites(username,name,url)values('"&Request.Cookies("username")&"','"&Request("name")&"','"&url&"')")

message="<li>��ӳɹ�<li><a href=favorites.asp>�����ղؼ�</a><li><a href=usercp.asp>���ؿ������</a><li><a href=main.asp>����������ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=favorites.asp>")


case "del"
conn.execute("delete from [favorites] where username='"&Request.Cookies("username")&"' and id="&id&"")


end select

top
%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body> <title>�������</title><table width=97% align="center" border="0"><tr><td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="����ҳ"></a></td><td vAlign="center" align="top">��<img src="images/closedfold.gif" border="0">��<a href=main.asp><%=clubname%></a><br>��<img src="images/bar.gif" border="0"><img src="images/openfold.gif">���������</td></tr></table>
<br>
<table cellspacing="1" cellpadding="1" width="97%" align="center" border="0" class="a2">
  <TR id=TableTitleLink class=a1 height="25">
      <Td width="14%" align="center"><b><a href="usercp.asp">���������ҳ</a></b></td>
      <TD width="14%" align="center"><b><a href="EditProfile.asp">���������޸�</a></b></td>
      <TD width="14%" align="center"><b><a href="EditProfile.asp?menu=contact">�༭��̳ѡ��</a></b></td>
      <TD width="14%" align="center"><b><a href="EditProfile.asp?menu=pass">�û������޸�</a></b></td>
      <TD width="14%" align="center"><b><a href="message.asp">�û����ŷ���</a></b></td>
      <TD width="14%" align="center"><b><a href="friend.asp">�༭�����б�</a></b></td>
      <TD width="16%" align="center"><b><a href="favorites.asp">�û��ղع���</a></b></td></TR></TABLE>

<br>
<SCRIPT>valigntop()</SCRIPT>
<table width="97%" border=0 align=center cellPadding=3 cellSpacing=1 class=a2><tr><td vAlign="top" colSpan="3"  class=a1 align="center"><b>
&gt;&gt; �����ղؼ� &lt;&lt;</b></td></tr><tr><td width="69%" align="center" bgcolor="FFFFFF"><b>�� ��</b></td><td width="20%" align="center" bgcolor="FFFFFF"><b>ʱ��</b></td><td width="11%" align="center" bgcolor="FFFFFF"><b>����</b></td></tr><%
  

  
sql="select * from favorites where username='"&Request.Cookies("username")&"' order by id Desc"
rs.Open sql,Conn,1


pagesetup=10 '�趨ÿҳ����ʾ����
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '��ҳ��
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '��ת��ָ��ҳ��


i=0 
Do While Not RS.EOF and i<pagesetup 
i=i+1 


%> <tr><td bgcolor="FFFFFF"><img src=images/ie.gif> <a target=_blank href="<%=rs("url")%>"><%=rs("name")%></a></td><td align=center bgcolor="FFFFFF"><%=rs("addtime")%>��</td><td align=center bgcolor="FFFFFF"><a href="?menu=del&id=<%=rs("id")%>"><img src=images/del.gif border=0></a></td></tr><%          
RS.MoveNext
loop
RS.Close      
%> <tr><td vAlign="top" colSpan="3"  class=a1 align="center"><b>&gt;&gt; ��������� &lt;&lt;</b></td></tr><tr><td align="center" colspan="3" bgcolor="FFFFFF"><form method=post name=form action=favorites.asp><input type=hidden name=menu value=add><b>�� �⣺</b><INPUT size=20 name=name>��<b>���ӵ�ַ��</b><INPUT size=40 name=url value="http://">��<input type=submit value="�� ��" name="Submit"> </td></form></tr></table>
<SCRIPT>valignbottom()</SCRIPT>
<center><br>Page��[ 

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
 

 ]<br><%
htmlend
%>