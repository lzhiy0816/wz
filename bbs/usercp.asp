<!-- #include file="setup.asp" -->

<%
top
if Request.Cookies("username")=empty then error("<li>����δ<a href=login.asp>��¼</a>����")

%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>








<title>�������</title>
<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="����ҳ"></a></td>
<td vAlign="center" align="top">��<img src="images/closedfold.gif" border="0">��<a href=main.asp><%=clubname%></a><br>
��<img src="images/bar.gif" border="0"><img src="images/openfold.gif">���������</td>
</tr>
</table>
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
<TABLE cellSpacing=0 cellPadding=0 width=97% align=center>

  <TR>
    <TD vAlign=top width="28%">
      <TABLE style="WIDTH: 95%" height="100%" cellSpacing=1 cellPadding=3 border=0 class=a2>
        <TBODY>
        <TR>
          <TD height=25 class=a1 align="center" colspan="2">
          <b>�û�ͷ��</TH></b></TR>
        <TR align=middle>
          <TD colspan="2" bgcolor="FFFFFF"><img src="<%=userface%>" onload='javascript:if(this.width>200)this.width=200'></TD></TR>
        <TR>
          <TD height=25 class=a1 align="center" colspan="2">
          <b>��������</TH></b></TR>


       <%
sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn
response.write "<TR><TD bgcolor=FFFFFF>�û��ȼ��� <Script>document.write(level("&rs("experience")&","&rs("membercode")&",'','')+levelname)</Script><br>�û����ɣ� "&rs("faction")&"<br>�û���ż�� "&rs("consort")&"<br>�û������� "&rs("userlife")&"<br>����ԭ���� "&rs("posttopic")&"<br>��������� "&rs("postrevert")&"<br>���������� "&rs("goodtopic")&"<br>��ɾԭ���� "&rs("deltopic")&"<br>������ң� "&rs("money")&"<br>�û����飺 "&rs("experience")&"<br>��¼������ "&rs("degree")&"<br>ע��ʱ�䣺 "&rs("regtime")&"<br>�ϴε�¼�� "&rs("landtime")&"</TD></TR>"

rs.close
%>   


</TABLE></TD>
    <TD vAlign=top>
      <TABLE style="WIDTH: 100%" height=29 cellSpacing=1 cellPadding=3 align=center border=0 class=a2>

        <TR>
          <TD align=left height=25 class=a1>
          <b>-=&gt; ��̳��ѶϢ</b></TD></TR>
        <TR>
          <TD align="center" bgcolor="FFFFFF"><a href="message.asp"><FONT 
            color=ff0000>�ռ���</FONT></a><font color="FF0000">&nbsp; </font>���� <B>[<%=conn.execute("Select count(id)from message where incept='"&Request.Cookies("username")&"'")(0)%>]</B> ��ѶϢ��<a href="message.asp?send=1"><FONT 
            color=ff0000>�ѷ���ѶϢ</FONT></a><B> [<%=conn.execute("Select count(id)from message where author='"&Request.Cookies("username")&"'")(0)%>]</B> ����<BR></TD></TR></TBODY></TABLE><BR>
      <TABLE  style="WIDTH: 100%" cellSpacing=1 cellPadding=3 align=center border=0 class=a2>
        <TBODY>
        <TR>
          <TD align=left colSpan=5 height=25 class=a1>
          <b>-=&gt; �����յ��Ķ�Ѷ</TH></b></TR>
        <TR>
          <TD vAlign=center align=middle 
          width=30 bgcolor="FFFFFF">��</TD>
          <TD vAlign=center align=middle 
          width=100 bgcolor="FFFFFF"><B>������</B></TD>
          <TD vAlign=center align=middle 
          width=300 bgcolor="FFFFFF"><B>����</B></TD>
          <TD align=middle 
          width=160 bgcolor="FFFFFF"><B>����</B></TD>
          <TD vAlign=center align=middle 
          width=50 bgcolor="FFFFFF"><B>��С</B></TD></TR>
          
<%

sql="select top 5 author,content,time from message where incept='"&Request.Cookies("username")&"' order by time Desc"
rs.Open sql,Conn
Do While Not RS.EOF 
size=Len(""&rs("content")&"")
if size>20 then
content=left(""&rs("content")&"",16)&"..."
else
content=rs("content")
end if
%>
   
        <TR>
          <TD vAlign=center align=middle bgcolor="FFFFFF">
          <IMG 
            src="images/f_norm.gif"></TD>
          <TD vAlign=center align=middle bgcolor="FFFFFF"><A href="Profile.asp?username=<%=rs("author")%>" ><%=rs("author")%></A>��</TD>
          <TD bgcolor="FFFFFF"><A href=message.asp><%=content%></A>��</TD>
          <TD align="center" bgcolor="FFFFFF"><%=rs("time")%></FONT>��</TD>
          <TD align="center" bgcolor="FFFFFF"><%=size%>��</TD></TR>
          
<%          
RS.MoveNext
loop
RS.Close      
%>          
          
          
          
          </TD></TR></TABLE><BR>
      <TABLE style="WIDTH: 100%" cellSpacing=1 cellPadding=3 align=center border=0 class=a2>

        <TR>
          <TD align=left height=25 class=a1>
          <b>-=&gt; 
  ������������</TH></b></TR>
    <%
sql="select top 5 * from forum where deltopic<>1 and username='"&Request.Cookies("username")&"' order by lasttime Desc  "
rs.Open sql,Conn

Do While Not RS.EOF 

%>

  <tr>
    <td align="left" bgcolor="FFFFFF">&nbsp;����<a onclick=min_yuzi() target=message href=ShowPost.asp?id=<%=rs("id")%>><%=rs("topic")%></a> -- <%=rs("lasttime")%></td>
  </tr>


<%


RS.MoveNext
loop
RS.Close

  %>
</table></TD></TR></TBODY></TABLE>
  
<%
htmlend
%>