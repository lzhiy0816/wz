<!-- #include file="setup.asp" -->
<%
if Request.Cookies("username")=empty then error("<li>����δ<a href=login.asp>��¼</a>����")


id=int(Request("id"))





if Request("menu")="post" then
%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<META http-equiv=Content-Type content=text/html;charset=gb2312>
<link href=images/skins/<%=Request.Cookies("skins")%>/bbs.css rel=stylesheet>
<SCRIPT>
function check(theForm) {

if(theForm.incept.value == "" ) {
alert("�ǳƲ���û����д��");
return false;
}

if(theForm.incept.value == "<%=Request.Cookies("username")%>" ) {
alert("��������Ҫ���͵Ķ��󣬲��ܷ�ѶϢ���Լ���");
return false;
}

if(theForm.content.value == "" ) {
alert("���ܷ���ѶϢ��");
return false;
}

if (theForm.content.value.length > 255){
alert("�Բ�������ѶϢ���ܳ��� 255 ���ֽڣ�");
return false;
}
}

function presskey(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.form.submit();}}


function DoTitle(addTitle) {
document.form.incept.value=addTitle
}

function Check(){var Name=document.form.incept.value;
window.open("friend.asp?menu=post&log=1&incept="+Name,"_top","width=320,height=170");
}

</SCRIPT><TITLE>������Ϣ</TITLE><body topmargin=0 bgcolor="ECE9D8">
<TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0><TR><form name=form action="friend.asp" method="post">
<input type="hidden" name="menu" value="addpost">
<TD height="35">
&nbsp;�ǳƣ�<input name="incept" type="text" size="15">
</TD><TD align=right height="35">

<select onchange=DoTitle(this.options[this.selectedIndex].value)>
<option>�����б�</option>
<SCRIPT>
var moderated="<%=Conn.Execute("Select friend From [user] where username='"&Request.Cookies("username")&"'")(0)%>"
var list= moderated.split ('|'); 
for(i=0;i<list.length;i=i+1) {
if (list[i] !=""){document.write("<option value="+list[i]+">"+list[i]+"</option>")}
}
</SCRIPT>
</select>
</TD></TR><TR><TD VALIGN=top ALIGN=right colspan="2" bgcolor="F8F8F8">
    <textarea name="content" cols="39" rows="6" onkeydown=presskey()></textarea>
</TD></TR></TABLE><TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0 height="30">
<tr ALIGN=center><TD>
<input onclick=javascript:Check() type="button" value="�����¼">
</td><TD><input type="reset" value="ȡ������" OnClick="window.close();"> </td><TD><input name=submit1 type="submit" value="����ѶϢ" onclick="return check(this.form)"></td>
</TR></form>
</TABLE>
<body onload=resizeTo(330,206)>
<%
responseend

end if


if Request("menu")="del" then

conn.execute("delete from [message] where id="&id&" and incept='"&Request.Cookies("username")&"'")
end if




top

if newmessage>0 then
conn.execute("update [user] set newmessage=0 where username='"&Request.Cookies("username")&"'")
end if

%>

<script>function checkclick(msg){if(confirm(msg)){event.returnValue=true;}else{event.returnValue=false;}}</script>


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

<center>

<br>
<SCRIPT>valigntop()</SCRIPT>
<TABLE  cellSpacing=1 border=0 width="97%" class=a2 cellpadding="3">
<TR><TD align="center" bgcolor="FFFFFF" colspan=5>

          <a style="text-decoration: none; color: #000000" href="message.asp">
          <img alt src="images/m_inbox.gif" border="0" dypop="�ռ���"></a>&nbsp;
          <a style="text-decoration: none; color: #000000" href="message.asp?send=1">
          <img alt src="images/M_issend.gif" border="0" dypop="�ѷ����ʼ�"></a>&nbsp;
          <a href="friend.asp">
          <img alt src="images/M_address.gif" border="0" dypop="����ѶϢ"></a>&nbsp;
          <a style=cursor:hand onclick="javascript:open('?menu=post','','width=320,height=170')"><img alt src="images/m_write.gif" border="0" dypop="����ѶϢ"></a></td></TR>

  <TR>
          
    <TD vAlign=center align=middle 
          width=5% class=a1 height="20">��</TD>
          
    <TD vAlign=center align=middle 
          width=12% class=a1 height="20"><B><span id=send>������</span></B></TD>
          
    <TD vAlign=center align=middle 
          width=50% class=a1 height="20"><B>����</B></TD>
          
    <TD align=middle 
          width=17% class=a1 height="20"><B>����</B></TD>
          
    <TD vAlign=center align=middle 
          width=16% class=a1 height="20"><B>����</B></TD>
  </TR>



<%



if Request("send")="1" then
sql="select * from message where author='"&Request.Cookies("username")&"' order by time Desc"
response.write "<script>send.innerText='�ռ���'</script>"
else
sql="select * from message where incept='"&Request.Cookies("username")&"' order by time Desc"

end if

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


if Request("send")="1" then
author=rs("incept")
del="<a style=cursor:hand onclick=javascript:open('friend.asp?menu=post&log=1&incept="&author&"','','width=320,height=170')>�����¼</a>"

else
author=rs("author")
del="<a style=cursor:hand onclick=javascript:open('friend.asp?menu=post&incept="&rs("author")&"','','width=320,height=170')><img alt='�ظ��������' src=images/replynow.gif border=0></a> <a onclick=checkclick('��ȷ��Ҫɾ������ѶϢ?') href=?menu=del&id="&rs("id")&"><img alt='ɾ��������' src=images/del.gif border=0></a>"
end if


%>



        <TR>
          
    <TD vAlign=center align=middle width="5%" bgcolor="FFFFFF"> <IMG 
            src="images/f_norm.gif"></TD>
          
    <TD vAlign=center align=middle width="12%" bgcolor="FFFFFF"> <A href="Profile.asp?username=<%=author%>"><%=author%></A> 
      ��</TD>
          
    <TD width="50%" bgcolor="FFFFFF"><%=rs("content")%>��</TD>
          
    <TD align="center" width="17%" bgcolor="FFFFFF"><%=rs("time")%></FONT>��</TD>
          
    <TD align="center" width="16%" bgcolor="FFFFFF"> <%=del%> ��</TD>
  </TR>


<%
RS.MoveNext
loop
RS.Close
%>
</TD></TR></TABLE>
<SCRIPT>valignbottom()</SCRIPT>
<BR>

Page��[

<script>
PageCount=<%=TotalPage%> //��ҳ��
topage=<%=PageCount%>   //��ǰͣ��ҳ
for (var i=1; i <= PageCount; i++) {
if (i <= topage+3 && i >= topage-3 || i==1 || i==PageCount){
if (i > topage+4 || i < topage-2 && i!=1 && i!=2 ){document.write(" ... ");}
if (topage==i){document.write(" "+ i +" ");}
else{
document.write("<a href=?send=<%=Request("send")%>&topage="+i+">"+ i +"</a> ");
}
}
}
</script>

]


</TABLE><br>

<%htmlend%>