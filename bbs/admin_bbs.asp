<!-- #include file="setup.asp" -->
<%
if adminpassword<>session("pass") then response.redirect "admin.asp?menu=login"
id=server.htmlencode(Request("id"))


log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")


%>
<META http-equiv=Content-Type content=text/html;charset=gb2312>
<link href=images/skins/<%=Request.Cookies("skins")%>/bbs.css rel=stylesheet>
<br><center>
<p></p>

<%


select case Request("menu")
case "applymanage"
applymanage


case "bbsmanage"
bbsmanage

case "bbsmanagexiu"
bbsmanagexiu


case "bbsmanagexiuok"
bbsmanagexiuok

case "bbsadd"
bbsadd

case "bbsaddok"
bbsaddok

case "classs"
classs

case "classxiu"
classxiu

case "upclubconfig"
upclubconfig

case "upclubconfigok"
upclubconfigok

case "classadd"
if id="" then id=1
if not isnumeric(id) then error2("�������������д")
if Request("classname")="" then error2("����д�������")
conn.Execute("insert into class(id,classname) values ('"&id&"','"&Request("classname")&"')")
classs

case "classdel"
conn.execute("delete from [class] where id="&id&"")
classs

case "classxiuok"
conn.execute("update [class] set classname='"&Request("classname")&"',id="&id&" where id="&Request("oldid")&"")
classs

case "bbsmanagedel"
conn.execute("delete from [bbsconfig] where id="&id&"")
error2("�Ѿ�������̳����������ɾ���ˣ�")

case "delforumok"

if request("jh")="1" then
jh=" and goodtopic<>1"
end if

if request("bbsid")<>"" then
bbsid="and forumid="&request("bbsid")&""
end if


conn.execute("delete from [forum] where lasttime<"&SqlNowString&"-"&request("TimeLimit")&" "&jh&" "&bbsid&"")

error2("�Ѿ���"&request("TimeLimit")&"��û���û��ظ�������ɾ���ˣ�")


case "delretopicok"
conn.execute("delete from [reforum] where posttime<"&SqlNowString&"-"&request("TimeLimit")&"")
error2("�Ѿ���"&request("TimeLimit")&"��Ļ���ɾ���ˣ�")


case "delbbsconfigok"

if request("hide")="1" then
hide=" and hide=1"
end if

conn.execute("delete from [bbsconfig] where lasttime<"&SqlNowString&"-"&request("TimeLimit")&" "&hide&"")
error2("�Ѿ���"&request("TimeLimit")&"��û�������ӵ���̳ɾ���ˣ�")



case "deltopicok"
if request("topic")="" then
error2("��û�������ַ���")
end if
conn.execute("delete from [forum] where topic like '%"&request("topic")&"%' ")
error2("�Ѿ�������������� "&request("topic")&" ������ȫ��ɾ���ˣ�")


case "uniteok"
if Request("hbbs") = Request("ybbs") then
error2("����ѡ����ͬ��̳��")
end if
conn.execute("update [forum] set forumid="&Request("hbbs")&" where forumid="&Request("ybbs")&"")
error2("�Ѿ��ɹ���2����̳�����Ϻϲ��ˣ�")



end select

sub applymanage
%>
<script>function checkclick(msg){if(confirm(msg)){event.returnValue=true;}else{event.returnValue=false;}}</script>

<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">



  <tr class=a1 id=TableTitleLink>
<td align="center" height="25"><a href="?menu=applymanage&fashion=id">
ID</a></td>
<td width="20%" align="center" height="25">
<a href="?menu=applymanage&fashion=bbsname">��̳</a></td>
<td align=center height="25">����</td>
<td align=center height="25"><a href="?menu=applymanage&fashion=toltopic">
����</a></td>
<td align=center height="25"><a href="?menu=applymanage&fashion=tolrestore">
����</a></td>

<td align="center" height="25">����</td>
    
  </tr>


<%

if Request("fashion")=empty then
fashion="tolrestore"
else
fashion=Request("fashion")
end if


sql="select * from bbsconfig where hide=1 order by "&fashion&" Desc"
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
  <tr class=a3>
<td align="center" height="25"><%=rs("id")%></td>
<td><a target=_blank href=ShowForum.htm?forumid=<%=rs("id")%>><%=rs("bbsname")%></a></td>
<td align=center><%=rs("moderated")%></td>
<td align=center><b><font color=red><%=rs("toltopic")%></font></b></td>
<td align=center><b><font color=red><%=rs("tolrestore")%></font></b></td>

<td align="center"><a href=?menu=bbsmanagexiu&id=<%=rs("id")%>>�༭��̳</a> | <a onclick=checkclick('��ȷ��Ҫɾ������̳����������?') href=?menu=bbsmanagedel&id=<%=rs("id")%>>ɾ����̳</a>
��</td>
    
  </tr>
<%
RS.MoveNext
loop
RS.Close

%>
</table>

[<b>
<script>
PageCount=<%=TotalPage%> //��ҳ��
topage=<%=PageCount%>   //��ǰͣ��ҳ
for (var i=1; i <= PageCount; i++) {
if (i <= topage+3 && i >= topage-3 || i==1 || i==PageCount){
if (i > topage+4 || i < topage-2 && i!=1 && i!=2 ){document.write(" ... ");}
if (topage==i){document.write(" "+ i +" ");}
else{
document.write("<a href=?menu=applymanage&fashion=<%=Request("fashion")%>&topage="+i+">"+ i +"</a> ");
}
}
}
</script>

</b>]
<%

end sub


sub classs
%>
<script>function checkclick(msg){if(confirm(msg)){event.returnValue=true;}else{event.returnValue=false;}}</script>
<table border="0" width="90%">
	<tr>
		<td align="center"><form method="post" action="?menu=classadd">
������ƣ������磺�������磩<INPUT size=20 name=classname> ��ţ�<INPUT size=2 name=id value=<%=conn.execute("Select max(id)from class")(0)+1%>>
<input type=submit value="����"></form></td>
		<td align="center"><form method="post" action="?menu=bbsmanagexiu">
������̳��<INPUT size=2 name=id value="ID" onfocus="this.value = ''" >
<input type=submit value="ȷ��"></form></td>
	</tr>
</table>





<%


sql="select * from class order by id"
rs.Open sql,Conn
do while not rs.eof

%>


<table border="0" cellpadding="5" cellspacing="0"  width="90%" class=a2>
  <tr class=a1 id=TableTitleLink>
    <td width="50%">
<%=rs("id")%>.<%=rs("classname")%>
    </td>

    <td width="50%" align="right">
    <a href=?menu=bbsadd&id=<%=rs("id")%>>������̳</a> | <a href=?menu=classxiu&id=<%=rs("id")%>>
    �༭����</a> | <a href=?menu=classdel&id=<%=rs("id")%>>
    ɾ������</a>

    ��</td>
  </tr>
</table>
<%


sql="select * from bbsconfig  where classid="&rs("id")&" and hide=0"
rs1.Open sql,Conn
do while not rs1.eof
%>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="90%" class=a4>
  <tr>

<td width="50%">

<li><%=rs1("bbsname")%></td>
    
<td width="50%" align="right"><a href=?menu=bbsmanagexiu&id=<%=rs1("id")%>>�༭��̳</a> | <a onclick=checkclick('��ȷ��Ҫɾ������̳����������?') href=?menu=bbsmanagedel&id=<%=rs1("id")%>>ɾ����̳</a>
��</td>
    
  </tr>
</table>
<%
rs1.movenext
loop
rs1.close




rs.movenext
loop
rs.close

end sub

sub classxiu
sql="select * from class where id="&id&""
rs.Open sql,Conn
%>
<form method="post" action="?menu=classxiuok">
<input type=hidden name=oldid  value=<%=id%>>
������ƣ�<INPUT size=20 name=classname value=<%=rs("classname")%>> ��ţ�<INPUT size=2 name=id value=<%=id%>>
<input type=submit value="�޸�"><center><br><br><a href=javascript:history.back()>< �� �� ></a>
<%
rs.close
end sub

sub bbsadd

%>


<table cellspacing="1" cellpadding="2" width="80%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan="2">������̳����</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle>

<form name="form" method="post" action="?menu=bbsaddok">
<input type=hidden name=classid value=<%=id%>>

��̳����</td>
    <td class=a3>

<input size="20" name="bbsname"></td></tr>
   <tr height=25>
    <td class=a3 align=middle>

��̳����</td>
    <td class=a3>

<input size="30" name=moderated> ������������|�ָ����磺yuzi|ԣԣ
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

��̳����</td>
    <td class=a3>

<textarea rows="5" name="intro" cols="50"></textarea></td></tr>
   <tr height=25>
    <td class=a3 align=middle>

��̳״̬</td>
    <td class=a3>

<select size="1" name="pass">
<option value=0>�ر���̳</option>
<option value=1 selected>������̳</option>
<option value=2>��Ա��̳</option>
<option value=3>�α���̳</option>
</select>
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

Сͼ��URL</td>
    <td class=a3>

<input size="30" name="icon"> ��ʾ��������ҳ��̳�������
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

��ͼ��URL</td>
    <td class=a3>

<input size="30" name="logo"> ��ʾ����̳���Ͻ�
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

�Ƿ���ʾ����̳�б�����������������������</td>
    <td class=a3>

<input type="radio" CHECKED value="0" name="hide" id=hide1><label for=hide1>��ʾ</label>
 
<input type="radio" value="1" name="hide" id=hide2><label for=hide2>����</label> ��</td></tr>
   <tr height=25>
    <td class=a3 align=middle colspan="2">

�� <input type="submit" value=" �� �� " name="Submit"><br></td></tr></table>
</form>
<center><br><a href=javascript:history.back()>< �� �� ></a>


<%
end sub
sub bbsaddok
if Request("bbsname")="" then error2("��������̳����")


temp="|"&Request("moderated")&"|"
temp=replace(temp,"||","|")
master=split(temp,"|")
for i = 1 to ubound(master)-1
If conn.Execute("Select id From [user] where username='"&HTMLEncode(master(i))&"'" ).eof and master(i)<>"" Then error2(""&master(i)&"���û����ϻ�δע��")
next



rs.Open "bbsconfig",Conn,1,3
rs.addnew
rs("classid")=Request("classid")
rs("bbsname")=server.htmlencode(Request("bbsname"))
rs("moderated")=temp
rs("pass")=Request("pass")
rs("intro")=server.htmlencode(Request("intro"))
rs("icon")=server.htmlencode(Request("icon"))
rs("logo")=server.htmlencode(Request("logo"))
rs.update
id=rs("id")

rs.close

classs

end sub




sub bbsmanagexiuok

if Request("bbsname")="" then error2("��������̳����")

if Request("hide")=0 and Request("classid")="" then
error2("��ָ����̳�����")
end if


temp="|"&Request("moderated")&"|"
temp=replace(temp,"||","|")
master=split(temp,"|")
for i = 1 to ubound(master)-1
If conn.Execute("Select id From [user] where username='"&HTMLEncode(master(i))&"'" ).eof and master(i)<>"" Then error2(""&master(i)&"���û����ϻ�δע��")
next


sql="select * from bbsconfig where id="&id&""
rs.Open sql,Conn,1,3


if Request("classid")<>empty then rs("classid")=Request("classid")
rs("bbsname")=server.htmlencode(Request("bbsname"))
rs("moderated")=temp
rs("pass")=Request("pass")
rs("intro")=server.htmlencode(Request("intro"))
rs("icon")=server.htmlencode(Request("icon"))
rs("logo")=server.htmlencode(Request("logo"))
rs("hide")=Request("hide")
rs.update

rs.close
%>
�༭�ɹ�<br><br><a href=javascript:history.back()>�� ��</a>
<%
end sub


sub bbsmanagexiu
rs.open "class order by id",conn
do while not rs.eof
classlist=""+classlist+"<option value="&rs("id")&">"&rs("classname")&"</option>"
rs.movenext
loop
rs.close
sql="select * from bbsconfig where id="&id&""
rs.Open sql,Conn
%>


<form method="post" action="?menu=bbsmanagexiuok"><input type=hidden name=id value=<%=rs("id")%>>
<table cellspacing="1" cellpadding="2" width="80%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan="2">�༭��̳����</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle>

��̳����</td>
    <td class=a3>


<input size="10" name="bbsname" value="<%=rs("bbsname")%>"> ��� 
<select name="classid">
<option value="<%=rs("classid")%>">Ĭ��</option>
<%=classlist%>
</select></td></tr>
   <tr height=25>
    <td class=a3 align=middle>


��̳����</td>
    <td class=a3>


<input size="30" name=moderated value="<%=rs("moderated")%>"> ������������|�ָ����磺yuzi|ԣԣ
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>


��̳����</td>
    <td class=a3>


<textarea rows="5" name="intro" cols="50"><%=rs("intro")%></textarea></td></tr>
   <tr height=25>
    <td class=a3 align=middle>


��̳״̬</td>
    <td class=a3>


<select size="1" name="pass">
<option value=0 <%if rs("pass")=0 then%>selected<%end if%>>�ر���̳</option>
<option value=1 <%if rs("pass")=1 then%>selected<%end if%>>������̳</option>
<option value=2 <%if rs("pass")=2 then%>selected<%end if%>>��Ա��̳</option>
<option value=3 <%if rs("pass")=3 then%>selected<%end if%>>�α���̳</option>
</select>
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>


Сͼ��URL</td>
    <td class=a3>


<input size="30" name="icon" value="<%=rs("icon")%>"> ��ʾ��������ҳ��̳�������
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>


��ͼ��URL</td>
    <td class=a3>


<input size="30" name="logo" value="<%=rs("logo")%>"> ��ʾ����̳���Ͻ�</td></tr>
   <tr height=25>
    <td class=a3 align=middle>


�Ƿ���ʾ����̳�б�</td>
    <td class=a3>


<input type="radio" <%if rs("hide")=0 then%>CHECKED <%end if%>value="0" name="hide" value="0" id=hide1><label for=hide1>��ʾ</label> 
<input type="radio" <%if rs("hide")=1 then%>CHECKED <%end if%>value="1" name="hide" value="1" id=hide2><label for=hide2>����</label> </td></tr>
   <tr height=25>
    <td class=a3 align=middle colspan="2">


<input type="submit" value=" �� �� " name="Submit1"></td></tr></table><br></form>
<center><br><a href=javascript:history.back()>< �� �� ></a>
<%
end sub



sub bbsmanage

%>

��̳����<b><font color=red><%=conn.execute("Select count(id)from bbsconfig")(0)%></font></b>������������<b><font color=red><%=conn.execute("Select count(id)from forum")(0)%></font></b>������������<b><font color=red><%=conn.execute("Select count(id)from reforum")(0)%></font></b><br>

<form method="post" action="?menu=delforumok">

<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle>����ɾ������</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle>
    
ɾ�� <select name=TimeLimit>
<option value="90">90</option>
<option value="180" selected>180</option>
<option value="360">360</option>
</select> ��û���û��ظ������⡡<input type="checkbox" value="1" name="jh" id=jh checked><label for="jh">�������ӳ���</label>
<br>
<select name="bbsid">
<option value="">������̳</option>

<%
sql="select id,bbsname,classid from bbsconfig where hide=0 order by classid,id"
rs.Open sql,Conn
do while not rs.eof
Classid=Trim(Rs("classid"))
		if TClass <> Classid Then
		Response.write "<option style=BACKGROUND-COLOR:ECF5FF value=''>�� "&Conn.Execute("Select classname From class where id="&Classid)(0)&"</option>"
		TClass = Classid
		end if
		
	Response.write "<option value="&rs("id")&""&selected&">������"&rs("bbsname")&"��</option>"
	rs.movenext
loop
rs.close
%>

</select>
 <input type="submit" value=" ȷ �� ">
 


</td></form></tr>
   <tr height=25>
    <td class=a4 align=middle>
    <form method="post" action="?menu=deltopicok">
ɾ������������� <input size="20" name="topic"> ���������� <input type="submit" value="ȷ��">

��</td></tr></table></form>


<form method="post" action="?menu=delretopicok">
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>����ɾ������</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>


ɾ�� <select name=TimeLimit>
<option value="90">90</option>
<option value="180" selected>180</option>
<option value="360">360</option>
</select> ����ǰ�Ļ���  <input type="submit" value=" ȷ �� ">
</td></tr></table></form>

<form method="post" action="?menu=delbbsconfigok">
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>����ɾ����̳</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
ɾ�� <select name=TimeLimit>
<option value="30">30</option>
<option value="60" selected>60</option>
<option value="90">90</option>
</select> ��û�������ӵ���̳ <input type="checkbox" value="1" name="hide" id=hide checked><label for="hide">�б��е���̳����</label><br><input type="submit" value=" ȷ �� ">
</td></tr></table></form>



<form method="post" action="?menu=uniteok">
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>�ϲ���̳����</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>

�� <select name="ybbs">
<%

rs.Open "bbsconfig where hide=0",Conn

do while not rs.eof
%>
<option value="<%=rs("id")%>">��<%=rs("bbsname")%>��</option>
<%
rs.movenext
loop
rs.close
%>
</select> �ϲ��� <select name="hbbs">
<%
rs.Open "bbsconfig where hide=0",Conn
do while not rs.eof
%>
<option value="<%=rs("id")%>">��<%=rs("bbsname")%>��</option>
<%
rs.movenext
loop
rs.close
%>
</select>
<br>
<INPUT type=submit value=" ȷ �� " name=submit>
</td></tr></table></form>






<%
end sub

sub upclubconfig
%>

<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>�������ϸ���</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
    
�˲�����������̳���ϡ��û�����<br>����ֹ��̳ͳ�ƴ�����û�������ʾ����ȷ�����������<br>
<a href="?menu=upclubconfigok">������������̳ͳ������</a><br>
</td></tr></table><br>

<%
end sub

sub upclubconfigok

rs.Open "bbsconfig",Conn
do while not rs.eof

allarticle=conn.execute("Select count(forumid) from forum where forumid="&rs("id")&"")(0)
if allarticle>0 then
allrearticle=conn.execute("Select sum(replies) from forum where forumid="&rs("id")&"")(0)
else
allrearticle=0
end if

conn.execute("update [bbsconfig] set toltopic="&allarticle&",tolrestore="&allarticle+allrearticle&" where ID="&rs("id")&"")


rs.movenext
loop
rs.close
%>
�����ɹ���<br>
<br>
�Ѿ���������������̳��ͳ������<br>

<%
end sub

htmlend



%>