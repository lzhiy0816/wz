<!-- #include file="setup.asp" -->
<%
if adminpassword<>session("pass") then response.redirect "admin.asp?menu=login"
id=int(Request("id"))

log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")


%>
<META http-equiv=Content-Type content=text/html;charset=gb2312>
<link href=images/skins/<%=Request.Cookies("skins")%>/bbs.css rel=stylesheet>
<br><center>
<p></p>

<%

select case Request("menu")

case "link"
link

case "linkok"
linkok

case "dellink"
conn.execute("delete from [link] where id="&id&"")
link



case "variable"
variable
case "variableok"
variableok
end select


sub link





%>

<FORM action=?menu=linkok method=post>
<table cellspacing="1" cellpadding="5" width="97%" border="0" class="a2" align="center"><tr>
	<td height="7" class="a1" colspan="2">����<b> </b>�������ӹ���</td></tr><tr>
	<td height="6" class="a3">��̳���ƣ�<INPUT size=40 name=name></td>
	<td height="6" class="a3">��ַURL��<INPUT size=40 name=url value="http://"></td></tr><tr>
	<td height="6" class="a3">��̳��飺<INPUT size=40 name=intro></td>
	<td height="6" class="a3">ͼ��URL��<INPUT size=40 name=logo value="http://"></td></tr><tr>
	<td height="6" class="a4" colspan="2" align="center"><INPUT type=submit value=" �� �� " name=Submit>
<input type="reset" value=" �� �� " name="Submit2">

</td></tr></table>
</FORM>

<table cellspacing="1" cellpadding="5" width="97%" border="0" class="a2" align="center"><tr><td height="25" colspan="2" class="a1">����<b> </b>��������</td></tr><tr>
<td align="center" bgcolor=FFFFFF width="5%"><img src="images/shareforum.gif"></td>
<td class="a4"><%
rs.Open "link",Conn
do while not rs.eof
if rs("logo")="" or rs("logo")="http://" then
link1=link1+"<a title='"&rs("intro")&"' href="&rs("url")&" target=_blank>"&rs("name")&"</a><a href=?menu=dellink&id="&rs("id")&"><font color=red>��ɾ����</font></a>����"
else
link2=link2+"<a title='"&rs("name")&""&chr(10)&""&rs("intro")&"' href="&rs("url")&" target=_blank><img src="&rs("logo")&" border=0 width=88 height=31></a><a href=?menu=dellink&id="&rs("id")&"><font color=red>��ɾ����</font></a>����"
end if
rs.movenext
loop
rs.close
%>
<%=link1%>
<table cellspacing=0 width=100% align=center class=a2><tr><td></td></tr></table>
<%=link2%>
</td></tr></table>


<%




end sub

sub linkok

if Request("url")="http://" or Request("url")="" then
error2("��̳URLû����д")
end if
conn.Execute("insert into link(name,logo,url,intro) values ('"&Request("name")&"','"&Request("logo")&"','"&Request("url")&"','"&Request("intro")&"')")
link
end sub




sub variable
if ""&cluburl&""=empty then cluburl="http://"&Request.ServerVariables("server_name")&""&replace(Request.ServerVariables("script_name"),"admin_setup.asp","")&""
%>

<form method="post" action="?menu=variableok">
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>����������Ϣ</td>
  </tr>


<tr class=a3>
<td width="50%">�������ƣ�</td>
<td><input size="30" name="clubname" value="<%=clubname%>"></td>
</tr>
<tr class=a4>
<td>����URL��<br>���磺 <font color="FF0000">http://<%=Request.ServerVariables("server_name")%><%=replace(Request.ServerVariables("script_name"),"admin_setup.asp","")%></font></td>
<td><input size="30" name="cluburl" value="<%=cluburl%>"></td>
</tr>
<tr class=a3>
<td>��ҳ���ƣ��ײ���ʾ����</td>
<td><input size="30" name="homename" value="<%=homename%>"></td>
</tr>
<tr class=a4>
<td>��ҳ��ַ���ײ���ʾ����</td>
<td><input size="30" value="<%=homeurl%>" name="homeurl"></td>
</tr>
<tr class=a3>
<td>�����루�ײ���ʾ��֧��HTML����</td>
<td>
<textarea name="adcode" rows="3" cols="30"><%=adcode%></textarea>


</td>
</tr>
</table>



<br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>������Ϣ����</td>
  </tr>
  
  <tr class=a3>
<td>�û�����ʱ�䣺�����趨������֮����Զ��ͷţ�</td>
<td><input size="4" value="<%=prison%>" name="prison"> ��&nbsp; [Ĭ��:15]</td>
</tr>

  <tr class=a4>
<td>�ű���ʱʱ�䣺</td>
<td><input size="4" value="<%=Timeout%>" name="Timeout"> ��&nbsp; [Ĭ��:60]</td>
</tr>
  
  <tr class=a3>
<td width="50%">ͳ���û�����ʱ�䣺</td>
<td><input size="4" value="<%=OnlineTime%>" name="OnlineTime"> ��&nbsp; [Ĭ��:1200]</td>
	</tr>
  
  <tr class=a4>
<td>���������֣��������á�|���ָ�����<br>���磺 <font color="FF0000">fuck|shit|����</font></td>
<td><input size="30" value="<%=badwords%>" name="badwords"></td>
</tr>

  <tr class=a3>
<td>�����û����ӣ����û����á�|���ָ�����<br>���磺 <font color="FF0000">yuzi|ԣԣ</font></td>
<td><input size="30" value="<%=badlist%>" name="badlist"></td>
</tr>

<tr class=a4>
<td>��ֹIP��ַ������̳����IP���á�|���ָ�����<br>���磺 <font color="FF0000">127.0.0.1|192.168.0.1</font></td>
<td><input size="30" value="<%=badip%>" name="badip"></td>
</tr>
</table>




<br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>Email��Ϣ����</td>
  </tr>
  

<tr class=a3>
<td width="50%"0>�����ʼ������</td>
<td>
<select name="selectmail">
<option value="">�ر�</option>
<option value="JMail" <%if selectmail="JMail" then%>selected<%end if%>>JMail</option>
<option value="CDONTS" <%if selectmail="CDONTS" then%>selected<%end if%>>CDONTS</option>
</select>
</td>
</tr>
<tr class=a4>
<td>SMTP Server��ַ��</td>
<td><input size="30" value="<%=smtp%>" name="smtp"></td>
</tr>
<tr class=a3>
<td>�ʼ���������¼����</td>
<td><input size="30" value="<%=MailServerusername%>" name="MailServerusername"></td>
</tr>
<tr class=a4>
<td>�ʼ���������¼���룺</td>
<td>
<input size="30" value="<%=MailServerPassword%>" name="MailServerPassword" type="password"></td>
</tr>
<tr class=a3>
<td>������Email��ַ��</td>
<td><input size="30" value="<%=smtpmail%>" name="smtpmail"></td>
</tr>  
</table>




<br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>��������ѡ��</td>
  </tr>

<tr class=a3>
<td>����������̳���ܣ�</td>
<td>
<input type=radio name="apply" value="0" <%if apply=0 then%>checked<%end if%> id=apply><label for=apply>�ر�</label>
<input type=radio name="apply" value="1" <%if apply=1 then%>checked<%end if%> id=apply_1><label for=apply_1>����</label>
</td>
</tr>
<tr class=a3>
<td width="50%">������ҳ�Զ�չ��������̳��</td>
<td>
<input type=radio name="allclass" value="0" <%if allclass=0 then%>checked<%end if%> id=allclass><label for=allclass>�ر�</label>
<input type=radio name="allclass" value="1" <%if allclass=1 then%>checked<%end if%> id=allclass_1><label for=allclass_1>��</label>
</td>
</tr>

<tr class=a4>
<td>ע���û�����ͨ��Email���ͣ�<br><font color="FF0000">����������֧���ʼ����͹���</font></td>
<td>
<input type=radio name="SendPassword" value="0" <%if SendPassword=0 then%>checked<%end if%> id=SendPassword><label for=SendPassword>�ر�</label>
<input type=radio name="SendPassword" value="1" <%if SendPassword=1 then%>checked<%end if%> id=SendPassword_1><label for=SendPassword_1>��</label>
</td>
</tr>
  
</table>



<br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>�û�����ѡ��</td>
  </tr>
  
<tr class=a3>
<td width="50%">���û�ע�᣺</td>
<td>
<input type=radio name="CloseRegUser" value="1" <%if CloseRegUser=1 then%>checked<%end if%> id=CloseRegUser><label for=CloseRegUser>�ر�</label>
<input type=radio name="CloseRegUser" value="0" <%if CloseRegUser=0 then%>checked<%end if%> id=CloseRegUser_1><label for=CloseRegUser_1>����</label>
</td>
</tr>

<tr class=a4>
<td>ע����Ƿ���Ҫͣ��10���Ӳ��ܷ���</td>
<td>
<input type=radio name="Reg10" value="0" <%if Reg10=0 then%>checked<%end if%> id=Reg10><label for=Reg10>�ر�</label>
<input type=radio name="Reg10" value="1" <%if Reg10=1 then%>checked<%end if%> id=Reg10_1><label for=Reg10_1>��</label>
</td>
</tr>

<tr class=a3>
<td>һ��Emailֻ��ע��һ���ʺ�</td>
<td>
<input type=radio name="RegOnlyMail" value="0" <%if RegOnlyMail=0 then%>checked<%end if%> id=RegOnlyMail><label for=RegOnlyMail>�ر�</label>
<input type=radio name="RegOnlyMail" value="1" <%if RegOnlyMail=1 then%>checked<%end if%> id=RegOnlyMail_1><label for=RegOnlyMail_1>��</label>
</td>
</tr>

  
</table>

<br>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>�ϴ��ļ�ѡ��</td>
  </tr>
  
<tr class=a3>
<td width="50%">����ͷ���ļ��Ĵ�С��</td>
<td>


<input size="6" value="<%=MaxFace%>" name="MaxFace"> byte&nbsp; [Ĭ��:10240]</td>
</tr>

<tr class=a4>
<td>������Ƭ�ļ��Ĵ�С��</td>
<td>
<input size="6" value="<%=MaxPhoto%>" name="MaxPhoto"> byte&nbsp; [Ĭ��:30720]</td>
</tr>
  
 <tr class=a3>
<td>�������ļ��Ĵ�С��</td>
<td>
<input size="6" value="<%=MaxFile%>" name="MaxFile"> byte&nbsp; [Ĭ��:102400]</td>
</tr>
   
  
  
  
</table>
<input type="submit" value=" �� �� ">
<%
end sub


sub variableok

if Request("clubname")="" then
error2("��������������")
end if



filtrate=split(Request("badip"),"|")
for i = 0 to ubound(filtrate)
if instr("|"&remoteaddr&"","|"&filtrate(i)&"") > 0 then error2("�����������IP��ַ�Ƿ���ȷ")
next




rs.Open "clubconfig",Conn,1,3
rs("badwords")=Request("badwords")
rs("clubname")=Request("clubname")
rs("cluburl")=Request("cluburl")
rs("homename")=Request("homename")
rs("homeurl")=Request("homeurl")
rs("selectmail")=Request("selectmail")
rs("smtp")=Request("smtp")
rs("smtpmail")=Request("smtpmail")
rs("MailServerusername")=Request("MailServerusername")
rs("MailServerPassword")=Request("MailServerPassword")
rs("adcode")=Request("adcode")
rs("badlist")=Request("badlist")
rs("badip")=Request("badip")


rs("allclass")=""&Request("allclass")&","&Request("CloseRegUser")&","&Request("Reg10")&","&Request("RegOnlyMail")&","&Request("SendPassword")&","&Request("apply")&","&Request("prison")&","&Request("Timeout")&","&Request("OnlineTime")&","&Request("MaxFace")&","&Request("MaxPhoto")&","&Request("MaxFile")&""
rs.update

rs.close


on error resume next
if Request("selectmail")="JMail" then
Set JMail=Server.CreateObject("JMail.Message")
If -2147221005 = Err Then response.write "����������֧�� JMail.Message �������رշ����ʼ����ܣ�<br><br>"
elseif Request("selectmail")="CDONTS" then
Set MailObject = Server.CreateObject("CDONTS.NewMail")
If -2147221005 = Err Then response.write "����������֧�� CDONTS.NewMail �������رշ����ʼ����ܣ�<br><br>"
end if


%>
���³ɹ�
<br><br><a href=javascript:history.back()>�� ��</a>
<%
end sub


htmlend

%>