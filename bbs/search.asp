<!-- #include file="setup.asp" -->
<%
top
if Request.Cookies("username")=empty then error("<li>����δ<a href=login.asp>��¼</a>����")

%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<title>��������</title>
<script>
var i=0;
function formCheck(){i++;if (i>1) {document.form.submit1.disabled = true;}return true;}
</script>
<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%">
<a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="����ҳ"></a></td>
<td vAlign="center" align="top">��<img src="images/closedfold.gif" border="0">��<a href=main.asp><%=clubname%></a><br>
��<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">����������</td>
</tr>
</table>




<form method="post" action="searchok.asp" name=form>

<center>

<SCRIPT>valigntop()</SCRIPT>
<table height="207" cellSpacing="1" cellPadding="0" width="97%" class=a2 border="0">
<tr>
<td colSpan="2" height=25 class=a1 align="center">������Ҫ�����Ĺؼ���</td>
</tr>
<tr>
<td vAlign="top" bgColor="#FFFFFF" colSpan="2" height="8">
<p align="center"><input size="40" name="content"></td>
</tr>
<tr>
<td class=a1 colSpan="2" height=25 align="center">����ѡ��</td>
</tr>
<tr>
<td width="30%" height="24" bgcolor="FFFFFF">
<p align="right"><font style="FONT-SIZE: 9pt"><label for=search>��������</label></font><input type="radio" value="author" name="search" id=search></td>
<td height=25 bgcolor="FFFFFF">
&nbsp;<select size="1" name="searchxm">
<option selected value="username">������������</option>
<option value="lastname">�������ظ�����</option>
</select></td>
</tr>
<tr>
<td width="30%" height="21" bgcolor="FFFFFF">
<p align="right"><span style="FONT-SIZE: 9pt"><label for=search_1>�ؼ�������</label></span><input type="radio" value="key" name="search" checked id=search_1></td>
<td height=25 bgcolor="FFFFFF">
&nbsp;<select size="1" name="searchxm2">
<option selected value="topic">�������������ؼ���</option>

</select></td>
</tr>
<tr>
<td width="30%" height="23" bgcolor="FFFFFF">
<p align="right"><font style="FONT-SIZE: 9pt" color="000000">���ڷ�Χ</font></td>
<td height=25 bgcolor="FFFFFF">
&nbsp;<select size="1" name="TimeLimit">
<option value="">��������</option>
<option value="1">��������</option>
<option value="5" selected>5������</option>
<option value="10">10������</option>
<option value="30">30������</option>
</select></td>
</tr>
<tr>
<td width="30%" height="26" align="right" bgcolor="FFFFFF"><font style="FONT-SIZE: 9pt" color="000000">��ѡ��Ҫ��������̳</font></td>
<td height="26" bgcolor="FFFFFF">


&nbsp;<select name="forumid">
<option value="" selected>ȫ����̳</option>

<%

sql="select id,bbsname,classid from bbsconfig where hide=0 order by classid,id"
rs.Open sql,Conn
do while not rs.eof
Classid=Trim(Rs("classid"))
		if TClass <> Classid Then
		Response.write "<option style=BACKGROUND-COLOR:ECF5FF value=''>�� "&Conn.Execute("Select classname From class where id="&Classid)(0)&"</option>"
		TClass = Classid
		end if
		
if Request("forumid")=""&rs("id")&"" then
selected=" selected"
else
selected=""
end if
		
	Response.write "<option value="&rs("id")&""&selected&">������"&rs("bbsname")&"��</option>"
	rs.movenext
loop
rs.close
%>
</select> <input type="submit" name=submit1 value="��ʼ����" onclick="return formCheck()"></td>
</tr>

</table>
<SCRIPT>valignbottom()</SCRIPT> &nbsp;</center>







<%


htmlend
%>