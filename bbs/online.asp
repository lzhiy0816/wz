<!-- #include file="setup.asp" -->


<%
top
count=conn.execute("Select count(sessionid)from online")(0)
regonline=conn.execute("Select count(sessionid)from online where username<>''")(0)
toltopic=conn.execute("Select SUM(toltopic)from bbsconfig")(0)
tolretopic=conn.execute("Select SUM(tolrestore)from bbsconfig")(0)
%>
<body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="����ҳ"></a></td>
<td>��<img src="images/closedfold.gif" border="0">��<a href=main.asp><%=clubname%></a><br>
��<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">���鿴��̳״̬</td>
</tr>
</table><BR>


<SCRIPT>valigntop()</SCRIPT>
<TABLE cellspacing="1" cellpadding="4" width="97%" align="center" border="0" class="a2">
  <TBODY>
  <TR class="a1" id=TableTitleLink>
      <Td width="16%" align="center" height="10"><b><font color="FFFFFF"><a href="online.asp"> 
        �������</a></font></TH> </b> 
      <TD width="16%" align="center" height="10"><b>
		<a href="online.asp?menu=sex">�Ա�ͼ��</a></b><TD width="16%" align="center" height="10">
		<b><a href="online.asp?menu=level">����ͼ��</a></b><TD width="16%" align="center" height="10"><b><font color="FFFFFF"> <a href="online.asp?menu=cutline">����ͼ��</TH> 
        </a> </font> </b> 
      <TD width="16%" align="center" height="10"><b><font color="FFFFFF"><a href="online.asp?menu=board"> 
        ������ͼ��</a></font></TH> </b> 
      <TD width="16%" align="center" height="10"><b><font color="FFFFFF"> <a href="online.asp?menu=tolrestore">������ͼ��</TH> 
        </a> </font> </b> 
      <TR class=a3>
      <Td width="48%" align="center" height="10" colspan="3">������ <%=tolretopic%> ƪ���������� <%=toltopic%> ƪ������ <%=tolretopic-toltopic%> 
		ƪ��<Td width="48%" align="center" height="10" colspan="3">������ <%=count%> �ˡ�����ע���û� <%=regonline%> �ˣ��ÿ� <%=count-regonline%> �ˡ�</TBODY></TABLE>
<SCRIPT>valignbottom()</SCRIPT>
<br>
<%

select case Request("menu")
case ""
index
case "cutline"
cutline

case "board"
board

case "tolrestore"
tolrestore

case "sex"
sex
case "level"
level

end select

sub index



sql="select * from online order by lasttime Desc"
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




if Request.Cookies("lookip")<>"1" and membercode<4 then
ips=split(rs("ip"),".")
ip=""&ips(0)&"."&ips(1)&".*.*"
else
ip=""&rs("ip")&""
end if


if rs("username")="" then
username="�� ��"
elseif rs("eremite")=1 and membercode<4 then
username="�� ��"
else
username="<img src="&rs("userface")&" width=16 height=16> <a href=Profile.asp?username="&rs("username")&">"&rs("username")&"</a>"
end if



place2=""
if rs("act")<>"" then
place2 = "�����<a onclick=min_yuzi() target=message href="&rs("acturl")&">"&rs("act")&"</a>��"
place = "�� "&rs("bbsname")&" ��"
else
place = "�� <a href="&rs("acturl")&">"&rs("bbsname")&"</a> ��"
end if
allline=""&allline&"<TR align=middle class=a4><TD height=24>"&ip&"</TD><TD height=24>"&rs("cometime")&"</TD><TD height=24>"&username&"</TD><TD height=24>"&place&"��</TD><TD height=24>"&place2&"��</TD><TD height=24>"&rs("lasttime")&"</TD></TR>"

rs.movenext
loop
rs.close



%>

<SCRIPT>valigntop()</SCRIPT>
<TABLE cellSpacing=1 cellPadding=1 width=97% align=center border=0 class=a2>



<TR align=middle class=a1>

<TD height=23>IP��ַ</TD>
<TD height=23>��¼ʱ��</TD>
<TD height=23>�û���</TD>
<TD height=23>��ǰλ��</TD>
<TD height=23>��ǰ����</TD>
<TD height=23>�ʱ��</TD>

</TR>

<%=allline%>

</TABLE>
<SCRIPT>valignbottom()</SCRIPT>
<center>


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


end sub


sub cutline


sql="select * from bbsconfig where hide=0"
rs.Open sql,Conn



%>


<SCRIPT>valigntop()</SCRIPT>
<table class="a2" cellSpacing="1" cellPadding="4" width="97%" align="center" border="0">
		<tr>
			<td class="a1" vAlign="bottom" align="middle" height="20">��̳����</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">ͼ�α���</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">��������</td>
		</tr>
		


<%

i=0
Do While Not RS.EOF
i=i+1
onlinemany=conn.execute("Select count(sessionid)from online  where forumid="&rs("id")&"")(0)

%>

		<tr class="a3">
			<td width="21%" height="2" align="center">
			<a href="ShowForum.asp?forumid=<%=rs("id")%>"><%=rs("bbsname")%></a></td>
			<td width="65%" height="2"><IMG height=8 src="images/bar<%=i%>.gif" width="<%=onlinemany/count*100%>%"></td>
			<td align="center" width="12%" height="2"><%=onlinemany%></td></tr>
			
			
			<%
			
			
if i=10 then
i=0
end if
			
RS.MoveNext
loop
RS.Close   


%></table>

<SCRIPT>valignbottom()</SCRIPT>
<%


end sub

sub board

sql="select * from bbsconfig where hide=0 order by toltopic Desc"
rs.Open sql,Conn
%>
<SCRIPT>valigntop()</SCRIPT>
<table class="a2" cellSpacing="1" cellPadding="4" width="97%" align="center" border="0">
		<tr>
			<td class="a1" vAlign="bottom" align="middle" height="20">��̳����</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">ͼ�α���</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">������</td>
		</tr>
		


<%
i=0
Do While Not RS.EOF
i=i+1
%>

		<tr class="a3">
			<td width="21%" height="2" align="center"><a href="ShowForum.asp?forumid=<%=rs("id")%>"><%=rs("bbsname")%></a></td>
			<td width="65%" height="2"><IMG height=8 src="images/bar<%=i%>.gif" width="<%=rs("toltopic")/toltopic*100%>%"></td>
			<td align="center" width="12%" height="2"><%=rs("toltopic")%></td></tr><%
			
if i=9 then
i=0
end if

			
RS.MoveNext
loop
RS.Close   


%></table>
<SCRIPT>valignbottom()</SCRIPT>
<%
end sub

sub tolrestore

sql="select * from bbsconfig where hide=0 order by tolrestore Desc"
rs.Open sql,Conn
%>
<SCRIPT>valigntop()</SCRIPT>
<table class="a2" cellSpacing="1" cellPadding="4" width="97%" align="center" border="0">
		<tr>
			<td class="a1" vAlign="bottom" align="middle" height="20">��̳����</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">ͼ�α���</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">������</td>
		</tr>
		


<%
i=0
Do While Not RS.EOF
i=i+1
%>

		<tr class="a3">
			<td width="21%" height="2" align="center"><a href="ShowForum.asp?forumid=<%=rs("id")%>"><%=rs("bbsname")%></a></td>
			<td width="65%" height="2"><IMG height=8 src="images/bar<%=i%>.gif" width="<%=rs("tolrestore")/tolretopic*100%>%"></td>
			<td align="center" width="12%" height="2"><%=rs("tolrestore")%></td></tr><%
			
if i=9 then
i=0
end if

			
RS.MoveNext
loop
RS.Close   


%></table>
<SCRIPT>valignbottom()</SCRIPT>
<%
end sub





sub sex
count=conn.execute("Select count(id)from [user]")(0)
male=conn.execute("Select count(id)from [user] where sex='male'")(0)
female=conn.execute("Select count(id)from [user] where sex='female'")(0)

%>

<SCRIPT>valigntop()</SCRIPT>
<table class="a2" cellSpacing="1" cellPadding="4" width="97%" align="center" border="0">
		<tr>
			<td class="a1" vAlign="bottom" align="middle" height="20">�Ա�</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">ͼ�α���</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">����</td>
		</tr>
		



		<tr class="a3">
			<td width="10%" height="2" align="center">
			<img src=images/male.gif></td>
			<td width="75%" height="2"><IMG height=8 src="images/bar3.gif" width="<%=male/count*100%>%"></td>
			<td align="center" width="12%" height="2"><%=male%></td></tr>		
			
		<tr class="a3">
			<td width="10%" height="2" align="center">
			<img src=images/female.gif></td>
			<td width="75%" height="2"><IMG height=8 src="images/bar1.gif" width="<%=female/count*100%>%"></td>
			<td align="center" width="12%" height="2"><%=female%></td></tr>		
			
		<tr class="a3">
			<td width="10%" height="2" align="center">δ֪</td>
			<td width="75%" height="2"><IMG height=8 src="images/bar2.gif" width="<%=(count-male-female)/count*100%>%"></td>
			<td align="center" width="12%" height="2"><%=count-male-female%></td></tr>		
	

</table>
<SCRIPT>valignbottom()</SCRIPT>
<%
end sub


sub level
count=conn.execute("Select count(id)from [user]")(0)
membercode1=conn.execute("Select count(id)from [user] where membercode=1")(0)
membercode2=conn.execute("Select count(id)from [user] where membercode=2")(0)
membercode4=conn.execute("Select count(id)from [user] where membercode=4")(0)
membercode5=conn.execute("Select count(id)from [user] where membercode=5")(0)

%>
<SCRIPT>valigntop()</SCRIPT>
<table class="a2" cellSpacing="1" cellPadding="4" width="97%" align="center" border="0">
		<tr>
			<td class="a1" vAlign="bottom" align="middle" height="20">��Ա����</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">ͼ�α���</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">����</td>
		</tr>

		<tr class="a3">
			<td width="10%" height="2" align="center">��ͨ��Ա</td>
			<td width="75%" height="2"><IMG height=8 src="images/bar1.gif" width="<%=membercode1/count*100%>%"></td>
			<td align="center" width="12%" height="2"><%=membercode1%></td></tr>		
			
		<tr class="a3">
			<td width="10%" height="2" align="center">�����α�</td>
			<td width="75%" height="2"><IMG height=8 src="images/bar2.gif" width="<%=membercode2/count*100%>%"></td>
			<td align="center" width="12%" height="2"><%=membercode2%></td></tr>		
			
		<tr class="a3">
			<td width="10%" height="2" align="center">�� �� Ա</td>
			<td width="75%" height="2"><IMG height=8 src="images/bar4.gif" width="<%=membercode4/count*100%>%"></td>
			<td align="center" width="12%" height="2"><%=membercode4%></td></tr>		
	
		<tr class="a3">
			<td width="10%" height="2" align="center">��������</td>
			<td width="75%" height="2"><IMG height=8 src="images/bar5.gif" width="<%=membercode5/count*100%>%"></td>
			<td align="center" width="12%" height="2"><%=membercode5%></td></tr></table>
<SCRIPT>valignbottom()</SCRIPT>
<%
end sub



htmlend
%>