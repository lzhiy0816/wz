<!-- #include file="setup.asp" -->
<%
top
sex=server.htmlencode(Request("sex"))

if sex="" then
sql="select * from [user] where userphoto<>'' order by id Desc"
else
sql="select * from [user] where userphoto<>'' and sex='"&sex&"' order by id Desc"
end if
rs.Open sql,Conn,1
count=rs.recordcount
if count=0 then error("<li>�û���δ�ϴ���Ƭ��")
%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<script src="inc/birth.js"></script>
<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="����ҳ"></a></td>
<td vAlign="center" align="top">&nbsp;<font color=000000><img src="images/closedfold.gif">��<a href=main.asp><%=clubname%></a><br>
&nbsp;<img src="images/coner.gif"><img src="images/openfold.gif">��<a href="star.asp">�������</a></font></td>
</tr>
</table><br>

<SCRIPT>valigntop()</SCRIPT>

<table class="a2" cellSpacing="1" cellPadding="3" align="center" width=97%>
	<tr class=a1>
		<td width="50%" align="center">
		��Ṳ�� <b><%=count%></b> λ��Ա</td>
		<td width="50%" align="right">
		<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
		<option selected>���Ա�������</option>
		<option value="star.asp?sex=male">��</option>
		<option value="star.asp?sex=female">Ů</option>
				</select>&nbsp; </td>
	</tr>
	<tr>

		<td bgcolor="FFFFFF" colspan="2"  width="33%" align="center">
		<table border="0" width="100%">
			<tr>
			
<%





pagesetup=10 '�趨ÿҳ����ʾ����
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '��ҳ��
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '��ת��ָ��ҳ��


i=0
ti=0
Do While Not RS.EOF and i<pagesetup
i=i+1

select case rs("sex")
case "male"
sex1="��"
case "female"
sex1="Ů"
end select

ti=ti+1
%>
			
			
				<td width="33%">
				<table border="0" width="100%" cellspacing="1" cellpadding="4" class=a2 style=TABLE-LAYOUT:fixed>
					<tr>
						<TD align=middle bgColor=#ffffff rowspan="2">
						<a target="_blank" href="<%=rs("userphoto")%>"><IMG 
                  height=112 src="<%=rs("userphoto")%>" width=150 
                  border=0></a> </TD>
                <TD bgColor=#ffffff class=a3 ><b>�ǡ����ƣ�</b><A href="Profile.asp?username=<%=rs("username")%>"><%=rs("username")%></a><br>
				<b>��ʵ������</b><%=rs("realname")%> <br>
				<b>�ԡ�����</b><%=sex1%> <br>
				<b>Ѫ�����ͣ�</b><%=rs("blood")%><br>
				<b>�ǡ�������</b><script>document.write(astro("<%=rs("birthday")%>"));</script><br>
				<b>������Ф��</b><script>document.write(getpet("<%=rs("birthday")%>"));</script></TD></TR>
              <TR>
                <TD bgColor=#ffffff class=a4 valign="top" height="32"><b>���˼�飺</b><%=rs("personal")%></TD>
					</tr>
				</table>
</td>
<%

if ti=2 then
response.write "</tr><tr>"
ti=0
end if


RS.MoveNext
loop
RS.Close
%></table>
[
<script>
PageCount=<%=TotalPage%> //��ҳ��
topage=<%=PageCount%>   //��ǰͣ��ҳ
for (var i=1; i <= PageCount; i++) {
if (i <= topage+3 && i >= topage-3 || i==1 || i==PageCount){
if (i > topage+4 || i < topage-2 && i!=1 && i!=2 ){document.write(" ... ");}
if (topage==i){document.write(" "+ i +" ");}
else{
document.write("<a href=?topage="+i+"&sex=<%=sex%>>"+ i +"</a> ");
}
}
}
</script>
	
]<br>
				
			
			</tr>
		</table>
		</td>
	</tr>
	</table>
	<SCRIPT>valignbottom()</SCRIPT>
<%htmlend%>