<!-- #include file="setup.asp" -->
<%
id=int(Request("id"))

top

types=server.htmlencode(Request("types"))
filename=server.htmlencode(Request("filename"))

if Request("menu")="del" and membercode>3 then
conn.execute("delete from [upfile] where filename='"&filename&"' and types='"&types&"'")
on error resume next  '��֧��FSOʱ���Դ���
set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
MyFileObject.DeleteFile""&Server.MapPath("./images/upfile/"&filename&"."&types&"")&""
error2("ɾ���ɹ���")
end if

%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="����ҳ"></a></td>
<td vAlign="center" align="top">&nbsp;<font color=#000000><img src="images/closedfold.gif">��<a href=main.asp><%=clubname%></a><br>
&nbsp;<img src="images/coner.gif"><img src="images/openfold.gif">��<a href="ShowFile.asp">����չ��</a></font></td>
</tr>
</table><br>
<%
if Request("menu")="show" then

sql="select * from [upfile] where filename='"&filename&"'"
rs.Open sql,Conn
%>
<SCRIPT>valigntop()</SCRIPT>
<table class="a2" cellSpacing="1" cellPadding="3" align="center" width=97%>
	<tr class=a1>
		<td width="50%" align="center" colspan="2">
		�ļ�չʾ</td>
	</tr>

	<tr>
		<td width="50%" align="center" colspan="2" bgcolor="#FFFFFF">
<br>
<img src="images/upfile/<%=rs("filename")%>.<%=rs("types")%>"><br><br></td>
	</tr>

	<tr class=a3>
		<td width="19%" align="right">���ߣ�</td>
		<td width="79%" align="left">&nbsp;<%=rs("username")%></td>
	</tr>

	<tr class=a4>
		<td width="19%" align="right">�ļ�����</td>
		<td width="79%" align="left">&nbsp;<a style=text-decoration:none;color:000000 href="images/upfile/<%=rs("filename")%>.<%=rs("types")%>" target="_blank"><%=rs("filename")%>.<%=rs("types")%></a></td>
	</tr>

	<tr class=a3>
		<td width="19%" align="right">�ļ���С��</td>
		<td width="79%" align="left">&nbsp;<%=rs("len")%> byte</td>
	</tr>

	<tr class=a4>
		<td width="19%" align="right">˵����</td>
		<td width="79%" align="left">&nbsp;<%=rs("topic")%></td>
	</tr>

	<tr class=a3>
		<td width="19%" align="right">������ӣ�</td>
		<td width="79%" align="left">&nbsp;<a style=text-decoration:none;color:000000 onclick="min_yuzi()" target="message" href="ShowPost.asp?id=<%=rs("id")%>">�鿴�������...</a></td>
	</tr>
</table>
<SCRIPT>valignbottom()</SCRIPT>
<center>
<br><a href="javascript:history.back()">
<img src="images/plus/back.gif" border="0"></a>
<%
RS.Close
htmlend
end if





if types="" then
sql="select * from [upfile] order by id Desc"
else
sql="select * from [upfile] where types ='"&types&"' order by id Desc"
end if

rs.Open sql,Conn,1
count=rs.recordcount
if count=0 then error2("�û���δ�ϴ���Ʒ��")

%>



<SCRIPT>valigntop()</SCRIPT>

<table class="a2" cellSpacing="1" cellPadding="3" align="center" width=97%>
	<tr class=a1>
		<td width="50%" align="center">
		չ������ <b><%=count%></b> ����Ʒ</td>
		<td width="50%" align="right">
		<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
		<option selected>���ļ��������</option>
		<option value="ShowFile.asp?types=jpg">JPGͼ��</option>
		<option value="ShowFile.asp?types=gif">GIFͼ��</option>
		<option value="ShowFile.asp?types=swf">FLASH����</option>
		<option value="ShowFile.asp?types=zip">ZIPѹ���ļ�</option>
		<option value="ShowFile.asp?types=MID">MID�����ļ�</option>
				</select>&nbsp; </td>
	</tr>
	<tr>

		<td bgcolor="#FFFFFF" colspan="2"  width="33%" align="center">
		<table border="0" width="100%">
			<tr>
			
<%





pagesetup=9 '�趨ÿҳ����ʾ����
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
			
			
				<td width="33%">
				<table border="0" width="100%" cellspacing="1" cellpadding="4" class=a2>
					<tr>
						<td align="center" height="150" bgcolor="#FFFFFF">
<%
if LCase(rs("types")) ="gif" or LCase(rs("types")) ="jpg" then
%>
<a href="?menu=show&filename=<%=rs("filename")%>">
<img src="images/upfile/<%=rs("filename")%>.<%=rs("types")%>" width="150" height="112" border="0"></a>
<%	
else
%>
<a href="images/upfile/<%=rs("filename")%>.<%=rs("types")%>"><img src="images/plus/info.gif" border="0"></a>
<%		
end if
%>
</td>
					</tr>
					<tr>
						<td class=a3>����: <a style=text-decoration:none;color:000000 href="Profile.asp?username=<%=rs("username")%>"><%=rs("username")%></a><br>�ļ���: <a style=text-decoration:none;color:000000 href="images/upfile/<%=rs("filename")%>.<%=rs("types")%>" target="_blank">
						<%=rs("filename")%>.<%=rs("types")%></a><br>
						�ļ���С: <%=rs("len")%> byte<br>
						˵��: <%=rs("topic")%><br>
						<a style=text-decoration:none;color:000000 onclick="min_yuzi()" target="message" href="ShowPost.asp?id=<%=rs("id")%>">�鿴�������...</a>
<%if membercode > 3 then%>
������<a style=text-decoration:none;color:000000 onclick="{if(confirm('��ȷ��Ҫɾ������Ʒ?')){return true;}return false;}" href="?menu=del&filename=<%=rs("filename")%>&types=<%=rs("types")%>">ɾ������Ʒ...</a>
<%end if%>
						
						</td>
					</tr>
				</table>

</td>
<%
if i=3 or i=6 then
response.write "</tr><tr>"
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
document.write("<a href=?topage="+i+"&types=<%=Request("types")%>>"+ i +"</a> ");
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