<!-- #include file="setup.asp" -->
<%
id=int(Request("id"))

top

types=server.htmlencode(Request("types"))
filename=server.htmlencode(Request("filename"))

if Request("menu")="del" and membercode>3 then
conn.execute("delete from [upfile] where filename='"&filename&"' and types='"&types&"'")
on error resume next  '不支持FSO时忽略错误
set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
MyFileObject.DeleteFile""&Server.MapPath("./images/upfile/"&filename&"."&types&"")&""
error2("删除成功！")
end if

%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
<td vAlign="center" align="top">&nbsp;<font color=#000000><img src="images/closedfold.gif">　<a href=main.asp><%=clubname%></a><br>
&nbsp;<img src="images/coner.gif"><img src="images/openfold.gif">　<a href="ShowFile.asp">社区展区</a></font></td>
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
		文件展示</td>
	</tr>

	<tr>
		<td width="50%" align="center" colspan="2" bgcolor="#FFFFFF">
<br>
<img src="images/upfile/<%=rs("filename")%>.<%=rs("types")%>"><br><br></td>
	</tr>

	<tr class=a3>
		<td width="19%" align="right">作者：</td>
		<td width="79%" align="left">&nbsp;<%=rs("username")%></td>
	</tr>

	<tr class=a4>
		<td width="19%" align="right">文件名：</td>
		<td width="79%" align="left">&nbsp;<a style=text-decoration:none;color:000000 href="images/upfile/<%=rs("filename")%>.<%=rs("types")%>" target="_blank"><%=rs("filename")%>.<%=rs("types")%></a></td>
	</tr>

	<tr class=a3>
		<td width="19%" align="right">文件大小：</td>
		<td width="79%" align="left">&nbsp;<%=rs("len")%> byte</td>
	</tr>

	<tr class=a4>
		<td width="19%" align="right">说明：</td>
		<td width="79%" align="left">&nbsp;<%=rs("topic")%></td>
	</tr>

	<tr class=a3>
		<td width="19%" align="right">相关帖子：</td>
		<td width="79%" align="left">&nbsp;<a style=text-decoration:none;color:000000 onclick="min_yuzi()" target="message" href="ShowPost.asp?id=<%=rs("id")%>">查看相关讨论...</a></td>
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
if count=0 then error2("用户尚未上传作品！")

%>



<SCRIPT>valigntop()</SCRIPT>

<table class="a2" cellSpacing="1" cellPadding="3" align="center" width=97%>
	<tr class=a1>
		<td width="50%" align="center">
		展区共有 <b><%=count%></b> 幅作品</td>
		<td width="50%" align="right">
		<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
		<option selected>按文件分类浏览</option>
		<option value="ShowFile.asp?types=jpg">JPG图像</option>
		<option value="ShowFile.asp?types=gif">GIF图像</option>
		<option value="ShowFile.asp?types=swf">FLASH动画</option>
		<option value="ShowFile.asp?types=zip">ZIP压缩文件</option>
		<option value="ShowFile.asp?types=MID">MID音乐文件</option>
				</select>&nbsp; </td>
	</tr>
	<tr>

		<td bgcolor="#FFFFFF" colspan="2"  width="33%" align="center">
		<table border="0" width="100%">
			<tr>
			
<%





pagesetup=9 '设定每页的显示数量
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '总页数
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '跳转到指定页数


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
						<td class=a3>作者: <a style=text-decoration:none;color:000000 href="Profile.asp?username=<%=rs("username")%>"><%=rs("username")%></a><br>文件名: <a style=text-decoration:none;color:000000 href="images/upfile/<%=rs("filename")%>.<%=rs("types")%>" target="_blank">
						<%=rs("filename")%>.<%=rs("types")%></a><br>
						文件大小: <%=rs("len")%> byte<br>
						说明: <%=rs("topic")%><br>
						<a style=text-decoration:none;color:000000 onclick="min_yuzi()" target="message" href="ShowPost.asp?id=<%=rs("id")%>">查看相关讨论...</a>
<%if membercode > 3 then%>
　　　<a style=text-decoration:none;color:000000 onclick="{if(confirm('您确定要删除该作品?')){return true;}return false;}" href="?menu=del&filename=<%=rs("filename")%>&types=<%=rs("types")%>">删除该作品...</a>
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
PageCount=<%=TotalPage%> //总页数
topage=<%=PageCount%>   //当前停留页
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