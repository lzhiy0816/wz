<!-- #include file="setup.asp" -->
<%top

if Request.Cookies("username")=empty then error("<li>����δ<a href=login.asp>��¼</a>����")
title=server.htmlencode(Request("title"))
content=replace(server.htmlencode(Request.Form("content")), "'", "&#39;")
content=replace(content,vbCrlf,"<br>")
content=replace(content,"\","&#92;")

lookdate=server.htmlencode(Request("lookdate"))
adddate=server.htmlencode(Request("adddate"))
hide=int(Request("hide"))
id=int(Request("id"))

if Request("menu")="add" then
if title=empty then error("<li>��û�������ռ�����")

if content=empty then error("<li>��û�������ռ�����")


sql="insert into calendar(title,username,content,hide,adddate) values ('"&title&"','"&Request.Cookies("username")&"','"&content&"','"&hide&"','"&adddate&"')"
conn.Execute(SQL)


message="<li>����ռǳɹ�<li><a href=calendar.asp?menu=show&lookdate="&adddate&">�����ռ�</a><li><a href=calendar.asp>��������</a><li><a href=main.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=calendar.asp?menu=show&lookdate="&adddate&">")

elseif Request("menu")="del" then
if membercode > 3 then
conn.execute("delete from [calendar] where id="&id&"")
else
conn.execute("delete from [calendar] where id="&id&" and username='"&Request.Cookies("username")&"'")
end if

message="<li>ɾ���ɹ�<li><a href=calendar.asp>��������</a><li><a href=main.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=calendar.asp>")

elseif Request("menu")="editok" then
conn.execute("update [calendar] set title='"&title&"',content='"&content&"',hide='"&hide&"' where id="&id&" and username='"&Request.Cookies("username")&"'")

message="<li>�༭�ռǳɹ�<li><a href=calendar.asp?menu=show&lookdate="&adddate&">�����ռ�</a><li><a href=calendar.asp>��������</a><li><a href=main.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=calendar.asp?menu=show&lookdate="&adddate&">")



end if

%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>


<script>function checkclick(msg){if(confirm(msg)){event.returnValue=true;}else{event.returnValue=false;}}</script>

<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="����ҳ"></a></td>
<td>��<img src="images/closedfold.gif" border="0">��<a href=main.asp><%=clubname%></a><br>
��<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">��<a href="calendar.asp">�ҵ��ռ�</a></td>
</tr>
</table><BR>




<%
if Request("menu")="show" then
%>
<script>var badwords= "<%=badwords%>";</script>
<script src="inc/ybb.js"></script>
<br>
<SCRIPT>valigntop()</SCRIPT>
<div align="center">
<table cellspacing="1" cellpadding="6" width="97%" border="0" class="a2"><tr class="a1">
	<td><b>���¹����ռ�</b></td>
	<td align="center" width="75%"><b><%=lookdate%> �ռ�</b></td></tr><tr class="a3">
	<td valign="top">
	
	
<%
sql="select top 15 * from calendar where hide=0 order by id Desc"

rs.Open sql,Conn
Do While Not RS.EOF
%>
<a href="calendar.asp?menu=show&lookdate=<%=rs("adddate")%>"><%=rs("title")%></a><br>
<%
RS.MoveNext
loop
RS.Close   
%>
</td>
	<td align="center" valign="top">


<%
sql="select * from calendar where (hide=0 or username='"&Request.Cookies("username")&"') and adddate='"&lookdate&"' order by id Desc"

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
%>
<table cellspacing="1" cellpadding="1" width="97%" border="0" class="a4">

			<tr>
				<td align="center" colspan="2"><b><%=rs("title")%></b></td>
			</tr>
			<tr>
				<td colspan="2">
				<br>
				
<SCRIPT>
document.write(ybbcode("<%=rs("content")%>"))
</SCRIPT>
				
				<br><br>
</td>
			</tr>
			<tr>
				<td>
				<a href="calendar.asp?menu=edit&id=<%=rs("id")%>"><img src="images/edit.gif" border="0"></a>
				<a href="calendar.asp?menu=del&id=<%=rs("id")%>" onclick="checkclick('��ȷ��Ҫɾ�������ռ�?')">
				<img src="images/del.gif" border="0"></a></td>
				<td align="right">
				���ߣ�<%=rs("username")%>&nbsp; ʱ�䣺<%=rs("addtime")%></td>
			</tr>
		</table>
		<hr size=0>
<%

RS.MoveNext
loop
RS.Close   
%>



	</td></tr>
	

</table>

<SCRIPT>valignbottom()</SCRIPT>
<b>[ <script>
PageCount=<%=TotalPage%> //��ҳ��
topage=<%=PageCount%>   //��ǰͣ��ҳ
for (var i=1; i <= PageCount; i++) {
if (i <= topage+3 && i >= topage-3 || i==1 || i==PageCount){
if (i > topage+4 || i < topage-2 && i!=1 && i!=2 ){document.write(" ... ");}
if (topage==i){document.write(" "+ i +" ");}
else{
document.write("<a href=?menu=show&lookdate=<%=lookdate%>&topage="+i+">"+ i +"</a> ");
}
}
}
</script> ]</b>
	<br>
<SCRIPT>valigntop()</SCRIPT>
<table cellspacing="1" cellpadding="6" width="97%" border="0" class="a2" align="center">
	
	<tr class="a1">
	<td>
		<b>���ڸ����ռ�</b></td>
	<td align="center" valign="top" width="75%">


<b>����ռ�</b></td>
	</tr><tr class="a3">
	<td valign="top">
<%
sql="select top 10 * from calendar where username='"&Request.Cookies("username")&"' order by id Desc"
rs.Open sql,Conn
Do While Not RS.EOF
%>
<a href="calendar.asp?menu=show&lookdate=<%=rs("adddate")%>"><%=rs("title")%></a><br>
<%
RS.MoveNext
loop
RS.Close   
%>
</td>
	<td align="center" valign="top">




<table cellspacing="0" cellpadding="4" width="97%" align="center" border="0" class="a2">
    
<form method=post name=form action=calendar.asp?menu=add>
<input type=hidden name=adddate value="<%=lookdate%>">
		<tr class=a3>
			<td width="14%"><b>�ռǱ��⣺</b></td>
		<td width="83%">
		<input maxLength="30" size="54" name="title"> <select name="hide" size="1">
<option value="0">����</option>
<option value="1" selected>����</option>
</select></td>
		</tr>
		<tr class=a3>
			<td width="14%"><b>�ռ����ݣ�</b></td>
			<td width="83%">
			<textarea name="content" rows="6" style="width:95%" cols="1" ></textarea>
		</td>
		</tr>
		<tr class=a3>
			<td colspan="2" align="center">��<input tabIndex="4" type="submit" value=" �� �� " name="submit1">&nbsp;
			<input onclick="{if(confirm('�������Ҫ���ȫ�������ݣ���ȷ��Ҫ�����?')){return true;}return false;}" type="reset" value=" �� д "></td>
		</tr></form>
	</table>



	</td>
	</tr></table><SCRIPT>valignbottom()</SCRIPT>
<br>



<%
htmlend

elseif Request("menu")="edit" then
sql="select * from calendar where id="&id&""
rs.Open sql,Conn
if ""&rs("username")&""<>""&Request.Cookies("username")&"" then error2("���ռǷ�����д��")

%>

	<br>
<SCRIPT>valigntop()</SCRIPT>
<table cellspacing="1" cellpadding="6" width="97%" border="0" class="a2" align="center">
	
	<tr class="a1">
	<td align="center" valign="top" width="75%">


<b>�༭�ռ�</b></td>
	</tr><tr class="a3">
	<td align="center" valign="top">




<table cellspacing="0" cellpadding="4" width="97%" align="center" border="0" class="a2">
    
<form method=post name=form action=calendar.asp?menu=editok>
<input type=hidden name=adddate value="<%=rs("adddate")%>">
<input type=hidden name=id value="<%=rs("id")%>">
		<tr class=a3>
			<td width="14%"><b>�ռǱ��⣺</b></td>
		<td width="83%">
		<input maxLength="30" size="70" name="title" value="<%=rs("title")%>"> <select name="hide" size="1">
<option value="0" <%if rs("hide")=0 then%>selected<%end if%>>����</option>
<option value="1" <%if rs("hide")=1 then%>selected<%end if%>>����</option>
</select></td>
		</tr>
		<tr class=a3>
			<td width="14%"><b>�ռ����ݣ�</b></td>
			<td width="83%">
			<textarea name="content" rows="6" style="width:95%" cols="1" ><%=rs("content")%></textarea>
		</td>
		</tr>
		<tr class=a3>
			<td colspan="2" align="center">��<input tabIndex="4" type="submit" value=" �� �� " name="submit1">&nbsp;
			<input onclick="{if(confirm('�������Ҫ���ȫ�������ݣ���ȷ��Ҫ�����?')){return true;}return false;}" type="reset" value=" �� д "></td>
		</tr></form>
	</table>



	</td>
	</tr></table><SCRIPT>valignbottom()</SCRIPT>
<br>
<Script>
document.form.content.value = unybbcode(document.form.content.value);
document.form.content.focus();
function unybbcode(temp) {
temp = temp.replace(/<br>/ig,"\n");
return (temp);
}
</Script>
<%
RS.Close
htmlend
end if

%>













<script src="inc/calendar.js"></script>

<SCRIPT language=VBScript>
<!--
'===== ������ʱ��
Function TimeAdd(UTC,T)
   Dim PlusMinus, DST, y
   If Left(T,1)="-" Then PlusMinus = -1 Else PlusMinus = 1
   UTC=Right(UTC,Len(UTC)-5)
   UTC=Left(UTC,Len(UTC)-4)
   y = Year(UTC)
   TimeAdd=DateAdd("n", (Cint(Mid(T,2,2))*60 + Cint(Mid(T,4,2))) * PlusMinus, UTC)
   '�����չ��Լ�ڼ�: 4�µ�һ������00:00 �� 10������һ��������00:00
   If Mid(T,6,1)="*" And DateSerial(y,4,(9 - Weekday(DateSerial(y,4,1)) mod 7) ) <= TimeAdd And DateSerial(y,10,31 - Weekday(DateSerial(y,10,31))) >= TimeAdd Then
      TimeAdd=CStr(DateAdd("h", 1, TimeAdd))
      tSave.innerHTML = "R"
   Else
      tSave.innerHTML = ""
   End If
   TimeAdd = CStr(TimeAdd)
End Function
'-->
</SCRIPT>




<BODY onload=initial()>



<DIV id=detail 
style="LEFT: 12px; WIDTH: 200px; POSITION: absolute; TOP: 0px; HEIGHT: 19px"></DIV>



<FORM name=CLD>


<SCRIPT>valigntop()</SCRIPT>

<TABLE cellSpacing=1 cellPadding=0 width=97% border=0 class=a2 align="center">
  <TBODY>

  <TR>
    <TD align=middle width=240 class=a4>
      <SCRIPT language=JavaScript>
var enabled = 0; today = new Date();
var day; var date;
if(today.getDay()==0) day = "������"
if(today.getDay()==1) day = "����һ"
if(today.getDay()==2) day = "���ڶ�"
if(today.getDay()==3) day = "������"
if(today.getDay()==4) day = "������"
if(today.getDay()==5) day = "������"
if(today.getDay()==6) day = "������"
document.fgColor = "black";
date = " ��Ԫ " + (today.getYear()) + " �� " +
(today.getMonth() + 1 ) + "�� " + today.getDate() + "�� " +
day +"";
document.write(date)
    </SCRIPT>
      </FONT></TD>
    <TD align=middle width=418 class=a3 rowspan="2">

      <TABLE border=0>
        <TBODY>
        <TR>
          <TD class=a1 colSpan=7 align="center" height="25">��Ԫ<SELECT style="FONT-SIZE: 9pt" 
            onchange=changeCld() name=SY> 
              <SCRIPT language=JavaScript><!--        
            for(i=1900;i<2050;i++) document.write('<option>'+i)        
            //--></SCRIPT>
            </SELECT>��<SELECT style="FONT-SIZE: 9pt" onchange=changeCld() 
            name=SM> 
              <SCRIPT language=JavaScript><!--        
            for(i=1;i<13;i++) document.write('<option>'+i)        
            //--></SCRIPT>
            </SELECT>��</FONT>&nbsp;&nbsp;  <FONT id=GZ></FONT><BR></TD></TR>
        <TR align=middle class=a4>
          <TD width=54 height="20">��</TD>
          <TD width=54 height="20">һ</TD>
          <TD width=54 height="20">��</TD>
          <TD width=50 height="20">��</TD>
          <TD width=54 height="20">��</TD>
          <TD width=54 height="20">��</TD>
          <TD width=54 height="20">��</TD></TR>
        <SCRIPT language=JavaScript><!--        
            var gNum        
            for(i=0;i<6;i++) {        
               document.write('<tr align=center>')        
               for(j=0;j<7;j++) {        
                  gNum = i*7+j        
                  document.write('<td ><font id="SD' + gNum +'" size=5 face="Arial Black"></font><br><font id="LD' + gNum + '" size=2 style="font-size:9pt"> </font></td>')
               }        
               document.write('</tr>')        
            }        
            //--></SCRIPT>
        </TBODY></TABLE></TD>
    <TD align=middle  class=a4 height="25">
    
    ��<font color="#FF0000"> <%=conn.execute("Select count(id)from [calendar]")(0)%> 
	</font>ƪ�ռ�</TD></TR>

  <TR>
    <TD vAlign=top align=middle width=240 class=a3><BR><FONT 
      id=Clock face=Arial color=000080 size=4 align="center"></FONT>
      <P><!--ʱ�� *��ʾ�Զ�����Ϊ�չ��Լʱ��--><FONT style="FONT-SIZE: 9pt" size=2><SELECT 
      style="FONT-SIZE: 9pt" onchange=changeTZ() name=TZ> <OPTION 
        value="+0800 ���������졢������" selected>�й�</OPTION> <OPTION 
        value="-1100 ��;������Ħ��Ⱥ��">��Ħ��</OPTION> <OPTION 
        value="-1000 ������">������</OPTION> <OPTION value=-0900*����˹��>����˹��</OPTION> 
        <OPTION value=-0800*̫ƽ��ʱ��(����)��������>̫ƽ��</OPTION> <OPTION 
        value="-0700 ����ɣ��">����ɽ��</OPTION> <OPTION 
        value=-0700*ɽ��ʱ��(����)>����ɽ��</OPTION> <OPTION 
        value=-0600*��������(���ô�)>���ô��в�</OPTION> <OPTION 
        value=-0600*ī�����С�����������>ī����</OPTION> <OPTION 
        value=-0600*�в�ʱ��(����)>�����в�</OPTION> <OPTION 
        value=-0500*���������>������̫ƽ��</OPTION> <OPTION 
        value=-0500*����ʱ��(����)>���Ӷ���</OPTION> <OPTION 
        value=-0500*ӡ�ڰ���(����)>����</OPTION> <OPTION 
        value=-0400*������˹��������>����������</OPTION> <OPTION 
        value="-0400*������ʱ�� ���ô�)">������</OPTION> <OPTION 
        value="-0330 �µ�(���ô󶫰�)">Ŧ����</OPTION> <OPTION 
        value="-0300 ��������">��������</OPTION> <OPTION 
        value="-0300 ��³ŵ˹����˹�����γ�">�����޶���</OPTION> <OPTION 
        value=-0200*�������в�>�������в�</OPTION> <OPTION 
        value=-0100*���ٶ�Ⱥ����ά�½�Ⱥ��>���ٶ�</OPTION> <OPTION 
        value="+0000 ��������ʱ�䡢�����֡����������׶�">Ӣ������</OPTION> <OPTION 
        value="+0000 Ī��ά��(��������)������������">�������α�׼</OPTION> <OPTION 
        value="+0100 ���衢�����">����</OPTION> <OPTION 
        value="+0100 ������, ��ɳ, ������˹">��ŷ</OPTION> <OPTION 
        value="+0100 ���֡�˹�¸��Ħ��������������³������άҲ��">��ŷ</OPTION> <OPTION 
        value="+0200 ��ɫ��">��ɫ��</OPTION> <OPTION value=+0200*��ŷ>��ŷ</OPTION> 
        <OPTION value=+0200*����>����</OPTION> <OPTION 
        value=+0200*�ŵ䡢�ն���������˹̹��>GFT</OPTION> <OPTION 
        value=+0200*�����ס�Ƥ����>�Ϸ�</OPTION> <OPTION 
        value=+0300*�͸������ء������(����)�����ŵ�(ɳ�ڵ�)>ɳ�ڵذ�����</OPTION> <OPTION 
        value=+0300*Ī˹�ơ�ʥ�˵ñ�����ռ�����߸����>����˹</OPTION> <OPTION 
        value=+0330*�º���>����</OPTION> <OPTION 
        value=+0400*�������(��������)��Ī˹����������˹(�����ǹ���)>������</OPTION> <OPTION 
        value=+0430*������>������</OPTION> <OPTION 
        value="+0500 ��˹����͵¡������桢�������ֱ��������">����</OPTION> <OPTION 
        value="+0530 ���򡢼Ӷ�������˹���µ��������">ӡ��</OPTION> <OPTION 
        value="+0600 �����ᡢ�￨">����</OPTION> <OPTION 
        value="+0700 ���ȡ��ǼӴ��־����">����</OPTION> <OPTION 
        value="-1200 ���������ˡ��ϼ���">���ʻ�����</OPTION> <OPTION 
        value="+0800 ̨�塢��ۡ��¼���">̨��</OPTION> <OPTION 
        value="+0900 ���������桢���ϡ����ǡ��ǿ���(����������)">����</OPTION> <OPTION 
        value="+0930 �����">�����в�</OPTION> <OPTION 
        value="+1000 ����˹����ī������ϯ����">ϯ����</OPTION> <OPTION 
        value="+1000 ������">��˹÷����</OPTION> <OPTION 
        value="+1000 �ص���Ī˹�ȸۡ�������">��̫ƽ��</OPTION> <OPTION 
        value=+1100*����������Ⱥ�����¿��׶�����>̫ƽ���в�</OPTION> <OPTION 
        value="+1200 ����١��¿���">Ŧ����</OPTION> <OPTION 
        value="+1200 쳼á��ϼ�������Ъ��Ⱥ��">쳼�</OPTION></SELECT>ʱ��</FONT> <FONT id=tSave 
      style="FONT-SIZE: 18pt; COLOR: red; FONT-FAMILY: Wingdings"></FONT><BR><BR><FONT 
      style="FONT-SIZE: 120pt; COLOR: #13b0f4; FONT-FAMILY: Webdings">&ucirc;</FONT><BR><BR><FONT id=CITY></FONT></TD>
    <TD align=middle  class=a3>
    
    <BUTTON style="FONT-SIZE: 9pt" 
      onclick="pushBtm('YU')">���</BUTTON><br>
	<BUTTON style="FONT-SIZE: 9pt" 
      onclick="pushBtm('YD')">���</BUTTON> 
      <br>
��<p> 
      <BUTTON style="FONT-SIZE: 9pt" 
      onclick="pushBtm('MU')">�¡�</BUTTON><br>
	<BUTTON style="FONT-SIZE: 9pt" 
      onclick="pushBtm('MD')">�¡�</BUTTON> 
      <br>
��</p>
	<p> 
      <BUTTON style="FONT-SIZE: 9pt" onclick="pushBtm('')">����</BUTTON> 
	<p>��</TD></TR>

</TABLE>
<SCRIPT>valignbottom()</SCRIPT>

    
</form>



    
<%
htmlend
%>