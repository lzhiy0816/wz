<!-- #include file="setup.asp" -->
<%top

if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")
title=server.htmlencode(Request("title"))
content=replace(server.htmlencode(Request.Form("content")), "'", "&#39;")
content=replace(content,vbCrlf,"<br>")
content=replace(content,"\","&#92;")

lookdate=server.htmlencode(Request("lookdate"))
adddate=server.htmlencode(Request("adddate"))
hide=int(Request("hide"))
id=int(Request("id"))

if Request("menu")="add" then
if title=empty then error("<li>您没有输入日记主题")

if content=empty then error("<li>您没有输入日记内容")


sql="insert into calendar(title,username,content,hide,adddate) values ('"&title&"','"&Request.Cookies("username")&"','"&content&"','"&hide&"','"&adddate&"')"
conn.Execute(SQL)


message="<li>添加日记成功<li><a href=calendar.asp?menu=show&lookdate="&adddate&">返回日记</a><li><a href=calendar.asp>返回日历</a><li><a href=main.asp>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=calendar.asp?menu=show&lookdate="&adddate&">")

elseif Request("menu")="del" then
if membercode > 3 then
conn.execute("delete from [calendar] where id="&id&"")
else
conn.execute("delete from [calendar] where id="&id&" and username='"&Request.Cookies("username")&"'")
end if

message="<li>删除成功<li><a href=calendar.asp>返回日历</a><li><a href=main.asp>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=calendar.asp>")

elseif Request("menu")="editok" then
conn.execute("update [calendar] set title='"&title&"',content='"&content&"',hide='"&hide&"' where id="&id&" and username='"&Request.Cookies("username")&"'")

message="<li>编辑日记成功<li><a href=calendar.asp?menu=show&lookdate="&adddate&">返回日记</a><li><a href=calendar.asp>返回日历</a><li><a href=main.asp>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=calendar.asp?menu=show&lookdate="&adddate&">")



end if

%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>


<script>function checkclick(msg){if(confirm(msg)){event.returnValue=true;}else{event.returnValue=false;}}</script>

<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
<td>　<img src="images/closedfold.gif" border="0">　<a href=main.asp><%=clubname%></a><br>
　<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">　<a href="calendar.asp">我的日记</a></td>
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
	<td><b>最新公开日记</b></td>
	<td align="center" width="75%"><b><%=lookdate%> 日记</b></td></tr><tr class="a3">
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
pagesetup=10 '设定每页的显示数量
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
				<a href="calendar.asp?menu=del&id=<%=rs("id")%>" onclick="checkclick('您确定要删除此条日记?')">
				<img src="images/del.gif" border="0"></a></td>
				<td align="right">
				作者：<%=rs("username")%>&nbsp; 时间：<%=rs("addtime")%></td>
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
PageCount=<%=TotalPage%> //总页数
topage=<%=PageCount%>   //当前停留页
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
		<b>近期个人日记</b></td>
	<td align="center" valign="top" width="75%">


<b>添加日记</b></td>
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
			<td width="14%"><b>日记标题：</b></td>
		<td width="83%">
		<input maxLength="30" size="54" name="title"> <select name="hide" size="1">
<option value="0">公开</option>
<option value="1" selected>保密</option>
</select></td>
		</tr>
		<tr class=a3>
			<td width="14%"><b>日记内容：</b></td>
			<td width="83%">
			<textarea name="content" rows="6" style="width:95%" cols="1" ></textarea>
		</td>
		</tr>
		<tr class=a3>
			<td colspan="2" align="center">　<input tabIndex="4" type="submit" value=" 发 送 " name="submit1">&nbsp;
			<input onclick="{if(confirm('该项操作要清除全部的内容，您确定要清除吗?')){return true;}return false;}" type="reset" value=" 重 写 "></td>
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
if ""&rs("username")&""<>""&Request.Cookies("username")&"" then error2("该日记非您所写！")

%>

	<br>
<SCRIPT>valigntop()</SCRIPT>
<table cellspacing="1" cellpadding="6" width="97%" border="0" class="a2" align="center">
	
	<tr class="a1">
	<td align="center" valign="top" width="75%">


<b>编辑日记</b></td>
	</tr><tr class="a3">
	<td align="center" valign="top">




<table cellspacing="0" cellpadding="4" width="97%" align="center" border="0" class="a2">
    
<form method=post name=form action=calendar.asp?menu=editok>
<input type=hidden name=adddate value="<%=rs("adddate")%>">
<input type=hidden name=id value="<%=rs("id")%>">
		<tr class=a3>
			<td width="14%"><b>日记标题：</b></td>
		<td width="83%">
		<input maxLength="30" size="70" name="title" value="<%=rs("title")%>"> <select name="hide" size="1">
<option value="0" <%if rs("hide")=0 then%>selected<%end if%>>公开</option>
<option value="1" <%if rs("hide")=1 then%>selected<%end if%>>保密</option>
</select></td>
		</tr>
		<tr class=a3>
			<td width="14%"><b>日记内容：</b></td>
			<td width="83%">
			<textarea name="content" rows="6" style="width:95%" cols="1" ><%=rs("content")%></textarea>
		</td>
		</tr>
		<tr class=a3>
			<td colspan="2" align="center">　<input tabIndex="4" type="submit" value=" 发 送 " name="submit1">&nbsp;
			<input onclick="{if(confirm('该项操作要清除全部的内容，您确定要清除吗?')){return true;}return false;}" type="reset" value=" 重 写 "></td>
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
'===== 算世界时间
Function TimeAdd(UTC,T)
   Dim PlusMinus, DST, y
   If Left(T,1)="-" Then PlusMinus = -1 Else PlusMinus = 1
   UTC=Right(UTC,Len(UTC)-5)
   UTC=Left(UTC,Len(UTC)-4)
   y = Year(UTC)
   TimeAdd=DateAdd("n", (Cint(Mid(T,2,2))*60 + Cint(Mid(T,4,2))) * PlusMinus, UTC)
   '美国日光节约期间: 4月第一个星日00:00 至 10月最後一个星期日00:00
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
if(today.getDay()==0) day = "星期日"
if(today.getDay()==1) day = "星期一"
if(today.getDay()==2) day = "星期二"
if(today.getDay()==3) day = "星期三"
if(today.getDay()==4) day = "星期四"
if(today.getDay()==5) day = "星期五"
if(today.getDay()==6) day = "星期六"
document.fgColor = "black";
date = " 公元 " + (today.getYear()) + " 年 " +
(today.getMonth() + 1 ) + "月 " + today.getDate() + "日 " +
day +"";
document.write(date)
    </SCRIPT>
      </FONT></TD>
    <TD align=middle width=418 class=a3 rowspan="2">

      <TABLE border=0>
        <TBODY>
        <TR>
          <TD class=a1 colSpan=7 align="center" height="25">公元<SELECT style="FONT-SIZE: 9pt" 
            onchange=changeCld() name=SY> 
              <SCRIPT language=JavaScript><!--        
            for(i=1900;i<2050;i++) document.write('<option>'+i)        
            //--></SCRIPT>
            </SELECT>年<SELECT style="FONT-SIZE: 9pt" onchange=changeCld() 
            name=SM> 
              <SCRIPT language=JavaScript><!--        
            for(i=1;i<13;i++) document.write('<option>'+i)        
            //--></SCRIPT>
            </SELECT>月</FONT>&nbsp;&nbsp;  <FONT id=GZ></FONT><BR></TD></TR>
        <TR align=middle class=a4>
          <TD width=54 height="20">日</TD>
          <TD width=54 height="20">一</TD>
          <TD width=54 height="20">二</TD>
          <TD width=50 height="20">三</TD>
          <TD width=54 height="20">四</TD>
          <TD width=54 height="20">五</TD>
          <TD width=54 height="20">六</TD></TR>
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
    
    共<font color="#FF0000"> <%=conn.execute("Select count(id)from [calendar]")(0)%> 
	</font>篇日记</TD></TR>

  <TR>
    <TD vAlign=top align=middle width=240 class=a3><BR><FONT 
      id=Clock face=Arial color=000080 size=4 align="center"></FONT>
      <P><!--时区 *表示自动调整为日光节约时间--><FONT style="FONT-SIZE: 9pt" size=2><SELECT 
      style="FONT-SIZE: 9pt" onchange=changeTZ() name=TZ> <OPTION 
        value="+0800 北京、重庆、黑龙江" selected>中国</OPTION> <OPTION 
        value="-1100 中途岛、萨摩亚群岛">萨摩亚</OPTION> <OPTION 
        value="-1000 夏威夷">夏威夷</OPTION> <OPTION value=-0900*阿拉斯加>阿拉斯加</OPTION> 
        <OPTION value=-0800*太平洋时间(美加)、提亚纳>太平洋</OPTION> <OPTION 
        value="-0700 亚历桑那">美国山区</OPTION> <OPTION 
        value=-0700*山区时间(美加)>美加山区</OPTION> <OPTION 
        value=-0600*萨克其万(加拿大)>加拿大中部</OPTION> <OPTION 
        value=-0600*墨西哥市、塔克西卡帕>墨西哥</OPTION> <OPTION 
        value=-0600*中部时间(美加)>美加中部</OPTION> <OPTION 
        value=-0500*波哥大、里玛>南美洲太平洋</OPTION> <OPTION 
        value=-0500*东部时间(美加)>美加东部</OPTION> <OPTION 
        value=-0500*印第安纳(东部)>美东</OPTION> <OPTION 
        value=-0400*加拉卡斯、拉帕兹>南美洲西部</OPTION> <OPTION 
        value="-0400*大西洋时间 加拿大)">大西洋</OPTION> <OPTION 
        value="-0330 新岛(加拿大东岸)">纽芬兰</OPTION> <OPTION 
        value="-0300 波西尼亚">东南美洲</OPTION> <OPTION 
        value="-0300 布鲁诺斯爱丽斯、乔治城">南美洲东部</OPTION> <OPTION 
        value=-0200*大西洋中部>大西洋中部</OPTION> <OPTION 
        value=-0100*亚速尔群岛、维德角群岛>亚速尔</OPTION> <OPTION 
        value="+0000 格林威治时间、都柏林、爱丁堡、伦敦">英国夏令</OPTION> <OPTION 
        value="+0000 莫洛维亚(赖比瑞亚)、卡萨布兰卡">格林威治标准</OPTION> <OPTION 
        value="+0100 巴黎、马德里">罗马</OPTION> <OPTION 
        value="+0100 布拉格, 华沙, 布达佩斯">中欧</OPTION> <OPTION 
        value="+0100 柏林、斯德哥尔摩、罗马、伯恩、布鲁赛尔、维也纳">西欧</OPTION> <OPTION 
        value="+0200 以色列">以色列</OPTION> <OPTION value=+0200*东欧>东欧</OPTION> 
        <OPTION value=+0200*开罗>埃及</OPTION> <OPTION 
        value=+0200*雅典、赫尔辛基、伊斯坦堡>GFT</OPTION> <OPTION 
        value=+0200*赫拉雷、皮托里>南非</OPTION> <OPTION 
        value=+0300*巴格达、科威特、奈洛比(肯亚)、里雅德(沙乌地)>沙乌地阿拉伯</OPTION> <OPTION 
        value=+0300*莫斯科、圣彼得堡、贺占、窝瓦格瑞德>俄罗斯</OPTION> <OPTION 
        value=+0330*德黑兰>伊朗</OPTION> <OPTION 
        value=+0400*阿布达比(东阿拉伯)、莫斯凯、塔布理斯(乔治亚共和)>阿拉伯</OPTION> <OPTION 
        value=+0430*喀布尔>阿富汗</OPTION> <OPTION 
        value="+0500 伊斯兰马巴德、克洛奇、伊卡特林堡、塔须肯">西亚</OPTION> <OPTION 
        value="+0530 孟买、加尔各答、马垂斯、新德里、可伦坡">印度</OPTION> <OPTION 
        value="+0600 阿马提、达卡">中亚</OPTION> <OPTION 
        value="+0700 曼谷、亚加达、胡志明市">曼谷</OPTION> <OPTION 
        value="-1200 安尼威土克、瓜甲兰">国际换日线</OPTION> <OPTION 
        value="+0800 台湾、香港、新加坡">台北</OPTION> <OPTION 
        value="+0900 东京、大阪、扎幌、汉城、亚库兹(东西伯利亚)">东京</OPTION> <OPTION 
        value="+0930 达尔文">澳洲中部</OPTION> <OPTION 
        value="+1000 布里斯本、墨尔本、席德尼">席德尼</OPTION> <OPTION 
        value="+1000 霍巴特">塔斯梅尼亚</OPTION> <OPTION 
        value="+1000 关岛、莫斯比港、海　威">西太平洋</OPTION> <OPTION 
        value=+1100*马哥大、所罗门群岛、新卡伦多尼亚>太平洋中部</OPTION> <OPTION 
        value="+1200 威灵顿、奥克兰">纽西兰</OPTION> <OPTION 
        value="+1200 斐济、肯加塔、马歇尔群岛">斐济</OPTION></SELECT>时间</FONT> <FONT id=tSave 
      style="FONT-SIZE: 18pt; COLOR: red; FONT-FAMILY: Wingdings"></FONT><BR><BR><FONT 
      style="FONT-SIZE: 120pt; COLOR: #13b0f4; FONT-FAMILY: Webdings">&ucirc;</FONT><BR><BR><FONT id=CITY></FONT></TD>
    <TD align=middle  class=a3>
    
    <BUTTON style="FONT-SIZE: 9pt" 
      onclick="pushBtm('YU')">年↑</BUTTON><br>
	<BUTTON style="FONT-SIZE: 9pt" 
      onclick="pushBtm('YD')">年↓</BUTTON> 
      <br>
　<p> 
      <BUTTON style="FONT-SIZE: 9pt" 
      onclick="pushBtm('MU')">月↑</BUTTON><br>
	<BUTTON style="FONT-SIZE: 9pt" 
      onclick="pushBtm('MD')">月↓</BUTTON> 
      <br>
　</p>
	<p> 
      <BUTTON style="FONT-SIZE: 9pt" onclick="pushBtm('')">当月</BUTTON> 
	<p>　</TD></TR>

</TABLE>
<SCRIPT>valignbottom()</SCRIPT>

    
</form>



    
<%
htmlend
%>