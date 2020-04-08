<%@language=vbscript%>
<%if not session("checked")="yes" then 
id=request.querystring("id")
response.Redirect "mdflogin.asp?id="&id&""
else
%>
<html>
<head>
<title>共享的世界</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<STYLE type=text/css>

<!--

body { font-family:"Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 9pt}

p {  font-family:"Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 9pt}

a:link {  color: #0000FF; text-decoration: none}

a:visited {  text-decoration: none}

a:hover {  color: #FF0033; text-decoration: underline}

td {  font-family: "Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 9pt}

.CompanyName {font-family: "Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 23pt}

-->

</STYLE>
</head>

<body bgcolor="#FFFFFF" text="#000000" background="img/background/lzybg01.jpg">

<!--#include file="conn.asp"-->
<%id=request.querystring("id")
exec="select * from userleave where id="&request.querystring("id")
set rs=server.createobject("adodb.recordset")
rs.open exec,conn
huifu=rs("huifu")
%>

<TD vAlign=middle> 
  <p>&nbsp;</p>
	  <p><STYLE type=text/css>BODY {
	FONT-SIZE: 12px; COLOR: #000000; BACKGROUND-COLOR: #ffffff
}
TD {
	FONT-SIZE: 12px; COLOR: #000000
}
TR {
	FONT-SIZE: 12px; COLOR: #000000
}
A:link {
	COLOR: #336699; TEXT-DECORATION: none
}
A:visited {
	COLOR: #6699ff; TEXT-DECORATION: none
}
A:hover {
	COLOR: red
}
FONT {
	TEXT-DECORATION: none
}
INPUT.text {
	BORDER-RIGHT: #101010 1px solid; BORDER-TOP: #101010 1px solid; BORDER-LEFT: #101010 1px solid; COLOR: #000000; BORDER-BOTTOM: #101010 1px solid; BACKGROUND-COLOR: #ffffff
}
INPUT.file {
	BORDER-RIGHT: #101010 1px solid; BORDER-TOP: #101010 1px solid; BORDER-LEFT: #101010 1px solid; COLOR: #000000; BORDER-BOTTOM: #101010 1px solid; BACKGROUND-COLOR: #ffffff
}
SELECT {
	BORDER-RIGHT: #101010 1px solid; BORDER-TOP: #101010 1px solid; BORDER-LEFT: #101010 1px solid; COLOR: #000000; BORDER-BOTTOM: #101010 1px solid; BACKGROUND-COLOR: #ffffff
}
TEXTAREA {
	BORDER-RIGHT: #101010 1px solid; BORDER-TOP: #101010 1px solid; BORDER-LEFT: #101010 1px solid; COLOR: #000000; BORDER-BOTTOM: #101010 1px solid; BACKGROUND-COLOR: #ffffff
}
INPUT.buttom {
	BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px; BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 12px; BORDER-LEFT-COLOR: #000000; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #000000; PADDING-BOTTOM: 1px; COLOR: #000000; BORDER-TOP-COLOR: #000000; PADDING-TOP: 1px; HEIGHT: 18px; BACKGROUND-COLOR: #dedfdf; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #000000
}
</STYLE>
    </p>

  <!-- form.cjs -->
<SCRIPT language=JavaScript1.2>
function check(theForm,textlengh) {

	if(theForm.xingming.value == "" || theForm.password.value == "" || theForm.liuyanneirong.value == "") {
		alert("对不起!姓名、密码和留言内容必须填写!");
		return false;
	}
	if (theForm.liuyanneirong.value.length > textlengh){
		alert("对不起！留言长度不能超过1000字符");
		return false;
	}	
}
</SCRIPT>

  <form name="form1" method="post" action="mdfsave.asp">
    <TABLE cellSpacing=1 cellPadding=0 width="75%" align=center border=0 height="60">
      <TBODY> 
      <TR>
    <td><img src="img/lzy/lzylogo.gif" width="128" height="60"></td>        
    <td> 
          <div align="right">
            <input onClick="window.location='leaveword.asp'" type=button value=查看留言 name=Button3>
          </div>
    </td>        
    <td> 
          <div align="right">
            <input onClick="window.location='index.htm'" type="button" name="Button1" value="回 首 页">
          </div>
    </td>
   </TR>
  </TBODY>
</TABLE>       
    
    <TABLE cellSpacing=1 cellPadding=0 width="75%" align=center border=1 bordercolor="#999999" height="30">
      <TBODY> 
      <TR bgColor=#e7e7e7>
            
        <TD align=middle colSpan=2 bgcolor="#dddddd"><font color="#009900"><b>请 
          修 改 您 的 留 言 及 相 关 信 息</b></font></TD>
          </TR></TBODY></TABLE>
        
    <TABLE cellSpacing=1 cellPadding=5 width="75%" align=center border=1 bordercolor="#999999" bordercolorlight="#999999" bordercolordark="#ffffff">
      <TBODY> 
      <%if huifu<>"" then%>
      <TR>
              <%else%>
      <TR> 
        <%end if%>
        <TD align=middle width="19%">头　　像：</TD>
        <TD vAlign=top width="81%"><input type=hidden name=face value="<%=rs("face")%>">
		   <a href=#f onclick="window.open('facelist.htm','tmpWin','menubar=0,scrollbars=yes');"> 
              <img src="face/<%=rs("face")%>.gif" name="picShow" border="0" alt="选择一个新头像">
		   </a>&nbsp;&nbsp;&nbsp;&nbsp; 
           <INPUT class=text type=hidden maxLength=16 size=10 name=aicq>
           <font color="green">（点头像选择其他头像）</font></TD></TR>
      <%if huifu<>"" then%>
      <TR>
              <%else%>
      <TR> 
        <%end if%>
        <TD align=middle width="19%">来自哪里：</TD>
                <TD vAlign=top width="81%"><INPUT class="text" maxLength="15" 
                  size="15" name="dizhi" value="<%=rs("dizhi")%>"> &nbsp;&nbsp; <INPUT class=text type=hidden
                  maxLength=15 size=10 name=icq> </TD></TR>
      <%if huifu<>"" then%>
	  <TR>
              <%else%>
      <TR> 
        <%end if%>
        <TD align=middle width="19%">电子邮件：</TD>
                <TD vAlign=top width="81%"><INPUT class="text" maxLength="52" 
                  size="52" name="email" value="<%=rs("email")%>"> &nbsp;&nbsp; <INPUT class=text type=hidden
                  maxLength=25 size=10 name=oicq> </TD></TR>
      <%if huifu<>"" then%>
	  <TR>
              <%else%>
      <TR> 
        <%end if%>
        <TD align=middle width="19%">您的网址：</TD>
                <TD vAlign=top width="81%">
                <INPUT class="text" maxLength="70"
                  size="52" value="<%=rs("zhuye")%>" name="zhuye">
                   </TD></TR>
      <%if huifu<>"" then%>
	  <TR>
              <%else%>
      <TR> 
        <%end if%>
        <TD align=middle width="19%">表情图标：</TD>
                <TD><BR>&nbsp; 
              <INPUT type=radio value=1 
                  name=icon <%if rs("icon")=1 then%>CHECKED<%end if%>>
              &nbsp;<IMG height=15 alt=没有表情 
                  src="image/icon/icon1.gif" width=15 
                  align=ABSCENTER>&nbsp;&nbsp; 
              <INPUT type=radio value=2 <%if rs("icon")=2 then%>CHECKED<%end if%>
                  name=icon>&nbsp;&nbsp;<IMG height=15 alt=快看！ 
                  src="image/icon/icon2.gif" width=15 
                  align=ABSCENTER>&nbsp;&nbsp; <INPUT type=radio value=3 <%if rs("icon")=3 then%>CHECKED<%end if%>
                  name=icon>&nbsp;&nbsp;<IMG height=15 alt=嗨嗨，我有主意了！
                  src="image/icon/icon3.gif" width=15 
                  align=ABSCENTER>&nbsp;&nbsp; <INPUT type=radio value=4 <%if rs("icon")=4 then%>CHECKED<%end if%>
                  name=icon>&nbsp;&nbsp;<IMG height=15 alt=警告！ 
                  src="image/icon/icon4.gif" width=15 
                  align=ABSCENTER>&nbsp;&nbsp; <INPUT type=radio value=5 <%if rs("icon")=5 then%>CHECKED<%end if%>
                  name=icon>&nbsp;&nbsp;<IMG height=15 alt=问题！ 
                  src="image/icon/icon5.gif" width=15 
                  align=ABSCENTER>&nbsp;&nbsp; <INPUT type=radio value=6 <%if rs("icon")=6 then%>CHECKED<%end if%>
                  name=icon>&nbsp;&nbsp;<IMG height=15 alt=哇！真高兴！ 
                  src="image/icon/icon6.gif" width=15 
                  align=ABSCENTER>&nbsp;&nbsp; <INPUT type=radio value=7 <%if rs("icon")=7 then%>CHECKED<%end if%>
                  name=icon>&nbsp;&nbsp;<IMG height=15 alt=好难过！ 
                  src="image/icon/icon7.gif" width=15 
                  align=ABSCENTER>&nbsp;&nbsp;<BR>&nbsp; <INPUT type=radio 
                  value=8 <%if rs("icon")=8 then%>CHECKED<%end if%> name=icon>&nbsp;&nbsp;<IMG height=15 alt=这下惨了！
                  src="image/icon/icon8.gif" width=15 
                  align=ABSCENTER>&nbsp;&nbsp; <INPUT type=radio value=9 <%if rs("icon")=9 then%>CHECKED<%end if%>
                  name=icon>&nbsp;&nbsp;<IMG height=15 alt=含情 
                  src="image/icon/icon9.gif" width=15 
                  align=ABSCENTER>&nbsp;&nbsp; <INPUT type=radio value=10 <%if rs("icon")=10 then%>CHECKED<%end if%>
                  name=icon>&nbsp;&nbsp;<IMG height=15 alt=嘻嘻！ 
                  src="image/icon/icon10.gif" width=15 
                  align=ABSCENTER>&nbsp;&nbsp; <INPUT type=radio value=11 <%if rs("icon")=11 then%>CHECKED<%end if%>
                  name=icon>&nbsp;&nbsp;<IMG height=15 alt=呵呵！ 
                  src="image/icon/icon11.gif" width=15 
                  align=ABSCENTER>&nbsp;&nbsp; <INPUT type=radio value=12 <%if rs("icon")=12 then%>CHECKED<%end if%>
                  name=icon>&nbsp;&nbsp;<IMG height=15 alt=眨眼－－对你有意思哦！ 
                  src="image/icon/icon12.gif" width=15 
                  align=ABSCENTER>&nbsp;&nbsp; <INPUT type=radio value=13 <%if rs("icon")=13 then%>CHECKED<%end if%>
                  name=icon>&nbsp;&nbsp;<IMG height=15 alt=强烈反对 
                  src="image/icon/icon13.gif" width=15 
                  align=ABSCENTER>&nbsp;&nbsp; <INPUT type=radio value=14 <%if rs("icon")=14 then%>CHECKED<%end if%>
                  name=icon>&nbsp;&nbsp;<IMG height=15 alt=同意！ 
                  src="image/icon/icon14.gif" width=15 
                  align=ABSCENTER>&nbsp;&nbsp; </TD></TR>
      <%if huifu<>"" then%>
	  <TR>
              <%else%>
      <TR> 
        <%end if%>
        <TD align=middle width="19%">
                 您的留言：
                </TD>
                <TD width="81%"><TEXTAREA name=liuyanneirong rows=6 cols=51><%=rs("liuyanneirong")%></TEXTAREA>
                <br>
                 <font color="#009900">
                 〈Br〉：表示换行，也可按“Enter”回车键换行
                 </font>
                </TD>
                </TR>
              <%if huifu<>"" then%>
              
      <TR height="100"> 
        <TD align=middle width="19%">版主回复：</TD>
                <td width="81%" height="100" valign=top>
                 <table width="82%" hight="90%" border="0" cellspacing="1" cellpadding="0">
                  <tr>
                   <TD><font color="#336699"><%=huifu%></font></TD>
                  </tr>
                 </table>
                </td>
              </TR>
              <%end if%>
              </TBODY></TABLE>
            
    <table width="75%" border="1" cellspacing="1" cellpadding="0" align="center" height="40" bordercolor="#999999" bordercolorlight="#999999" bordercolordark="#ffffff">
      <tr> 
        <td> 
          <div align="center">
            <table width="100%" border="0">
              <tr>
                <td>
                  <div align="center">
                    <input class=buttom onClick="return check(this.form,1000)" type=submit value=确定 name=submit>
                  </div>
                </td>
                <td>
                  <div align="center">
                    <input class=buttom type=reset value=取消 name=reset>
                  </div>
                </td>
              </tr>
            </table>
          </div>
          <div align="center"></div>
        </td>
      </tr>
    </table>
    <input type="hidden" name="id" value="<%=request.querystring("id")%>"> 
  </FORM> 
	</TD>
<%
session("checked")=""
session("check")=""
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<br><br>
<table width="60%" border="0" cellspacing="0" cellpadding="0" height="5%" align="center">
  <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
      
<table width="60%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
      <td width="5%">&nbsp;</td>
      <td width="18%"><font size="2"><b>共享的世界</b></font></td>
      <td width="72%"><font color="#FF0000"><marquee><i><b><font color="#0000FF">Lzhiy0816</font></b> 
        感谢您的光临！</i></marquee></font></td>
      <td>&nbsp;</td>
    </tr>
  </table>
  
<table width="60%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
      <td height="4"><img src="img/line/0173.gif" width="100%" height="4"></td>
    </tr>
  </table>
  
<table width="60%" border="0" cellspacing="0" cellpadding="0" height="8%" align="center">
  <tr>
      <td>
        <div align="center"><font color="#009900"> 
          <script language="JavaScript" >
<!--
tmpDate = new Date();
date = tmpDate.getDate();
month= tmpDate.getMonth() + 1 ;
year= tmpDate.getYear();
document.write(year);
document.write("年");
document.write(month);
document.write("月");
document.write(date);
document.write("日 ");

myArray=new Array(6);
myArray[0]="星期日"
myArray[1]="星期一"
myArray[2]="星期二"
myArray[3]="星期三"
myArray[4]="星期四"
myArray[5]="星期五"
myArray[6]="星期六"
weekday=tmpDate.getDay();
if (weekday==0 | weekday==6)
{
document.write(myArray[weekday])
}
else
{document.write(myArray[weekday])
};
// -->
</script>
          </font></div>
      </td>
    </tr>
  </table>
</body>
</html>
<%end if%>
