<%@language=vbscript%>
<%if not session("checked")="yes2" then
id=request.querystring("id")
response.Redirect "dellogin.asp?id="&id&""
else

session("checked")=""
session("check")=""
%>
<html><!-- #BeginTemplate "/Templates/000.dwt" -->
<head>
<title>〖自我介绍网站（临时）〗</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<STYLE type=text/css>
<!--
body { font-family:"Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 10pt; line-height:18px}
p {  font-family:"Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 10pt; line-height:18px}
a:link {  color: #0033cc; text-decoration: none}
a:visited {  color: #0033cc;  text-decoration: none}
a:hover {  color: #FF0033; text-decoration: underline}
td {  font-family: "Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 10pt; line-height:18px}
.CompanyName {font-family: "Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 23pt}
-->
</STYLE>

<!-- #BeginEditable "head" -->
<!-- #EndEditable -->
</head>

<body bgcolor="#FFFFFF" text="#000000">
<!-- #BeginEditable "doctitle" --> 
<!--#include file="conn.asp"-->
<%id=request.querystring("id")
exec="select * from userleave where id="&request.querystring("id")
set rs=server.createobject("adodb.recordset")
rs.open exec,conn
%>
    <TD vAlign=center>
	  <STYLE type=text/css>BODY {
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
<br>
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

            <form name="form1" method="post" action="delanssave.asp">
            
    <TABLE cellSpacing=1 cellPadding=0 width="75%" align=center border=0 height="60">
      <TBODY> 
      <TR>
        <td><img src="image/lyb.gif" width="150" height="50"></td>        
        <td> 
          <div align="right">
            <input onClick="window.location='leaveword.asp'" type=button value=查看留言 name=Button3>
          </div>
        </td>        
       </TR>
      </TBODY>
     </TABLE>       
    
    <TABLE cellSpacing=1 cellPadding=0 width="75%" align=center border=1 bordercolor="#999999" height="33">
      <TBODY> 
      <TR bgColor=#e7e7e7>
            
        <TD align=middle colSpan=2 bgcolor="#dddddd"><font color="#990000"><b>网 
          站 管 理 员 管 理 窗 口</b></font></TD>
          </TR></TBODY></TABLE>
        
    <TABLE cellSpacing=1 cellPadding=5 width="75%" align=center border=1 bordercolor="999999" bordercolorlight="999999" bordercolordark="ffffff">
      <TBODY> 
      <TR> 
        <TD align=middle width="19%">用户姓名：</TD>
                
        <TD vAlign=top width="81%"><%=rs("xingming")%>&nbsp;</TD>
      </TR>
      <TR> 
        <TD align=middle width="19%">头　　像：</TD>
                
        <TD vAlign=top width="81%"><img src="face/<%=rs("face")%>.gif"></TD>
      </TR>
              
      <TR> 
        <TD align=middle width="19%">来自哪里：</TD>
                
        <TD vAlign=top width="81%"><%=rs("dizhi")%>&nbsp;</TD>
      </TR>
              
      <TR> 
        <TD align=middle width="19%">IP 地 址：</TD>
                
        <TD vAlign=top width="81%"><%=rs("userip")%>&nbsp;</TD>
      </TR>
              
      <TR> 
        <TD align=middle width="19%">电子邮件：</TD>
                
        <TD vAlign=top width="81%"><%=rs("email")%>&nbsp;</TD>
      </TR>
              
      <TR> 
        <TD align=middle width="19%">用户网址：</TD>
                
        <TD vAlign=top width="81%"><%=rs("zhuye")%>&nbsp;</TD>
      </TR>
              
      <TR> 
        <TD align=middle width="19%">表情图标：</TD>
                
        <TD><br>
          &nbsp;&nbsp;
                  <%if rs("icon")=1 then%>
                  <IMG height=15 alt=没有表情
                  src="image/icon/icon1.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=2 then%>
                  <IMG height=15 alt=快看！
                  src="image/icon/icon2.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=3 then%>
                  <IMG height=15 alt=嗨嗨，我有主意了！
                  src="image/icon/icon3.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=4 then%>
                  <IMG height=15 alt=警告！
                  src="image/icon/icon4.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=5 then%>
                  <IMG height=15 alt=问题！
                  src="image/icon/icon5.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=6 then%>
                  <IMG height=15 alt=哇！真高兴！
                  src="image/icon/icon6.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=7 then%>
                  <IMG height=15 alt=好难过！
                  src="image/icon/icon7.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=8 then%>
                  <IMG height=15 alt=这下惨了！
                  src="image/icon/icon8.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=9 then%>
                  <IMG height=15 alt=含情
                  src="image/icon/icon9.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=10 then%>
                  <IMG height=15 alt=嘻嘻！
                  src="image/icon/icon10.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=11 then%>
                  <IMG height=15 alt=呵呵！
                  src="image/icon/icon11.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=12 then%>
                  <IMG height=15 alt=眨眼－－对你有意思哦！
                  src="image/icon/icon12.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=13 then%>
                  <IMG height=15 alt=强烈反对
                  src="image/icon/icon13.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=14 then%>
                  <IMG height=15 alt=同意！
                  src="image/icon/icon14.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>&nbsp;&nbsp; </TD></TR>
              
      <TR height="100"> 
        <TD align=middle width="19%">用户留言：</TD>
                
        <td width="81%" height="100" valign=top> 
          <table width="82%" hight="90%" border="0" cellspacing="1" cellpadding="0">
                  <tr>
                   <TD>
                   <%=rs("liuyanneirong")%>&nbsp;
                   </TD>
                  </tr>
                 </table>
                </td>
              </TR>
              
      <TR> 
        <TD align=middle width="19%"> 版主回复： </TD>
                
        <TD width="81%"> 
          <TEXTAREA name=huifu rows=6 cols=51><%=rs("huifu")%></TEXTAREA>
                <br>
                 <font color="#990000">
                 〈Br〉：表示换行，也可按“Enter”回车键换行
                 </font>
                </TD>
              </TR>
              </TBODY></TABLE>
      <table width="75%" border="1" cellspacing="1" cellpadding="0" align="center" height="40" bordercolor="#999999" bordercolorlight="#999999" bordercolordark="#ffffff">
      <tr> 
        <td> 
          <div align="center">
            <table width="100%" border="0">
              <tr> 
                <td>
                  <div align="center">
                    <input class=buttom onclick="return check(this.form,1000)" type=submit value=回复 name=submit>
                  </div>
                </td>
                <td>
                  <div align="center">
                    <input class=buttom onclick="return check(this.form,1000)" type=submit value=删除 name=submit>
                  </div>
                </td>
                <td>
                  <div align="center">
                    <input class=buttom onclick="return check(this.form,1000)" type=submit value=显IP name=submit>
                  </div>
                </td>
                <td>
                  <div align="center">
                    <INPUT class=buttom type=reset value=取消 name=reset>
                  </div>
                </td>
              </tr>
            </table>
          </div>
          <div align="center"></div>
        </td>
      </tr>
    </table>
	<CENTER><BR>
    <input type="hidden" name="id" value="<%=request.querystring("id")%>">
    </CENTER></FORM>
	</TD>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<br>
<!-- #EndEditable -->
<table width="80%" border="0" align="center">
  <tr> 
    <td> 
      <div align="center"><font size="2">Copyright&copy;2004&nbsp;&nbsp;&nbsp;&nbsp;〖自我介绍网站（临时）〗&nbsp;&nbsp;&nbsp;&nbsp;站长：梁志英</font></div>
    </td>
  </tr>
</table>
<table width="80%" border="1" cellspacing="1" cellpadding="1" align="center" bordercolor="#999999" bordercolorlight="#999999" bordercolordark="#ffffff">
  <tr>
    <td></td>
  </tr>
</table>
<table width="80%" border="0" align="center">
  <tr> 
    <td> 
      <div align="center"><font size="2">建站日期：2005-3-1&nbsp;&nbsp;&nbsp;&nbsp;更新日期：2005-4-23</font></div>
    </td>
  </tr>
</table>
<!-- #BeginEditable "count" -->
<!-- #EndEditable -->
</body>
<!-- #EndTemplate --></html>
<%end if%>
