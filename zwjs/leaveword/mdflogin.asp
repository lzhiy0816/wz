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

<%@language=vbscript%>
<!--#include file="conn.asp"-->
<%id=request.querystring("id")
exec="select * from userleave where id="&request.querystring("id")
set rs=server.createobject("adodb.recordset")
rs.open exec,conn
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
  <%id=request.querystring("id")%>
  <form name="form1" method="post" action="mdfcheck.asp?id=<%=id%>" >
    <TABLE cellSpacing=1 cellPadding=0 width="65%" align=center border=0 height="60">
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
    
    <TABLE cellSpacing=1 cellPadding=0 width="65%" align=center border=1 bordercolor="#999999" height="30">
      <TBODY> 
      <TR bgColor=#e7e7e7>
            
        <TD align=middle colSpan=2 bgcolor="#dddddd"><font color="#009900"><b>用 
          户 修 改 留 言 入 口</b></font></TD>
          </TR></TBODY></TABLE>
        
    <TABLE cellSpacing=1 cellPadding=5 width="65%" align=center border=1 bordercolor="#999999" bordercolorlight="#999999" bordercolordark="#ffffff">
      <TBODY> 
      <TR> 
        <TD align=middle width="19%">您的姓名：</TD>
            
        <TD width="81%"> <%=rs("xingming")%> </TD>
          </TR>
          
      <TR> 
        <TD align=middle width="19%">密　　码：</TD>
            
        <TD vAlign=top width="81%"> 
          <INPUT class="text" maxLength="16" 
                  size="16" name="password" type="password">
            </TD>
          </TR>
          <%if session("check")="wrong" then%>
          
      <TR> 
        <TD align=middle colspan="2" height="32"> 
          <div align="center"><font color="#FF0000">对不起！密码验证错误！</font></div>
            </TD>
          </TR>
          <%end if%>
<%
session("checked")=""
session("check")=""
%>
          </TBODY>
        </TABLE>
            <table width="65%" border="1" cellspacing="1" cellpadding="0" align="center" height="40" bordercolor="#999999" bordercolorlight="#999999" bordercolordark="#ffffff">
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
     
  </FORM> 
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
