<html>
<head>
<title>《柳州市二中74-4班 论坛》登录入口</title>
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

<body background="images/lzybg01.jpg" text="#000000">
<TD vAlign=middle> 
  <p>&nbsp;</p>
  <p>&nbsp;</p>
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
  <%@language=vbscript%>
  <form name="form1" method="post" action="ez74check.asp" >
    <TABLE cellSpacing=2 cellPadding=0 width="500" align=center border=1 bordercolor="#999999" height="40">
      <TBODY>
        <TR bgColor=#e7e7e7> 
          <TD align=center colSpan=2 bgcolor="#88A8D7"><font color="#FFFFFF" size="3"><strong>《柳州市二中74-4班 论坛》登录入口</strong></font></TD>
        </TR>
      </TBODY>
    </TABLE>
	<TABLE width="500" height="60" border=1 align=center cellPadding=5 cellSpacing=1 bordercolor="#999999" bordercolorlight="#999999" bordercolordark="#ffffff" bgcolor="#f3f3f3">
      <TBODY>
        <TR> 
          <td>
		  <table width="100%" border="0">
              <TBODY>
        <TR> 
          <td><img src="images/logo.gif" width="162" height="72"></td>
            <td align="center"> 
              <table width="90%" border="0">
			    <tr>
				        <td><font color="#FF0000" size="2">愿你享有期望中的全部喜悦，每一件微小的事物都能带给你甜美的感受和无穷的快乐，祝春节快乐，万事如意! 
</font>
                        </td>
			    </tr>
			  </table>
		    </td>
        </TR>
      </TBODY>
    </TABLE>
		  </td>
        </TR>
      </TBODY>
    </TABLE>
    <TABLE width="500" border=1 align=center cellPadding=0 cellSpacing=1 bordercolor="#999999" bordercolorlight="#999999" bordercolordark="#ffffff" bgcolor="#F3F3F3">
      <TBODY> 
      <TR> 
          <TD width="19%" height="25" align=center>用 户 名：</TD>
            
          <TD width="81%" height="25"> 　
<INPUT class="text" maxLength="16" size="16" 
                  name="ez74xingming">
            </TD>
          </TR>
          
      <TR> 
          <TD width="19%" height="25" align=center>密　　码：</TD>
            
          <TD width="81%" height="25">　 
            <INPUT class="text" maxLength="16" 
                  size="16" name="ez74password" type="password">
            </TD>
          </TR>
          <%if session("check")="ez74wrong" then%>
          
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
        
    <table width="500" height="40" border="1" align="center" cellpadding="5" cellspacing="1" bordercolor="#999999" bordercolorlight="#999999" bordercolordark="#ffffff" bgcolor="#D8E8F5">
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
</body>
</html>
