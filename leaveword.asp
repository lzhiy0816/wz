<html>
<head>
<title>���������</title>
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
<%

exec="select * from admin where id=1"    ''''''''��ȡlzhiy0816��"face"ֵ(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1
adminface=rs("face")
PageSize=rs("PageSizeleave")
rs.Close
set rs=nothing

exec="select * from userleave Order by id desc"    ''''''''��ȡid���û�������(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1
rs.PageSize=PageSize
pagecount=rs.PageCount 
recordcount=rs.recordcount
page=int(request.QueryString ("page"))
if page<=0 then page=1
if request.QueryString("page")="" then
page=1
end if
rs.AbsolutePage=page
%>

<BR>
<TABLE cellSpacing=1 cellPadding=0 width="75%" align=center border=0 height="60">
  <TBODY> 
  <TR>
    <td><img src="img/lzy/lzylogo.gif" width="128" height="60"></td>        
    <td> 
      <div align="right">
        <input onclick="window.location='leave.htm'" type="button" name="Button" value="��Ҫ����">
      </div>
    </td>        
    <td> 
      <div align="right">
        <input onclick="window.location='index.htm'" type="button" name="Button1" value="�� �� ҳ">
      </div>
    </td>
   </TR>
  </TBODY>
</TABLE>       
<TABLE cellSpacing=1 cellPadding=0 width="75%" align=center border=1 bordercolor="#999999">
  <TBODY> 
   <TR>
                <TD vAlign=bottom align=left colSpan=2>
                  
      <TABLE cellSpacing=0 cellPadding=5 width="100%" align=center 
                  border=0 bordercolor="#cccccc">
        <TBODY> 
        <TR bgcolor="#dddddd"> 
          <TD vAlign=bottom align=left width="55%"><FONT 
                        color=#000000>&nbsp;��ǰ��ʾ�� <%=page%> ҳ&nbsp;&nbsp;���� <%=pagecount%> 
            ҳ&nbsp;&nbsp;���� <%=recordcount%> ����¼</FONT></TD>
                      
          <TD align=right 
              width="45%"></TD>
                    </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
            
<table cellspacing=1 cellpadding=4 width="75%" align=center border=1 bordercolor="#999999" bordercolorlight="#999999" bordercolordark="#ffffff">
  <tbody> 
  <tr> 
    <td valign=top align=middle width="20%" rowspan=2> 
      <table cellspacing=0 cellpadding=0 width="100%" border=0>
        <tbody> 
        <tr> 
          <td align=middle width="100%" height="23"><font color=red>lzhiy0816</font></td>
        </tr>
        <tr> 
          <td align=middle width="100%" height="47"><img src="face/<%=adminface%>.gif"></td>
        </tr>
        <tr> 
          <td align=middle width="100%" height="35">����:��������</td>
        </tr>
        </tbody>
      </table>
    </td>
    <td valign=top align=left height="140"> 
      <table class=contenta cellpadding=2 width="100%" height="20" valign="top">
        <tr> 
          <td valign="top">
		    <img src="image/icon/icon1.gif" align=abscenter border=0 width="15" height="15" alt=û�б���> 
            &nbsp;&nbsp;<font color="#006600">����ʱ��: 2004-6-23 23:52:59</font>
		  </td>
        </tr>
      </table>
      <table class=contenta cellpadding=2 width="100%" height="60" valign="top">
        <tr> 
          <td valign="top">
			  ���ã�<br>
          ������ӭ����¼ <a href="index.htm"><b>���������</b></a> ��վ������ӭ���ڴ����ԣ����ᱦ���������лл�� 
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr width="80%" height=10> 
	<td class=contentb valign=center height=20>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
	   <tr>
	     <td width="50%">
                        <%if session("checked")="yes3" then %>
                        <font color="#666666">IP: <%=Request.ServerVariables("Remote_Addr")%></font>
                        <%end if %>
						&nbsp;
         </td>
	     <td width="50%"><div align="right">	 
	       <a href="mailto:lzhiy0816@163.com">
	         <img height=16 src="image/icon/email.gif" alt="��lzhiy0816���ʼ�" width=16 align=absMiddle border=0>
	       </a>
         </div></td>
	   </tr>
	  </table>
	</td>
  </tr>
  </tbody>
</table>
<%for i=1 to rs.PageSize%>
<%if rs.EOF then 
        exit for
        end if%>
<table cellspacing=1 cellpadding=4 width="75%" align=center border=1 bordercolor="#999999" bordercolorlight="#999999" bordercolordark="#ffffff">
  <tbody> 
  <tr> 
    <td valign=top align=middle width="20%" rowspan=2> 
      <table cellspacing=0 cellpadding=0 width="100%" border=0>
        <tbody> 
        <tr> 
          <td align=middle width="100%" height="23"><%=rs("xingming")%></td>
        </tr>
        <tr> 
          <td align=middle width="100%" height="47"><img src="face/<%=rs("face")%>.gif"></td>
        </tr>
        <tr> 
          <td align=middle width="100%" height="35">
		    <%if rs("dizhi")<>"" then%>
              ����:<%=rs("dizhi")%>
            <%end if%>
		  </td>
        </tr>
        </tbody> 
      </table>
    </td>
    <%if rs("face")>100 then%>
	<td valign=top align=left height="80">
    <%else%>
	<td valign=top align=left height="140">
	<%end if%>
      <table class=contenta cellpadding=2 width="100%" height="20" valign="top">
        <tr> 
          <td valign="top"> 
		          <%if rs("icon")=1 then%>
                  <IMG height=15 alt=û�б���
                  src="image/icon/icon1.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=2 then%>
                  <IMG height=15 alt=�쿴��
                  src="image/icon/icon2.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=3 then%>
                  <IMG height=15 alt=���ˣ����������ˣ�
                  src="image/icon/icon3.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=4 then%>
                  <IMG height=15 alt=���棡
                  src="image/icon/icon4.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=5 then%>
                  <IMG height=15 alt=���⣡
                  src="image/icon/icon5.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=6 then%>
                  <IMG height=15 alt=�ۣ�����ˣ�
                  src="image/icon/icon6.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=7 then%>
                  <IMG height=15 alt=���ѹ���
                  src="image/icon/icon7.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=8 then%>
                  <IMG height=15 alt=���²��ˣ�
                  src="image/icon/icon8.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=9 then%>
                  <IMG height=15 alt=����
                  src="image/icon/icon9.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=10 then%>
                  <IMG height=15 alt=������
                  src="image/icon/icon10.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=11 then%>
                  <IMG height=15 alt=�Ǻǣ�
                  src="image/icon/icon11.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=12 then%>
                  <IMG height=15 alt=գ�ۣ�����������˼Ŷ��
                  src="image/icon/icon12.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=13 then%>
                  <IMG height=15 alt=ǿ�ҷ���
                  src="image/icon/icon13.gif" width=15 
                  align=ABSCENTER>
                  <%end if%>
                  <%if rs("icon")=14 then%>
                  <IMG height=15 alt=ͬ�⣡
                  src="image/icon/icon14.gif" width=15 
                  align=ABSCENTER> 
				  <%end if%>
            &nbsp;&nbsp;<font color="#006600">����ʱ��: <%=rs("time")%></font> 
          </td>
        </tr>
      </table>
      <table class=contenta cellpadding=2 width="100%" height="60" valign="top">
        <tr> 
          <td valign="top"><%=rs("liuyanneirong")%></td>
        </tr>
      </table>
      <%if rs("huifu")<>"" then%>
      <TABLE class=contenta cellPadding=2 width="96%" valign="top">
        <TBODY> 
        <TR> 
          <TD width=20></TD>
          <TD class=tbCss  height=1 bgcolor="#ffcc00"></TD>
        </TR>
        <TR> 
          <TD width=20></TD>
          <TD><img src="image/icon/dot.gif" width="21" height="10"><font color="#FF0000">&nbsp;lzhiy0816�ظ�:</font></TD>
        </TR>
        <TR> 
          <TD width=20></TD>
          <TD class=guestdate vAlign=top><font color="#006600"><%=rs("huifu")%></font></TD>
        </TR>
        </TBODY> 
      </TABLE>
      <%end if%>

    </td>
  </tr>
  <tr width="80%" height=10> 
	<td class=contentb valign=center height=20>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
	   <tr>
	     <td width="50%">
                        <%if session("checked")="yes3" then %>
                        <font color="#666666">IP: <%=rs("userip")%></font>
                        <%end if %>
						&nbsp;
         </td>
         <td width="50%"><div align="right">
	                    <%if rs("password")<>"" then%>
                        <a
                        href="mdflogin.asp?id=<%=rs("id")%>"><img
                        height=16 src="image/icon/reply.gif" alt="�������޸�����" width=16
                        align=absMiddle border=0></a>&nbsp;
                        <%end if%>
                        <%if rs("email")<>"" then%>
                        <a
                        href="mailto:<%=rs("email")%>"><img height=16
                        src="image/icon/email.gif" alt="�������ʼ�" width=16 align=absMiddle
                        border=0></a>&nbsp;
                        <%end if%>
                        <%if rs("zhuye")<>"http://" then%>
                        <a
                        href="<%=rs("zhuye")%>"
                        target=_blank><img height=16 src="image/icon/home.gif" alt="���ʸ�����ҳ"
                        width=16 align=absMiddle border=0></a>&nbsp;
                        <%end if%> 
                        <a 
                        href="dellogin.asp?id=<%=rs("id")%>"><img
                        height=16 src="image/icon/recycle.gif" alt="����ظ�" width=16
                        align=absMiddle border=0></a>
         </div></td>
		</tr>
	   </table>
	 </td>
  </tr>
  </tbody> 
</table>
<%rs.movenext
next
%>
            
<TABLE cellSpacing=1 cellPadding=0 width="75%" align=center border=1 bordercolor="#999999">
  <TBODY> 
  <TR>
                <TD vAlign=bottom align=left colSpan=2>
                  
      <TABLE cellSpacing=0 cellPadding=5 width="100%" align=center 
                  border=0 bordercolor="#cccccc">
        <TBODY> 
        <TR bgColor=#dddddd> 
          <TD vAlign=bottom align=left width="76%"><FONT 
                        color=#000000>&nbsp;��ǰ��ʾ�� <%=page%> ҳ&nbsp;&nbsp;���� <%=pagecount%> 
            ҳ&nbsp;&nbsp;���� <%=recordcount%> ����¼ </FONT></TD>
                      
          <TD align=right 
              width="24%"></TD>
                    </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
            
			
<TABLE cellSpacing=0 cellPadding=0 width="75%" align=center border=0 height="18">
  <TBODY> 
  <TR>
                <TD vAlign=bottom align=left colSpan=2 height="14"> 
                  <TABLE cellSpacing=0 cellPadding=5 width="100%" align=center 
                  border=0>
                    <TBODY> 
                    <TR> 
                      <%if page=1 and not page=pagecount then%>
                      <TD align=left width="76%" >&nbsp;
                        <img src="image/icon/head.gif" width="13" height="11" alt="��ҳ">
                        <img src="image/icon/pre.gif" width="13" height="11" alt="��һҳ">
                        &nbsp;
                       <font color=#000000>[ <%=page%> ]</font>&nbsp;&nbsp;
                       <a href="leaveword.asp?page=<%=page+1%>">
                        <img src="image/icon/next.gif" width="13" height="11" alt="��һҳ">
                       </a>
                       <a href="leaveword.asp?page=<%=pagecount%>">
                        <img src="image/icon/end.gif" width="13" height="11" alt="βҳ">
                       </a>
                      </TD>
                      <TD width="24%" align=right>
                        <form name="form1" method="post" action="go.asp">
                          <input type="image" border="0" name="imageField" src="image/icon/go.gif" alt="����">
                          �� 
                          <input type="text" name="page" size="6" maxlength="6">
                          ҳ 
                        </form>
                      </TD>
                      <%elseif page=pagecount and not page=1 then%>
                      <TD align=left width="76%" >&nbsp;
                       <a href="leaveword.asp?page=1">
                        <img src="image/icon/head.gif" width="13" height="11" alt="��ҳ">
                       </a>
                       <a href="leaveword.asp?page=<%=page-1%>">
                        <img src="image/icon/pre.gif" width="13" height="11" alt="��һҳ">
                       </a>&nbsp;
                       <font color=#000000>[ <%=page%> ]</font>&nbsp;&nbsp;
                        <img src="image/icon/next.gif" width="13" height="11" alt="��һҳ">
                        <img src="image/icon/end.gif" width="13" height="11" alt="βҳ">
                      </TD>
                      <TD width="24%" align=right>
                        <form name="form1" method="post" action="go.asp">
                          <input type="image" border="0" name="imageField" src="image/icon/go.gif" alt="����">
                          �� 
                          <input type="text" name="page" size="6" maxlength="6">
                          ҳ 
                        </form>
                      </TD>
                      <%elseif page<1 then%>
                        <TD align=left width="76%" ><div class="font"><font color=red>û���κμ�¼��</font></div></td>
                        <TD width="24%" align=right>&nbsp;</td>
                      <%elseif page>pagecount then%>
                        <TD align=left width="76%" ><div class="font"><font color=red>û���κμ�¼��</font></div></td>
                        <TD width="24%" align=right>&nbsp;</td>
                      <%elseif page=1 and page=pagecount then%>
                      <TD align=left width="76%" >&nbsp;
                       
                      </TD>
                      <TD width="24%" align=right>&nbsp;
                       
                      </TD>
                      <%else%>
                      <TD align=left width="76%" >&nbsp;
                       <a href="leaveword.asp?page=1">
                        <img src="image/icon/head.gif" width="13" height="11" alt="��ҳ">
                       </a>
                       <a href="leaveword.asp?page=<%=page-1%>">
                        <img src="image/icon/pre.gif" width="13" height="11" alt="��һҳ">
                       </a>&nbsp;
                       <font color=#000000>[ <%=page%> ]</font>&nbsp;&nbsp;
                       <a href="leaveword.asp?page=<%=page+1%>">
                        <img src="image/icon/next.gif" width="13" height="11" alt="��һҳ">
                       </a>
                       <a href="leaveword.asp?page=<%=pagecount%>">
                        <img src="image/icon/end.gif" width="13" height="11" alt="βҳ">
                       </a>
                      </TD>
                      <TD width="24%" align=right>
                        <form name="form1" method="post" action="go.asp">
                          <input type="image" border="0" name="imageField" src="image/icon/go.gif" alt="����">
                          �� 
                          <input type="text" name="page" size="6" maxlength="6">
                          ҳ 
                        </form>
                      </TD>
                      <%end if%>
                    </TR>

                    </TBODY></TABLE></TD></TR></TBODY></TABLE>
            
<table width="60%" border="0" cellspacing="0" cellpadding="0" height="5%" align="center">
  <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
  
<table width="60%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
      <td width="5%">&nbsp;</td>
      <td width="18%"><font size="2"><b>���������</b></font></td>
      <td width="72%"><font color="#FF0000"><marquee><i><b><font color="#0000FF">Lzhiy0816</font></b> 
        ��л���Ĺ��٣�</i></marquee></font></td>
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
document.write("��");
document.write(month);
document.write("��");
document.write(date);
document.write("�� ");

myArray=new Array(6);
myArray[0]="������"
myArray[1]="����һ"
myArray[2]="���ڶ�"
myArray[3]="������"
myArray[4]="������"
myArray[5]="������"
myArray[6]="������"
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
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
          
</body>
</html>
