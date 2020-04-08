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
<p><!--#include file="conn.asp"-->
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

<%

exec="select * from admin where id=1"    ''''''''获取lzhiy0816的"face"值(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1
adminface=rs("face")
PageSize=rs("PageSizeleave")
rs.Close
set rs=nothing

exec="select * from userleave Order by id desc"    ''''''''获取id号用户的数据(lzhiy0816 design)''''''''
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
<table width="750" border="1" align="center" height="25" bordercolor="#999999" bordercolorlight="#999999" bordercolordark="#ffffff" bgcolor="#ffffff">
  <tr> 
    <td width="11%"> 
      <div align="center"><a href="../index.htm">首页</a></div>
    </td>
    <td width="11%"> 
      <div align="center"><a href="../aboutme/aboutme.htm" title="自我介绍、相关证件、联系方法等">关于我</a></div>
    </td>
    <td width="11%"> 
      <div align="center"><a href="../job/job.htm" title="就业单位及工作情况的介绍">从业概况</a></div>
    </td>
    <td width="11%"> 
      <div align="center"><a href="../article/article.htm" title="在某杂志、论坛上发表过的技术文章">发表论著</a></div>
    </td>
    <td width="11%"> 
      <div align="center"><a href="../prize/prize.htm" title="技术成果奖励情况及证书">获奖情况</a></div>
    </td>
    <td width="11%"> 
      <div align="center"><a href="http://www.zy816.com/liute" title="-= 特柴之窗 =-    这是我独自为本单位做的动态网站" target="_blank">单位网站</a></div>
    </td>
    <td width="11%"> 
      <div align="center"><a href="http://www.zy816.com" title="共享的世界    个人动态网站，由于精力有限，暂时如此了！" target="_blank">本人网站</a></div>
    </td>
    <td width="11%"> 
      <div align="center"><a href="../lzyjl.doc" title="通常格式的表格简历" target="_blank">简历</a></div>
    </td>
    <td width="11%"> 
      <div align="center">给我忠告</div>
    </td>
  </tr>
</table>

<BR>
<TABLE cellSpacing=1 cellPadding=0 width="75%" align=center border=0 height="60">
  <TBODY> 
  <TR>
    <td><img src="image/lyb.gif" width="150" height="50"></td>        
    <td> 
      <div align="right">
        <input onclick="window.location='leave.htm'" type="button" name="Button" value="我要留言">
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
                        color=#000000>&nbsp;当前显示第 <%=page%> 页&nbsp;&nbsp;共有 <%=pagecount%> 
            页&nbsp;&nbsp;共有 <%=recordcount%> 条记录</FONT></TD>
                      
          <TD align=right 
              width="45%"></TD>
                    </TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
            
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
              来自:<%=rs("dizhi")%>
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
				  <%end if%>
            &nbsp;&nbsp;<font color="#006600">留言时间: <%=rs("time")%></font> 
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
          <TD><img src="image/icon/dot.gif" width="21" height="10"><font color="#FF0000">&nbsp;lzhiy0816回复:</font></TD>
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
                        height=16 src="image/icon/reply.gif" alt="留言者修改留言" width=16
                        align=absMiddle border=0></a>&nbsp;
                        <%end if%>
                        <%if rs("email")<>"" then%>
                        <a
                        href="mailto:<%=rs("email")%>"><img height=16
                        src="image/icon/email.gif" alt="给他发邮件" width=16 align=absMiddle
                        border=0></a>&nbsp;
                        <%end if%>
                        <%if rs("zhuye")<>"http://" then%>
                        <a
                        href="<%=rs("zhuye")%>"
                        target=_blank><img height=16 src="image/icon/home.gif" alt="访问个人主页"
                        width=16 align=absMiddle border=0></a>&nbsp;
                        <%end if%> 
                        <a 
                        href="dellogin.asp?id=<%=rs("id")%>"><img
                        height=16 src="image/icon/recycle.gif" alt="管理回复" width=16
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
                        color=#000000>&nbsp;当前显示第 <%=page%> 页&nbsp;&nbsp;共有 <%=pagecount%> 
            页&nbsp;&nbsp;共有 <%=recordcount%> 条记录 </FONT></TD>
                      
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
                        <img src="image/icon/head.gif" width="13" height="11" alt="首页">
                        <img src="image/icon/pre.gif" width="13" height="11" alt="上一页">
                        &nbsp;
                       <font color=#000000>[ <%=page%> ]</font>&nbsp;&nbsp;
                       <a href="leaveword.asp?page=<%=page+1%>">
                        <img src="image/icon/next.gif" width="13" height="11" alt="下一页">
                       </a>
                       <a href="leaveword.asp?page=<%=pagecount%>">
                        <img src="image/icon/end.gif" width="13" height="11" alt="尾页">
                       </a>
                      </TD>
                      <TD width="24%" align=right>
                        <form name="form1" method="post" action="go.asp">
                          <input type="image" border="0" name="imageField" src="image/icon/go.gif" alt="跳至">
                          第 
                          <input type="text" name="page" size="6" maxlength="6">
                          页 
                        </form>
                      </TD>
                      <%elseif page=pagecount and not page=1 then%>
                      <TD align=left width="76%" >&nbsp;
                       <a href="leaveword.asp?page=1">
                        <img src="image/icon/head.gif" width="13" height="11" alt="首页">
                       </a>
                       <a href="leaveword.asp?page=<%=page-1%>">
                        <img src="image/icon/pre.gif" width="13" height="11" alt="上一页">
                       </a>&nbsp;
                       <font color=#000000>[ <%=page%> ]</font>&nbsp;&nbsp;
                        <img src="image/icon/next.gif" width="13" height="11" alt="下一页">
                        <img src="image/icon/end.gif" width="13" height="11" alt="尾页">
                      </TD>
                      <TD width="24%" align=right>
                        <form name="form1" method="post" action="go.asp">
                          <input type="image" border="0" name="imageField" src="image/icon/go.gif" alt="跳至">
                          第 
                          <input type="text" name="page" size="6" maxlength="6">
                          页 
                        </form>
                      </TD>
                      <%elseif page<1 then%>
                        <TD align=left width="76%" ><div class="font"><font color=red>没有任何记录！</font></div></td>
                        <TD width="24%" align=right>&nbsp;</td>
                      <%elseif page>pagecount then%>
                        <TD align=left width="76%" ><div class="font"><font color=red>没有任何记录！</font></div></td>
                        <TD width="24%" align=right>&nbsp;</td>
                      <%elseif page=1 and page=pagecount then%>
                      <TD align=left width="76%" >&nbsp;
                       
                      </TD>
                      <TD width="24%" align=right>&nbsp;
                       
                      </TD>
                      <%else%>
                      <TD align=left width="76%" >&nbsp;
                       <a href="leaveword.asp?page=1">
                        <img src="image/icon/head.gif" width="13" height="11" alt="首页">
                       </a>
                       <a href="leaveword.asp?page=<%=page-1%>">
                        <img src="image/icon/pre.gif" width="13" height="11" alt="上一页">
                       </a>&nbsp;
                       <font color=#000000>[ <%=page%> ]</font>&nbsp;&nbsp;
                       <a href="leaveword.asp?page=<%=page+1%>">
                        <img src="image/icon/next.gif" width="13" height="11" alt="下一页">
                       </a>
                       <a href="leaveword.asp?page=<%=pagecount%>">
                        <img src="image/icon/end.gif" width="13" height="11" alt="尾页">
                       </a>
                      </TD>
                      <TD width="24%" align=right>
                        <form name="form1" method="post" action="go.asp">
                          <input type="image" border="0" name="imageField" src="image/icon/go.gif" alt="跳至">
                          第 
                          <input type="text" name="page" size="6" maxlength="6">
                          页 
                        </form>
                      </TD>
                      <%end if%>
                    </TR>

                    </TBODY></TABLE></TD></TR></TBODY></TABLE>
            
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
