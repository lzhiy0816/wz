<%@language=vbscript%>
<%
''''''''----====    lzhiy0816 本地站点     ====----''''''''
''''''''         http://lzhiy0816.vicp.net         ''''''''
''''''''         http://lzhiy0816.3322.org         ''''''''

''''''''----====    lzhiy0816 网上空间     ====----''''''''
''''''''      http://free.glzc.net/lzhiy0816       ''''''''
''''''''      http://lzhiy0816.go.nease.net        ''''''''
%>

<%if not session("checked")="yes" then 
id=request.querystring("id")
response.Redirect "mdflogin.asp?id="&id&""
else
%>
<HTML><!-- #BeginTemplate "/Templates/index.dwt" -->
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<TITLE>-= 柳特之窗 =-</TITLE>

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
</HEAD>
<BODY topmargin="0">
<!-- #BeginEditable "body" --><!-- #EndEditable --> 
<DIV align=center>
    
  <table width="760" height="10" bgcolor="#ffffff" align="center">
    <tr>
	  <td>
	  </td>
	</tr>
  </table>
  <TABLE border=0 cellPadding=0 cellSpacing=0 width="760" bgcolor="#ffffff">
    <TR> 
      <TD height=80 width="200"> 
        <DIV align=center class=CompanyName><img src="image/lflogo.gif" width="200" height="80"></DIV>
      </TD>
      <TD background="image/lflogobg.gif"> 
        <div align="center"><b><img src="image/TechaiImg.gif" width="390" height="65"> 
          </b></div>
      </TD>
      <TD background="image/lflogobg.gif"> 
        <table width="100%">
          <tr> 
            <td height="20"> 
              <div align="left"><img src="image/home.gif" width="20" height="20"></div>
            </td>
            <td height="20" width="70"> 
              <div align="left"><b><font color="#336699"><a href="http://www.liute.com/#" onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('http://www.liute.com/');return(fulse);" style="BEHAVIOR;url(#defaule#homepage)">设为首页</a></font></b></div>
            </td>
          </tr>
          <tr>
		  <tr> 
            <td height="5"> 
              <div align="center"></div>
            </td>
            <td height="5"> 
              <div align="right"></div>
            </td>
          </tr> 
            <td height="20"> 
              <div align="left"><img src="image/fai.gif" width="20" height="20"></div>
            </td>
            <td height="20"> 
              <div align="left"><font color="#336699"><b><a href="javascript:window.external.AddFavorite('http://www.liute.com/','-= 柳特之窗 =-')">加入收藏</a></b></font></div>
            </td>
          </tr>
        </table>
      </TD>
    </TR>
    <TR> 
      <TD colSpan=3> 
        <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
          <TR> 
            <TD bgColor=#cccccc height=1> 
              <DIV align=center><IMG height=1 src="./image/pixel.gif" width=1></DIV>
            </TD>
          </TR>
        </TABLE>
      </TD>
    </TR>
    <TR bgcolor="#f7f7f7"> 
      <TD colSpan=3 height=12> 
        <div align="left">
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td height="1"></td>
              <td width="60" height="1"></td>
              <td width="1" height="1"></td>
              <td width="60" height="1"></td>
              <td width="1" height="1"></td>
              <td width="60" height="1"></td>
              <td width="1" height="1"></td>
              <td width="60" height="1"></td>
              <td width="1" height="1"></td>
              <td width="60" height="1"></td>
              <td width="1" height="1"></td>
              <td width="60" height="1"></td>
            </tr>
            <tr> 
              <td height="12"> 
                <div align="center"></div>
              </td>
              <td height="12" width="60"> 
                <div align="center"><!-- #BeginEditable "index" --><a href="index.htm"><font color="#336699"><b>首　页</b></font></a><!-- #EndEditable --></div>
              </td>
              <td height="12" width="2" bgcolor="#cccccc"></td>
              <td height="12" width="60"> 
                <div align="center"><!-- #BeginEditable "intro" --><A href="about/introduce.htm"><font color="#3169a5"><b>简　介</b></font></A><!-- #EndEditable --></div>
              </td>
              <td height="12" width="2" bgcolor="#cccccc"> 
                <div align="center"></div>
              </td>
              <td height="12" width="60"> 
                <div align="center"><!-- #BeginEditable "prod" --><A href="product/power.htm"><font color="#3169a5"><b>产　品</b></font></A><!-- #EndEditable --></div>
              </td>
              <td height="12" width="2" bgcolor="#cccccc"> 
                <div align="center"></div>
              </td>
              <td height="12" width="60"> 
                <div align="center"><!-- #BeginEditable "serve" --><a href="serve/userbook.htm"><font color="#3169a5"><b>服　务</b></font></A><!-- #EndEditable --></div>
              </td>
              <td height="12" width="2" bgcolor="#cccccc"> 
                <div align="center"></div>
              </td>
              <td height="12" width="60"> 
                <div align="center"><!-- #BeginEditable "email" --><a href="Email.htm"><font color="#3169a5"><b>联　系</b></font></a><!-- #EndEditable --></div>
              </td>
              <td height="12" width="2" bgcolor="#cccccc"> 
                <div align="center"></div>
              </td>
              <td height="12" width="60"> 
                <div align="center"><A href="leaveword.asp"><font color="#3169a5"><b>留言板</b></font></A></div>
              </td>
            </tr>
            <tr> 
              <td height="1"><img src="image/pixel.gif" width="1" height="1"></td>
              <td width="60" height="1"></td>
              <td width="1" height="1"></td>
              <td width="60" height="1"></td>
              <td width="1" height="1"></td>
              <td width="60" height="1"></td>
              <td width="1" height="1"></td>
              <td width="60" height="1"></td>
              <td width="1" height="1"></td>
              <td width="60" height="1"></td>
              <td width="1" height="1"></td>
              <td width="60" height="1"></td>
            </tr>
          </table>
        </div>
      </TD>
    </TR>
    <TR> 
      <TD colSpan=3 bgcolor="#f7f7f7"> 
        <DIV align=left> 
          <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
            <TR> 
              <TD bgColor=#CCCCCC height=2> 
                <DIV align=center><IMG height=1 src="./image/pixel.gif" width=1></DIV>
              </TD>
            </TR>
          </TABLE>
        </DIV>
      </TD>
    </TR>
  </TABLE>
</DIV>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width="760" bgcolor="#ffffff">
  <TR> 
    <!-- #BeginEditable "文本区" --> 
<!--#include file="conn.asp"-->
<%id=request.querystring("id")
exec="select * from userleave where id="&request.querystring("id")
set rs=server.createobject("adodb.recordset")
rs.open exec,conn
huifu=rs("huifu")
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

            <TABLE cellSpacing=1 cellPadding=4 width="95%" align=center 
border=0>
              <TBODY>
              <TR>
                <TD vAlign=top align=middle><br><br><INPUT onclick="window.location='leaveword.asp'" type=button value=查看留言 name=Button3>
                </TD></TR></TBODY></TABLE><!-- form.cjs -->
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
            
        <TABLE cellSpacing=1 cellPadding=0 width="95%" align=center border=1 bordercolor="#cccccc" height="30">
          <TBODY> 
          <TR bgColor=#e7e7e7>
            <TD align=middle colSpan=2><font color="#009900"><b>请 修 改 您 的 留 言 及 相 关 信 息</b></font></TD>
          </TR></TBODY></TABLE>
        <TABLE cellSpacing=1 cellPadding=5 width="95%" align=center border=0 
            bgColor=#CCCCCC>
          <TBODY> 
          
		      <%if huifu<>"" then%><TR bgColor=#f5f5f5>
              <%else%><TR bgColor=#ffffff>
              <%end if%>
                <TD align=middle width="20%">头　　像：</TD>
                <TD vAlign=top width="80%"><input type=hidden name=face value="<%=rs("face")%>">
				 <a href=#f onclick="window.open('facelist.htm','tmpWin','menubar=0,scrollbars=yes');"> 
                  <img src="face/<%=rs("face")%>.gif" name="picShow" border="0" alt="选择一个新头像"> </a>&nbsp;&nbsp;&nbsp;&nbsp; 
                  <INPUT class=text type=hidden maxLength=16 size=10 name=aicq>
                  <font color="green">（点头像选择其他头像）</font>
    			</TD>
              </TR>
		      <%if huifu<>"" then%><TR bgColor=#ffffff>
              <%else%><TR bgColor=#f5f5f5>
              <%end if%>
                <TD align=middle width="20%">来自哪里：</TD>
                <TD vAlign=top width="80%"><INPUT class="text" maxLength="15" 
                  size="15" name="dizhi" value="<%=rs("dizhi")%>"> &nbsp;&nbsp; <INPUT class=text type=hidden
                  maxLength=15 size=10 name=icq> </TD></TR>
              <%if huifu<>"" then%><TR bgColor=#f5f5f5>
              <%else%><TR bgColor=#ffffff>
              <%end if%>
                <TD align=middle width="20%">电子邮件：</TD>
                <TD vAlign=top width="80%"><INPUT class="text" maxLength="52" 
                  size="52" name="email" value="<%=rs("email")%>"> &nbsp;&nbsp; <INPUT class=text type=hidden
                  maxLength=25 size=10 name=oicq> </TD></TR>
              <%if huifu<>"" then%><TR bgColor=#ffffff>
              <%else%><TR bgColor=#f5f5f5>
              <%end if%>
                <TD align=middle width="20%">您的网址：</TD>
                <TD vAlign=top width="80%">
                <INPUT class="text" maxLength="70"
                  size="52" value="<%=rs("zhuye")%>" name="zhuye">
                   </TD></TR>
              <%if huifu<>"" then%><TR bgColor=#f5f5f5>
              <%else%><TR bgColor=#ffffff>
              <%end if%>
                <TD align=middle width="20%">表情图标：</TD>
                <TD>
              <table width="80%" height="80%">
                <tr>
                  <td width=15>
				    <INPUT type=radio value=1 name=icon <%if rs("icon")=1 then%>CHECKED<%end if%>>
				  </td>
                  <td width=20>
				    <IMG height=15 alt="没有表情" src="image/icon/icon1.gif" width=15 align=ABSCENTER>
				  </td>
                  <td width=15>
				    <INPUT type=radio value=2 <%if rs("icon")=2 then%>CHECKED<%end if%> name=icon>
				  </td>
                  <td width=20>
				    <IMG height=15 alt="快看！" src="image/icon/icon2.gif" width=15 align=ABSCENTER>
				  </td>
                  <td width=15>
				    <INPUT type=radio value=3 <%if rs("icon")=3 then%>CHECKED<%end if%> name=icon>
				  </td>
                  <td width=20>
				    <IMG height=15 alt="嗨嗨，我有主意了！" src="image/icon/icon3.gif" width=15 align=ABSCENTER>
				  </td>
                  <td width=15>
				    <INPUT type=radio value=4 <%if rs("icon")=4 then%>CHECKED<%end if%> name=icon>
				  </td>
                  <td width=20>
				   <IMG height=15 alt=警告！ src="image/icon/icon4.gif" width=15 align=ABSCENTER>
				  </td>
                  <td width=15>
				    <INPUT type=radio value=5 <%if rs("icon")=5 then%>CHECKED<%end if%> name=icon>
				  </td>
                  <td width=20>
				    <IMG height=15 alt=问题！ src="image/icon/icon5.gif" width=15 align=ABSCENTER>
				  </td>
                  <td width=15>
				    <INPUT type=radio value=6 <%if rs("icon")=6 then%>CHECKED<%end if%> name=icon>
				  </td>
                  <td width=20>
				    <IMG height=15 alt=哇！真高兴！ src="image/icon/icon6.gif" width=15 align=ABSCENTER>
				  </td>
                  <td width=15>
				    <INPUT type=radio value=7 <%if rs("icon")=7 then%>CHECKED<%end if%> name=icon>
				  </td>
                  <td width=20>
				    <IMG height=15 alt=好难过！ src="image/icon/icon7.gif" width=15 align=ABSCENTER>
				  </td>
                </tr>
                <tr>
                  <td width=15>
                    <input type=radio value=8 <%if rs("icon")=8 then%>CHECKED<%end if%> name=icon>
                  </td>
                  <td width=20>
				    <img height=15 alt=这下惨了！ src="image/icon/icon8.gif" width=15 align=ABSCENTER>
				  </td>
                  <td width=15>
                    <input type=radio value=9 <%if rs("icon")=9 then%>CHECKED<%end if%> name=icon>
                  </td>
                  <td width=20>
				    <img height=15 alt=含情 src="image/icon/icon9.gif" width=15 align=ABSCENTER>
				  </td>
                  <td width=15>
                    <input type=radio value=10 <%if rs("icon")=10 then%>CHECKED<%end if%> name=icon>
                  </td>
                  <td width=20>
				    <img height=15 alt=嘻嘻！ src="image/icon/icon10.gif" width=15 align=ABSCENTER>
				  </td>
                  <td width=15>
                    <input type=radio value=11 <%if rs("icon")=11 then%>CHECKED<%end if%> name=icon>
                  </td>
                  <td width=20>
				    <img height=15 alt=呵呵！ src="image/icon/icon11.gif" width=15 align=ABSCENTER>
				  </td>
                  <td width=15>
                    <input type=radio value=12 <%if rs("icon")=12 then%>CHECKED<%end if%> name=icon>
                  </td>
                  <td width=20>
				    <img height=15 alt=眨眼－－对你有意思哦！ src="image/icon/icon12.gif" width=15 align=ABSCENTER>
				  </td>
                  <td width=15>
                    <input type=radio value=13 <%if rs("icon")=13 then%>CHECKED<%end if%> name=icon>
                  </td>
                  <td width=20>
				    <img height=15 alt=强烈反对 src="image/icon/icon13.gif" width=15 align=ABSCENTER>
				  </td>
                  <td width=15>
                    <input type=radio value=14 <%if rs("icon")=14 then%>CHECKED<%end if%> name=icon>
                  </td>
                  <td width=20>
				    <img height=15 alt=同意！ src="image/icon/icon14.gif" width=15 align=ABSCENTER>
				  </td>
                </tr>
              </table>
            </TD>
			  </TR>
              <%if huifu<>"" then%><TR bgColor=#ffffff>
              <%else%><TR bgColor=#f5f5f5>
              <%end if%>
                <TD align=middle width="20%">
                 您的留言：
                </TD>
                <TD width="80%"><TEXTAREA name=liuyanneirong rows=6 cols=51><%=rs("liuyanneirong")%></TEXTAREA>
                <br>
                 <font color="#009900">
                 〈Br〉：表示换行，也可按“Enter”回车键换行
                 </font>
                </TD>
                </TR>
              <%if huifu<>"" then%>
              <TR bgColor=#f5f5f5 height="100">
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
            <CENTER><BR>
          <input class=buttom onclick="return check(this.form,1000)" type=submit value=确定 name=submit>
          <INPUT class=buttom type=reset value=取消 name=reset> 
          <input type="hidden" name="id" value="<%=request.querystring("id")%>">
         </CENTER></FORM>
	</TD>
<%
session("checked")=""
session("check")=""
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
    <!-- #EndEditable -->
    <TD bgColor=#cccccc height=500 width=2><IMG height=1 src="./image/pixel.gif" width=1></TD>
    <!-- #BeginEditable "栏目区" --> 
    <TD bgColor=#f7f7f7 vAlign=top width=150 align="center"> 
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <DIV align=right> 
        <table border="0" cellpadding="0" cellspacing="0" width="80" align="center">
          <!-- fwtable fwsrc="Welcome.png" fwbase="Welcome.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
          <tr> 
            <td><img src="image/spacer.gif" width="80" height="1" border="0" name="undefined_2"></td>
            <td><img src="image/spacer.gif" width="1" height="1" border="0" name="undefined_2"></td>
          </tr>
          <tr> 
            <td><img name="Welcome_r1_c1" src="image/Welcome_r1_c1.gif" width="80" height="82" border="0"></td>
            <td><img src="image/spacer.gif" width="1" height="82" border="0" name="undefined_2"></td>
          </tr>
          <tr> 
            <td><img name="Welcome_r2_c1" src="image/Welcome_r2_c1.gif" width="80" height="52" border="0"></td>
            <td><img src="image/spacer.gif" width="1" height="52" border="0" name="undefined_2"></td>
          </tr>
          <tr> 
            <td><img name="Welcome_r3_c1" src="image/Welcome_r3_c1.gif" width="80" height="27" border="0"></td>
            <td><img src="image/spacer.gif" width="1" height="27" border="0" name="undefined_2"></td>
          </tr>
          <tr> 
            <td><img name="Welcome_r4_c1" src="image/Welcome_r4_c1.gif" width="80" height="53" border="0"></td>
            <td><img src="image/spacer.gif" width="1" height="53" border="0" name="undefined_2"></td>
          </tr>
          <tr> 
            <td><img name="Welcome_r5_c1" src="image/Welcome_r5_c1.gif" width="80" height="51" border="0"></td>
            <td><img src="image/spacer.gif" width="1" height="51" border="0" name="undefined_2"></td>
          </tr>
          <tr> 
            <td><img name="Welcome_r6_c1" src="image/Welcome_r6_c1.gif" width="80" height="80" border="0"></td>
            <td><img src="image/spacer.gif" width="1" height="80" border="0" name="undefined_2"></td>
          </tr>
          <tr> 
            <td><img name="Welcome_r7_c1" src="image/Welcome_r7_c1.gif" width="80" height="55" border="0"></td>
            <td><img src="image/spacer.gif" width="1" height="55" border="0" name="undefined_2"></td>
          </tr>
        </table>
        <P align=left>&nbsp;</P>
      </DIV>
	</TD>
  
    <!-- #EndEditable -->
  </TR>
  
</TABLE>
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#ffffff">
  <tr>
    <TD colSpan=3 valign="top"> 
      <DIV align=right> 
        <TABLE border=0 cellPadding=0 cellSpacing=0 width="80%">
          <TR bgColor=#009999> 
            <TD height=1 bgcolor="#cccccc"> 
              <DIV align=right><IMG height=1 src="file:///D|/./image/pixel.gif" width=1></DIV>
            </TD>
          </TR>
        </TABLE>
      </DIV>
      <DIV align=center> <BR>
        <table width="60%">
          <tr> 
            <td width="80"> 
              <div align="center"><b><font color="#990000" size="3" face="黑体">柳特之窗</font></b></div>
            </td>
            <td><font color="#FF0000"><marquee><b><i><font size="3">欢迎您经常光临！热枕为您服务！</font></i></b></marquee></font></td>
          </tr>
        </table>
      </DIV>
      <DIV align=right> 
        <TABLE border=0 cellPadding=0 cellSpacing=0 width="90%">
          <TR bgColor=#949231> 
            <TD height=2> 
              <DIV align=right><IMG height=1 src="file:///D|/./image/pixel.gif" width=1></DIV>
            </TD>
          </TR>
        </TABLE>
        <DIV align=center><font color="#FF9900">
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
		</font> <BR>
        </DIV>
        <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
          <TR bgColor=#006500> 
            <TD height=3> 
              <DIV align=right><IMG height=1 src="file:///D|/./image/pixel.gif" width=1></DIV>
            </TD>
          </TR>
        </TABLE>
      </DIV>
      <div align="center"><font color=#3169a5 face="Arial, Helvetica, sans-serif" size=2>Copyright&copy; 
        <b>LiuTe</b> All rights reserved.</font><font color=#3169a5> <br>
        柳州市特种柴油机厂　版权所有</font></div>
    </TD>
  
  </tr>
</table>
<table width="760" height="15" bgcolor="#ffffff" align="center">
  <tr>
	  <td>
	  </td>
	</tr>
</table>

</BODY>
<!-- #EndTemplate --></HTML>
<HTML>
<%end if%>
