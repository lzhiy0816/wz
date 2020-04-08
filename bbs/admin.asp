<!-- #include file="setup.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")

if Conn.Execute("Select membercode From [user] where username='"&Request.Cookies("username")&"'")(0)<5 then error("<li>你不是社区区长，无权登录后台管理")


if Request("menu")="out" then
session("pass")=""
response.redirect "Default.asp"
end if
%>

<title><%=clubname%>后台管理 - Powered By BBSxp</title>
<META http-equiv=Content-Type content=text/html;charset=gb2312>
<link href=images/skins/<%=Request.Cookies("skins")%>/bbs.css rel=stylesheet>



<%
select case Request("menu")
case ""
index
case "left"
left
case "top"
top2
case "login"
login

case "pass"
pass



end select


sub top2


%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg" topmargin=0  rightmargin=0  leftmargin=0><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
  <tr>
  <td width="100%" class=a1 height="20">
  
 <center>
<B><FONT face=幼圆>第五代 BBS 系统 -- BBSxp</FONT> <FONT color=ffff66>安全</FONT> <FONT 
color=ff0033>快速</FONT> <FONT color=33ff33>方便 </FONT><FONT color=0000ff>可靠 
</FONT><FONT color=000000>可升级</FONT></B> 
  </td>
  </tr>
  </table>

<%
end sub




sub pass

session("pass")=md5(""&Request("pass")&"")
if adminpassword<>session("pass") then error2("社区管理密码错误")


log("登录后台管理")



conn.execute("delete from [log] where logtime<"&SqlNowString&"-7")


Dim theInstalledObjects(18)
theInstalledObjects(0) = "MSWC.AdRotator"
theInstalledObjects(1) = "MSWC.BrowserType"
theInstalledObjects(2) = "MSWC.NextLink"
theInstalledObjects(3) = "MSWC.Tools"
theInstalledObjects(4) = "MSWC.Status"
theInstalledObjects(5) = "MSWC.Counters"
theInstalledObjects(6) = "IISSample.ContentRotator"
theInstalledObjects(7) = "IISSample.PageCounter"
theInstalledObjects(8) = "MSWC.PermissionChecker"
theInstalledObjects(9) = "Scripting.FileSystemObject"
theInstalledObjects(10) = "adodb.connection"
theInstalledObjects(11) = "SoftArtisans.FileUp"
theInstalledObjects(12) = "SoftArtisans.FileManager"
theInstalledObjects(13) = "JMail.Message"
theInstalledObjects(14) = "CDONTS.NewMail"
theInstalledObjects(15) = "Persits.MailSender"
theInstalledObjects(16) = "LyfUpload.UploadFile"
theInstalledObjects(17) = "Persits.Upload.1"
theInstalledObjects(18) = "w3.upload"




%>
<link href=images/skins/<%=Request.Cookies("skins")%>/bbs.css rel=stylesheet>






<html>
<head>


<title>BBSxp论坛--管理页面</title>

<table cellpadding="2" cellspacing="1" border="0" width="95%" align=center class="a2">
<tr>
<td class="a1" colspan=2 height=25 align="center"><b>论坛信息统计</b></td>
</tr>
<tr class=a4>
<td height=23 colspan=2>

系统信息：主题数 <B><%=conn.execute("Select count(id)from [forum]")(0)%></B> 回帖数 <B><%=conn.execute("Select count(id)from [reforum]")(0)%></B> 用户数 <B><%=conn.execute("Select count(id)from [user]")(0)%></B>

</td></tr>
<tr class=a4>
<td width="50%" height=23>服务器类型：<%=Request.ServerVariables("OS")%> ( IP:<%=Request.ServerVariables("LOCAL_ADDR")%> )</td>
<td width="50%">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
</tr>
<tr class=a4>
<td width="50%" height=23>服务器软件的名称：<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
<td width="50%">ACCESS 数据库路径：<a target="_blank" href="<%=datapath%><%=datafile%>"><%=datapath%><%=datafile%></a></td>
</tr>
<tr class=a4>
<td width="50%" height=23>FSO 文本读写：<%If Not IsObjInstalled(theInstalledObjects(9)) Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%></td>
<td width="50%">SA-FileUp 文件上传：<%If Not IsObjInstalled(theInstalledObjects(11)) Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%></td>
</tr>
<tr class=a4>
<td width="50%" height=23>Jmail 组件支持：<%If Not IsObjInstalled(theInstalledObjects(13)) Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%></td>
<td width="50%">CDONTS 组件支持：<%If Not IsObjInstalled(theInstalledObjects(14)) Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%></td>
</tr>
</table>
<br><br>
<table cellpadding="2" cellspacing="1" border="0" width="95%" align=center class="a2">
<tr>
<td class="a1" colspan=2 height=25 align="center"><b>管理快捷方式</b></td>
</tr><tr class="a4"><td width="20%"  height=23>快速查找用户</td>
<td width="80%" height=23>
<form method="post" action="admin_user.asp?menu=user2">
<input size="13" name="username"> <input type="submit" value="立刻查找">
</td></form>
</tr>

<tr class="a4"><td width="20%" height=23>快捷功能链接</td>
<td width="80%" height=23>
<a href="admin_bbs.asp?menu=classs"><font color=000000>建立论坛数据</font></a> | 
<a href="admin_bbs.asp?menu=bbsmanage"><font color=000000>管理论坛资料</font></a> |
<a href="admin_bbs.asp?menu=upclubconfigok"><font color=000000>更新论坛资料</font></a></td>
</tr>
</table>





<p></p>

<p></p>

<script language='javascript'> function jumpto(url) { if (url != '') { window.open(url); } } </script>
<p></p>

<table cellpadding="2" cellspacing="1" border="0" width="95%" align=center class="a2">
<tr class=a4>
<td class="a1" colspan=2 height=25 align="center"><b>BBSxp论坛</b></td>
</tr><tr class=a4><td width="20%">产品开发</td>
<td width="80%">
Yuzi工作室</td>
</tr>
<tr class=a4>
<td width="20%"  height=23>产品负责</td>
<td width="80%">
Yuzi网络科技发展有限公司</td>
</tr>
<tr class=a4>
<td width="20%" height=23>联系方法</td>
<td width="80%">
电话：0595-2251981 网络部：0595-2251981-11<br>
传真：0595-2205903 开发部：0595-2251981-17<br>
网址：http://www.bbsxp.com<br>
地址：中国 福建省 泉州市 温陵南路 商业城 231 号</td>
</tr>
<tr class=a4>
<td width="20%" height=23>插件开发</td>
<td width="80%">
BBSxp插件组织（BBSxp Plus Organization）
</td>
</tr>
</table>





<%
htmlend
end sub




sub login





%>


<br><br><br>


<form action="admin.asp" method="post">
<input type="hidden" value="pass" name="menu">


<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class=a2>
<tr>
<td width="328" class=a1 height="25"><div align="center">
登录社区管理</div>
</td></tr><tr>
<td height="19" width="328" valign="top" class=a3>
<div align="center">请输入管理密码:
<input size="15" name="pass" type="password"><br>
<input type="submit" value=" 登录 " name="Submit1">　<input type="reset" value=" 取消 " name="Submit">
</div></td></tr> </table>
</form>
<%
htmlend
end sub

sub left

%>




<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>BODY {
	BACKGROUND: #799ae1; MARGIN: 0px; FONT: 9pt 宋体
}
TABLE {
	BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px
}
TD {
	FONT: 12px 宋体
}
IMG {
	BORDER-RIGHT: 0px; BORDER-TOP: 0px; VERTICAL-ALIGN: bottom; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px
}
A {
	FONT: 12px 宋体; COLOR: #000000; TEXT-DECORATION: none
}
A:hover {
	COLOR: #428eff; TEXT-DECORATION: underline
}
.sec_menu {
	BORDER-RIGHT: white 1px solid; BACKGROUND: #d6dff7; OVERFLOW: hidden; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 1px solid
}
.menu_title {
	
}
.menu_title SPAN {
	FONT-WEIGHT: bold; LEFT: 8px; COLOR: #215dc6; POSITION: relative; TOP: 2px
}
.menu_title2 {
	
}
.menu_title2 SPAN {
	FONT-WEIGHT: bold; LEFT: 8px; COLOR: #428eff; POSITION: relative; TOP: 2px
}
</STYLE>

<SCRIPT>
function aa(Dir)
{tt.doScroll(Dir);Timer=setTimeout('aa("'+Dir+'")',100)}//这里100为滚动速度
function StopScroll(){if(Timer!=null)clearTimeout(Timer)}

function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}

function message()
{
window.open('loading.htm','message')
}
</SCRIPT>

</HEAD>
<center>
  <TBODY>
  <TR>
    <TD vAlign=top>
      <TABLE cellSpacing=0 cellPadding=0 width=158 align=center>
        <TBODY>
        <TR>
          <TD vAlign=bottom height=42><IMG height=38 
            src="images/title.gif" width=158> </TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=158 align=center>
        <TBODY>
        <TR>
          <TD class=menu_title onmouseover="this.className='menu_title2';" 
          onmouseout="this.className='menu_title';" 
          background=images/title_bg_quit.gif height=25><SPAN><B>
          <font color="215DC6"><%=Request.Cookies("username")%></font></B> | 
          
          
          
<a target=_top href=?menu=out><font color=215DC6><B>退出</a>
</font></a></B></SPAN></TD></TR>

<tr>
<td align="center" onMouseOver=aa('up') onMouseOut=StopScroll()><font face="Webdings" color="#FFFFFF">5</font>
</td>
</tr>

</TBODY></TABLE>
          
<script>
var he=document.body.clientHeight-105
document.write("<div id=tt style=height:"+he+";overflow:hidden>")
</script>
<TABLE cellSpacing=0 cellPadding=0 width=158 align=center>
        <TBODY>
        <TR>
          <TD class=menu_title id=menuTitle1 
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(0) 
          onmouseout="this.className='menu_title';"  style=cursor:hand
          background=images/admin_left_3.gif height=25><SPAN>用户管理
          </SPAN></TD></TR>
        <TR>
          <TD id=submenu0 style="DISPLAY: none">
            <DIV class=sec_menu style="WIDTH: 158px">
            <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
              <TBODY>
              <TR>
                <TD height=20>
               <a target="main" href="admin_user.asp?menu=showall"><font color="000000">注册用户管理</font></a></TD></TR>
               
                <TD height=20>
               <a target="main" href="admin_user.asp?menu=user"><font color="000000">用户发放工资</font></a></TD></TR>
              </TBODY></TABLE></DIV>
            <DIV style="WIDTH: 158px">
            <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
              <TBODY>
              <TR>
                <TD height=20></TD></TR></TBODY></TABLE></DIV></TD></TR></TBODY></TABLE>
          
          
          
          

                
                
                
      <TABLE cellSpacing=0 cellPadding=0 width=158 align=center>
        <TBODY>
        <TR>
          <TD class=menu_title id=menuTitle1 
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(3) 
          onmouseout="this.className='menu_title';"  style=cursor:hand
          background=images/admin_left_4.gif height=25><SPAN>论坛管理 
          </SPAN></TD></TR>
        <TR>
          <TD id=submenu3 style="DISPLAY: none">
            <DIV class=sec_menu style="WIDTH: 158px">
            <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
              <TBODY>
              <TR>
                <TD height=20>
                <a target="main" href="admin_bbs.asp?menu=classs"><font color="000000">
                建立论坛数据</font></a></TD></TR>
              <TR>
                <TD height=20>
                <a target="main" href="admin_bbs.asp?menu=bbsmanage"><font color="000000">管理论坛资料</font></a></TD></TR>
              <TR>
                <TD height=20>
                <a target="main" href="admin_bbs.asp?menu=applymanage"><font color="000000">管理申请论坛</font></a></TD></TR>

                
              <TR>
                <TD height=20>
                <a target="main" href="admin_bbs.asp?menu=upclubconfig"><font color="000000">更新论坛资料</font></a></TD>
                </TR>
                
              </TBODY></TABLE></DIV>
            <DIV style="WIDTH: 158px">
            <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
              <TBODY>
              <TR>
                <TD height=20></TD></TR></TBODY></TABLE></DIV></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=158 align=center>
        <TBODY>
        <tr>
          <TD class=menu_title id=menuTitle1 
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(2) 
          onmouseout="this.className='menu_title';"  style=cursor:hand
          background=images/admin_left_2.gif 
            height=25><SPAN>设置管理</SPAN> </TD>
        </tr>
        <tr>
          <TD id=submenu2 style="DISPLAY: none">
            <DIV class=sec_menu style="WIDTH: 158px">
            <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
              <TBODY>
              <TR>
                <TD height=20>
                <a target="main" href="admin_setup.asp?menu=link"><font color="000000">友情链接管理</font></a></TD>
                </TR>
              <TR>
                <TD height=20>
                <a target="main" href="admin_setup.asp?menu=variable"><font color="000000">社区变量设置</font></a></TD>
                </TR>
              </TBODY></TABLE></DIV>
            <DIV style="WIDTH: 158px">
            <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
              <TBODY>
              <TR>
                <TD height=20></TD></TR></TBODY></TABLE></DIV></TD>
        </tr>
        <TR>
          <TD class=menu_title id=menuTitle1 
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(6) 
          onmouseout="this.className='menu_title';"  style=cursor:hand
          background=images/admin_left_5.gif height=25><SPAN>社区管理</SPAN> 
          </TD></TR>
        <TR>
          <TD id=submenu6 style="DISPLAY: none">
            <DIV class=sec_menu style="WIDTH: 158px">
            <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
              <TBODY>
              
                
              <TR>
                <TD height=19>
                <a target="main" href="admin_club.asp?menu=sendmail"><font color="000000">群发邮件</font></a></TD>
                </TR>
                
              <TR>
                <TD height=20>
                <a target="main" href="admin_club.asp?menu=message"><font color="000000">短讯息管理</font></a></TD>
                </TR>

              <TR>
                <TD height=19>
                <a target="main" href="admin_club.asp?menu=affiche"><font color="000000">发布社区公告</font></a></TD>
                </TR>
              </TBODY></TABLE></DIV>
            <DIV style="WIDTH: 158px">
            <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
              <TBODY>
              <TR>
                <TD height=20></TD></TR></TBODY></TABLE></DIV></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=158 align=center>
        <TBODY>
        <TR>
          <TD class=menu_title id=menuTitle1 
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(4) 
          onmouseout="this.className='menu_title';"  style=cursor:hand
          background=images/admin_left_7.gif height=25><SPAN>数据管理 
            </SPAN></TD></TR>
        <TR>
          <TD id=submenu4 style="DISPLAY: none">
            <DIV class=sec_menu style="WIDTH: 158px">
            <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
              <TBODY>


              <TR>
                <TD height=20>
                <a target="main" href="admin_fso.asp?menu=css"><font color="000000">编辑CSS模板</font></a></TD>
                </TR>
              <TR>
                <TD height=20>
                <a target="main" href="admin_fso.asp?menu=statroom"><font color="000000">统计占用空间</font></a></TD>
                </TR>
              <TR>
                <TD height=20>
                <a target="main" href="admin_fso.asp?menu=files"><font color="000000">上传附件管理</font></a></TD>
                </TR>
              <TR>
                <TD height=20>
                <a target="main" href="admin_fso.asp?menu=bak"><font color="000000">ACCESS数据库</font></a>
                </TD>
                </TR>
                
              </TBODY></TABLE></DIV>
            <DIV style="WIDTH: 158px">
            <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
              <TBODY>
              <TR>
                <TD height=20></TD></TR></TBODY></TABLE></DIV></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width=158 align=center>
        <TBODY>
        <TR>
          <TD class=menu_title id=menuTitle1 
          onmouseover="this.className='menu_title2';" onclick=showsubmenu(5) 
          onmouseout="this.className='menu_title';"  style=cursor:hand
          background=images/admin_left_6.gif height=25><SPAN>其他功能</SPAN> 
          </TD></TR>
        <TR>
          <TD id=submenu5 style="DISPLAY: none">
            <DIV class=sec_menu style="WIDTH: 158px">
            <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
              <TBODY>
              <TR>
                <TD height=20>
                <a target="main" href="admin_other.asp?menu=circumstance"><font color="000000">主机环境变量</font></a></TD>
                </TR>
              <TR>
                <TD height=20>
                <a target="main" href="admin_other.asp?menu=discreteness"><font color="000000">组件支持情况</font></a></TD>
                </TR>
              <TR>
                <TD height=20>
                <a target="main" href="admin_other.asp?menu=log"><font color="000000">系统操作记录</font></a></TD>
                </TR>
              </TBODY></TABLE></DIV>
            </TD></TR></TBODY></TABLE>&nbsp; 

      <TABLE cellSpacing=0 cellPadding=0 width=158 align=center>
        <TBODY>
        <TR>
          <TD class=menu_title id=menuTitle1 
          onmouseover="this.className='menu_title2';" 
          onmouseout="this.className='menu_title';" 
          background=images/admin_left_9.gif height=25><SPAN>BBSxp信息</SPAN> 
          </TD></TR>
        <TR>
          <TD>
            <DIV class=sec_menu style="WIDTH: 158px">
            <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
              <TBODY>
              <TR>
                <TD height=20><BR>版本：BBSxp 2004<BR><BR>版权所有：<BR>Yuzi工作室(<a target="_blank" href="http://www.yuzi.net/"><font color="000000">Yuzi<font color="FF0000">.Net</font></font></a><font color="000000">)<BR><BR>
                </font></TD></TR></TBODY></TABLE></DIV></TD></TR></TBODY></TABLE>&nbsp; 
  </TR></TBODY></TABLE>
  
</div>

<table cellspacing="0" cellpadding="0" width="158" align="center">
<tr>
<td align="center" onMouseOver=aa('Down') onMouseOut=StopScroll()><font face="Webdings" color="#FFFFFF">6</font>
</td>
</tr>
</table>

<%
end sub
sub index


%>


<html>

<link REL="SHORTCUT ICON" href="images/ybb.ico">
<meta http-equiv="Content-Type" content="text/html;charset=gb2312">
<style type="text/css">.navPoint {COLOR: white; CURSOR: hand; FONT-FAMILY: Webdings; FONT-SIZE: 9pt}</style>

<script>
function switchSysBar(){
if (switchPoint.innerText==3){
switchPoint.innerText=4
document.all("frmTitle").style.display="none"
}else{
switchPoint.innerText=3
document.all("frmTitle").style.display=""
}}
</script>
<link href="images/skins/<%=Request.Cookies("skins")%>/bbs.css" rel="stylesheet">


<body scroll=no topmargin=0  rightmargin=0  leftmargin=0>

<table border="0" cellPadding="0" cellSpacing="0" height="100%" width="100%">
  <tr>
    <td align="middle" id="frmTitle" noWrap vAlign="center" name="frmTitle">
    
    
    <iframe frameBorder="0" id="carnoc" name="carnoc" scrolling=no  src="?menu=left" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 170px; Z-INDEX: 2">
    </iframe>
    
    

    </td>
    <td class=a2 style="WIDTH: 9pt">
    <table border="0" cellPadding="0" cellSpacing="0" height="100%">
      <tr>
        <td style="HEIGHT: 100%" onclick="switchSysBar()">
        <font style="FONT-SIZE: 9pt; CURSOR: default; COLOR: #ffffff">
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <span class="navPoint" id="switchPoint" title="关闭/打开左栏">3</span><br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        屏幕切换 </font></td>
      </tr>
    </table>
    </td>
    <td style="WIDTH: 100%">
    
    <iframe frameBorder="0" id="main" name="top" scrolling="no" scrolling="yes" src="?menu=top" style="HEIGHT: 20; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1"></iframe>
    
    <iframe frameBorder="0" id="main" name="main" scrolling="yes" src="?menu=login" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1">
    </iframe></td>
  </tr>
</table>
</html>



<%
end sub

Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function

%>