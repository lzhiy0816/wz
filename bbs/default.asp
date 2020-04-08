<!-- #include file="setup.asp" -->
<%
id=Request("id")
if isnumeric(""&id&"") then
bbsname=Conn.Execute("Select bbsname From bbsconfig where id="&id)(0)
url="ShowForum.htm?forumid="&id&""
Response.Cookies("forumid")=id
else
bbsname=clubname
url="main.htm"
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html;charset=gb2312">
<meta name="keywords" content="BBSxp,Board,yuzi,ASP,Access,MSSQL,FORUM">
<meta name="description" content="<%=clubname%> - Powered by BBSxp">
<title><%=bbsname%> - Powered By BBSxp</title>
</head>
<link href="images/skins/<%=Request.Cookies("skins")%>/bbs.css" rel="stylesheet">
<link REL="SHORTCUT ICON" href="images/ybb.ico">
<script>
if(self!=top){top.location=self.location;}
function switchSysBar(){
if (switchPoint.innerText==3){
switchPoint.innerText=4
document.all("frmTitle").style.display="none"
}else{
switchPoint.innerText=3
document.all("frmTitle").style.display=""
}}
</script>

<style type="text/css">.navPoint {COLOR: white; CURSOR: hand; FONT-FAMILY: Webdings; FONT-SIZE: 9pt}
</style>
<body style="MARGIN: 0px" scroll=no>

<table border="0" cellPadding="0" cellSpacing="0" height="100%" width="100%">
  <tr>
    <td align="middle" id="frmTitle" noWrap vAlign="center" name="frmTitle">
    
    
    <iframe frameBorder="0" id="carnoc" name="carnoc" scrolling=no src="left.htm" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 170px; Z-INDEX: 2">
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

    <iframe frameBorder="0" id="main" name="main" scrolling="yes" src="<%=url%>" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1">
    </iframe>

    </td>
  </tr>
</table>
</html>
<script>if (window.screen.width<'1024'){switchSysBar()}</script>
<%responseend%>