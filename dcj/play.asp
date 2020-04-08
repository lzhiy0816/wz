<html>
<head>
<title>==在线播放==</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body topmargin="0" leftmargin="0">

  <%@language=vbscript%>
  <%id=request.querystring("id")%>
<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA height=450 
    id=video1 width=500 VIEWASTEXT>
    <param name=_ExtentX value=5503>
    <param name=_ExtentY value=1588>
    <param name=AUTOSTART value=-1>
    <param name=SHUFFLE value=0>
    <param name=PREFETCH value=0>
    <param name=NOLABELS value=0>
<%if id>00 and id<41 then %>
    <param name=SRC value=../dcj1/<%=id%>.rmvb>
<%elseif id>40 and id<71 then %>
    <param name=SRC value=../dcj2/<%=id%>.rmvb>
<% end if %>
    <param name=CONTROLS value=Imagewindow,StatusBar,ControlPanel>
    <param name=CONSOLE value=RAPLAYER>
    <param name=LOOP value=0>
    <param name=NUMLOOP value=0>
    <param name=CENTER value=0>
    <param name=MAINTAINASPECT value=0>
    <param name=BACKGROUNDCOLOR value=#000000>
    </OBJECT>
    </body>
</html>
