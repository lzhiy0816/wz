<%
''''''''----====    lzhiy0816 ����վ��     ====----''''''''
''''''''         http://lzhiy0816.vicp.net         ''''''''
''''''''         http://lzhiy0816.3322.org         ''''''''

''''''''----====    lzhiy0816 ���Ͽռ�     ====----''''''''
''''''''      http://free.glzc.net/lzhiy0816       ''''''''
''''''''      http://lzhiy0816.go.nease.net        ''''''''
%>

<!--#include file="conn.asp"-->
<%    ''''''''====�����ܣ��û�����������ӵ����ݿ⡷(lzhiy0816 design)====''''''''

xingming=request.form("xingming")    ''''''''�ӱ���ȡ����'''''''
password=request.form("password")
mingchen=request.form("mingchen")
shuoming=request.form("shuoming")
lianjie=request.form("lianjie")
logoimg=request.form("logoimg")

exec="select xingming from logo"    ''''''''��Ѱxingming�ֶμ�¼(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1
xingming1=""                        ''''''''�Ƚ�xingming��Ϊ�գ�lzhiy0816 design)''''''''
do while not rs.eof                 ''''''''����ύ���û����ڼ�¼����ڣ���xingming1��Ϊ��¼ֵ��lzhiy0816 design)''''''''
  if xingming=rs("xingming") then
  xingming1=rs("xingming")
  end if
rs.movenext
loop
if xingming1<>"" then
response.redirect "logoback.asp?err="&xingming1&""
rs.close
set rs=nothing
else

''''''''������ӵ����ݿ�'''''''
exec="insert into logo(xingming,password,mingchen,shuoming,lianjie,logoimg) values('"+xingming+"','"+password+"','"+mingchen+"','"+shuoming+"','"+lianjie+"','"+logoimg+"')"
conn.execute exec
conn.close
set conn=nothing
response.redirect "logo.asp"
end if
%>
