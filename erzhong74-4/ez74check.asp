<%
''''''''====《柳州市二中74-4班 论坛密码检查》(lzhiy0816 design)====''''''''
dim inputxm2,inputpw2    ''''''''获取lzhiy0816数据(lzhiy0816 design)''''''''
inputxm2=request.form("ez74xingming")
inputpw2=request.form("ez74password")

''''''''如果是本班同学则进入《柳州市二中74-4班 论坛》的主界面(lzhiy0816 design)''''''''
if inputxm2="erzhong74-4"  and inputpw2="19741979" then
session("checked")="ez74yes"
session("check")="ez74right"
response.Redirect "index.asp"

''''''''如果不是lzhiy0816回到《柳州市二中74-4班 论坛》登录入口(lzhiy0816 design)''''''''
else
session("checked")="ez74no"
session("check")="ez74wrong"
response.Redirect "ez74login.asp"
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
 
