<%
''''''''====《自动化79密码检查》(lzhiy0816 design)====''''''''
dim inputxm2,inputpw2    ''''''''获取lzhiy0816数据(lzhiy0816 design)''''''''
inputxm2=request.form("z79xingming")
inputpw2=request.form("z79password")

''''''''如果是本班同学则进入《自动化79论坛》的主界面(lzhiy0816 design)''''''''
if inputxm2="zidonghua79"  and inputpw2="19791983" then
session("checked")="z79yes"
session("check")="z79right"
response.Redirect "index.asp"

''''''''如果不是lzhiy0816回到《自动化79论坛》登录入口(lzhiy0816 design)''''''''
else
session("checked")="z79no"
session("check")="z79wrong"
response.Redirect "z79login.asp"
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
 
