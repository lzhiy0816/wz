<%
''''''''====���Զ���79�����顷(lzhiy0816 design)====''''''''
dim inputxm2,inputpw2    ''''''''��ȡlzhiy0816����(lzhiy0816 design)''''''''
inputxm2=request.form("z79xingming")
inputpw2=request.form("z79password")

''''''''����Ǳ���ͬѧ����롶�Զ���79��̳����������(lzhiy0816 design)''''''''
if inputxm2="zidonghua79"  and inputpw2="19791983" then
session("checked")="z79yes"
session("check")="z79right"
response.Redirect "index.asp"

''''''''�������lzhiy0816�ص����Զ���79��̳����¼���(lzhiy0816 design)''''''''
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
 
