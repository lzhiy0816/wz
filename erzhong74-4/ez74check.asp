<%
''''''''====�������ж���74-4�� ��̳�����顷(lzhiy0816 design)====''''''''
dim inputxm2,inputpw2    ''''''''��ȡlzhiy0816����(lzhiy0816 design)''''''''
inputxm2=request.form("ez74xingming")
inputpw2=request.form("ez74password")

''''''''����Ǳ���ͬѧ����롶�����ж���74-4�� ��̳����������(lzhiy0816 design)''''''''
if inputxm2="erzhong74-4"  and inputpw2="19741979" then
session("checked")="ez74yes"
session("check")="ez74right"
response.Redirect "index.asp"

''''''''�������lzhiy0816�ص��������ж���74-4�� ��̳����¼���(lzhiy0816 design)''''''''
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
 
