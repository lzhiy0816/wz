<!--#include file="../conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
Head()
Dim admin_flag
FoundErr=False 
admin_flag=",14,"
If Not Dvbbs.Master Or InStr(","&Session("flag")&",",admin_flag)=0 Then
	Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=../admin_login.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	Dvbbs_Error()
Else
	Main()
	If FoundErr Then Dvbbs_Error()
	Footer()
End if

'��ʾ����
'�û�����ְ����������棬������������������ɾ������½����������½ʱ���IP��������־��������־����������30���ڲ���������������ֿɲ鿴�����������־�б���30���ڷ������
'����ָ��
'�г�����������б����а�����Ӧ����ð���
Sub Main()
End Sub
%>