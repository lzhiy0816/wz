<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="../inc/md5.asp" -->
<!-- #include file="../inc/myadmin.asp" -->
<%
	Rem ���ļ�������̳�Ͷ����ٷ���������ͨѶ��֤��;
	Rem ������Ϣ����<˫�ز��ɽ���>�ļ����㷨����,������ϢΪ��ǰ��������̳����Ա����½IP
	Rem ��ѯ�û�Ϊǰ̨�Ĺ���Ա�û���,��Ϊ������<˫�ز��ɽ���>����Ϣ
	Rem �Ƿ��ĶԸ��ļ�����ֻ�ܲµ�ǰ̨����Ա�ʻ�(����ǰ̨������Ϣ)
	Rem �Ƿ��ĶԸ��ļ����󲢲��ܻ�֪��̨��½�˺�,Ҳֻ�ܻ�ȡ���õ�IP���ɽ��ܼ�����Ϣ
	Dim UserName
	UserName = Request("Name")
	If UserName = "" Then
		Response.Write "�Ƿ��Ĳ�����"
		Response.End
	End If
	UserName = Replace(UserName,"'","")
	Set Rs = Dvbbs.Execute("Select LastLoginIP From "&AdminTable&" Where AddUser = '"&UserName&"'")
	If Rs.Eof And Rs.Bof Then
		Response.Write "Null"
	Else
		Response.Write Md5(Md5(Rs(0),32),32)
	End If
	Rs.Close
	Set Rs=Nothing
	Conn.Close
	Set Conn=Nothing
%>