<!--#include file="../conn.asp"-->
<!--#include file="../inc/Dv_ClsMain.asp"-->
<%
Dim iCacheName,iCache,mCacheName
MyDbPath = "../"
'�����̳������Ϣ�ͼ���û���½״̬
'Dvbbs.GetForum_Setting
'Dvbbs.CheckUserLogin
'���¸����û��Ƿ�ɽ����̨Ȩ��
'If Dvbbs.GroupSetting(70)="1" Then Dvbbs.Master = True
Dim admin_flag
admin_flag=","
Dim CacheName
CacheName=Dvbbs.CacheName
'If InStr(session("flag"),admin_flag) >0  Then 
	Call delallcache()
'End If

Function  GetallCache()
	Dim Cacheobj
	For Each Cacheobj in Application.Contents
	If CStr(Left(Cacheobj,Len(CacheName)+1))=CStr(CacheName&"_") Then	
		GetallCache=GetallCache&Cacheobj&","
	End If
	Next
End Function
Sub delallcache()
	Dim cachelist,i
	Cachelist=split(GetallCache(),",")
	On Error Resume Next
	If UBound(cachelist)>1 Then
		For i=0 to UBound(cachelist)-1
			Response.Write "<b>"&Replace(cachelist(i),CacheName&"_","")&"</b><br>"		
			Response.Write "����"
			If IsObject(Application(Cachelist(i))) Then
				Response.Write "����,ռ��"
				Response.Write Len(Application(Cachelist(i)).xml)&"�ֽ�<br>"
			ElseIf IsArray(Application(Cachelist(i))) Then
				Response.Write "����,ռ��δ֪<br>"
			Else
				Response.Write "�ִ�,ռ��"&Len(Application(Cachelist(i)))&"�ֽ�<br>"
			End If
		Next
		Response.Write "��"
		Response.Write UBound(cachelist)-1
		Response.Write "���������<br>"	
	Else
		Response.Write "���ж����Ѿ����¡�"
	End If
End Sub 
Sub DelCahe(MyCaheName)
	Application.Lock
	Application.Contents.Remove(MyCaheName)
	Application.unLock
End Sub
%>