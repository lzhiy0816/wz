<!--#include file="../conn.asp"-->
<!--#include file="../inc/Dv_ClsMain.asp"-->
<%
Dim iCacheName,iCache,mCacheName
MyDbPath = "../"
'获得论坛基本信息和检测用户登陆状态
'Dvbbs.GetForum_Setting
'Dvbbs.CheckUserLogin
'重新赋予用户是否可进入后台权限
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
			Response.Write "类型"
			If IsObject(Application(Cachelist(i))) Then
				Response.Write "对象,占用"
				Response.Write Len(Application(Cachelist(i)).xml)&"字节<br>"
			ElseIf IsArray(Application(Cachelist(i))) Then
				Response.Write "数组,占用未知<br>"
			Else
				Response.Write "字串,占用"&Len(Application(Cachelist(i)))&"字节<br>"
			End If
		Next
		Response.Write "共"
		Response.Write UBound(cachelist)-1
		Response.Write "个缓存对象<br>"	
	Else
		Response.Write "所有对象已经更新。"
	End If
End Sub 
Sub DelCahe(MyCaheName)
	Application.Lock
	Application.Contents.Remove(MyCaheName)
	Application.unLock
End Sub
%>