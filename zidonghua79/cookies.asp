<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/Dv_ClsOther.asp" -->
<%
Dim action
action=Request("action")
Select Case action
	Case "hidden"
		Call hidden()
	Case "online"
		Call online()
	Case "stylemod"
		Call stylemod()
	Case "setlistmod"
		Call SetListmod
	Case "setlistmoda"
		Call SetListmoda
	Case "ReGroup"
		Call ReGroup
	Case Else
End Select

If IsNull(Request.ServerVariables("HTTP_REFERER")) Or Request.ServerVariables("HTTP_REFERER")="" Then
	Response.Redirect "index.asp"
Else
	Response.Redirect Request.ServerVariables("HTTP_REFERER")
End If

Sub hidden()
	If Dvbbs.UserID=0 Then Exit Sub
	Dvbbs.execute("update [Dv_online] set userhidden=1 where userid="&Dvbbs.userid)
	Dvbbs.execute("update [Dv_user] set userhidden=1 where userid="&Dvbbs.userid)
	Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userhidden").text=1
	Dim usercookies
	usercookies=request.cookies(Dvbbs.Forum_sn)("usercookies")
	If IsNull(usercookies) or usercookies="" then usercookies="0"
	Select Case usercookies
		Case "0"
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
		Case 1
   			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+1
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
		Case 2
			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+31
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
		Case 3
			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+365
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
	End Select 
	Response.Cookies(Dvbbs.Forum_sn)("userhidden") = 1
	Response.Cookies(Dvbbs.Forum_sn).path=Dvbbs.cookiepath
End Sub

Sub online()
	If Dvbbs.UserID=0 Then Exit Sub
	Dvbbs.execute("update [dv_online] set userhidden=2 where userid="&Dvbbs.userid)
	Dvbbs.execute("update [Dv_user] set userhidden=2,lastlogin=" & SqlNowString & " where userid="&Dvbbs.userid)
	Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userhidden").text=2
	Dim  usercookies
	usercookies=request.cookies(Dvbbs.Forum_sn)("usercookies")
	If IsNull(usercookies) or usercookies="" Then usercookies="0"
	Select Case usercookies
		Case "0"
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
		Case 1
   			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+1
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
		Case 2
			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+31
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
		Case 3
			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+365
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
	End select
	Response.Cookies(Dvbbs.Forum_sn)("userhidden") = 2
	Response.Cookies(Dvbbs.Forum_sn).path=Dvbbs.cookiepath
End Sub

Sub Stylemod()
	Dim skinid
	skinid=Request("skinid")
	Response.Write skinid
	Response.Cookies("skin").expires= date+7
	Response.Cookies("skin")("skinid_"&Dvbbs.boardid)=skinid
	Response.Cookies("skin").path=Dvbbs.cookiepath
End Sub
Sub SetListmod()
	Response.Cookies("List").path=Dvbbs.cookiepath
	Response.Cookies("List").expires= date+7
	Response.Cookies("List")("list"&Request("id"))=request("thisvalue")
	Response.Write "<script language=""javascript"" type=""text/javascript"">parent.location.reload();</script>"
	Response.End 
End Sub

Sub ReGroup()
	If Dvbbs.UserID = 0 Then Exit Sub
	If Dvbbs.UserGroupParent <> 4 And Cint(Dvbbs.IsUserPermissionOnly) = 0 Then Exit Sub
	Dim ReGroupID,i,FoundGroupID,iUserInfo
	ReGroupID = Request("GroupID")
	If Not IsNumeric(ReGroupID) Or ReGroupID="" Then Exit Sub
	ReGroupID = Cint(ReGroupID)
	FoundGroupID = False
	For i = 0 To Ubound(Dvbbs.UserGroupParentID)
		If ReGroupID = Cint(Dvbbs.UserGroupParentID(i)) Then
			FoundGroupID = True
			Exit For
		End If
	Next
	If Cint(Dvbbs.IsUserPermissionOnly) > 0 Then FoundGroupID = True
	If Not FoundGroupID Then Exit Sub
	If Not (Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usergroupid2") is Nothing )  Then
		If CLng(ReGroupID)=CLng(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usergroupid2").text) Then
			Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usergroupid2").text=0
		Else
			Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usergroupid2").text=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usergroupid").text
		End If
	Else
		Dvbbs.UserSession.documentElement.selectSingleNode("userinfo").attributes.setNamedItem(Dvbbs.UserSession.createNode(2,"usergroupid2","")).text=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usergroupid").text
	End If
	Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usergroupid").text=ReGroupID
End Sub
%>
