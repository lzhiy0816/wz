<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/Class_Mobile.asp"-->
<!--#include file="inc/chan_const.asp"-->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="inc/Email_Cls.asp"-->
<!--#include file="inc/md5.asp"-->
<%
Dvbbs.LoadTemplates("login")
DvbbsWap.ShowXMLStar
If Request("t")="1" Then
	Wap_UserLogin
Else
	Wap_UserReg
End If
DvbbsWap.ShowXMLEnd

Sub Wap_UserReg()
	Dim ErrCodes
	Dim TempLateStr
	Dim RegSplitWords,i
	Dim Name,Sex,Email,Mobile,Password

	Name = Dvbbs.CheckStr(Request("name"))
	Sex = Trim(Request("sex"))
	Email = Dvbbs.CheckStr(Trim(Request("email")))
	Password = Dvbbs.CheckStr(Trim(Request("password")))
	Mobile = Dvbbs.CheckStr(Trim(Request("mobile")))

	'传递信息验证
	If Name = "" or StrLength(Name)>Cint(Dvbbs.Forum_Setting(41)) or StrLength(Name)<Cint(Dvbbs.Forum_Setting(40)) Then
		TempLateStr=template.Strings(28)
		TempLateStr=Replace(TempLateStr,"{$RegMaxLength}",Dvbbs.Forum_Setting(41))
		TempLateStr=Replace(TempLateStr,"{$RegLimLength}",Dvbbs.Forum_Setting(40))
		ErrCodes=ErrCodes+TempLateStr
		TempLateStr=""
		DvbbsWap.ShowErr 0,ErrCodes
		Exit Sub
	End If

	If Instr(Name,"=")>0 or Instr(Name,"%")>0 or Instr(Name,chr(32))>0 or Instr(Name,"?")>0 or Instr(Name,"&")>0 or Instr(Name,";")>0 or Instr(Name,",")>0 or Instr(Name,"'")>0 or Instr(Name,",")>0 or Instr(Name,chr(34))>0 or Instr(Name,chr(9))>0 or Instr(Name,"")>0 or Instr(Name,"$")>0 or Instr(Name,"|")>0 Then
		DvbbsWap.AddErrCode(19)
		Exit sub
	End If

	If Trim(Dvbbs.cachedata(1,0))<>"" Then
	RegSplitWords=split(Dvbbs.cachedata(1,0),"|||")(4)
	RegSplitWords=split(RegSplitWords,",")
		For i = 0 to ubound(RegSplitWords)
			If Trim(RegSplitWords(i))<>"" Then
				If instr(Name,RegSplitWords(i))>0 Then
					DvbbsWap.AddErrCode(19)
					Exit sub
				End If
			End If 
		next
	End If

	If Sex="0" or Sex="1" Then
		Sex=Cint(Sex)
	Else
		Sex=1
	End If

	If IsValidEmail(Email) = False Then
		ErrCodes = template.Strings(30)
		DvbbsWap.ShowErr 0,ErrCodes
		Exit Sub
	Else
		If not Isnull(Dvbbs.Forum_Setting(52)) and Dvbbs.Forum_Setting(52)<>"" and Dvbbs.Forum_Setting(52)<>"0" Then
			Dim SplitUserEmail
			SplitUserEmail=Split(Dvbbs.Forum_Setting(52),"|")
			For i=0 to Ubound(SplitUserEmail)
				If Instr(Email,SplitUserEmail(i))>0 Then
					ErrCodes = template.Strings(31)
					DvbbsWap.ShowErr 0,ErrCodes
					Exit Sub
				End If
			Next
		End If
	End If

	If Password="" or Len(Password)>10 or Len(Password)<6 Then
		ErrCodes = template.Strings(13)
		DvbbsWap.ShowErr 0,ErrCodes
		Exit Sub
	Else
		Password=Md5(Password,16)
	End If

	If Mobile<>"" and IsNumeric(Mobile) Then
		Mobile = cCur(Mobile)
	Else
		ErrCodes = "请正确填写您的手机信息！"
		DvbbsWap.ShowErr 0,ErrCodes
		Exit Sub
	End If

	'论坛用户信息
	Dim UserIM,UserInfo,UserSetting,UserFav,UserToday
	Dim Forum_userface,FaceDefault,FaceWidth,FaceHeight
	Dim UserClass,TitlePic
	Dim Rs,Sql

	UserIM			= "||||||||||||||||||"
	UserInfo		= "||||||||||||||||||||||||||||||||||||||||||"
	UserSetting		= "1|||0|||0"
	UserFav			= "陌生人,我的好友,黑名单"
	UserToday		= "0|0|0|0|0"
	FaceWidth		= CInt(Dvbbs.Forum_setting(38))
	FaceHeight		= CInt(Dvbbs.Forum_setting(39))
	Forum_userface	= Split(Dvbbs.Forum_userface,"|||")
	FaceDefault		= Forum_userface(0) & Forum_userface(1)

	Sql = "select usertitle,grouppic,UserGroupID,IsSetting,ParentGID from Dv_UserGroups where ParentGID=3 order by MinArticle"
	Set Rs = Dvbbs.execute(Sql)
	UserClass = Rs(0)
	TitlePic = Rs(1)
	Dvbbs.UserGroupID = Rs(2)
	Rs.Close

	Sql = "select * from [Dv_user] where username='"&Name&"' or usermobile='"&Mobile&"'"
	If Cint(Dvbbs.Forum_Setting(24))=1 Then
		Sql = Sql & " or useremail='"&Email&"'"
	End If
	'Response.Write Sql

	Set Rs = Server.createobject("adodb.recordset")
	Rs.open Sql,conn,1,3
	If not Rs.eof and not Rs.bof Then
		If Dvbbs.Forum_Setting(24)="1" Then
			DvbbsWap.AddErrCode(20)
			Exit sub
		Else
			DvbbsWap.AddErrCode(21)
			Exit Sub
		End If
	Else
	Rs.addnew
		Rs("UserName") = Name
		Rs("UserPassword") = Password
		Rs("UserEmail") = Email
		Rs("UserMobile") = Mobile
		Rs("Userclass") = UserClass
		Rs("TitlePic") = TitlePic
		Rs("UserIM") = UserIM
		Rs("UserPost") = 0
		If Dvbbs.Forum_Setting(25) = "1" Then
			Rs("UserGroupID") = 5
		Else
		   	Rs("UserGroupID") = Dvbbs.UserGroupID
		End If
		Rs("Lockuser") = 0
		Rs("UserSex") = Sex
		Rs("JoinDate") = NOW()
		Rs("UserFace") = Dvbbs.CheckStr(FaceDefault)
		Rs("UserWidth") = FaceWidth
		Rs("UserHeight") = FaceHeight
		Rs("UserLogins") = 1
		Rs("LastLogin") = NOW()
		Rs("userWealth") = Dvbbs.Forum_user(0)
		Rs("userEP") = Dvbbs.Forum_user(5)
		Rs("usercP") = Dvbbs.Forum_user(10)
		Rs("UserInfo") = UserInfo
		Rs("UserSetting") = UserSetting
		Rs("UserPower") = 0
		Rs("UserDel") = 0
		Rs("UserIsbest") = 0
		Rs("UserMoney") = 0
		Rs("UserTicket") = 0
		Rs("UserFav") = UserFav
		Rs("IsChallenge") = 1
		Rs("UserLastIP") = Request.ServerVariables("REMOTE_ADDR")
	Rs.update
	Dvbbs.execute("UpDate Dv_Setup Set Forum_UserNum=Forum_UserNum+1,Forum_lastUser='"&Name&"'")
	Dvbbs.ReloadSetupCache Name,14
	Dvbbs.ReloadSetupCache (CLng(Dvbbs.CacheData(10,0))+1),10
	End If
	Rs.Close
	Dvbbs.Userid = Dvbbs.execute("select top 1 userid from [Dv_user] order by userid desc")(0)

	If Dvbbs.Forum_Setting(46)="1" Then
		'发送注册短信
		Dim sender,title,body,UserMsg,MsgID
		sender=Dvbbs.Forum_Info(0)
		title=Dvbbs.lanstr(2)&Dvbbs.Forum_Info(0)
		body = template.html(18)
		body = Replace(body,"{$Forumname}",Dvbbs.Forum_Info(0))
		sql="insert into dv_message(incept,sender,title,content,sendtime,flag,issend) values('"&name&"','"&sender&"','"&title&"','"&body&"',"&SqlNowString&",0,1)"
		Dvbbs.Execute(sql)
		MsgID = Dvbbs.execute("select top 1 ID from [Dv_message] order by ID desc")(0)
		UserMsg="1||"& MsgID &"||"& sender
		Dvbbs.execute("UPDATE [Dv_User] Set UserMsg='"&Dvbbs.CheckStr(UserMsg)&"' WHERE UserID="&Dvbbs.userid)
	End If
	If Dvbbs.Forum_Setting(47)=1 and Cint(Dvbbs.Forum_Setting(2))>0 Then
		'on error resume next
		'发送注册邮件
		Dim getpass,Topic,sendmsg,Mailbody
		topic = Replace(template.Strings(35),"{$Forumname}",Dvbbs.Forum_Info(0))
		getpass = Trim(Request("password"))
		mailbody = template.html(17)
		mailbody = Replace(mailbody,"{$username}",Dvbbs.HtmlEncode(Name))
		mailbody = Replace(mailbody,"{$password}",getpass)
		mailbody = Replace(mailbody,"{$copyright}",Dvbbs.Forum_Copyright)
		mailbody = Replace(mailbody,"{$version}",Dvbbs.Forum_Version)
		Dim DvEmail
		Set DvEmail = New Dv_SendMail
		DvEmail.SendObject = Cint(Dvbbs.Forum_Setting(2))	'设置选取组件 1=Jmail,2=Cdonts,3=Aspemail
		DvEmail.ServerLoginName = Dvbbs.Forum_info(12)	'您的邮件服务器登录名
		DvEmail.ServerLoginPass = Dvbbs.Forum_info(13)	'登录密码
		DvEmail.SendSMTP = Dvbbs.Forum_info(4)			'SMTP地址
		DvEmail.SendFromEmail = Dvbbs.Forum_info(5)		'发送来源地址
		DvEmail.SendFromName = Dvbbs.Forum_info(0)		'发送人信息
		If DvEmail.ErrCode = 0 Then
			DvEmail.SendMail Email,topic,mailbody	'执行发送邮件
			If DvEmail.Count>0 Then
				If Cint(Dvbbs.Forum_Setting(23))=1 Then
					sendmsg=template.Strings(38)
				Else
					sendmsg=template.Strings(39)
				End If
			Else
				sendmsg=template.Strings(37)
			End If
		Else
			sendmsg=template.Strings(37)
		End If
		Set DvEmail = Nothing
	Else
		sendmsg = template.Strings(36)
	End If
	ErrCodes = "恭喜您，注册成功。"
	DvbbsWap.ShowErr 1,ErrCodes
End Sub

Sub Wap_UserLogin()
	Dim ErrCodes
	Dim Name,Mobile,Password
	Dim Rs,Sql
	Name = Dvbbs.CheckStr(Trim(Request("name")))
	Password = Dvbbs.CheckStr(Trim(Request("password")))
	Mobile = DvbbsWap.Mobile
	If Dvbbs.UserID>0 Then
		ErrCodes = "恭喜您，登陆成功。"&Dvbbs.UserID
		DvbbsWap.ShowErr 1,ErrCodes
		Exit Sub
	End If
	If Name="" or Password="" Then
		DvbbsWap.ShowErr 0,"用户帐号或密码不能为空！"
		Exit Sub
	Else
		Password=Md5(Password,16)
	End If
	If Not IsObject(Conn) Then ConnectionDatabase
	Sql = "select UserID,UserName,usermobile,UserLogins,LastLogin,IsChallenge,UserLastIP,Lockuser,UserGroupID from [Dv_user] where username='"&Name&"' and UserPassword='"&Password&"'"
	'Response.Write sql
	Set Rs = Server.createobject("adodb.recordset")
	Rs.open Sql,conn,1,3
	If Rs.eof Then
		DvbbsWap.ShowErr 0,"你的帐号和密码不符！"
		Exit Sub
	Else
		If Rs("UserGroupID") = 5 Then
			DvbbsWap.ShowErr 0,"你是等待验证的(COPPA)会员，暂时不能登陆！"
			Exit Sub
		End If
		If Rs("Lockuser") = 1 Then
			DvbbsWap.ShowErr 0,"你的帐号已被锁定，不允许登陆！"
			Exit Sub
		End If
		Rs("UserLogins") = Rs("UserLogins")+1
		Rs("LastLogin") = NOW()
		Rs("IsChallenge") = 1
		if Mobile>0 Then
			Rs("usermobile") = Mobile
		Else
			Mobile = Rs("usermobile")
		End If
		Rs("UserLastIP") = Request.ServerVariables("REMOTE_ADDR")
		Rs.update
		Response.Write "<mobile>"&Mobile&"</mobile>"
		ErrCodes = "恭喜您，登陆成功。"
		DvbbsWap.ShowErr 1,ErrCodes
	End If
	Rs.Close
	Set Rs=Nothing
End Sub
%>