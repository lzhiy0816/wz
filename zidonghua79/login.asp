<!--#include file="Conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/chan_const.asp"-->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="inc/Email_Cls.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="dv_dpo/cls_dvapi.asp"-->
<%
Dim comeurl
Dim TruePassWord
session("flag")=empty
Dvbbs.LoadTemplates("login")
Dvbbs.stats=template.Strings(1)
Dvbbs.Nav()
Dvbbs.Head_var 0,0,template.Strings(0),"login.asp"
TruePassWord=Dvbbs.Createpass
Select Case request("action")
Case "chk"
	Dvbbs_ChkLogin
	Dvbbs.Showerr()
Case "redir"
	redir
	Dvbbs.Showerr()
Case "save_redir_reg"
	call save_redir_reg()
	Dvbbs.Showerr()
Case Else
	Main
End Select
Dvbbs.ActiveOnline
Dvbbs.Footer()

Function Main()
	Dim TempStr
	TempStr = template.html(0)
	If Dvbbs.forum_setting(79)="0" Then
		TempStr = Replace(TempStr,"{$getcode}","")
	Else
		Template.html(23)=Replace(template.html(23),"{$codestr}",Dvbbs.GetCode())
		TempStr = Replace(TempStr,"{$getcode}",template.html(23))
	End If
	TempStr = Replace(TempStr,"{$rayuserlogin}",template.html(1))
	Dim Comeurl,tmpstr
	If Request("f")<>"" Then
		Comeurl=Request("f")
	ElseIf Request.ServerVariables("HTTP_REFERER")<>"" Then 
		tmpstr=split(Request.ServerVariables("HTTP_REFERER"),"/")
		Comeurl=tmpstr(UBound(tmpstr))
	Else
		Comeurl="index.asp"
	End If
	TempStr = Replace(TempStr,"{$comeurl}",Comeurl)
	Response.Write TempStr
	TempStr=""
End Function

Function Dvbbs_ChkLogin
	Dim UserIP
	Dim username
	Dim userclass
	Dim password
	Dim article
	Dim usercookies
	Dim mobile
	Dim chrs,i
	UserIP=Dvbbs.UserTrueIP
	mobile=trim(Dvbbs.CheckStr(request("passport")))
	'if mobile<>"" and request("username")="" then
	'	if len(mobile)>12 then
	'		Dvbbs.AddErrCode(9)
	'	end if
	'end if
	'if mobile<>"" then
	'	if len(mobile)>12 And Not IsNumeric(mobile) then mobile=""
	'end if
	If Request("t")="1" And Mobile = "" Then
			 Response.redirect "showerr.asp?ErrCodes=<li>������������̳ͨ��֤��&action=OtherErr"
	End If
	If Dvbbs.forum_setting(79)="1" Then
		If mobile="" And Not Dvbbs.CodeIsTrue() Then
			 Response.redirect "showerr.asp?ErrCodes=<li>��֤��У��ʧ�ܣ��뷵��ˢ��ҳ�����������֤�롣&action=OtherErr"
		End If
	End If
	If Request("username")="" Then
		If Request("passport")="" Then
			Dvbbs.AddErrCode(10)
		End If
	Else
		username=trim(Dvbbs.CheckStr(request("username")))
	End If
	If request("password")="" and mobile="" Then
		Dvbbs.AddErrCode(11)
	Else
		password=md5(trim(Dvbbs.CheckStr(request("password"))),16)
		If Request("password") = "" Then password = ""
	End If

	If Dvbbs.ErrCodes<>"" Then Exit Function

	'-----------------------------------------------------------------
	'ϵͳ����
	'-----------------------------------------------------------------
	Dim DvApi_Obj,DvApi_SaveCookie,SysKey
	If DvApi_Enable Then
		Set DvApi_Obj = New DvApi
			'DvApi_Obj.NodeValue "syskey",SysKey,0,False
			DvApi_Obj.NodeValue "action","login",0,False
			DvApi_Obj.NodeValue "username",UserName,1,False
			Md5OLD = 1
			SysKey = Md5(DvApi_Obj.XmlNode("username")&DvApi_SysKey,16)
			Md5OLD = 0
			DvApi_Obj.NodeValue "syskey",SysKey,0,False
			DvApi_Obj.NodeValue "password",Request("password"),0,False
			DvApi_Obj.SendHttpData
			If DvApi_Obj.Status = "1" Then
				Response.redirect "showerr.asp?ErrCodes="& DvApi_Obj.Message &"&action=OtherErr"	
			Else
				DvApi_SaveCookie = DvApi_Obj.SetCookie(SysKey,UserName,Password,request("CookieDate"))
			End If
		Set DvApi_Obj = Nothing
	End If
	'-----------------------------------------------------------------

	usercookies=request("CookieDate")
	'�жϸ���cookiesĿ¼
	Dim cookies_path_s,cookies_path_d,cookies_path
	cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
	cookies_path_d=ubound(cookies_path_s)
	cookies_path="/"
	For i=1 to cookies_path_d-1
		If not (cookies_path_s(i)="upload" or cookies_path_s(i)="admin") Then cookies_path=cookies_path&cookies_path_s(i)&"/"
	Next
	If dvbbs.cookiepath<>cookies_path Then
		cookies_path=replace(cookies_path,"'","")
		Dvbbs.execute("update dv_setup set Forum_Cookiespath='"&cookies_path&"'")
		Dim setupData 
		Dvbbs.CacheData(26,0)=cookies_path
		Dvbbs.Name="setup"
		Dvbbs.value=Dvbbs.CacheData
	End If
	
	If ChkUserLogin(username,password,mobile,usercookies,1)=false Then
		'������֤δͨ����ʹ���ֻ��ŵ�¼��
		If mobile<>"" Then
			challenge_check mobile,password
			Exit Function
		'������֤δͨ����ʹ���û�����¼�ģ������Ǹ߼��û����������������֤����
		Else
			set chrs=Dvbbs.Execute("select Passport,IsChallenge from [Dv_User] where username='"&username&"' and IsChallenge=1")
			If chrs.eof and chrs.bof Then
				Dvbbs.AddErrCode(12)
				Exit Function
			Else
				challenge_check chrs("Passport"),password
				Exit Function
			End If
			set chrs=nothing
		End If
	End If

	Dim comeurlname
	If instr(lcase(request("comeurl")),"reg.asp")>0 or instr(lcase(request("comeurl")),"login.asp")>0 or trim(request("comeurl"))="" Then
		comeurlname=""
		comeurl="index.asp"
	Else
		comeurl=request("comeurl")
		comeurlname="<li><a href="&request("comeurl")&">"&request("comeurl")&"</a></li>"
	End If

	Dim TempStr
	TempStr = template.html(2)
	'If Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(10)=1 And Dvbbs.Forum_ChanSetting(12)=1 Then
	'	TempStr = Replace(TempStr,"{$ray_logininfo}",template.html(3))
	'Else
	'	TempStr = Replace(TempStr,"{$ray_logininfo}","")
	'End If
	'-----------------------------------------------------------------
	'ϵͳ����
	'-----------------------------------------------------------------
	If DvApi_Enable Then
		Response.Write DvApi_SaveCookie
		Response.Flush
	End If
	'-----------------------------------------------------------------
	TempStr = Replace(TempStr,"{$ray_logininfo}","")
	TempStr = Replace(TempStr,"{$comeurl}",comeurl)
	TempStr = Replace(TempStr,"{$comeurlinfo}",comeurlname)
	TempStr = Replace(TempStr,"{$forumname}",Dvbbs.Forum_Info(0))
	Response.Write TempStr
	TempStr=""
End Function

'ȫ����֤
Function challenge_check(mobile,password)
	'If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(10)=1) Then
	'	Dvbbs.AddErrCode(13)
	'	Exit Function
	'End If
	Dim rs,iUserID
	Dim MyForumID
	Dim PostChanWord
	'���ɶ�����:01+yyyyMMddhhmmss+��λ�����
	'���������ִ�
	Dim NowTimes,PayMonth,PayDay,PayHour,PayMin,PaySe,PayDayStr,RandomizeStr,num2
	Dim PayCode,PayCodeEnCode
	NowTimes = Now()
	PayMonth = Month(NowTimes)
	If Len(PayMonth)=1 Then PayMonth = "0" & PayMonth
	PayDay = Day(NowTimes)
	If Len(PayDay)=1 Then PayDay = "0" & PayDay
	PayHour = Hour(NowTimes)
	If Len(PayHour)=1 Then PayHour = "0" & PayHour
	PayMin = Minute(NowTimes)
	If Len(PayMin)=1 Then PayMin = "0" & PayMin
	PaySe = Second(NowTimes)
	If Len(PaySe)=1 Then PaySe = "0" & PaySe
	PayDayStr = Year(NowTimes) & PayMonth & PayDay & PayHour & PayMin & PaySe
	'��������ִ�
	Randomize
	Do While Len(RandomizeStr)<5
		num2 = CStr(Chr((57-48)*rnd+48))
		RandomizeStr = RandomizeStr & num2
	Loop
	PayCode = PayDayStr & RandomizeStr & Left(MD5(Dvbbs.Forum_ChanSetting(4)&Dvbbs.Forum_ChanSetting(6),32),8)
	Dim FoundMobile,UserAnswer,UserJoinTime
	Set Rs=Dvbbs.Execute("Select UserID,Passport,UserAnswer,JoinDate From Dv_User Where Passport = '"&Dvbbs.CheckStr(Mobile)&"'")
	If Rs.Eof And Rs.Bof Then
		FoundMobile = False
		Rs.Close:Set Rs=Nothing
		Set Rs=Dvbbs.Execute("Select Top 1 UserID,Passport,UserAnswer,JoinDate From Dv_User Order By UserID")
		iUserID = "-" & Rs(0)
		UserAnswer = Rs(2)
		UserJoinTime = Rs(3)
	Else
		FoundMobile = True
		iUserID = Rs(0)
		UserAnswer = Rs(2)
		UserJoinTime = Rs(3)
	End If
	Rs.Close
	Set Rs=Nothing
	Session("challengeWord_key") = MD5(PayCode & ":" & MD5(UserAnswer & ":" & FormatDateTime(UserJoinTime,2),32),32)
	Session("challengeUserID") = iUserID

	Dim TempStr,TempArray
	TempArray = Split(template.html(19),"||")
	TempStr = TempArray(0)
	TempStr = Replace(TempStr,"{$Dvbbs_Server}","http://www.dvbbs.net/passport/login.asp")
	TempStr = Replace(TempStr,"{$passport}",mobile)
	TempStr = Replace(TempStr,"{$userid}",iUserID)
	'TempStr = Replace(TempStr,"{$password}",password)
	'TempStr = Replace(TempStr,"{$MyForumID}",MyForumID)
	TempStr = Replace(TempStr,"{$serverurl}",Dvbbs.Get_ScriptNameUrl())
	TempStr = Replace(TempStr,"{$PostChanWord}",PayCode)
	TempStr = Replace(TempStr,"{$remobile}",mobile)
	TempStr = Replace(TempStr,"{$usermobile}",mobile)
	If FoundMobile Then
		TempStr = Replace(TempStr,"{$ifpassnull}","�������ڽ�����̳ͨ��֤�û�<B>���ٵ�¼</B>��������һ��������")
		TempStr = Replace(TempStr,"{$ifpassnull1}","�����ϣ���ô���̳ͨ��֤ע�����û������¼��̳���޸ĵ�ǰ�û��󶨵���̳ͨ��֤Ϊ����ͨ��֤�ʺŻ�ȡ��ͨ��֤�󶨡�")
	Else
		TempStr = Replace(TempStr,"{$ifpassnull}","�������ڽ�����̳ͨ��֤�û�<B>����ע��</B>��������һ��������")
		TempStr = Replace(TempStr,"{$ifpassnull1}","���������������ڱ���̳ע�ᣬ����ͬ��������̳ͨ��֤�������ϵ��û�������Ϣ��")
	End If
	Response.Write TempStr
	TempStr = ""
	set rs=nothing
	If not IsObject(Application(Dvbbs.CacheName & "_iplist")) Then
		SendData()
	ElseIf DateDiff("D",Application(Dvbbs.CacheName & "_iplist").documentElement.selectSingleNode("@date").text,Date())<> 0 Then
		SendData()
	End If
	'Response.Write Application(Dvbbs.CacheName & "_iplist").documentElement.selectSingleNode("@date").text
End Function

Function strAnsi2Unicode(asContents)
	Dim len1,i,varchar,varasc
	strAnsi2Unicode = ""
	len1=LenB(asContents)
	If len1=0 Then Exit Function
	  For i=1 to len1
	  	varchar=MidB(asContents,i,1)
	  	varasc=AscB(varchar)
	  	If varasc > 127  Then
	  		If MidB(asContents,i+1,1)<>"" Then
	  			strAnsi2Unicode = strAnsi2Unicode & chr(ascw(midb(asContents,i+1,1) & varchar))
	  		End If
	  		i=i+1
	     Else
	     	strAnsi2Unicode = strAnsi2Unicode & Chr(varasc)
	     End If	
	  Next
End Function
Sub SendData()
	Dim xmlhttp,xml,DataToSend,xmlserverurl
  On Error Resume Next
  Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP"&MsxmlVersion)
	xmlserverurl="http://server.dvbbs.net/dvbbs/iplist.asp"
	xmlhttp.setTimeouts 65000, 65000, 65000, 65000
  xmlhttp.Open "POST",xmlserverurl,false
  xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  xmlhttp.send
  Set XML=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
  If XML.loadxml(strAnsi2Unicode(xmlhttp.responseBody)) Then
  	Xml.documentElement.selectSingleNode("@date").text=Date()
		Set Application(Dvbbs.CacheName & "_iplist")=Xml.cloneNode(true)
	End If
	Set xmlhttp = Nothing
End Sub

Function redir()

	Dim ErrorCode,ErrorMsg
	Dim remobile,rechallengeWord,retokerWord,reuserpassword
	Dim resex,reqq,reemail,reusername
	Dim challengeWord_key,rechallengeWord_key
	Dim userclass
	Dim rs,iUserID

	ErrorCode=trim(request("ErrorCode"))
	ErrorMsg=trim(request("ErrorMsg"))
	remobile=trim(Dvbbs.CheckStr(request("passport")))
	reuserpassword=trim(Dvbbs.CheckStr(request("password")))
	rechallengeWord=trim(Dvbbs.CheckStr(request("seqno")))
	retokerWord=trim(request("token"))
	'reemail=trim(Dvbbs.CheckStr(request("email")))
	'resex=trim(Dvbbs.CheckStr(request("sex")))
	'If resex="F" Then 
	'	resex=1
	'Else
	'	resex=0
	'End If
	'reqq=trim(Dvbbs.CheckStr(request("qq")))
	'reusername=trim(Dvbbs.CheckStr(request("username")))

	Session("re_challenge_reg_temp")=checkreal(remobile) & "|||" & checkreal(remobile)
	iUserID = Session("challengeUserID")
	If iUserID = "" Or Not IsNumeric(iUserID) Then
		Response.Redirect "index.asp"
		Exit Function
	End If
	iUserID = cCur(iUserID)

	If ErrorCode = "1" Then
		challengeWord_key=Session("challengeWord_key")
		If challengeWord_key=retokerWord Then
			Set Rs=Dvbbs.Execute("Select Passport,IsChallenge,UserID,UserClass,UserName,UserPassword From [Dv_User] Where Passport='"&remobile&"'")
			'����̳ͨ��֤���û�ע�����û�
			If Rs.Eof And Rs.Bof Then
				redir_reg_1()
				Exit Function
			'�Ѱ�ͨ��֤�û����е�¼,�˴��������û�Ϊ��¼״̬�����������ʺ���Ϣ
			Else
				Dvbbs.UserID=Rs(2)
				UserClass=Rs(5)
				reUserName=Rs(4)
				If Rs("IsChallenge")=0 Then Dvbbs.Execute("Update Dv_User Set IsChallenge = 1 Where UserID = " & Rs(2))
			End If
		Else
			'Response.Write session("challengeWord")&"||"&rechallengeWord
			'Response.End
			Response.Redirect "showerr.asp?ErrCodes=<li>������֤ʧ��2�����ܵ�ԭ���У����糬ʱ���Ƿ����ύ����&action=OtherErr"
			'challengeWord_key & "," & retokerWord & "," & md5(Session("challengeWord") & ":" & "raynetwork",32) & "<br>ԭʼ�������"&Session("challengeWord")&",�����������"&rechallengeWord&""
			Exit Function
		End If
	Else
		Response.redirect "showerr.asp?ErrCodes=<li>"&ErrorMsg&"&action=OtherErr"
		Exit Function
	End If

	Dim TempStr
	TempStr = template.html(20)
	If Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(10)=1 And Dvbbs.Forum_ChanSetting(12)=1 Then
		TempStr = Replace(TempStr,"{$ray_logininfo}",template.html(3))
	Else
		TempStr = Replace(TempStr,"{$ray_logininfo}","")
	End If
	TempStr = Replace(TempStr,"{$reuserpassword}",reuserpassword)
	TempStr = Replace(TempStr,"{$forumname}",Dvbbs.Forum_Info(0))
	Response.Write TempStr
	TempStr=""
	Dim StatUserID,UserSessionID
	StatUserID = Dvbbs.checkStr(Trim(Request.Cookies(Dvbbs.Forum_sn)("StatUserID")))
	If IsNumeric(StatUserID) = 0 or StatUserID = "" Then
		StatUserID = Replace(Dvbbs.UserTrueIP,".","")
		UserSessionID = Replace(Startime,".","")
		If IsNumeric(StatUserID) = 0 or StatUserID = "" Then StatUserID = 0
		StatUserID = Ccur(StatUserID) + Ccur(UserSessionID)
	End If
	StatUserID = Ccur(StatUserID)
	If ChkUserLogin(reusername,userclass,"",0,1) Then userclass=""
	Session("challengeUserID") = Empty
	Session("challengeWord_key") = Empty
	Session("re_challenge_reg_temp") = Empty
	
End Function

Sub redir_reg_1()

	If Session("re_challenge_reg_temp")="" Then
		Dvbbs.AddErrCode(14)
		exit sub
	End If

	Dim re_challenge_reg_temp
	re_challenge_reg_temp=split(Session("re_challenge_reg_temp"),"|||")

	Dim TempStr
	TempStr = template.html(21)
	TempStr = Replace(TempStr,"{$maxuserlength}",Dvbbs.Forum_Setting(41))
	TempStr = Replace(TempStr,"{$minuserlength}",Dvbbs.Forum_Setting(40))
	TempStr = Replace(TempStr,"{$reusername}",re_challenge_reg_temp(0))
	TempStr = Replace(TempStr,"{$passport}",re_challenge_reg_temp(1))
	TempStr = Replace(TempStr,"{$width}",Dvbbs.mainsetting(0))
	Response.Write TempStr
End Sub

Sub save_redir_reg()
	If Session("re_challenge_reg_temp")="" Then
		Dvbbs.AddErrCode(14)
		Exit Sub
	End If

	Dim username,sex,pass1,pass2,password,ErrCodes
	Dim useremail,face,width,height
	Dim oicq,sign,showRe,birthday
	Dim mailbody,sendmsg,rndnum,num1
	Dim quesion,answer,topic
	Dim userinfo,usersetting
	Dim userclass,UserIM
	Dim re_challenge_reg_temp
	Dim rs,sql,i,namebadword,SplitWords
	Dim t
	Dim StatUserID,UserSessionID
	Dim TempStr
	t = Request("t")
	If t = "" Or Not IsNumeric(t) Then t = 1
	t = Cint(t)
	If t <> 1 And t <> 2 Then t = 1
	re_challenge_reg_temp=split(Session("re_challenge_reg_temp"),"|||")

	If Request("name")="" or strLength(Request("name"))>Cint(Dvbbs.Forum_Setting(41)) or strLength(Request("name"))<Cint(Dvbbs.Forum_Setting(40)) Then
		Dvbbs.AddErrCode(17)
	Else
		username=Dvbbs.CheckStr(Trim(Request("name")))
	End If

	If Instr(username,"=")>0 or Instr(username,"%")>0 or Instr(username,chr(32))>0 or Instr(username,"?")>0 or Instr(username,"&")>0 or Instr(username,";")>0 or Instr(username,",")>0 or Instr(username,"'")>0 or Instr(username,",")>0 or Instr(username,chr(34))>0 or Instr(username,chr(9))>0 or Instr(username,"��")>0 or Instr(username,"$")>0 Then
		Dvbbs.AddErrCode(19)
	End If

	If Request.form("psw")="" or len(Request.form("psw"))>10 or len(Request.form("psw"))<6 Then
		ErrCodes=ErrCodes+"<li>�������������룬���볤��Ϊ6-10�ֽڡ�"
	Else
		pass1=Request.form("psw")
	End If
	'���û�����
	If t = 2 Then
		If ErrCodes<>"" Then Response.redirect "showerr.asp?ErrCodes="&ErrCodes&"&action=OtherErr"
		password = MD5(pass1,16)
		If Dvbbs.ErrCodes<>"" Then Exit Sub
		If ChkUserLogin(username,password,"",0,1)=False Then
			Dvbbs.AddErrCode(12)
		End If
		If Dvbbs.ErrCodes<>"" Then Exit Sub
		Conn.Execute("Update Dv_User Set Passport = '"&re_challenge_reg_temp(0)&"',IsChallenge=1 Where UserName = '"&username&"'")
		StatUserID = Dvbbs.checkStr(Trim(Request.Cookies(Dvbbs.Forum_sn)("StatUserID")))
		If IsNumeric(StatUserID) = 0 or StatUserID = "" Then
			StatUserID = Replace(Dvbbs.UserTrueIP,".","")
			UserSessionID = Replace(Startime,".","")
			If IsNumeric(StatUserID) = 0 or StatUserID = "" Then StatUserID = 0
			StatUserID = Ccur(StatUserID) + Ccur(UserSessionID)
		End If
		StatUserID = Ccur(StatUserID)
		TempStr = template.html(22)
		TempStr = Replace(TempStr,"{$ray_logininfo}","")
		TempStr = Replace(TempStr,"{$reuserpassword}",re_challenge_reg_temp(1))
		TempStr = Replace(TempStr,"{$sendmsg}","<li>��̳ͨ��֤����̳�û��ɹ���")
		TempStr = Replace(TempStr,"{$forumname}",Dvbbs.Forum_Info(0))
		Response.Write TempStr
		Session("challengeUserID") = Empty
		Session("challengeWord_key") = Empty
		Session("re_challenge_reg_temp") = Empty
		Exit Sub
	End If
	If Request.form("pswc")="" or strLength(Request.form("pswc"))>10 or len(Request.form("pswc"))<6 Then
		ErrCodes=ErrCodes+"<li>"+template.Strings(13)
	Else
		pass2=Request.form("pswc")
	End If
	If pass1<>pass2 Then
		ErrCodes=ErrCodes+"<li>"+template.Strings(29)
	Else
		password=md5(pass2,16)
	End If

	Dim RegSplitWords
	If Trim(Dvbbs.cachedata(1,0))<>"" Then
	RegSplitWords=split(Dvbbs.cachedata(1,0),"|||")(4)
	RegSplitWords=split(RegSplitWords,",")
		For i = 0 to ubound(RegSplitWords)
			If Trim(RegSplitWords(i))<>"" Then
				If instr(username,RegSplitWords(i))>0 Then
					Dvbbs.AddErrCode(19)
					Exit For
				End If
			End If 
		next
	End If 
	sex=1
	'password=md5(re_challenge_reg_temp(1),16)
	useremail=re_challenge_reg_temp(0) & "@dvbbs.net"
	showRe=1
	face="images/userface/image1.gif"
	width=32
	height=32

	If request.Form("birthyear")="" or request.form("birthmonth")="" or request.form("birthday")="" Then
		birthday=""
	Else
		birthday=trim(Request.Form("birthyear"))&"-"&trim(Request.Form("birthmonth"))&"-"&trim(Request.Form("birthday"))
		If not isdate(birthday) Then birthday=""
	End If

	userinfo=checkreal(request.Form("realname")) & "|||" & checkreal(request.Form("character")) & "|||" & checkreal(request.Form("personal")) & "|||" & checkreal(request.Form("country")) & "|||" & checkreal(request.Form("province")) & "|||" & checkreal(request.Form("city")) & "|||" & request.Form("shengxiao") & "|||" & request.Form("blood") & "|||" & request.Form("belief") & "|||" & request.Form("occupation") & "|||" & request.Form("marital") & "|||" & request.Form("education") & "|||" & checkreal(request.Form("college")) & "|||" & checkreal(request.Form("userphone")) & "|||" & checkreal(request.Form("address"))
	usersetting=request.Form("setuserinfo") & "|||" & request.Form("setusertrue") & "|||" & showRe

	If ErrCodes<>"" Then
		Response.redirect "showerr.asp?ErrCodes="&ErrCodes&"&action=OtherErr"
		Exit Sub
	End If
	If Dvbbs.ErrCodes<>"" Then Exit Sub
	Dim titlepic,iUserGroupID
	set rs=Dvbbs.Execute("select usertitle,grouppic,UserGroupID from Dv_UserGroups where ParentGID=3 order by minarticle")
	userclass=rs(0)
	titlepic=rs(1)
	iUserGroupID=rs(2)
	UserIM = "||||||||||||||||||"
	set rs=server.createobject("adodb.recordset")
	sql="select * from [Dv_User] where username='"&username&"' or Passport='"&re_challenge_reg_temp(0)&"'"
	rs.open sql,conn,1,3
	If not rs.eof and not rs.bof Then
		Dvbbs.AddErrCode(21)
		Exit Sub
	Else
		rs.addnew
		rs("IsChallenge")=1
		rs("username")=username
		rs("userpassword")=password
		rs("TruePassWord")=TruePassWord
		rs("useremail")=useremail
		rs("userclass")=userclass
		rs("titlepic")=titlepic
		rs("Passport")=re_challenge_reg_temp(0)
		Rs("UserIM")=UserIM
		Rs("UserPost")=0
		Rs("usergroupid")=iUserGroupID
		rs("lockuser")=0
		Rs("Usersex")=sex
		rs("JoinDate")=NOW()
		rs("Userface")=replace(face,"'","")
		rs("UserWidth")=width
		rs("UserHeight")=height
		rs("UserLogins")=1
		Rs("lastlogin")=NOW()
		rs("userWealth")=Dvbbs.Forum_user(0)
		rs("userEP")=Dvbbs.Forum_user(5)
		rs("usercP")=Dvbbs.Forum_user(10)
		rs("userinfo")=userinfo
		rs("usersetting")=usersetting
		rs("UserFav")="İ����,�ҵĺ���,������"
		rs.update
		Dvbbs.Execute("update Dv_Setup set Forum_usernum=Forum_usernum+1,Forum_lastuser='"&username&"'")
	End If
	rs.close
	set rs=Dvbbs.Execute("select top 1 userid from [Dv_User] order by userid desc")
	dvbbs.userid=rs(0)
	set rs=nothing
	Dvbbs.ReloadSetupCache username,14
	Dvbbs.ReloadSetupCache (CLng(Dvbbs.CacheData(10,0))+1),10 

	If Dvbbs.Forum_Setting(47)=1 and Cint(Dvbbs.Forum_Setting(2))>0 Then
		'on error resume next
		'����ע���ʼ�
		Dim getpass
		topic=Replace(template.Strings(35),"{$Forumname}",Dvbbs.Forum_Info(0))
		mailbody = template.html(17)
		mailbody = Replace(mailbody,"{$username}",Dvbbs.HtmlEncode(username))
		mailbody = Replace(mailbody,"{$password}",password)
		mailbody = Replace(mailbody,"{$copyright}",Dvbbs.Forum_Copyright)
		mailbody = Replace(mailbody,"{$version}",Dvbbs.Forum_Version)
		Dim DvEmail
		Set DvEmail = New Dv_SendMail
		DvEmail.SendObject = Cint(Dvbbs.Forum_Setting(2))	'����ѡȡ��� 1=Jmail,2=Cdonts,3=Aspemail
		DvEmail.ServerLoginName = Dvbbs.Forum_info(12)	'�����ʼ���������¼��
		DvEmail.ServerLoginPass = Dvbbs.Forum_info(13)	'��¼����
		DvEmail.SendSMTP = Dvbbs.Forum_info(4)			'SMTP��ַ
		DvEmail.SendFromEmail = Dvbbs.Forum_info(5)		'������Դ��ַ
		DvEmail.SendFromName = Dvbbs.Forum_info(0)		'��������Ϣ
		If DvEmail.ErrCode = 0 Then
			DvEmail.SendMail useremail,topic,mailbody	'ִ�з����ʼ�
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
		Dvbbs.ErrCodes=""
	Else
		sendmsg = template.Strings(36)
	End If

	If Dvbbs.Forum_Setting(46)=1 Then
		'����ע�����
		Dim sender,title,body,UserMsg,MsgID
		sender=Dvbbs.Forum_info(0)
		title=Dvbbs.Forum_info(0)&"��ӭ���ĵ���"

		body = template.html(18)
		body = Replace(body,"{$Forumname}",Dvbbs.Forum_Info(0))
		'response.write body
		sql="insert into dv_message(incept,sender,title,content,sendtime,flag,issend) values('"&username&"','"&sender&"','"&title&"','"&body&"',"&SqlNowString&",0,1)"
		Dvbbs.Execute(sql)
		Set rs=Dvbbs.execute("select top 1 ID from [Dv_message] order by ID desc")
		MsgID=rs(0)
		Rs.close:Set Rs=Nothing
		UserMsg="1||"& MsgID &"||"& sender
		Dvbbs.execute("UPDATE [Dv_User] Set UserMsg='"&Dvbbs.CheckStr(UserMsg)&"' WHERE UserID="&Dvbbs.userid)
	End If

	If cint(Dvbbs.Forum_Setting(25))=1 Then

	Else
		Response.Cookies(Dvbbs.Forum_sn).path=dvbbs.cookiepath
		Response.Cookies(Dvbbs.Forum_sn)("username")=""
		Response.Cookies(Dvbbs.Forum_sn)("password")=""
		Response.Cookies(Dvbbs.Forum_sn)("userclass")=""
		Response.Cookies(Dvbbs.Forum_sn)("userid")=""
		Response.Cookies(Dvbbs.Forum_sn)("userhidden")=""
		Response.Cookies(Dvbbs.Forum_sn)("usercookies")=""
		

		StatUserID = Dvbbs.checkStr(Trim(Request.Cookies(Dvbbs.Forum_sn)("StatUserID")))
		If IsNumeric(StatUserID) = 0 or StatUserID = "" Then
			StatUserID = Replace(Dvbbs.UserTrueIP,".","")
			UserSessionID = Replace(Startime,".","")
			If IsNumeric(StatUserID) = 0 or StatUserID = "" Then StatUserID = 0
			StatUserID = Ccur(StatUserID) + Ccur(UserSessionID)
		End If
		StatUserID = Ccur(StatUserID)
		Response.Cookies(Dvbbs.Forum_sn).path=Dvbbs.cookiepath
		Response.Cookies(Dvbbs.Forum_sn)("StatUserID") = StatUserID
 		Response.Cookies(Dvbbs.Forum_sn)("usercookies") = "0"
		Response.Cookies(Dvbbs.Forum_sn)("username") = username
		Response.Cookies(Dvbbs.Forum_sn)("password") = TruePassWord
		Response.Cookies(Dvbbs.Forum_sn)("userclass") = userclass
		Response.Cookies(Dvbbs.Forum_sn)("userid") = dvbbs.userid
		Response.Cookies(Dvbbs.Forum_sn)("userhidden") = 2
		Dvbbs.Execute("delete from dv_online where username='"&dvbbs.membername&"' Or id="&StatUserID&"")
	End If
	If ChkUserLogin(username,password,"",0,1) Then password=""

	TempStr = template.html(22)
	TempStr = Replace(TempStr,"{$ray_logininfo}","")
	TempStr = Replace(TempStr,"{$reuserpassword}",re_challenge_reg_temp(1))
	TempStr = Replace(TempStr,"{$sendmsg}","<li>��̳ͨ��֤����ע����̳�û��ɹ���")
	TempStr = Replace(TempStr,"{$forumname}",Dvbbs.Forum_Info(0))
	Response.Write TempStr
	TempStr=""
	Session("re_challenge_reg_temp")=""
	Session("challengeUserID") = Empty
	Session("challengeWord_key") = Empty
End Sub

Function checkreal(v)
Dim w
If not isnull(v) Then
	w=replace(v,"|||","����")
	checkreal=w
End If
End Function


Rem ==========��̳��¼����=========
Rem �ж��û���¼
Function ChkUserLogin(username,password,mobile,usercookies,ctype)

	Dim rsUser,article,userclass,titlepic
	Dim userhidden,lastip,UserLastLogin
	Dim GroupID,ClassSql,FoundGrade
	Dim regname,iMyUserInfo
	Dim sql,sqlstr,OLDuserhidden

	FoundGrade=False
	lastip=Dvbbs.UserTrueIP
	userhidden=request.form("userhidden")
	If userhidden <> "1" Then userhidden=2
	ChkUserLogin=false
	If mobile<>"" Then
		sqlstr=" Passport='"&mobile&"'"
	Else
		sqlstr=" UserName='"&username&"'"
	End If
	Sql="Select UserID,UserName,UserPassword,UserEmail,UserPost,UserTopic,UserSex,UserFace,UserWidth,UserHeight,JoinDate,LastLogin,lastlogin as cometime , LastLogin as activetime,UserLogins,Lockuser,Userclass,UserGroupID,UserGroup,userWealth,userEP,userCP,UserPower,UserBirthday,UserLastIP,UserDel,UserIsBest,UserHidden,UserMsg,IsChallenge,UserMobile,TitlePic,UserTitle,TruePassWord,UserToday,UserMoney,UserTicket,FollowMsgID,Vip_StarTime,Vip_EndTime,userid as boardid"
	Sql=Sql & " From [Dv_User] Where "&sqlstr&""
	set rsUser=Dvbbs.Execute(sql)
	If rsUser.eof and rsUser.bof Then
		ChkUserLogin=False
		Exit Function
	Else
		If rsUser("Lockuser") =1 Or rsUser("UserGroupID") =5 Then
			ChkUserLogin=False
			Exit Function
		Else
			If Trim(password)=Trim(rsUser("UserPassword")) Then
				ChkUserLogin=True
				Dvbbs.UserID=RsUser("UserID")
				RegName = RsUser("UserName")
				Article= RsUser("UserPost")
				UserLastLogin = RsUser("cometime")
				UserClass = RsUser("Userclass")		
				GroupID = RsUser("userGroupID")
				OLDuserhidden=RsUser("UserHidden")
				TitlePic = RsUser("UserTitle")
				If Article < 0  Then Article=0
				Set Dvbbs.UserSession=Dvbbs.RecordsetToxml(rsUser,"userinfo","xml")
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@cometime").text=Now()
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@activetime").text=DateAdd("s",-3600,Now())
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@boardid").text=0
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo").attributes.setNamedItem(Dvbbs.UserSession.createNode(2,"isuserpermissionall","")).text=Dvbbs.FoundUserPermission_All()
				If OLDuserhidden <> CLng(userhidden) Then
					Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userhidden").text=userhidden
					Dvbbs.Execute("update Dv_user set userhidden="&userhidden&" where UserId=" & Dvbbs.UserID)
				End If
				Dim BS
				Set Bs=Dvbbs.GetBrowser()
				Dvbbs.UserSession.documentElement.appendChild(Bs.documentElement)
				If EnabledSession Then 	Session(Dvbbs.CacheName & "UserID")=Dvbbs.UserSession.xml
			Else
				ChkUserLogin=False
				Exit Function
			End If
		End If
	End If
	If ChkUserLogin Then
	REM �ж��û���(�ȼ�)���ϣ����û�����Ϊ�����������������Զ������û���(�ȼ�)
	REM �Զ������û�����
	REM �������ϵͳ��������������
	Set rsUser=Dvbbs.Execute("Select MinArticle,IsSetting,ParentGID,UserTitle,GroupPic From Dv_UserGroups Where UserGroupID="&GroupID)
	If Not (rsUser.Eof And rsUser.Bof) Then
		If rsUser(2)=1 Or rsUser(2)=2 Or rsUser(2)=4 Or rsUser(2)=5 Then
			'�û��ȼ������������������û�Ϊϵͳ��������������
			UserClass=rsUser(3)
			TitlePic=rsUser(4)
			FoundGrade=True
		End If
	End If
	If Not FoundGrade Then
		'���������ϵͳ��������������,�򽫸��û�����ע���û����Ұ������������Զ��������û���(�ȼ�)
		Set rsUser=Dvbbs.Execute("Select Top 1 usertitle,GroupPic,UserGroupID From Dv_UserGroups Where ParentGID=3 And Minarticle<="&Article&" Order By MinArticle Desc,UserGroupID")
		If Not (rsUser.Eof And rsUser.Bof) Then
			UserClass=rsUser(0)
			TitlePic=rsUser(1)
			GroupID=rsUser(2)
			FoundGrade=True
		End If
	End If
	Set rsUser=nothing
	If Not FoundGrade Then Response.Redirect "showerr.asp?ErrCodes=<li>ϵͳû���ҵ�����ע���û������ϣ�����ϵ����Ա����������&action=OtherErr"
	select case ctype
	case 1
		If Datediff("d",UserLastLogin,Now())=0 Then
			sql="update [Dv_User] set LastLogin="&SqlNowString&",UserLogins=UserLogins+1,UserLastIP='"&lastip&"',userclass='"&userclass&"',titlepic='"&titlepic&"',UserGroupID="&GroupID&",TruePassWord='"&TruePassWord&"' where userid="&dvbbs.UserID
		Else
			sql="update [Dv_User] set userWealth=userWealth+"&Dvbbs.Forum_user(4)&",userEP=userEP+"&Dvbbs.Forum_user(9)&",userCP=userCP+"&Dvbbs.Forum_user(14)&",LastLogin="&SqlNowString&",UserLogins=UserLogins+1,UserLastIP='"&lastip&"',userclass='"&userclass&"',titlepic='"&titlepic&"',UserGroupID="&GroupID&",TruePassWord='"&TruePassWord&"' where userid="&dvbbs.UserID
		End If
	case 2
		sql="update [Dv_User] set UserPost=UserPost+1,UserTopic=UserTopic+1,userWealth=userWealth+"&Dvbbs.Forum_user(1)&",userEP=userEP+"&Dvbbs.Forum_user(6)&",userCP=userCP+"&Dvbbs.Forum_user(11)&",LastLogin="&SqlNowString&",UserLastIP='"&lastip&"',userclass='"&userclass&"',titlepic='"&titlepic&"',UserGroupID="&GroupID&",TruePassWord='"&TruePassWord&"' where userid="&dvbbs.UserID
	case 3
		sql="update [Dv_User] set UserPost=UserPost+1,userWealth=userWealth+"&Dvbbs.Forum_user(2)&",userEP=userEP+"&Dvbbs.Forum_user(7)&",userCP=userCP+"&Dvbbs.Forum_user(12)&",LastLogin="&SqlNowString&",UserLastIP='"&lastip&"',userclass='"&userclass&"',titlepic='"&titlepic&"',UserGroupID="&GroupID&",TruePassWord='"&TruePassWord&"' where userid="&dvbbs.UserID
	end select
	Dvbbs.Execute(sql)
	Dim StatUserID,UserSessionID
		StatUserID = Dvbbs.checkStr(Trim(Request.Cookies(Dvbbs.Forum_sn)("StatUserID")))
		If IsNumeric(StatUserID) = 0 or StatUserID = "" Then
			StatUserID = Replace(Dvbbs.UserTrueIP,".","")
			UserSessionID = Replace(Startime,".","")
			If IsNumeric(StatUserID) = 0 or StatUserID = "" Then StatUserID = 0
			StatUserID = Ccur(StatUserID) + Ccur(UserSessionID)
		End If
	StatUserID = Ccur(StatUserID)
	Dvbbs.Execute("delete from dv_online where  id="&StatUserID&"")
	If trim(username)<>trim(Dvbbs.membername) Then
		Response.Cookies(Dvbbs.Forum_sn)("username")=""
		Response.Cookies(Dvbbs.Forum_sn)("password")=""
		Response.Cookies(Dvbbs.Forum_sn)("userclass")=""
		Response.Cookies(Dvbbs.Forum_sn)("userid")=""
		Response.Cookies(Dvbbs.Forum_sn)("userhidden")=""
		Response.Cookies(Dvbbs.Forum_sn)("usercookies")=""
		Dvbbs.Execute("delete from dv_online where username='"&Dvbbs.membername&"'")
	End If
	If isnull(usercookies) or usercookies="" Then usercookies="0"
	select case usercookies
	case "0"
		Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
	case 1
   		Response.Cookies(Dvbbs.Forum_sn).Expires=Date+1
		Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
	case 2
		Response.Cookies(Dvbbs.Forum_sn).Expires=Date+31
		Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
	case 3
		Response.Cookies(Dvbbs.Forum_sn).Expires=Date+365
		Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
	end select
	Response.Cookies(Dvbbs.Forum_sn).path = Dvbbs.cookiepath
	Response.Cookies(Dvbbs.Forum_sn)("username") = regname
	Response.Cookies(Dvbbs.Forum_sn)("userid") = Dvbbs.UserID
	Response.Cookies(Dvbbs.Forum_sn)("password") = TruePassWord
	Response.Cookies(Dvbbs.Forum_sn)("userclass") = userclass
	Response.Cookies(Dvbbs.Forum_sn)("userhidden") = userhidden
	rem ���ͼƬ�ϴ���������
	Response.Cookies("upNum")=0
	Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@truepassword").text= TruePassWord
	Dvbbs.Membername=Dvbbs.Checkstr(regname)
	Dvbbs.Memberclass=Dvbbs.Checkstr(userclass)
	Dvbbs.UserGroupID=GroupID
	End If
End Function
%>