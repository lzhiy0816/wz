<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/chan_const.asp"-->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="inc/Email_Cls.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="dv_dpo/cls_dvapi.asp"-->
<%
Dim Selectinfo(5)
Dim XMLDom
Dim Stats,ErrCodes
session("flag")=empty
If request("action")="postipinfo" and Request.form("comfrom")<>"" Then
		saveipinfo
Else
	LoadRegSetting()
	If Request("t")="1" Then
		ChkReg_Main()
	Else
		Reg_Main()
	End If
End If
Sub Reg_Main()
	Dim PageSid
	PageSid = Dvbbs.Skinid
	Dvbbs.LoadTemplates("usermanager")
	Dvbbs.Skinid = PageSid
	Selectinfo(0)=chk_select("",template.Strings(11))
	Selectinfo(1)=chk_select("",template.Strings(12))
	Selectinfo(2)=chk_select("",template.Strings(13))
	Selectinfo(3)=chk_select("",template.Strings(14))
	Selectinfo(4)=Chk_KidneyType("character","",template.Strings(15))
	Selectinfo(5)=chk_select("",template.Strings(16))
	Dvbbs.LoadTemplates("login")
	Stats=Split(template.Strings(25),"||")
	Dvbbs.Stats=Stats(0)
	Dvbbs.Nav()
	Dvbbs.ActiveOnline
	If request("action")<>"" and Request.form("submit")="" Then
		 Response.redirect "showerr.asp?ErrCodes=您提交的参数错误&action=OtherErr"
	ElseIf request("action")<>"" Then
		If Not CheckFormID(Request.form(GetFormID())) Then
			Response.redirect "showerr.asp?ErrCodes=您提交的参数错误&action=OtherErr"
		End If
	End If
	If Cint(dvbbs.Forum_Setting(37))=0 Then
		ErrCodes=ErrCodes+"<li>"+template.Strings(26)
	Else	
		If request("action")="apply" Then
			Dvbbs.stats=Stats(2)
			Dvbbs.Head_var 0,0,Stats(0),"reg.asp"
			reg_2()
		ElseIf request("action")="save" Then
			Dvbbs.stats=Stats(3)
			Dvbbs.Head_var 0,0,Stats(0),"reg.asp"
			reg_3()
		ElseIf request("action")="redir" Then
			Dvbbs.stats=Stats(3)
			Dvbbs.Head_var 0,0,Stats(0),"reg.asp"
			redir()
		Else
			Dvbbs.stats=Stats(1)
			Dvbbs.Head_var 0,0,Stats(0),"reg.asp"
			reg_1()
		End If
	End If
	Dvbbs.Showerr()
	If ErrCodes<>"" Then Response.redirect "showerr.asp?ErrCodes="&ErrCodes&"&action=OtherErr"	
	Dvbbs.Footer()
End Sub
Sub saveipinfo()
	Dim Node,rs
	Set XMLDom=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	If XMLDom.loadxml(Dvbbs.CacheData(27,0)) Then
		If XMLDom.documentElement.selectSingleNode("checkip/@use").text = 1 Then
			Set Node=XMLDom.documentElement.selectSingleNode("checkip/iplist1")
			If Not Node.selectNodes("ip").length =0 Then
				If Not IpInList(Node) Then
				Set Rs=Dvbbs.Execute("select Forum_BirthUser From Dv_setup")
				If Not XMLDom.loadxml(Rs(0)) Then
					XMLDom.LoadXML "<?xml version=""1.0""?><regpost/>"
				Else
					Set Node=XMLDom.documentElement.selectNodes("ip")
					If Node.length > 200 Then
						XMLDom.documentElement.removeChild(XMLDom.documentElement.firstChild)
					End If
				End If
				If XMLDom.documentElement.selectSingleNode("ip[.='"&Dvbbs.userTrueIP&"']") Is Nothing Then
				Set node=XMLDom.documentElement.appendChild(XMLDom.createNode(1,"ip",""))
					node.text=Dvbbs.userTrueIP
					Node.attributes.setNamedItem(XMLDom.createNode(2,"description","")).text=Request.form("comfrom")
					Node.attributes.setNamedItem(XMLDom.createNode(2,"dateandtime","")).text=Now()
					Dvbbs.Execute("update Dv_setup set Forum_BirthUser='"&Dvbbs.checkstr(XMLDom.xml)&"'")
				End If
			End If
			Dvbbs.LoadTemplates("")
				Dvbbs.Stats="提交注册允许请求"
				Dvbbs.Nav()
				Dvbbs.ActiveOnline
				Dvbbs.Head_var 0,0,"提交成功","reg.asp"
				Dvbbs.Dvbbs_Suc("<li>您提交的信息已经成功保存,管理员会尽快处理,请在一个工作日后再次尝试注册.</li>")
				Dvbbs.Footer()
			End If
		End If
	End If
	
End Sub
Sub reg_1()
	Dim TempLateStr
	TempLateStr=template.html(12)
	TempLateStr=Replace(TempLateStr,"{$Forum_Name}",Dvbbs.Forum_Info(0))
	TempLateStr=Replace(TempLateStr,"{$hidden}",GetFormID())
	Response.Write TempLateStr
End Sub

Sub reg_2()
	Dim grouploopinfo,TempLateStr,Rs,FormID,fname
	TempLateStr=template.html(13)
	If Dvbbs.forum_setting(78)="0" Then
		TempLateStr=Replace(TempLateStr,"{$getcode}","")
	Else
		template.html(24)=Replace(template.html(24),"{$codestr}",Dvbbs.GetCode())
		TempLateStr=Replace(TempLateStr,"{$getcode}",template.html(24))
	End If
	Set Rs=Dvbbs.Execute("select * from DV_GroupName")
	If Rs.eof and Rs.bof Then
		grouploopinfo="<option value=""无门无派"">无门无派</option>"
	Else
		do while not Rs.eof
		grouploopinfo=grouploopinfo & "<option value="""&rs("Groupname")&""">"&rs("GroupName")&"</option>"
		Rs.movenext
		loop
	End If
	Rs.close:Set Rs=Nothing
	Dim userregface,i,Forum_userface,FaceDefault
	Forum_userface = split(Dvbbs.Forum_userface,"|||")
	FaceDefault=Forum_userface(0)&Forum_userface(1)
	For i = 1 to Ubound(Forum_userface)-1
		userregface = userregface & "<option value="""&Forum_userface(0)&Forum_userface(i)
		userregface = userregface & """>" & Forum_userface(i) & "</option>"
	Next
	TempLateStr=Replace(TempLateStr,"{$color}",Dvbbs.mainsetting(1))
	TempLateStr=Replace(TempLateStr,"{$FaceDefault}",FaceDefault)
	TempLateStr=Replace(TempLateStr,"{$Face_select}",userregface)
	TempLateStr=Replace(TempLateStr,"{$FaceMaxWidth}",Dvbbs.Forum_Setting(38))
	TempLateStr=Replace(TempLateStr,"{$FaceMaxHeight}",Dvbbs.Forum_Setting(39))
	TempLateStr=Replace(TempLateStr,"{$ForumFaceMax}",Dvbbs.Forum_Setting(57))
	TempLateStr=Replace(TempLateStr,"{$NameLimLength}",Dvbbs.Forum_Setting(40))
	TempLateStr=Replace(TempLateStr,"{$NameMaxLength}",Dvbbs.Forum_Setting(41))
	TempLateStr=Replace(TempLateStr,"{$Forum_Setting7}",Dvbbs.Forum_UploadSetting(0))
	TempLateStr=Replace(TempLateStr,"{$Forum_Setting23}",Dvbbs.Forum_Setting(23))
	TempLateStr=Replace(TempLateStr,"{$Forum_Setting32}",Dvbbs.Forum_Setting(32))
	TempLateStr=Replace(TempLateStr,"{$Forum_Setting54}",Dvbbs.Forum_Setting(54))
	TempLateStr=Replace(TempLateStr,"{$Forum_Setting42}",Dvbbs.Forum_Setting(42))
	TempLateStr=Replace(TempLateStr,"{$grouploopinfo}",grouploopinfo)
	TempLateStr=Replace(TempLateStr,"{$user_blood}",chk_select("","A,B,AB,O"))
	TempLateStr=Replace(TempLateStr,"{$user_shengxiao}",Selectinfo(0))
	TempLateStr=Replace(TempLateStr,"{$user_occupation}",Selectinfo(1))
	TempLateStr=Replace(TempLateStr,"{$user_marital}",Selectinfo(2))
	TempLateStr=Replace(TempLateStr,"{$user_education}",Selectinfo(3))
	TempLateStr=Replace(TempLateStr,"{$user_character}",Selectinfo(4))
	TempLateStr=Replace(TempLateStr,"{$user_belief}",Selectinfo(5))
	FormID=GetFormID()
	TempLateStr=Replace(TempLateStr,"{$hidden}",FormID)
	If XMLDom.documentElement.selectSingleNode("@usevarform").text = "1" Then
		fname="_"&Md5(FormID,16)
	End If
	TempLateStr=Replace(TempLateStr,"{$username}","username"&fname)
	TempLateStr=Replace(TempLateStr,"{$psw}","psw"&fname)
	TempLateStr=Replace(TempLateStr,"{$pswc}","pswc"&fname)
	If XMLDom.documentElement.selectSingleNode("@checktime").text = "1" Then
		TempLateStr=Replace(TempLateStr,"{$difference}",Replace(template.html(4),"{$options}",Getoptions()))
	Else
		TempLateStr=Replace(TempLateStr,"{$difference}","")
	End If
	Response.Write TempLateStr
End Sub
Function Getoptions()
	Dim xmltime_difference,node
	Set xmltime_difference=Server.CreateObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	xmltime_difference.load Server.MapPath(MyDbPath &"inc\Time_difference.xml")
	For each node in xmltime_difference.documentElement.selectnodes("time_difference")
		Getoptions=Getoptions& "<option value="""&node.selectSingleNode("@value").text&""">"&node.text&"</option>"&vbnewline
	Next
End Function
Function GetFormID()
	Dim i,sessionid
	sessionid = Session.SessionID
	For i=1 to Len(sessionid)
		GetFormID=GetFormID&Chr(Mid(sessionid,i,1)+97)
	Next
End Function
Function CheckFormID(id)
	CheckFormID=false
	Dim i,Str
	For i=1 to Len(id)
		Str=Str & Asc(Mid(id,i,1))-97
	Next
	If Session.SessionID=Str Then
		CheckFormID=True
	End If
End Function
'下拉菜单转换输出
Function Chk_select(str1,str2)
	Dim k
	str2=Split(str2,",")
	If  str1="" Then chk_select="<option value="""" selected=""selected"">...</option>"
	For k=0 to ubound(str2)
		chk_select=chk_select & "<option value=""" & str2(k)&""""
		If str2(k)=str1 Then chk_select=chk_select &" selected=""selected"" "
		chk_select=chk_select & " >" & str2(k) &"</option>"
	Next
End Function

'多项选取转换输出
Function Chk_KidneyType(str0,str1,str2)
	Dim k
	str2=split(str2,",")
	For k = 0 to ubound(str2)	
		chk_KidneyType=chk_KidneyType+"<input type=""checkbox"" name="""&str0&""" value="""&trim(str2(k))&""" "	 
		If instr(str1,trim(str2(k)))>0 Then '如果有此项性格
		chk_KidneyType=chk_KidneyType + "checked" 
		End If 
		chk_KidneyType=chk_KidneyType + ">"&trim(str2(k))&" "
	If ((k+1) mod 5)=0 Then chk_KidneyType=chk_KidneyType +  "<br>"  '每行显示六个性格进行换行
	Next
End Function
Function checktime(time_difference,time)
	Dim GMT,YGMT
	GMT=DateAdd("s",-(8*3600),Now())
	YGMT=DateAdd("s",time_difference*3600,GMT)
	checktime=( Hour(YGMT)=CLng(time))
End Function
Sub reg_3()
	Dim username,sex,pass1,pass2,password,FormID,fname
	Dim useremail,face,width,height
	Dim sign,showRe,birthday,UserIM
	Dim mailbody,sendmsg,rndnum,num1
	Dim quesion,answer,topic
	Dim userinfo,usersetting
	Dim userclass,UserJoinTime
	Dim rs,sql,i,TempLateStr
	Dim Qq
	'判断同一IP注册间隔时间
	If Not Isnull(Session("regtime")) Or Clng(Dvbbs.Forum_Setting(22)) > 0 Then
		If DateDiff("s",Session("regtime"),Now()) < Clng(Dvbbs.Forum_Setting(22)) Then
			ErrCodes = ErrCodes + "<li>" + Replace(Template.Strings(27), "{$RegTime}", Dvbbs.Forum_Setting(22))
			Exit Sub
		End If
	End If
	If Dvbbs.chkpost=false Then
		Dvbbs.AddErrCode(16)
		Exit sub
	End If
	If XMLDom.documentElement.selectSingleNode("@checktime").text = "1" Then
		If Trim(Request.form("time_difference"))="" Or Trim(Request.form("time"))="" Or Not IsNumeric(Trim(Request.form("time_difference"))) or Not IsNumeric(Trim(Request.form("time")))Then
			Response.redirect "showerr.asp?ErrCodes=<li>您必须选择时区和时间&action=OtherErr"
			Exit sub
		Else
			If not  checktime(Trim(Request.form("time_difference")),Trim(Request.form("time"))) Then
					Response.redirect "showerr.asp?ErrCodes=<li>您选择时区和时间不正确&action=OtherErr"
			End If
		End If
	End If
	FormID=GetFormID()
	If XMLDom.documentElement.selectSingleNode("@usevarform").text = "1" Then
		fname="_"&Md5(FormID,16)
	End If
	username=Request.form("username"&fname)
	If Trim(username)="" or strLength(username)>Cint(Dvbbs.Forum_Setting(41)) or strLength(username)<Cint(Dvbbs.Forum_Setting(40)) Then
		TempLateStr=template.Strings(28)
		TempLateStr=Replace(TempLateStr,"{$RegMaxLength}",Dvbbs.Forum_Setting(41))
		TempLateStr=Replace(TempLateStr,"{$RegLimLength}",Dvbbs.Forum_Setting(40))
		ErrCodes=ErrCodes+"<li>"+TempLateStr
		TempLateStr=""
		Exit Sub
	End If
	If XMLDom.documentElement.selectSingleNode("@checknumeric").text = "1" Then
		If IsNumeric(username) Then
			Response.redirect "showerr.asp?ErrCodes=<li>本论坛不接受全数字的用户名注册.&action=OtherErr"
		End If
	End If
	username=Dvbbs.CheckStr(username)
	If Instr(username,"=")>0 or Instr(username,"%")>0 or Instr(username,chr(32))>0 or Instr(username,"?")>0 or Instr(username,"&")>0 or Instr(username,";")>0 or Instr(username,",")>0 or Instr(username,"'")>0 or Instr(username,",")>0 or Instr(username,chr(34))>0 or Instr(username,chr(9))>0 or Instr(username,"")>0 or Instr(username,"$")>0 or Instr(username,"|")>0 Then
		Dvbbs.AddErrCode(19)
		Exit sub
	End If
	If Dvbbs.forum_setting(78)="1" Then
		If Not Dvbbs.CodeIsTrue() Then
			 Response.redirect "showerr.asp?ErrCodes=<li>验证码校验失败，请返回刷新页面后再输入验证码。&action=OtherErr"
		End If
	End If
	Dim RegSplitWords
	If Trim(Dvbbs.cachedata(1,0))<>"" Then
	RegSplitWords=split(Dvbbs.cachedata(1,0),"|||")(4)
	RegSplitWords=split(RegSplitWords,",")
		For i = 0 to ubound(RegSplitWords)
			If Trim(RegSplitWords(i))<>"" Then
				If instr(username,RegSplitWords(i))>0 Then
					Dvbbs.AddErrCode(19)
					Exit sub
				End If
			End If 
		next
	End If 
	If Request.form("sex")=0 or Request.form("sex")=1 Then
		sex=Cint(Request.form("sex"))
	Else
		sex=1
	End If
	
	If Request.form("showRe")=0 or Request.form("showRe")=1 Then
		showRe=Request.form("showRe")
	Else
		showRe=1
	End If

	If Cint(Dvbbs.Forum_Setting(23))=1 Then
		Randomize
		Do While Len(rndnum)<8
		num1=CStr(Chr((57-48)*rnd+48))
		rndnum=rndnum&num1
		loop
		pass2 = rndnum
		password=md5(rndnum,16)
	Else
		If Request.form("psw"&fname)="" or len(Request.form("psw"&fname))>10 or len(Request.form("psw"&fname))<6 Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(13)
		Else
			pass1=Request.form("psw"&fname)
		End If
		If Request.form("pswc"&fname)="" or strLength(Request.form("pswc"&fname))>10 or len(Request.form("pswc"&fname))<6 Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(13)
		Else
			pass2=Request.form("pswc"&fname)
		End If
		If pass1<>pass2 Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(29)
		Else
			password=md5(pass2,16)
		End If
	End If

	If Request.form("quesion")="" Then
		ErrCodes=ErrCodes+"<li>"+template.Strings(11)
	Else
		quesion=Request.form("quesion")
	End If
	If Request.form("answer")="" Then
  		ErrCodes=ErrCodes+"<li>"+template.Strings(11)
	ElseIf Request.form("answer")=Request.form("oldanswer") Then
		answer=Request.form("answer")
	Else
		answer=md5(Request.form("answer"),16)
	End If

	If IsValidEmail(Trim(Request.form("e_mail")))=false Then
		ErrCodes=ErrCodes+"<li>"+template.Strings(30)
	Else
		If not Isnull(Dvbbs.Forum_Setting(52)) and Dvbbs.Forum_Setting(52)<>"" and Dvbbs.Forum_Setting(52)<>"0" Then
			Dim SplitUserEmail
			SplitUserEmail=Split(Dvbbs.Forum_Setting(52),"|")
			For i=0 to Ubound(SplitUserEmail)
				If Instr(Request.form("e_mail"),SplitUserEmail(i))>0 Then
				ErrCodes=ErrCodes+"<li>"+template.Strings(31)
				Exit Sub
				End If
			Next
		End If
		useremail=Dvbbs.CheckStr(Trim(Request.form("e_mail")))
	End If

	If Request.form("myface")<>"" and Cint(Dvbbs.Forum_Setting(54))=0 Then
		If Request.form("width")="" or Request.form("height")="" Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(32)
		ElseIf Not IsNumeric(Request.form("width")) or not IsNumeric(Request.form("height")) Then
			Dvbbs.AddErrCode(18)
			Exit Sub
		ElseIf Cint(Request.form("width"))>Cint(Dvbbs.Forum_Setting(57)) Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(33)
		ElseIf Cint(Request.form("height"))>Cint(Dvbbs.Forum_Setting(57)) Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(33)
		Else
			If Cint(Dvbbs.Forum_Setting(55))=0 Then
				If instr(lcase(Request.form("myface")),"http://")>0 or instr(lcase(Request.form("myface")),"www.")>0 Then
					ErrCodes=ErrCodes+"<li>"+template.Strings(34)
				End If
			End If
			face=Request.form("myface")
		End If
	Else
		If Request.form("face")<>"" Then
			face=Request.form("face")
		End If
	End If
	face=Server.htmlencode(face)
	face=replace(face,"|","")
	width=Request.form("width")
	height=Request.form("height")
	If width="" Or Not IsNumeric(width) Then width=CInt(Dvbbs.forum_setting(57))
	If height="" Or Not IsNumeric(height) Then height=CInt(Dvbbs.forum_setting(57))
	width=CInt(width)
	height=CInt(height)
	If Width > CInt(Dvbbs.forum_setting(57)) Then width=CInt(Dvbbs.forum_setting(57))
	If height > CInt(Dvbbs.forum_setting(57)) Then height=CInt(Dvbbs.forum_setting(57))
	birthday=Dvbbs.Checkstr(Trim(Request.Form("birthday")))
	If not Isdate(birthday) Then birthday=""
	'防止填写QQ号码为非数字类型 2005-3-22 Dv.Yz
	If Isnumeric(Request.Form("OICQ")) Then
		Qq = Int(Request.Form("OICQ"))
	Else
		Qq = ""
	End If
	userinfo=checkreal(Request.Form("realname")) & "|||" & checkreal(Request.Form("character")) & "|||" & checkreal(Request.Form("personal")) & "|||" & checkreal(Request.Form("country")) & "|||" & checkreal(Request.Form("province")) & "|||" & checkreal(Request.Form("city")) & "|||" & Request.Form("shengxiao") & "|||" & Request.Form("blood") & "|||" & Request.Form("belief") & "|||" & Request.Form("occupation") & "|||" & Request.Form("marital") & "|||" & Request.Form("education") & "|||" & checkreal(Request.Form("college")) & "|||" & checkreal(Request.Form("userphone")) & "|||" & checkreal(Request.Form("address"))
	usersetting=Request.Form("setuserinfo") & "|||" & Request.Form("setusertrue") & "|||" & showRe
	UserIM=checkreal(Request.form("homepage")) &"|||" & Qq & "|||"& checkreal(Request.form("ICQ")) &"|||"& checkreal(Request.form("msn")) &"|||"& checkreal(Request.form("yahoo")) &"|||"& checkreal(Request.form("aim")) &"|||"& checkreal(Request.form("uc"))
	UserIM=Server.htmlencode(UserIM)



	If ErrCodes<>"" Then Exit Sub
	If Dvbbs.ErrCodes<>"" Then Exit Sub
	
	'-----------------------------------------------------------------
	'系统整合
	'-----------------------------------------------------------------
	Dim DvApi_Obj,DvApi_SaveCookie,SysKey
	If DvApi_Enable Then
		'SysKey = Md5(UserName&DvApi_SysKey,16)
		Set DvApi_Obj = New DvApi
			'DvApi_Obj.NodeValue "syskey",SysKey,0,False
			DvApi_Obj.NodeValue "action","reguser",0,False
			DvApi_Obj.NodeValue "username",UserName,1,False
			Md5OLD = 1
			SysKey = Md5(DvApi_Obj.XmlNode("username")&DvApi_SysKey,16)
			Md5OLD = 0
			DvApi_Obj.NodeValue "syskey",SysKey,0,False
			DvApi_Obj.NodeValue "password",pass2,0,False
			DvApi_Obj.NodeValue "email",UserEmail,1,False
			DvApi_Obj.NodeValue "question",quesion,1,False
			DvApi_Obj.NodeValue "answer",Request.form("answer"),1,False
			DvApi_Obj.NodeValue "truename",Request.Form("realname"),1,False
			DvApi_Obj.NodeValue "gender",sex,0,False
			DvApi_Obj.NodeValue "birthday",birthday,0,False
			DvApi_Obj.NodeValue "qq",Qq,1,False
			DvApi_Obj.NodeValue "msn",Request.form("msn"),1,False
			DvApi_Obj.NodeValue "mobile",Request.Form("userphone"),1,False
			DvApi_Obj.NodeValue "homepage",Request.form("homepage"),1,False
			DvApi_Obj.SendHttpData
			If DvApi_Obj.Status = "1" Then
				Response.redirect "showerr.asp?ErrCodes="& DvApi_Obj.Message &"&action=OtherErr"
				Exit Sub
			Else
				DvApi_SaveCookie = DvApi_Obj.SetCookie(SysKey,UserName,Password,Request("usercookies"))
			End If
		Set DvApi_Obj = Nothing
	End If
	'-----------------------------------------------------------------

	Dim titlepic
	Dim TruePassWord
	TruePassWord=Dvbbs.Createpass
	Set Rs=Dvbbs.Execute("Select UserTitle,GroupPic,UserGroupID,IsSetting,ParentGID From Dv_UserGroups Where ParentGID=3 Order By MinArticle")
	UserClass=rs(0)
	TitlePic=rs(1)
	Dvbbs.UserGroupID = Rs(2)
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	If Request("Forum_Passport")<>"" And Cint(Dvbbs.Forum_Setting(24))=1 Then
		Sql="Select * From [Dv_user] Where UserName='"&UserName&"' Or UserEmail='"&UserEmail&"' Or Passport='"&Dvbbs.CheckStr(Request("Forum_Passport"))&"'"
	ElseIf Request("Forum_Passport")<>"" Then
		Sql="Select * From [Dv_user] Where UserName='"&UserName&"' Or Passport='"&Dvbbs.CheckStr(Request("Forum_Passport"))&"'"
	ElseIf Cint(Dvbbs.Forum_Setting(24))=1 Then
		Sql="Select * From [Dv_user] Where Username='"&UserName&"' or useremail='"&UserEmail&"'"
	Else
		Sql="Select * From [Dv_user] Where Username='"&UserName&"'"
	End If
	'Response.Write sql
	'response.end
	Rs.Open Sql,Conn,1,3
	If Not Rs.Eof And Not Rs.Bof Then
		If Request("Forum_Passport")<>"" And Cint(Dvbbs.Forum_Setting(24))=1 Then
			Response.redirect "showerr.asp?ErrCodes=<li>您填写的用户名已经被注册或者已经有用户使用了您填写的电子邮件地址。<li>或者您选择了填写了论坛通行证帐号，但所使用的通行证在本论坛已经被使用，您可以选择使用该<a href=login.asp>通行证登录论坛</a>。&action=OtherErr"
		ElseIf Request("Forum_Passport")<>"" Then
			Response.redirect "showerr.asp?ErrCodes=<li>您填写的用户名已经被注册。<li>或者您选择了填写了论坛通行证帐号，但所使用的通行证在本论坛已经被使用，您可以选择使用该<a href=login.asp>通行证登录论坛</a>。&action=OtherErr"
		ElseIf Cint(Dvbbs.Forum_Setting(24))=1 Then
			Response.redirect "showerr.asp?ErrCodes=<li>您填写的用户名已经被注册或者已经有用户使用了您填写的电子邮件地址。&action=OtherErr"
		Else
			Response.redirect "showerr.asp?ErrCodes=<li>您填写的用户名已经被注册。&action=OtherErr"
		End If
		Exit Sub
	Else
	Rs.AddNew
		UserJoinTime = Now()
		Rs("UserName")=username
		Rs("UserPassword")=password
		Rs("UserEmail")=useremail
		Rs("Userclass")=userclass
		Rs("TitlePic")=titlepic
		Rs("UserQuesion")=quesion
		Rs("UserAnswer")=answer
		Rs("TruePassWord")=TruePassWord
		Rs("UserIM")=UserIM
		If Request.Form("Signature")<>"" Then Rs("UserSign")=Dvbbs.Htmlencode(Trim(Request.Form("Signature")))
		Rs("UserPost")=0
		If Dvbbs.Forum_Setting(25)="1" Then
			Rs("UserGroupID")=5
		Else
		   	Rs("UserGroupID")=Dvbbs.UserGroupID
		End If
		Rs("Lockuser")=0
		Rs("UserSex")=sex
		If birthday<>"" Then rs("UserBirthday")=birthday
		Rs("UserGroup")=Request.form("UserGroup")
		Rs("JoinDate")=UserJoinTime
		If Request.form("myface")<>"" Then
			Rs("UserFace")=replace(Dv_FilterJS(face),"'","")
		Else
			Rs("UserFace")=replace(Dv_FilterJS(face),"'","")
		End If
		Rs("UserWidth")=Width
		Rs("Usertoday")="0|0|0|0|0"
		Rs("UserHeight")=height
		Rs("UserLogins")=1
		Rs("LastLogin")=UserJoinTime
		Rs("userWealth")=dvbbs.Forum_user(0)
		Rs("userEP")=dvbbs.Forum_user(5)
		Rs("usercP")=dvbbs.Forum_user(10)
		Rs("UserInfo")=userinfo
		Rs("UserSetting")=usersetting
		Rs("UserPower")=0
		Rs("UserDel")=0
		Rs("UserIsbest")=0
		Rs("UserMoney")=0
		Rs("UserTicket")=0
		Rs("UserFav")="陌生人,我的好友,黑名单"
		Rs("IsChallenge")=0
		Rs("UserHidden")=0
		Rs("UserLastIP")=Dvbbs.UserTrueIP
		Rs.Update
		Dvbbs.Execute("UpDate Dv_Setup Set Forum_UserNum=Forum_UserNum+1,Forum_lastUser='"&Dvbbs.HtmlEncode(username)&"'")
		
	End If
	rs.close
	Dvbbs.ReloadSetupCache username,14
	Dvbbs.ReloadSetupCache (CLng(Dvbbs.CacheData(10,0))+1),10 
	Dim facename
	Set rs=Dvbbs.execute("select top 1 userid,UserFace from [Dv_user] order by userid desc")
		Dvbbs.userid=rs(0)
		facename=rs(1)
	rs.close
	set rs=nothing
	Saveregcount(username)
	
	'******************
	'对上传头象进行过滤与改名
	If Cint(Dvbbs.Forum_UploadSetting(0))=1 Then 
		on error resume next
		Dim objFSO,upface,newfilename
		facename=Replace(facename,"\","/")
		facename=Replace(facename,"//","/")
		facename=Replace(facename,"..","")
		facename=Replace(facename,"^","")
		facename=Replace(facename,"@","")
		facename=Replace(facename,"%","")
		If instr(Lcase(facename),"uploadface/") Then
			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
			facename=objFSO.GetFileName(facename)
			upface="uploadFace/"&facename
			newfilename="uploadFace/"&Dvbbs.userid&"_"&facename
			if	objFSO.fileExists(Server.MapPath(upface)) Then
				objFSO.movefile ""&Server.MapPath(upface)&"",""&Server.MapPath(newfilename)&""
				If Not Err Then
					Dvbbs.execute("update [Dv_user] set UserFace='"&replace(newfilename,"'","")&"' Where userid="&Dvbbs.userid)
				End If
			End If
			set objFSO=nothing
		End If
	End If
	'对上传头象进行过滤与改名结束
	'****************
	If Dvbbs.Forum_Setting(23) = 1 Then Dvbbs.Forum_Setting(47) = 1 'EMAIL通知密码当然要发邮件 2005-11-9 Dv.Yz
	If Dvbbs.Forum_Setting(47)=1 and Cint(Dvbbs.Forum_Setting(2))>0 Then
		'on error resume next
		'发送注册邮件
		Dim getpass
		topic=Replace(template.Strings(35),"{$Forumname}",Dvbbs.Forum_Info(0))
		If Cint(Dvbbs.Forum_Setting(23))=1 Then
			getpass=rndnum
		Else
			getpass=Request.form("psw"&fname)
		End If
		mailbody = template.html(17)
		mailbody = Replace(mailbody,"{$username}",Dvbbs.HtmlEncode(username))
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
			DvEmail.SendMail useremail,topic,mailbody	'执行发送邮件
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

	If Dvbbs.Forum_Setting(46)="1" Then
		'发送注册短信
		Dim sender,title,body,UserMsg,MsgID
		sender=Dvbbs.Forum_Info(0)
		title=Dvbbs.lanstr(2)&Dvbbs.Forum_Info(0)
		body = template.html(18)
		body = Replace(body,"{$Forumname}",Dvbbs.Forum_Info(0))
		sql="insert into dv_message(incept,sender,title,content,sendtime,flag,issend) values('"&username&"','"&sender&"','"&title&"','"&body&"',"&SqlNowString&",0,1)"
		Dvbbs.Execute(sql)
		Set rs=Dvbbs.execute("select top 1 ID from [Dv_message] order by ID desc")
		MsgID=rs(0)
		Rs.close:Set Rs=Nothing
		UserMsg="1||"& MsgID &"||"& sender
		Dvbbs.execute("UPDATE [Dv_User] Set UserMsg='"&Dvbbs.CheckStr(UserMsg)&"' WHERE UserID="&Dvbbs.userid)
	End If

	If Not (cint(Dvbbs.Forum_Setting(23))=1 and  CInt(Dvbbs.Forum_Setting(25))=1) Then
		If EnabledSession Then
			Set Dvbbs.UserSession = Nothing 
			Session(Dvbbs.CacheName & "UserID")= empty
		End If
		Dim StatUserID,UserSessionID
		StatUserID = Dvbbs.checkStr(Trim(Request.Cookies(Dvbbs.Forum_sn)("StatUserID")))
		If IsNumeric(StatUserID) = 0 or StatUserID = "" Then
			StatUserID = Replace(Dvbbs.UserTrueIP,".","")
			UserSessionID = Replace(Startime,".","")
			If IsNumeric(StatUserID) = 0 or StatUserID = "" Then StatUserID = 0
			StatUserID = Ccur(StatUserID) + Ccur(UserSessionID)
		End If
		StatUserID = Ccur(StatUserID)
		Dvbbs.Execute("delete from dv_online where username='"&dvbbs.membername&"' Or id="&StatUserID&"")
		Response.Cookies(Dvbbs.Forum_sn)("StatUserID") = StatUserID
		select case request("usercookies")
	 	case 0
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = request("usercookies")
		Case 1
			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+1
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = request("usercookies")
		Case 2
			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+31
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = request("usercookies")
	    case 3
			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+365
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = request("usercookies")
		end select
		Response.Cookies(Dvbbs.Forum_sn)("username") = username
		Response.Cookies(Dvbbs.Forum_sn)("password") = TruePassWord
		Response.Cookies(Dvbbs.Forum_sn)("userclass") = userclass
		Response.Cookies(Dvbbs.Forum_sn)("userid") = Dvbbs.userid
		Response.Cookies(Dvbbs.Forum_sn)("userhidden") = 2
		Response.Cookies(Dvbbs.Forum_sn).path=Dvbbs.cookiepath
		Dvbbs.membername=username
		Dvbbs.userhidden=2
		Dvbbs.MemberClass=userclass		
	End If
	'如果开邮件发送密码或认证则不是登录状态 2005-11-9 Dv.Yz
	If Cint(Dvbbs.Forum_Setting(23)) = 1 Or CInt(Dvbbs.Forum_Setting(25)) = 1 Then
		Response.Cookies(Dvbbs.Forum_sn).path = Dvbbs.cookiepath
		Response.Cookies(Dvbbs.Forum_sn)("username") = ""
		Response.Cookies(Dvbbs.Forum_sn)("password") = ""
		Response.Cookies(Dvbbs.Forum_sn)("userclass") = ""
		Response.Cookies(Dvbbs.Forum_sn)("userid") = ""
		Response.Cookies(Dvbbs.Forum_sn)("userhidden") = ""
		Response.Cookies(Dvbbs.Forum_sn)("usercookies") = ""
		If EnabledSession Then
			Session(Dvbbs.CacheName & "UserID") = Empty
		End If
		Set Dvbbs.UserSession = Nothing
	End If
	Session("regtime")=now()


	'-----------------------------------------------------------------
	'系统整合
	'-----------------------------------------------------------------
	If DvApi_Enable Then
		Response.Write DvApi_SaveCookie
		Response.Flush
	End If
	'-----------------------------------------------------------------

	'前往官方论坛通行证验证
	If Request("Forum_Passport")<>"" Then
		'Get_ChallengeWord
		'生成订单号:01+yyyyMMddhhmmss+六位随机数
		'生成日期字串
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
		'生成随机字串
		Randomize
		Do While Len(RandomizeStr)<5
			num2 = CStr(Chr((57-48)*rnd+48))
			RandomizeStr = RandomizeStr & num2
		Loop
		PayCode = PayDayStr & RandomizeStr & Left(MD5(Dvbbs.Forum_ChanSetting(4)&Dvbbs.Forum_ChanSetting(6),32),8)
		Session("challengeWord_key") = MD5(PayCode & ":" & MD5(answer & ":" & FormatDateTime(UserJoinTime,2),32),32)
		Session("challengeUserID") = Dvbbs.UserID
		Response.Write Replace(template.html(14),"{$Forumname}",Dvbbs.Forum_Info(0))
		If not IsObject(Application(Dvbbs.CacheName & "_iplist")) Then
			SendData()
		ElseIf DateDiff("D",Application(Dvbbs.CacheName & "_iplist").documentElement.selectSingleNode("@date").text,Date())<> 0 Then
			SendData()
		End If
%>
<form name="redir" action="http://www.dvbbs.net/passport/reg.asp" method="post">
<INPUT type=hidden name="password" value="">
<INPUT type=hidden name="passport" value="<%=request("Forum_Passport")%>">
<INPUT type=hidden name="username" value="<%=username%>">
<INPUT type=hidden name="userid" value="<%=Dvbbs.UserID%>">
<INPUT type=hidden name="email" value="<%=useremail%>">
<INPUT type=hidden name="forumname" value="<%=Dvbbs.Forum_Info(0)%>">
<INPUT type=hidden name="forumurl" value="<%=Dvbbs.Get_ScriptNameUrl%>">
<input type=hidden value="<%=PayCode%>" name="seqno">
<input type=hidden value="<%=Dvbbs.Get_ScriptNameUrl%>reg.asp?action=redir" name="backurl">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
Else
	TempLateStr=template.html(15)
	TempLateStr=Replace(TempLateStr,"{$Forumname}",Dvbbs.Forum_Info(0))
	TempLateStr=Replace(TempLateStr,"{$sendmsg}",sendmsg)
	Response.Write TempLateStr
End If
End Sub

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
	Dim remobile,rechallengeWord,retokerWord
	Dim challengeWord_key,rechallengeWord_key
	ErrorCode=trim(request("ErrorCode"))
	ErrorMsg=trim(request("ErrorMsg"))
	remobile=trim(Dvbbs.CheckStr(request("passport")))
	rechallengeWord=trim(Dvbbs.CheckStr(request("seqno")))
	retokerWord=trim(request("token"))

	If Session("challengeUserID") = "" Then Response.Redirect "index.asp"

	If ErrorCode = "1" Then
		challengeWord_key = Session("challengeWord_key")
		If challengeWord_key = retokerWord Then
			Dvbbs.Execute("Update [Dv_User] Set Passport='"&remobile&"',IsChallenge=1 Where UserID = " & Session("challengeUserID"))
		Else
			ErrCodes=ErrCodes+"<li>本地验证失败，可能的原因有：网络超时、非法的提交请求。"
			Exit Function
		End If
	Else
		Response.redirect "showerr.asp?ErrCodes="&ErrorMsg&"&action=OtherErr"
		Exit Function
	End If
	Session("challengeWord_key") = Empty
	Session("challengeUserID") = Empty
	Response.Write Replace(Replace(template.html(15),"{$Forumname}",Dvbbs.Forum_Info(0)),"{$sendmsg}",ErrorMsg)
End Function

Function checkreal(v)
Dim w
If not isnull(v) Then
	w=replace(v,"|","")
	checkreal=w
End If
End Function

Sub ChkReg_Main()
	Dvbbs.LoadTemplates("login")
	Dim Stats,TempLateStr
	Dim username,i,sql,Rs,useremail
	Stats=split(template.Strings(25),"||")
	Dvbbs.Stats=Stats(0)
	dvbbs.head()
	ErrCodes=""
	If Request.form("username")="" Then ErrCodes=ErrCodes+"<li>"+template.Strings(6)
	If strLength(Request.form("username"))>Cint(Dvbbs.Forum_Setting(41)) or strLength(Request.form("username"))<Cint(Dvbbs.Forum_Setting(40)) Then
		TempLateStr=template.Strings(28)
		TempLateStr=Replace(TempLateStr,"{$RegMaxLength}",Dvbbs.Forum_Setting(41))
		TempLateStr=Replace(TempLateStr,"{$RegLimLength}",Dvbbs.Forum_Setting(40))
		ErrCodes=ErrCodes+"<li>"+TempLateStr
		TempLateStr=""
	Else
		username=Dvbbs.CheckStr(Trim(Request.form("username")))
		If XMLDom.documentElement.selectSingleNode("@checknumeric").text = "1" Then
		If IsNumeric(username) Then
			ErrCodes=ErrCodes&"<li>本论坛不接受全数字的用户名注册."
		End If
	End If
		If Instr(username,"=")>0 or Instr(username,"%")>0 or Instr(username,chr(32))>0 or Instr(username,"?")>0 or Instr(username,"&")>0 or Instr(username,";")>0 or Instr(username,",")>0 or Instr(username,"'")>0 or Instr(username,",")>0 or Instr(username,chr(34))>0 or Instr(username,chr(9))>0 or Instr(username,"")>0 or Instr(username,"$")>0 Then
		ErrCodes=ErrCodes+"<li>"+template.Strings(46)
		End If
		Dim RegSplitWords
		RegSplitWords=split(Dvbbs.forum_setting(4),",")
		for i = 0 to ubound(RegSplitWords)
			If instr(username,RegSplitWords(i))>0 Then
				ErrCodes=ErrCodes+"<li>"+template.Strings(46)
			End If
		next
	End If
	If Request("action")="" Then
		If IsValidEmail(trim(Request.form("email")))=false then
			ErrCodes=ErrCodes+"<li>"+template.Strings(30)
		Else 
			useremail=Dvbbs.checkStr(Request.form("email"))
		End If
	End If


	'-----------------------------------------------------------------
	'系统整合
	'-----------------------------------------------------------------
	Dim DvApi_Obj,DvApi_SaveCookie,SysKey
	If DvApi_Enable Then
		'SysKey = Md5(username&DvApi_SysKey,16)
		Set DvApi_Obj = New DvApi
			DvApi_Obj.NodeValue "action","checkname",0,False
			DvApi_Obj.NodeValue "username",username,1,False
			Md5OLD = 1
			SysKey = Md5(Username&DvApi_SysKey,16)
			Md5OLD = 0
			DvApi_Obj.NodeValue "syskey",SysKey,0,False
			If Request("action")="" Then DvApi_Obj.NodeValue "email",UserEmail,1,False
			DvApi_Obj.SendHttpData
			If DvApi_Obj.Status = "1" Then
				ErrCodes = ErrCodes + DvApi_Obj.Message
			End If
		Set DvApi_Obj = Nothing
	End If
	'-----------------------------------------------------------------

	If ErrCodes<>"" Then Showerr()
	if ErrCodes="" then
		If cint(Dvbbs.Forum_Setting(24))=1 Then
		If Request("action")="" Then
			sql="select username,useremail from [Dv_user] where username='"&username&"' or useremail='"&useremail&"'"
		Else
			sql="select username,useremail from [Dv_user] where username='"&username&"'"
		End If
		Else 
		sql="select username,useremail from [Dv_user] where username='"&username&"'"
		End If
		Set Rs=Dvbbs.execute(sql)
		If Not rs.eof and not rs.bof then
			If cint(Dvbbs.Forum_Setting(24))=1 And Rs("useremail")=useremail Then
				If Request("action")="" Then
				ErrCodes=ErrCodes+"<li>"+template.Strings(44)
				Else
				ErrCodes=ErrCodes+"<li>"+template.Strings(43)
				End If
			Else 
				ErrCodes=ErrCodes+"<li>"+template.Strings(44)
			End If
		End If 
		Rs.close:Set Rs=Nothing
	
		If ErrCodes="" Then 
			ErrCodes=template.Strings(45)
		End If
		Response.Write Replace(template.html(16),"{$Reportmsg}",ErrCodes)
	End If
	Dvbbs.Footer()
End Sub
Sub Showerr()
	Dim Show_Errmsg
	If ErrCodes<>"" Then 
		Show_Errmsg=Dvbbs.mainhtml(14)
		ErrCodes=Replace(ErrCodes,"{$color}",Dvbbs.mainSetting(1))
		Show_Errmsg=Replace(Show_Errmsg,"{$color}",Dvbbs.mainSetting(1))
		Show_Errmsg=Replace(Show_Errmsg,"{$errtitle}",Dvbbs.Forum_Info(0)&"-"&Dvbbs.Stats)
		Show_Errmsg=Replace(Show_Errmsg,"{$action}",Dvbbs.Stats)
		Show_Errmsg=Replace(Show_Errmsg,"{$ErrString}",ErrCodes)
	End If
	Response.write Show_Errmsg
End Sub
Sub LoadRegSetting()
	Dim Node
	Set XMLDom=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	If XMLDom.loadxml(Dvbbs.CacheData(27,0)) Then
		If XMLDom.documentElement.nodeName<>"regsetting" Then
			ToDefaultsetting()
		End If
		If XMLDom.documentElement.selectSingleNode("checkip/@use").text = "1" Then
			Set Node=XMLDom.documentElement.selectSingleNode("checkip/iplist2")
			If IpInList(Node) Then
				Response.redirect "showerr.asp?ErrCodes=<li>您的IP地址 "&Dvbbs.userTrueIP&" 被禁止注册&action=OtherErr"
			End If
			Set Node=XMLDom.documentElement.selectSingleNode("checkip/iplist1")
			If Not Node.selectNodes("ip").length =0 Then
				If Not IpInList(Node) Then
					If not XMLDom.documentElement.selectSingleNode("@postipinfo").text ="1" Then
						Response.redirect "showerr.asp?ErrCodes=<li>您的IP地址 "&Dvbbs.userTrueIP&" 尚未被允许注册&action=OtherErr"
					Else
					postipinfo()
					End If
				End If
			End If
			If XMLDom.documentElement.selectSingleNode("@checkregcount").text <>"0" Then
				If (CLng(XMLDom.documentElement.selectSingleNode("@checkregcount").text) <= regcount()) Then
					Response.redirect "showerr.asp?ErrCodes=<li>您的IP地址 "&Dvbbs.userTrueIP&" 每日注册数超过了限制&action=OtherErr"
				End If
			End If	 
		End If
	End If
End Sub
Sub ToDefaultsetting()
	Dim Node
	Set XMLDom=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	XMLDom.appendChild(XMLDom.createElement("regsetting"))
	Set Node=XMLDom.documentElement.appendChild(XMLDom.createNode(1,"checkip",""))
	Node.attributes.setNamedItem(XMLDom.createNode(2,"use","")).text="0"
	Node.appendChild(XMLDom.createElement("iplist1"))
	Node.appendChild(XMLDom.createElement("iplist2"))
	XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"postipinfo","")).text="0"
	XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checknumeric","")).text="0"
	XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checktime","")).text="0"
	XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"usevarform","")).text="0"
	XMLDom.documentElement.attributes.setNamedItem(XMLDom.createNode(2,"checkregcount","")).text="0"
	Dvbbs.Execute("update dv_setup set Forum_Boards='"&Dvbbs.checkstr(XMLDom.XML)&"'")
	Dvbbs.loadSetup()
End Sub
Function IpInList(Node)
	Ipinlist=False
	Dim ip,iparray
	ip=Dvbbs.userTrueIP
	iparray=split(ip,".")
	If UBound(iparray)=3 Then
		If Not Node.selectSingleNode("ip[.='"&iparray(0)&".*.*.*']") Is Nothing Then
			Ipinlist=True
		ElseIf Not Node.selectSingleNode("ip[.='"&iparray(0)&"."&iparray(1)&".*.*']") Is Nothing Then
			Ipinlist=True
		ElseIf Not Node.selectSingleNode("ip[.='"&iparray(0)&"."&iparray(1)&"."&iparray(2)&".*']") Is Nothing Then
			Ipinlist=True
		ElseIf Not Node.selectSingleNode("ip[.='"&iparray(0)&"."&iparray(1)&"."&iparray(2)&"."&iparray(3)&"']") Is Nothing Then
			Ipinlist=True	
		End If
	End If
End Function
Function regcount()
	Dim Node,rs,nedupdate,XMLDom1
	Set XMLDom1=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	Set Rs=Dvbbs.Execute("select Forum_Ad From Dv_setup")
	If Not XMLDom1.loadxml(Rs(0)) Then
			regcount=0
	Else
		For Each Node in XMLDom1.documentElement.selectNodes("ip")
			If Datediff("d",Node.selectSingleNode("@datetime").text,Now()) > 0 Then
					XMLDom1.documentElement.removeChild(node)	
					nedupdate=True
			End If
		Next
		Set Node=XMLDom1.documentElement.selectNodes("ip[.='"&Dvbbs.userTrueIP&"']")
		regcount=Node.length
		If nedupdate Then
			Dvbbs.Execute("update Dv_setup set Forum_Ad='"&Dvbbs.checkstr(XMLDom1.xml)&"'")
		End If
	End If
End Function
Sub Saveregcount(username)
		Dim Node,rs,XMLDom1
	Set XMLDom1=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	Set Rs=Dvbbs.Execute("select Forum_Ad From Dv_setup")
	If Not XMLDom1.loadxml(Rs(0)) Then
		XMLDom1.LoadXML "<?xml version=""1.0""?><reglist/>"
	Else
		For Each Node in XMLDom1.documentElement.selectNodes("ip")
			If Datediff("d",Node.selectSingleNode("@datetime").text,Now()) > 0 Then
					XMLDom1.documentElement.removeChild(node)	
			End If
		Next
	End If
	Set Node=XMLDom1.documentElement.appendChild(XMLDom1.createNode(1,"ip",""))
	Node.text=Dvbbs.userTrueIP
	Node.attributes.setNamedItem(XMLDom1.createNode(2,"datetime","")).text=Now()
	Node.attributes.setNamedItem(XMLDom1.createNode(2,"username","")).text=username
	Dvbbs.Execute("update Dv_setup set Forum_Ad='"&Dvbbs.checkstr(XMLDom1.xml)&"'")
End Sub
Sub postipinfo()
	Dim TempLateStr
	Dvbbs.LoadTemplates("login")
	Stats=Split(template.Strings(25),"||")
	Dvbbs.Stats=Stats(0)
	Dvbbs.Nav()
	Dvbbs.ActiveOnline
	Dvbbs.Head_var 0,0,"提交注册允许请求","reg.asp"
	TempLateStr=template.html(3)
	TempLateStr=Replace(TempLateStr,"{$ip}",Dvbbs.usertrueIP)
	TempLateStr=Replace(TempLateStr,"{$hidden}",GetFormID())
	Response.Write TempLateStr
	Dvbbs.Footer
	Response.End
End Sub
%>