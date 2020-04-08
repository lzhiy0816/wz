<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<%
Dim ErrString,action,i,showstr,ShowErrType,ComeUrl,tmpstr
action = Request("action")
'过滤HTML标签 2005-12-29 Dv.Yz
action = Dvbbs.Replacehtml(action)
ShowErrType = Trim(Request("ShowErrType"))
Dim ErrCodes
If Request.ServerVariables("HTTP_REFERER")<>"" Then 
	tmpstr=Split(Request.ServerVariables("HTTP_REFERER"),"/")
	Comeurl=tmpstr(UBound(tmpstr))
Else
	Comeurl="index.asp"
End If
Dvbbs.LoadTemplates("showerr")
If Dvbbs.forum_setting(79)="0" Then
	Template.html(1) = Replace(Template.html(1),"{$getcode}","")
Else
	template.html(3)=Replace(template.html(3),"{$codestr}",Dvbbs.GetCode())
	Template.html(1) = Replace(Template.html(1),"{$getcode}",template.html(3))
End If
Select Case action
	Case "stop"		'论坛暂停
		Dvbbs.Stats=Template.Strings(1)
		Dvbbs.head()
		template.html(2)=Replace(template.html(2),"{$title}",Template.Strings(2)&Template.Strings(1))
		If Dvbbs.BoardID=0 Then
			If Dvbbs.Forum_Setting(69)="0" Or Dvbbs.Forum_Setting(21)="1" Then 
				template.html(2)=Replace(template.html(2),"{$stopreadme}",Stopreadme)
			Else
				Dvbbs.Forum_Setting(70)=Split(Dvbbs.Forum_Setting(70),"|")
				showstr="<br><b>&nbsp;&nbsp;"&Dvbbs.Forum_Info(0)&"</b>设置了是否定时开放，请按下面的时间访问：<hr size=""1""><ul>"
				For i=0 to UBound(Dvbbs.Forum_Setting(70))
					If i mod 6=0 Then showstr=showstr&"<li>"
					If i<10 Then showstr=showstr&"&nbsp;"
					showstr=showstr&i&"点："
					If Dvbbs.Forum_Setting(70)(i)="1" Then
						showstr=showstr&"<font color="""&Dvbbs.mainsetting(1)&""">是</font>&nbsp;&nbsp;"
					Else
						showstr=showstr&"否&nbsp;&nbsp;"
					End If
				Next
				showstr=showstr&"</ul>"
				template.html(2)=Replace(template.html(2),"{$stopreadme}",showstr)
			End If
		Else
			Dvbbs.Board_Setting(22)=Split(Dvbbs.Board_Setting(22),"|")
			showstr="<br><b>&nbsp;&nbsp;"&Dvbbs.boardtype&"</b>设置了是否定时开放，请按下面的时间访问：<hr size=""1""><ul>"
			For i=0 to UBound(Dvbbs.Board_Setting(22))
				If i mod 6=0 Then showstr=showstr&"<li>"
				If i<10 Then showstr=showstr&"&nbsp;"
				showstr=showstr&i&"点："
				If Dvbbs.Board_Setting(22)(i)="1" Then
					showstr=showstr&"<font color="""&Dvbbs.mainsetting(1)&""">是</font>&nbsp;&nbsp;"
				Else
					showstr=showstr&"否&nbsp;&nbsp;"
				End If
			Next
			showstr=showstr&"</ul>"
			template.html(2)=Replace(template.html(2),"{$stopreadme}",showstr)
		End If 
		Response.Write Template.html(2)
	Case "iplock"	'IP被限
		Dvbbs.Stats=Template.Strings(4)
		Dvbbs.head()
		Session(Dvbbs.CacheName & "UserID")=Empty 
		template.html(2)=Replace(template.html(2),"{$title}",Template.Strings(4))
		Template.Strings(5)=replace(Template.Strings(5),"{$ip}",Dvbbs.usertrueip)
		Template.Strings(5)=replace(Template.Strings(5),"{$email}",Dvbbs.Forum_Info(5))
		template.html(2)=Replace(template.html(2),"{$stopreadme}",Template.Strings(5))
		Response.Write Template.html(2)
	Case "limitedonline"	'在线被限
		Dvbbs.Stats=Template.Strings(4)
		Dvbbs.head()
		template.html(2)=Replace(template.html(2),"{$title}",Template.Strings(23))
		template.html(2)=Replace(template.html(2),"{$stopreadme}",Replace(Template.Strings(22),"{$onlinelimited}",Request("lnum")))
		Response.Write Template.html(2)
	Case "OtherErr"		'显示顶部及导航的错误信息提示
		ErrCodes=CheckErrCodes(Request("ErrCodes"),0)
		Dvbbs.Stats=action&"-"&Template.Strings(0)
		Dvbbs.head()
		Dvbbs.showtoptable()
		Dvbbs.Head_var 0,"",Template.Strings(0),""
		template.html(0)=Replace(template.html(0),"{$color}",Dvbbs.mainsetting(1))
		template.html(0)=Replace(template.html(0),"{$errtitle}",Dvbbs.Forum_Info(0)&"-"&Dvbbs.Stats)
		template.html(0)=Replace(template.html(0),"{$action}","访问论坛")
		template.html(0)=Replace(template.html(0),"{$ErrCount}",1)
		template.html(0)=Replace(template.html(0),"{$ErrString}",ErrCodes)
		If Request("autoreload")=1 Then	
			Response.Write "<meta http-equiv=refresh content=""2;URL="&Request.ServerVariables("HTTP_REFERER")&""">"		
		End If
		Response.Write Template.html(0)
		If dvbbs.userid=0 Then 
			Response.Write Replace(Template.html(1),"{$comeurl}",ComeUrl)
		End If
		Dvbbs.ActiveOnline()
		Dvbbs.footer()
	Case "iOtherErr"		'显示顶部及导航的错误信息提示，可使用html语法
		ErrCodes=CheckErrCodes(Request("ErrCodes"),1)
		Dvbbs.Stats=action&"-"&Template.Strings(0)
		Dvbbs.head()
		Dvbbs.showtoptable()
		Dvbbs.Head_var 0,"",Template.Strings(0),""
		template.html(0)=Replace(template.html(0),"{$color}",Dvbbs.mainsetting(1))
		template.html(0)=Replace(template.html(0),"{$errtitle}",Dvbbs.Forum_Info(0)&"-"&Dvbbs.Stats)
		template.html(0)=Replace(template.html(0),"{$action}","访问论坛")
		template.html(0)=Replace(template.html(0),"{$ErrCount}",1)
		template.html(0)=Replace(template.html(0),"{$ErrString}",ErrCodes)
		If Request("autoreload")=1 Then	
			Response.Write "<meta http-equiv=refresh content=""2;URL="&Request.ServerVariables("HTTP_REFERER")&""">"		
		End If
		Response.Write Template.html(0)
		If dvbbs.userid=0 Then 
			Response.Write Replace(Template.html(1),"{$comeurl}",ComeUrl)
		End If
		Dvbbs.ActiveOnline()
		Dvbbs.footer()
	Case "NoHeadErr"	'不显示顶部及导航的错误信息提示
		ErrCodes=CheckErrCodes(Request("ErrCodes"),0)
		Dvbbs.Stats=action&"-"&Template.Strings(0)
		Dvbbs.head()
		template.html(0)=Replace(template.html(0),"{$color}",Dvbbs.mainsetting(1))
		template.html(0)=Replace(template.html(0),"{$errtitle}",Dvbbs.Forum_Info(0)&"-"&Dvbbs.Stats)
		template.html(0)=Replace(template.html(0),"{$action}","访问论坛")
		template.html(0)=Replace(template.html(0),"{$ErrCount}",1)
		template.html(0)=Replace(template.html(0),"{$ErrString}",ErrCodes)
		If Request("autoreload")=1 Then	
			Response.Write "<meta http-equiv=refresh content=""2;URL="&Request.ServerVariables("HTTP_REFERER")&""">"		
		End If
		Response.Write Template.html(0)
		If dvbbs.userid=0 Then
			Response.Write Replace(Template.html(1),"{$comeurl}",ComeUrl)
		End If
		Dvbbs.ActiveOnline()
		Dvbbs.footer()
	Case "readonly"
		Dvbbs.Stats="当前论坛为只读"
		Dvbbs.head()
		Dvbbs.showtoptable()
		Dvbbs.Head_var 1,Dvbbs.boardtype,"",""
		template.html(2)=Replace(template.html(2),"{$title}",Template.Strings(2)&"当前时间论坛为只读")
		If Dvbbs.Board_Setting(21)="2" Then
			Dvbbs.Board_Setting(22)=Split(Dvbbs.Board_Setting(22),"|")
			showstr="<br><b>&nbsp;&nbsp;"&Dvbbs.boardtype&"</b>设置了当前时间为是否只读状态，请在规定的时间内发贴：<hr size=""1""><ul>"
			
			For i=0 to UBound(Dvbbs.Board_Setting(22))
				If i mod 6=0 Then showstr=showstr&"<li>"
				If i<10 Then showstr=showstr&"&nbsp;"
				showstr=showstr&i&"点："
				If Dvbbs.Board_Setting(22)(i)="1" Then
					showstr=showstr&"<font color="""&Dvbbs.mainsetting(1)&""">是</font>&nbsp;&nbsp;"
				Else
					showstr=showstr&"否&nbsp;&nbsp;"
				End If
			Next
			showstr=showstr&"</ul>"
		End If
		If Dvbbs.Forum_Setting(69) ="2" Then 
			Dvbbs.Forum_Setting(70)=Split(Dvbbs.Forum_Setting(70),"|")
				showstr="<br><b>&nbsp;&nbsp;"&Dvbbs.Forum_Info(0)&"</b>设置了当前时间为是否只读状态，请在规定的时间内发贴：<hr size=""1""><ul>"
				For i=0 to UBound(Dvbbs.Forum_Setting(70))
					If i mod 6=0 Then showstr=showstr&"<li>"
					If i<10 Then showstr=showstr&"&nbsp;"
					showstr=showstr&i&"点："
					If Dvbbs.Forum_Setting(70)(i)="1" Then
						showstr=showstr&"<font color="""&Dvbbs.mainsetting(1)&""">是</font>&nbsp;&nbsp;"
					Else
						showstr=showstr&"否&nbsp;&nbsp;"
					End If
				Next
				showstr=showstr&"</ul>"
				template.html(2)=Replace(template.html(2),"{$stopreadme}",showstr)
			End If
			template.html(2)=Replace(template.html(2),"{$stopreadme}",showstr)
		Response.Write Template.html(2)
		Dvbbs.ActiveOnline()
		Dvbbs.footer()
	Case "lock"
		Dvbbs.Stats="论坛已锁定"
		Dvbbs.head()
		Dvbbs.showtoptable()
		Dvbbs.Head_var 0,"",Dvbbs.boardtype,""
		template.html(2)=Replace(template.html(2),"{$title}",Template.Strings(2)&"论坛已锁定")
		template.html(2)=Replace(template.html(2),"{$stopreadme}","本论坛已经被锁定，不允许发贴回贴。")
		Response.Write Template.html(2)
		Dvbbs.ActiveOnline()
		Dvbbs.footer()
	Case "plus"
		ErrCodes=CheckErrCodes(Request("ErrCodes"),0)
		Dvbbs.Stats=action&"-"&Template.Strings(0)
		Dvbbs.head()
		Dvbbs.showtoptable()
		Dvbbs.Head_var 0,"",Template.Strings(0),""
		template.html(0)=Replace(template.html(0),"{$color}",Dvbbs.mainsetting(1))
		template.html(0)=Replace(template.html(0),"{$errtitle}",Dvbbs.Forum_Info(0)&"-"&Dvbbs.Stats)
		template.html(0)=Replace(template.html(0),"{$action}","使用论坛插件")
		template.html(0)=Replace(template.html(0),"{$ErrCount}",1)
		template.html(0)=Replace(template.html(0),"{$ErrString}",ErrCodes)
		Response.Write Template.html(0)
		If dvbbs.userid=0 Then 
			Response.Write Replace(Template.html(1),"{$comeurl}",ComeUrl)
		End If
		Dvbbs.ActiveOnline()
		Dvbbs.footer()
	Case Else
		Dvbbs.Stats = Action & Template.Strings(0)
		Dvbbs.head()
		If ShowErrType<>"1" Then
			Dvbbs.showtoptable()
			Dvbbs.Head_var 0,"",Template.Strings(0),""
		End If
		template.html(0)=Replace(template.html(0),"{$color}",Dvbbs.mainsetting(1))
		template.html(0)=Replace(template.html(0),"{$errtitle}",Dvbbs.Forum_Info(0)&"-"&Dvbbs.Stats)
		template.html(0)=Replace(template.html(0),"{$action}",action)
		template.html(0)=Replace(template.html(0),"{$ErrCount}",ErrCount)
		template.html(0)=Replace(template.html(0),"{$ErrString}",ErrString)
		Response.Write Template.html(0)
		If dvbbs.userid=0 Then 
			Response.Write Replace(Template.html(1),"{$comeurl}",ComeUrl)
		End If
		Dvbbs.ActiveOnline()
		Dvbbs.footer()
End Select
Function Stopreadme()
	Dim Setting
	Setting=Dvbbs.CacheData(1,0)
	Setting=split(Setting,"|||")
	Stopreadme=Setting(5)
End Function 
Function  ErrCount()
	Dim i
	ErrCount=0
	ErrCodes=Request("ErrCodes")
	If ErrCodes<>"" Then
		ErrCodes=Split(ErrCodes,",")
		For i=0 to UBound(ErrCodes)
			If IsNumeric(ErrCodes(i)) Then 
				If i=0 Then
					ErrString=ErrString&"<li>"&Template.Strings(ErrCodes(i))
				Else
					ErrString=ErrString&"<br><li>"&Template.Strings(ErrCodes(i))
				End If
				ErrCount=ErrCount+1
			End If 
		Next 
	End If 
End Function
Function CheckErrCodes(ErrCode,sType)
	If sType = 0 Then
		ErrCode=Server.htmlencode(ErrCode)
	End If
	Dim Re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="&lt;(li|br|B|\/li|\/B)&gt;"
	ErrCode=re.Replace(ErrCode,"<$1>")
	Set Re=Nothing
	CheckErrCodes=ErrCode
End Function
%>