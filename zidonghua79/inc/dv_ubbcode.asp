<%
'最后修改2006.6.11
Dim UserPointInfo(4)
'UBB代码勘套循环的最多次数，避免死循环加入此变量
Const MaxLoopcount=100
'部分服务器vbscript可能不支持SubMatches集合，请设置 Issupport=0
Const Issupport=1
'签名最大字体值
Const Maxsize=4
Rem 是否让管理员看到是贴子是否符合XHTML格式
Const showisxhtml=1
Const ubb_version="2006-6-20"
Rem can_Post_Style是不限制style使用的用户组别列表，你可以根据自己需要修改
Const can_Post_Style="1,2,3,35"
Dim Mtinfo
Mtinfo="<fieldset style=""border : 1px dotted #ccc;text-align : left;line-height:22px;text-indent:10px""><legend><b>媒体文件信息</b></legend><div>文件来源：$4</div>"&_
"<div>您可以点击控件上的播放按钮在线播放。注意，播放此媒体文件存在一些风险。</div>"&_
"<div>附加说明：动网论坛系统禁止了该文件的自动播放功能。</div>"&_
"<div>由于该用户没有发表自动播放多媒体文件的权限或者该版面被设置成不支持多媒体播放。</div></fieldset>"
Const DV_UBB_TITLE=" title=""dvubb"" "
Const UBB_TITLE="dvubb"
%>
<script language=vbscript runat=server>
Dim Ubblists
'[/img]编号:1.[/upload]编号:2.[/dir]编号:3.[/qt]编号:4.[/mp]编号:5.
'[/rm]编号:6.[/sound]编号:7.[/flash]编号:8.[/money]编号:9.[/point]编号:10.
'[/usercp]编号:11.[/power]编号:12.[/post]编号:13.[/replyview]编号:14.[/usemoney]编号:15.
'[/url]编号:16.[/email]编号:17.http编号:18.https编号:19.ftp编号:20.rtsp编号:21.
'mms编号:22.[/html]编号:23.[/code]编号:24.[/color]编号:25.[/face]编号:26.[/align]编号:27.
'[/quote]编号:28.[/fly]编号:29.[/move]编号:30.[/shadow]编号:31.[/glow]编号:32.[/size]编号:33.
'[/i]编号:34.[/b]编号:35.[/u]编号:36.[em编号:37.www.编号:38.[/payto]编号:40.[/username]编号:41.[/center]编号:42.

Class Dvbbs_UbbCode
	Public Re,reed,isgetreed,Board_Setting,WapPushUrl,xml,isxhtml
	Function xmlencode(Str)
		Dim s,i
		s=Str
		S=replace(S,"&","&amp;")
		For i=0 to 31
			S=Replace(S,Chr(i),"&amp;#"&i&";")
		Next
		For i=95 to 96
			S=Replace(S,Chr(i),"&amp;#"&i&";")
		Next
		xmlencode=S
	End Function
	Function Rexmlencode(Str)
		Dim i
		Str=replace(Str,"&amp;","&")
		For i=0 to 31
			Str=Replace(Str,"&#"&i&";",Chr(i))
		Next
		For i=95 to 96
			Str=Replace(Str,"&#"&i&";",Chr(i))
		Next
		Rexmlencode=Str
	End Function
	Public Property Let PostType(ByVal vNewvalue)
		If PostType=2 Then
			Board_Setting=Split("1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1",",")
			Board_Setting(6)=1
			Board_Setting(5)=0:Board_Setting(7)=1
			Board_Setting(8)=1:Board_Setting(9)=1
			Board_Setting(10)=0:Board_Setting(11)=0
			Board_Setting(12)=0:Board_Setting(13)=0
			Board_Setting(14)=0:Board_Setting(15)=0
			Board_Setting(23)=0:Board_Setting(44)=0
		Else
			If Dvbbs.BoardID >0 Then 
				Board_Setting=Dvbbs.Board_Setting
			Else
				Board_Setting=Split("1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1",",")
				Board_Setting(6)=1
				Board_Setting(5)=0:Board_Setting(7)=1
				Board_Setting(8)=1:Board_Setting(9)=1
				Board_Setting(10)=0:Board_Setting(11)=0
				Board_Setting(12)=0:Board_Setting(13)=0
				Board_Setting(14)=0:Board_Setting(15)=0
				Board_Setting(23)=0:Board_Setting(44)=0
			End If
		End If
	End Property
	Private Sub Class_Initialize()
		Set re=new RegExp
		re.IgnoreCase =true
		re.Global=true
		Set xml=Server.Createobject("msxml2.DOMDocument"& MsxmlVersion)
		If Dvbbs.UserID=0 Then
			UserPointInfo(0)=0:UserPointInfo(1)=0:UserPointInfo(2)=0:UserPointInfo(3)=0:UserPointInfo(4)=0
		Else
			UserPointInfo(0)=CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text)
			UserPointInfo(1)=CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text)
			UserPointInfo(2)=CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text)
			UserPointInfo(3)=CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpower").text)
			UserPointInfo(4)=CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text)
		End If
	End Sub
	Private Sub class_terminate()
		Set xml=Nothing
		Set Re=Nothing
	End Sub
	Function istext(Str)
			Dim text,text1
			text=Str
			text1=Str
		 If text1=Dvbbs.Replacehtml(text) Then
		 	istext=True
		 End If
	End Function
	Function TextFormt(Str)
		Dim tmp,i
		Str=replace(Str,Chr(13)& Chr(10),Chr(13))
		'Str=replace(Str," ",Chr(13))
		Str=replace(Str,Chr(10),Chr(13))
		Str = re.Replace(Str,"$1"& Chr(13) &"$3")
		TMP=Split(Str,Chr(13))
		Str=""
		For i=0 to UBound(tmp)
			If i=UBound(tmp) Then
				Str=Str & tmp(i) 
			Else
				Str=Str & tmp(i) &"<br />"
			End If
		Next
		're.Pattern="\[((.|\n^\]^\[)*?)<br />((.|\n^\]^\[)*?)\]"
		'Str = re.Replace(Str,"[$1 $3]")
		TextFormt=Str
	End Function
	Rem处理老DHTML贴子
	Public Function Dv_UbbCode_DHTML(s,PostUserGroup,PostType,sType)
		Dim matches,match,CodeStr
		If InStr(Ubblists,",39,")>0 And (InStr(Ubblists,",table,")>0 Or InStr(Ubblists,",td,")>0 Or InStr(Ubblists,",th,")>0 Or InStr(Ubblists,",tr,")>0 ) Then
				s = server.htmlencode(s)
				s="<form name=""scode"&replyid&""" method=""post"" action=""""><table class=""tableborder2"" cellspacing=""1"" cellpadding=""3"" width=""100%"" align=""center"" border=""0""><tr><th height=""22"">以下内容含错误标记</th></tr><tr><td class=""tablebody1"" align=""middle"" width=""98%""><textarea id=""CodeText"" style=""BORDER-RIGHT: 1px dotted; BORDER-TOP: 1px dotted; OVERFLOW-Y: visible; OVERFLOW: visible; BORDER-LEFT: 1px dotted; WIDth: 98%; COLOR: #000000; BORDER-BOTTOM: 1px dotted"" rows=""20"" cols=""120"">"&s&"</textarea></td></tr><tr><td class=""tablebody2"" align=""middle"" width=""98%""></td></tr></table></form>"
				Dv_UbbCode_DHTML=s
				Exit Function
		Else
			If Board_Setting(5)="0" Then
					re.Pattern ="<(\/?(i|b|p))>"
					s=re.Replace(s,Chr(1)&"$1"&Chr(2))
					re.Pattern="(>)("&vbNewLine&"){1,2}(<)"
					s=re.Replace(s,"$1$3")
					re.Pattern="(<div class=""quote"">)((.|\n)*?)(<\/div>)"
					Do While re.Test(s)
						s=re.Replace(s,"[quote]$2[/quote]")
					Loop
					re.Pattern = "(<\/tr>)"
					s = re.Replace(s,"[br]")
					re.Pattern = "(<br/>)"
					s = re.Replace(s,"[br]")
					re.Pattern = "(<br>)"
					s = re.Replace(s,"[br]")
					re.Pattern = "<(\/?s(ub|up|trike))>"
					s = re.Replace(s,"[$1]")
					re.Pattern = "(<)(\/?font[^>]*)(>)"
					s = re.Replace(s,CHR(1)&"$2"&CHR(2))
					re.Pattern="<([^<>]*?)>"
					Do while re.Test(s)
						s=re.Replace(s,"")
					Loop
					re.Pattern = "(\x01)(\/?font[^\x02]*)(\x02)"
					s = re.Replace(s,"<$2>")
					re.Pattern = "\[(\/?s(ub|up|trike))\]"
					s = re.Replace(s,"<$1>")
					re.Pattern="(\[quote\])((.|\n)*?)(\[\/quote\])"
					Do While re.Test(s)
						s=re.Replace(s,"<div class=""quote"">$2</div>")
					Loop
					re.Pattern="\x01(\/?(i|b|p))\x02"
					s=re.Replace(s,"<$1>")
					re.Pattern = "(\[br\])"
					s = re.Replace(s,"<br/>")
				End If
				re.Pattern="<((asp|\!|%))"
				s=re.Replace(s,"&lt;$1")
		End If
		If InStr(s,"<TABLE") =0 Then s="<pre>" & TextFormt(S) &"</pre>"
		Dv_UbbCode_DHTML=s
	End Function
	'论坛内容部分UBBCODE，入口：内容、用户组ID、模式(1=帖子/2=公告、短信等)、模式2(0=新版/1=老版)
	Public Function Dv_UbbCode(s,PostUserGroup,PostType,sType)
		Dim mt,i,tmp
		isxhtml=false
		mt=canusemt(PostUserGroup)
		Dim textonly
		textonly=istext(s)
		If textonly Then 
			re.Pattern="<"
			s= re.Replace(s,"&lt;")
			s=TextFormt(s)
		End If
		s=Rexmlencode(s)
		re.Pattern = "(\[br\])"
		s = re.Replace(s,"<br />")
		If xml.loadxml("<div>" & xmlencode(s) &"</div>") Then
			isxhtml=True
			's=checkXHTML(mt,PostUserGroup)
		Else
			Rem处理老DHTML贴子
			isxhtml=false
			s=Dv_UbbCode_DHTML(s,PostUserGroup,PostType,sType)
		End If
		'Ubb转换
			If InStr(Ubblists,",1,")>0 Or sType=1 Then 
				s=Dv_UbbCode_iS2(s,"img",_
				"<a href=""$1"" target=""_blank"" ><img "& DV_UBB_TITLE &" src=""$1"" border=""0"" /></a>",_
				"<img  "& DV_UBB_TITLE &" src=""skins/default/filetype/gif.gif"" border=""0"" alt="" /><a  href=""$1"" target=""_blank"" >$1</a>",_
				PostUserGroup,Cint(Board_Setting(7)),_
				"")
			End If
			'upload code
			If InStr(Ubblists,",2,")>0 Or sType=1 Then
				s=Dv_UbbCode_U(s,PostUserGroup,Cint(Board_Setting(7)))
			End If
			'media code
			If InStr(Ubblists,",3,")>0 Or sType=1 Then
				s=Dv_UbbCode_iS2(s,"DIR",_
				"<object "& DV_UBB_TITLE &" classid=""clsid:166B1BCA-3F9C-11CF-8075-444553540000"" "&_
				"codebase=""http://download.macromedia.com/pub/shockwave/cabs/director/sw.cab#version=7,0,2,0"" "&_
				"width=""$1"" height=""$2""><param name=""src"" value=""$3"" /><embed "& DV_UBB_TITLE &" src=""$3"""&_
				" pluginspage=""http://www.macromedia.com/shockwave/download/"" width=""$1"" height=""$2""></embed></object>",_
				"<a href=""$3"" target=""_blank"">$3</a>",_
				PostUserGroup,Cint(Board_Setting(9) * mt),_
				"=*([0-9]*),*([0-9]*)")
			End If
			'qt
			If InStr(Ubblists,",4,")>0 Or sType=1 Then
				s=Dv_UbbCode_iS2(s,"QT",_
				"<embed "& DV_UBB_TITLE &" src=""$3"" width=""$1"" height=""$2"" autoplay=""true"" loop=""false"" controller=""true"" playeveryframe=""false"" cache=""false"" scale=""TOFIT"" bgcolor=""#000000"" kioskmode=""false"" targetcache=""false"" pluginspage=""http://www.apple.com/quicktime/"" />",_
				"<embed "& DV_UBB_TITLE &" src=""$3"" width=""$1"" height=""$2"" autoplay=""false"" loop=""false"" controller=""true"" playeveryframe=""false"" cache=""false"" scale=""TOFIT"" bgcolor=""#000000"" kioskmode=""false"" targetcache=""false"" pluginspage=""http://www.apple.com/quicktime/"" />"&_
				 replace(Mtinfo,"$4","$3"),_
				PostUserGroup,Cint(Board_Setting(9) * mt),_
				"=*([0-9]*),*([0-9]*)")
			End If
			'mp
			If InStr(Ubblists,",5,")>0 Or sType=1 Then
				s=Dv_UbbCode_iS2(s,"mp",_
				"<object "& DV_UBB_TITLE &" align=""middle"" classid=""CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95"" class=""object"" id=""MediaPlayer"" width=""$1"" height=""$2"" >"&_
				"<param name=""ShowStatusBar"" value=""-1"" /><param name=""Filename"" value=""$3"" />"&_
				"<embed "& DV_UBB_TITLE &" type=""application/x-oleobject"" "&_
				"codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701"" flename=""mp"" src=""$3"" width=""$1"" height=""$2""></embed></object>",_
				"<object "& DV_UBB_TITLE &" align=""middle"" classid=""CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95"" class=""object"" id=""MediaPlayer"" width=""$1"" height=""$2"" >"&_
				"<param name=""ShowStatusBar"" value=""-1"" /><param name=""Filename"" value=""$3"" /><param name=""AUTOSTART"" value=""false"" />"&_
				"<embed "& DV_UBB_TITLE &" type=""application/x-oleobject"" "&_
				"codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701"" flename=""mp"" src=""$3"" width=""$1"" height=""$2""></embed></object>"&_
				replace(Mtinfo,"$4","$3"),_
				PostUserGroup,Cint(Board_Setting(9) * mt),"=*([0-9]*),*([0-9]*)")
				'Dv7 MediaPlayer自定义播放模式；
				s=Dv_UbbCode_iS2(s,"mp",_
				"<object "& DV_UBB_TITLE &" align=""middle"" classid=""CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95"" class=""object"" id=""MediaPlayer"" width=""$1"" height=""$2"" >"&_
				"<param name=""AUTOSTART"" value=""$3"" /><param name=""ShowStatusBar"" value=""-1"" /><param name=""Filename"" value=""$4"" />"&_
				"<embed "& DV_UBB_TITLE &" type=""application/x-oleobject"" codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701"" flename=""mp"" src=""$4"" width=""$1"" height=""$2""></embed></object>",_
				"<object "& DV_UBB_TITLE &" align=""middle"" classid=""CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95"" class=""object"" id=""MediaPlayer"" width=""$1"" height=""$2"" >"&_
				"<param name=""AUTOSTART"" value=""false"" /><param name=""ShowStatusBar"" value=""-1"" /><param name=""Filename"" value=""$4"" />"&_
				"<embed "& DV_UBB_TITLE &" type=""application/x-oleobject"" codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701"" flename=""mp"" src=""$4"" width=""$1"" height=""$2""></embed></object>"&_
				 Mtinfo,PostUserGroup,Cint(Board_Setting(9) * mt),"=*([0-9]*),*([0-9]*),*([0|1|true|false]*)")
			End If
			'rm
			If InStr(Ubblists,",6,")>0 Or sType=1 Then
				s=Dv_UbbCode_iS2(s,"rm",_
				"<div><object "& DV_UBB_TITLE &" classid=""clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA"" class=""object"" id=""RAOCX"" width=""$1"" height=""$2"">"&_
				"<param name=""src"" value=""$3"" />"&_
				"<param name=""CONSOLE"" value=""Clip1"" />"&_
				"<param name=""CONtrOLS"" value=""imagewindow"" />"&_
				"<param name=""AUTOSTART"" value=""true"" /></object></div>"&_
				"<div><object "& DV_UBB_TITLE &" classid=""CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA"" height=""32"" id=""video2"" width=""$1"">"&_
				"<param name=""src"" value=""$3"" /><param name=""AUTOSTART"" value=""-1"" />"&_
				"<param name=""CONtrOLS"" value=""controlpanel"" />"&_
				"<param name=""CONSOLE"" value=""Clip1"" /></object></div>",_
				"<div><object "& DV_UBB_TITLE &" classid=""clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA"" class=""object"" id=""RAOCX"" width=""$1"" height=""$2"">"&_
				"<param name=""src"" value=""$3"" />"&_
				"<param name=""CONSOLE"" value=""Clip1"" />"&_
				"<param name=""CONtrOLS"" value=""imagewindow"" />"&_
				"<param name=""AUTOSTART"" value=""false"" /></object></div>"&_
				"<div><object "& DV_UBB_TITLE &" classid=""CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA"" height=""32"" id=""video2"" width=""$1"">"&_
				"<param name=""src"" value=""$3"" /><param name=""AUTOSTART"" value=""false"" />"&_
				"<param name=""CONtrOLS"" value=""controlpanel"" />"&_
				"<param name=""CONSOLE"" value=""Clip1"" /></object></div>"& replace(Mtinfo,"$4","$3"),_
				PostUserGroup,Cint(Board_Setting(9) * mt),"=*([0-9]*),*([0-9]*)")
				'Dv7 RealPlayer自定义播放模式；
				s=Dv_UbbCode_iS2(s,"rm",_
				"<div><object "& DV_UBB_TITLE &" classid=""clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA"" class=""object"" id=""RAOCX"" width=""$1"" height=""$2"">"&_
				"<param name=""src"" value=""$4"" /><param name=""CONSOLE"" value=""$4"" /><param name=""CONtrOLS"" value=""imagewindow"" />"&_
				"<param name=""AUTOSTART"" value=""$3"" /></object></div>"&_
				"<div><object "& DV_UBB_TITLE &" classid=""CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA"" height=""32"" id=""video"" width=""$1"">"&_
				"<param name=""src"" value=""$4"" />"&_
				"<param name=""AUTOSTART"" value=""$3"" />"&_
				"<param name=""CONtrOLS"" value=""controlpanel"" /><param name=""CONSOLE"" value=""$4"" /></object></div>",_
				"<div><object "& DV_UBB_TITLE &" classid=""clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA"" class=""object"" id=""RAOCX"" width=""$1"" height=""$2"">"&_
				"<param name=""src"" value=""$4"" /><param name=""CONSOLE"" value=""$4"" /><param name=""CONtrOLS"" value=""imagewindow"" />"&_
				"<param name=""AUTOSTART"" value=""false"" /></object></div>"&_
				"<div><object "& DV_UBB_TITLE &" classid=""CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA"" height=""32"" id=""video"" width=""$1"">"&_
				"<param name=""src"" value=""$4"" />"&_
				"<param name=""AUTOSTART"" value=""false"" />"&_
				"<param name=""CONtrOLS"" value=""controlpanel"" /><param name=""CONSOLE"" value=""$4"" /></object></div>"&_
				Mtinfo,PostUserGroup,Cint(Board_Setting(9) * mt),"=*([0-9]*),*([0-9]*),*([0|1|true|false]*)")
			End If
			If InStr(Ubblists,",7,")>0 Or sType=1 Then
				s=Dv_UbbCode_iS2(s,"sound",_
				"<a href=""$1"" target=""_blank""><img "& DV_UBB_TITLE &" src=""skins/default/filetype/mid.gif"" border=""0"" alt=""背景音乐"" /></a><bgsound src=""$1"" loop=""-1"" />",_
				"<a href=""$1"" target=""_blank"">$1</a>"& replace(Mtinfo,"$4","$1"),_
				PostUserGroup,Cint(Board_Setting(9) * mt),"")
			End If
			'flash code
			If InStr(Ubblists,",8,")>0 Or sType=1 Then
				s=Dv_UbbCode_iS2(s,"flash",_
				"<a href=""$1"" target=""_blank""><img "& DV_UBB_TITLE &" src=""skins/default/filetype/swf.gif"" border=""0"" alt=""点击开新窗口欣赏该FLASH动画!"" height=""16"" width=""16"" />[全屏欣赏]</a><br/>"&_
				"<object "& DV_UBB_TITLE &" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0"" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000""  width=""500"" height=""400"">"&_
				"<param name=""movie"" value=""$1"" /><param name=""quality"" value=""high"" />"&_
				"<embed "& DV_UBB_TITLE &" src=""$1"" quality=""high"" pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" width=""500"" height=""400"">$1</embed></object>",_
				"<img "& DV_UBB_TITLE &" src=""skins/default/filetype/swf.gif"" border=""0""> <a href=$1 target=""_blank"">$1</a>"& replace(Mtinfo,"$4","$1"),_
				PostUserGroup,Cint(Board_Setting(44)),"")
				s=Dv_UbbCode_iS2(s,"flash",_
				"<a href=""$3"" target=""_blank""><img "& DV_UBB_TITLE &" src=""skins/default/filetype/swf.gif"" border=""0"" alt=""点击开新窗口欣赏该FLASH动画!"" height=""16"" width=""16"" />[全屏欣赏]</a><br/>"&_
				"<object "& DV_UBB_TITLE &" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0"" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000""  width=""$1"" height=""$2"">"&_
				"<param name=""movie"" value=""$3"" /><param name=""quality"" value=""high"" />"&_
				"<embed "& DV_UBB_TITLE &" src=""$3"" quality=""high"" pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" width=""$1"" height=""$2"">$3</embed></object>",_
				"<a href=""$3"" target=""_blank"">$3</a>",PostUserGroup,Cint(Board_Setting(44)),"=*([0-9]*),*([0-9]*)")
			End If
			'point view
			If InStr(Ubblists,",9,")>0 Or sType=1 Then
				s=Dv_UbbCode_Get(s,PostUserGroup,PostType,"money",_
				"<hr size=""1"" /><font color=""gray"">以下内容需要金钱数达到<b>$1</b>才可以浏览</font><br />$2<hr size=""1"" />",_
				"<hr size=""1"" /><font color="""&Dvbbs.Mainsetting(1)&""">以下内容需要金钱数达到<b>$1</b>才可以浏览</font><hr size=""1"" />",_
				UserPointInfo(0),Cint(Board_Setting(10)))
			End If
			If InStr(Ubblists,",10,")>0 Or sType=1 Then
				s=Dv_UbbCode_Get(s,PostUserGroup,PostType,"point",_
				"<hr size=""1"" /><font color=""gray"">以下内容需要积分达到<b>$1</b>才可以浏览</font><br/>$2<hr size=""1"" />",_
				"<hr size=""1"" /><font color="""&Dvbbs.Mainsetting(1)&""">以下内容需要积分达到<b>$1</b>才可以浏览</font><hr size=""1"" />",_
				UserPointInfo(1),Cint(Board_Setting(11)))
			End If
			If InStr(Ubblists,",11,")>0 Or sType=1 Then
				s=Dv_UbbCode_Get(s,PostUserGroup,PostType,_
				"UserCP","<hr size=""1"" /><font color=""gray"">以下内容需要魅力达到<b>$1</b>才可以浏览</font><br/>$2<hr size=""1"" />",_
				"<hr size=""1"" /><font color="""&Dvbbs.Mainsetting(1)&""">以下内容需要魅力达到<b>$1</b>才可以浏览</font><hr size=""1"" />",_
				UserPointInfo(2),Cint(Board_Setting(12)))
			End If
			If InStr(Ubblists,",12,")>0 Or sType=1 Then
				s=Dv_UbbCode_Get(s,PostUserGroup,PostType,_
				"Power","<hr size=""1"" /><font color=""gray"">以下内容需要威望达到<b>$1</b>才可以浏览</font><br/>$2<hr size=""1"" />",_
				"<hr size=""1"" /><font color="""&Dvbbs.Mainsetting(1)&""">以下内容需要威望达到<b>$1</b>才可以浏览</font><hr size=""1"" />",_
				UserPointInfo(3),Cint(Board_Setting(13)))
			End If
			If InStr(Ubblists,",13,")>0 Or sType=1 Then
				s=Dv_UbbCode_Get(s,PostUserGroup,PostType,"Post",_
				"<hr size=""1"" /><font color=""gray"">以下内容需要帖子数达到<b>$1</b>才可以浏览</font><br/>$2<hr size=""1"" />",_
				"<hr size=""1"" /><font color="""&Dvbbs.Mainsetting(1)&""">以下内容需要帖子数达到<b>$1</b>才可以浏览</font><hr size=""1"" />",_
				UserPointInfo(4),Cint(Board_Setting(14)))
			End If
			If InStr(Ubblists,",14,")>0 Or sType=1 Then
				s=UBB_REPLYVIEW(s,PostUserGroup,PostType)
			End If
			If InStr(Ubblists,",15,")>0 Or sType=1 Then
				s=UBB_USEMONEY(s,PostUserGroup,PostType)
			End If
			'url code
			If InStr(Ubblists,",16,")>0 Or sType=1 Then
				s=Dv_UbbCode_S1(s,"url","<a href=""$1"" target=""_blank"">$1</a>")
				s=Dv_UbbCode_UF(s,"url","<a href=""$1"" target=""_blank"">$2</a>","0")
			End If
			'email code
			If InStr(Ubblists,",17,")>0 Or sType=1 Then
				s=Dv_UbbCode_S1(s,"email","<img "& DV_UBB_TITLE &" align=""absmiddle"" src=""skins/default/email1.gif"" alt=""""/><a href=""mailto:$1"">$1</a>")
				s=Dv_UbbCode_UF(s,"email","<img "& DV_UBB_TITLE &" align=""absmiddle"" src=""skins/default/email1.gif"" alt=""""/><a href=""mailto:$1"" target=""_blank"">$2</a>","0")
			End If
			If InStr(Ubblists,",37,")>0 Or sType=1 Then
				If (Cint(Board_Setting(8)) = 1 Or PostUserGroup<4) And InStr(Lcase(s),"[em")>0 Then
					re.Pattern="\[em([0-9]+)\]"
					s=re.Replace(s,"<img "& DV_UBB_TITLE &" src="""&EmotPath&"em$1.gif"" border=""0"" align=""middle"" alt="""" />")
				End If
			End If
			If InStr(Ubblists,",23,")>0 Or sType=1 Then
				s=Dv_UbbCode_C(s,"html")
			End If
			If InStr(Ubblists,",24,")>0 Or sType=1 Then
				s=Dv_UbbCode_S1(s,"code","<pre class=""htmlcode""><b>以下内容为程序代码:</b><br/>$1</pre>")
			End If
			If InStr(Ubblists,",25,")>0 Or sType=1 Then
				s=Dv_UbbCode_UF(s,"color","<font color=""$1"">$2</font>","1")
			End If
			If InStr(Ubblists,",26,")>0 Or sType=1 Then
				s=Dv_UbbCode_UF(s,"face","<font face=""$1"">$2</font>","1")
			End If
			If InStr(Ubblists,",27,")>0 Or sType=1 Then
				s=Dv_UbbCode_Align(s)
			End If
			If InStr(Ubblists,",42,")>0 Or sType=1 Then
				s=Dv_UbbCode_S1(s,"center","<div align=""center"">$1</div>")
			End If
			If InStr(Ubblists,",28,")>0 Or sType=1 Then
				s=Dv_UbbCode_Q(s)
			End If
			If InStr(Ubblists,",29,")>0 Or sType=1 Then
				s=Dv_UbbCode_S1(s,"fly","<marquee width=""90%"" behavior=""alternate"" scrollamount=""3"">$1</marquee>")
			End If
			If InStr(Ubblists,",30,")>0 Or sType=1 Then
				s=Dv_UbbCode_S1(s,"move","<marquee scrollamount=""3"">$1</marquee>")
			End If
			If InStr(Ubblists,",31,")>0 Or sType=1 Then
				s=Dv_UbbCode_iS1(s,"shadow","<div style=""width:$1px;filter:shadow(color=$2, strength=$3)"">$4</div>")
			End If
			If InStr(Ubblists,",32,")>0 Or sType=1 Then
				s=Dv_UbbCode_iS1(s,"glow","<div style=""width:$1px;filter:glow(color=$2, strength=$3)"">$4</div>")
			End If
			If InStr(Ubblists,",33,")>0 Or sType=1 Then
				s=Dv_UbbCode_UF(s,"size","<font size=""$1"">$2</font>","1")
			End If
			If InStr(Ubblists,",34,")>0 Or sType=1 Then
				s=Dv_UbbCode_S1(s,"i","<i>$1</i>")
			End If
			If InStr(Ubblists,",35,")>0 Or sType=1 Then
				s=Dv_UbbCode_S1(s,"b","<b>$1</b>")
			End If
			If InStr(Ubblists,",36,")>0 Or sType=1 Then
				s=Dv_UbbCode_S1(s,"u","<u>$1</u>")
			End If
			If InStr(Ubblists,",41,")>0 Or sType=1 Then
				s= Dv_UbbCode_name(s)
			End If
			'如果没有更新过帖子数据，而定员帖失效的，请把下面的注释去掉，建议进行帖子数据更新，以提高性能 2005.10.10 By Winder.F
			'If InStr(Lcase(s),"[username")>0 Then s= Dv_UbbCode_name(s)
			
			If InStr(s,"payto:") = 0 Then
				s = Replace(s,"https://www.alipay.com/payt","https://www.alipay.com/payto:")
			End If
			If InStr(Ubblists,",40,")>0 Then
				s=Dv_Alipay_PayTo(s)
			End If
			'自动识别网址
			'If InStr(Ubblists,",18,")>0 Or InStr(Ubblists,",19,")>0 Or InStr(Ubblists,",20,")>0 Or InStr(Ubblists,",21,")>0 Or InStr(Ubblists,",22,")>0 Or sType=1 Then
			'	re.Pattern = "(^|[^<=""])(((http|https|ftp|rtsp|mms):(\/\/|\\\\))(([\w\/\\\+\-~`@:%])+\.)+([\w\/\\\.\=\?\+\-~`@\':!%#]|(&amp;)|&)+)"
			'	s = re.Replace(s,"$1<a target=""_blank"" href=""$2"">$2</a>")
			'End If
			'自动识别www等开头的网址
			'If InStr(Ubblists,",38,")>0 Or sType=1 Then
			'	re.Pattern = "(^|[^\/\\\w\=])((www|bbs)\.(\w)+\.([\w\/\\\.\=\?\+\-~`@\'!%#]|(&amp;))+)"
			'	s = re.Replace(s,"$1<a  href=""http://$2"">$2</a>")
			'End If
			If textonly Then
				If xml.loadxml("<div>" & xmlencode(s) &"</div>") Then
					s=checkXHTML(mt,PostUserGroup)
					s=checkimg(s)
					If showisxhtml=1 and Dvbbs.master Then
						s=s&"<p style=""color:green"" align=""right"">[符合XHML规范，内容为纯文本或UBB(UBB解释文件版本："&ubb_version&")]</p>"
					End If
				Else
					If Dv_FilterJS(s) Then
							re.Pattern="\[(br)\]"
							s=re.Replace(s,"<$1>")
							re.Pattern = "(&nbsp;)"
							s = re.Replace(s,Chr(9))
							re.Pattern = "(<br/>)"
							s = re.Replace(s,vbNewLine)
							re.Pattern = "(<br>)"
							s = re.Replace(s,vbNewLine)
							re.Pattern = "(<p>)"
							s = re.Replace(s,"")
							re.Pattern = "(<\/p>)"
							s = re.Replace(s,vbNewLine)
							s=server.htmlencode(s)
							s="<form name=""scode"&replyid&""" method=""post"" action=""""><table class=""tableborder2"" cellspacing=""1"" cellpadding=""3"" width=""100%"" align=""center"" border=""0""><tr><th height=""22"">以下内容含脚本,或可能导致页面不正常的代码</th></tr><tr><td class=""tablebody1"" align=""middle"" width=""98%""><textarea id=""CodeText"" style=""BORDER-RIGHT: 1px dotted; BORDER-TOP: 1px dotted; OVERFLOW-Y: visible; OVERFLOW: visible; BORDER-LEFT: 1px dotted; WIDth: 98%; COLOR: #000000; BORDER-BOTTOM: 1px dotted"" rows=""20"" cols=""120"">"&s&"</textarea></td></tr><tr><td class=""tablebody2"" align=""middle"" width=""98%""><b>说明：</b>上面显示的是代码内容。您可以先检查过代码没问题，或修改之后再运行.</td></tr><tr><td class=""tablebody1"" align=""middle"" width=""98%""><input type=""button"" name=""run"" value=""运行代码"" onclick=""Dvbbs_ViewCode("&replyid&");""></td></tr></table></form>"
							Dv_UbbCode=s
							Exit Function
						Else
							s=bbimg(s)
						End If
						If showisxhtml=1 and Dvbbs.master Then
								s=s&"<p style=""color:red"" align=""right"">[不符合XHML规范(UBB解释文件版本："&ubb_version&")]</p>"
						End If
				End If
			Else	
				If xml.loadxml("<div>" & xmlencode(s) &"</div>") Then
					s=checkXHTML(mt,PostUserGroup)
					s=checkimg(s)
					If showisxhtml=1 and Dvbbs.master Then
						 	s=s&"<p style=""color:green"" align=""right"">[符合XHML规范，内容含HTML代码(UBB解释文件版本："&ubb_version&")]</p>"
					End If
				Else
						If Dv_FilterJS(s) Then
							re.Pattern="\[(br)\]"
							s=re.Replace(s,"<$1>")
							re.Pattern = "(&nbsp;)"
							s = re.Replace(s,Chr(9))
							re.Pattern = "(<br/>)"
							s = re.Replace(s,vbNewLine)
							re.Pattern = "(<br>)"
							s = re.Replace(s,vbNewLine)
							re.Pattern = "(<p>)"
							s = re.Replace(s,"")
							re.Pattern = "(<\/p>)"
							s = re.Replace(s,vbNewLine)
							s=server.htmlencode(s)
							s="<form name=""scode"&replyid&""" method=""post"" action=""""><table class=""tableborder2"" cellspacing=""1"" cellpadding=""3"" width=""100%"" align=""center"" border=""0""><tr><th height=""22"">以下内容含脚本,或可能导致页面不正常的代码</th></tr><tr><td class=""tablebody1"" align=""middle"" width=""98%""><textarea id=""CodeText"" style=""BORDER-RIGHT: 1px dotted; BORDER-TOP: 1px dotted; OVERFLOW-Y: visible; OVERFLOW: visible; BORDER-LEFT: 1px dotted; WIDth: 98%; COLOR: #000000; BORDER-BOTTOM: 1px dotted"" rows=""20"" cols=""120"">"&s&"</textarea></td></tr><tr><td class=""tablebody2"" align=""middle"" width=""98%""><b>说明：</b>上面显示的是代码内容。您可以先检查过代码没问题，或修改之后再运行.</td></tr><tr><td class=""tablebody1"" align=""middle"" width=""98%""><input type=""button"" name=""run"" value=""运行代码"" onclick=""Dvbbs_ViewCode("&replyid&");""></td></tr></table></form>"
							Dv_UbbCode=s
							Exit Function
						Else
							s=bbimg(s)
						End If
						If showisxhtml=1 and Dvbbs.master Then
								s=s&"<p style=""color:red"" align=""right"">[不符合XHML规范(UBB解释文件版本："&ubb_version&")]</p>"
						End If
					End If
			End If
		
		Dv_UbbCode = Rexmlencode(s)
	End Function
	Private Sub checktag(mt,PostUserGroup)
		Dim node,newnode,nodetext,attributes1,attributes2,Fnode,iskill
		Rem  新xhtml 格式处理
		Rem 检索有害标记实行过滤
		Dim Stylestr,style,style1,newstyle,style_a,style_b
		For Each Node in xml.documentElement.getElementsByTagName("*")
			If LCase(Node.nodeName)="link" _
			Or LCase(Node.nodeName)="iframe" _
			Or LCase(Node.nodeName)="meta" _
			Or LCase(Node.nodeName)="script"  _
			Or LCase(Node.nodeName)="frameset"  _
			Or LCase(Node.nodeName)="layer"  _
			Or LCase(Node.nodeName)="style"  _
			Or LCase(Node.nodeName)="xss"  _
			Or LCase(Node.nodeName)="base"  _
			Or LCase(Node.nodeName)="html"  _
			Or LCase(Node.nodeName)="xhtml"  _
			Or LCase(Node.nodeName)="xml"  _
			Then
				Set newnode=xml.createTextNode(node.xml)		
				node.parentNode.replaceChild newnode,node
				checktag mt,PostUserGroup
				Exit Sub
			End If
			If LCase(Node.nodeName)="object" Then
				For Each newnode in node.attributes
					If LCase(newNode.nodeName)="data" Then
						node.attributes.removeNamedItem(newNode.nodeName)
						checktag mt,PostUserGroup
						Exit Sub
					End If
				Next
			End If
			If LCase(Node.nodeName)="a" Then
					Node.attributes.setNamedItem(xml.createNode(2,"target","")).text="_blank"
			End If
		Next
		'在style里吃掉xss和expression
		Dim i
		For Each node in xml.documentElement.selectNodes("//@*")
				If LCase(Node.nodeName)="style" Then
					Stylestr=node.text
					Stylestr=split(Stylestr,";")
					newstyle=""
				 	For each style in Stylestr
				 		style1=split(style,":")
				 		If UBound(style1)>0 Then
				 			style_a=LCase(Trim(style1(0)))
					 		style_b=LCase(Trim(style1(1)))
					 		If UBound(style1)>1 Then
					 				For i =2 to UBound(style1)
					 				style_b=style_b& ":"& style1(i)
					 				Next
					 			End If
					 		'吃掉POSITION:,top,left几个属性
					 		If (style_a<>"xss" and InStr(style_b,"expression")=0 and InStr(style_a,"script")=0) Then
					 			newstyle=newstyle&style_a&":"&style_b&";"
					 		End If
				 		End If
					Next
					node.text=newstyle
				End If
			Next
		'所有的属性的检查过滤
		For Each newnode in xml.documentElement.selectNodes("//*")
				For Each node in newnode.attributes
					If Left(LCase(Node.nodeName),2)="on" Then
						newnode.attributes.removeNamedItem(Node.nodeName)
						checktag mt,PostUserGroup
						Exit Sub
					Else
						nodetext=entity2Str(node.text)
						If InStr(nodetext,"script:")>0 or InStr(nodetext,"document.")>0 Or InStr(nodetext,"xss:") > 0 Or InStr(nodetext,"expression") > 0 Then
								newnode.attributes.removeNamedItem(Node.nodeName)
								checktag mt,PostUserGroup
								Exit Sub
						End If
					End If
				Next
		Next
		'改掉所有的id属性避免造成混乱
		For Each newnode in xml.documentElement.selectNodes("//*")
			For Each node in newnode.attributes
					If LCase(Node.nodeName)="id" or  LCase(Node.nodeName)="name" Then
						node.text=replace(node.text,"dv_","")
					End If
				Next
		Next
		Dim hasname,hasvalue
		For Each Node in xml.documentElement.getElementsByTagName("*")
			If LCase(Node.nodeName)="param" Then
				hasname=0
				hasvalue=0
				iskill=0
				For each attributes1 in Node.attributes
					If LCase(attributes1.nodeName)="name" Then
						If LCase(attributes1.text)="filename" or LCase(attributes1.text)="src" or LCase(attributes1.text)="console" Then
							hasname=1
						End If
					End If
					If LCase(attributes1.nodeName)="value" Then
						hasvalue=1
					End If
				Next
				If hasname=1 and hasvalue=1 Then
					For each attributes1 in Node.attributes
						If LCase(attributes1.nodeName)="value" Then
							nodetext=entity2Str(attributes1.text)
							If InStr(nodetext,".asp")>0 Then
								nodetext=replace(nodetext,"showimg.asp","")
								If InStr(nodetext,".asp")>0 Then
									iskill=1
								End If
							End If
						End If
					Next	
				End If
				If iskill=1 Then
					Set Fnode=node.parentNode.parentNode
					If Not Fnode is nothing Then
						Set newnode=xml.createTextNode(node.parentNode.xml)		
						node.parentNode.parentNode.replaceChild newnode,node.parentNode
					Else
						Set newnode=xml.createTextNode(node.xml)
						node.parentNode.replaceChild newnode,node
					End If
					checktag mt,PostUserGroup
					Exit Sub
				End If
			Else
				For each attributes1 in Node.attributes
				If LCase(attributes1.nodeName)="src" Then
						nodetext=entity2Str(attributes1.text)
						If InStr(nodetext,".asp")>0 Then
							nodetext=replace(nodetext,"showimg.asp","")
							If InStr(nodetext,".asp")>0 Then
								Set newnode=xml.createTextNode(node.xml)		
								node.parentNode.replaceChild newnode,node
								checktag mt,PostUserGroup
								Exit Sub
							End If
						End If
				End If	
				Next
			End If
		Next
		Dim XML1,titletext,thissrc,objcount
		If Cint(Board_Setting(9) * mt)=0 Then
			For Each Node in xml.documentElement.getElementsByTagName("*")
				'禁止所有自动播放的标签
				If LCase(Node.nodeName)="param" Then
					hasname=0
					hasvalue=0
					For each attributes1 in Node.attributes
						If LCase(attributes1.nodeName)="name" Then
							If LCase(attributes1.text)="autostart" Then
								hasname=1
							End If
						End If
						If LCase(attributes1.nodeName)="value" Then
								hasvalue=1
						End If
					Next
					If hasname=1 and hasvalue=1 Then
						For each attributes1 in Node.attributes
							If LCase(attributes1.nodeName)="value" Then
								Node.attributes.removeNamedItem(attributes1.nodeName)
							End If
						Next
						Node.attributes.setNamedItem(xml.createNode(2,"value","")).text="false"
					End If
				ElseIf LCase(Node.nodeName)="embed" Then
					For each attributes1 in Node.attributes
						If LCase(attributes1.nodeName)="autoplay" Then
							Node.attributes.removeNamedItem(attributes1.nodeName)
						End If
						If LCase(attributes1.nodeName)="src" Then
							thissrc=entity2Str(attributes1.text)
							If InStr(thissrc,".swf")> 0 Or InStr(thissrc,".swi") > 0 Then
								Node.attributes.removeNamedItem(attributes1.nodeName)
							End If
						End If
						
					Next
					Node.attributes.setNamedItem(xml.createNode(2,"autoplay","")).text="false"
				End If
			Next
		'加插媒体被禁止自动播放的标签
		
		For Each Node in xml.documentElement.getElementsByTagName("embed")
				Set titletext=node.attributes.getNamedItem("title")
				If titletext is nothing Then
					titletext=""
				Else
					titletext=titletext.text
				End If 
				If titletext<>UBB_TITLE Then
					Node.attributes.setNamedItem(xml.createNode(2,"title","")).text=UBB_TITLE
					Set xml1=Server.Createobject("msxml2.DOMDocument"& MsxmlVersion)
					Set thissrc=node.attributes.getNamedItem("src")
					If thissrc is nothing Then
						thissrc=""
					Else
						thissrc=entity2Str(thissrc.text)
					End If
					thissrc="<div>"& Node.xml& replace(Mtinfo,"$4",thissrc)&"</div>"
					If xml1.loadxml(thissrc) Then
						node.parentNode.replaceChild xml1.documentElement,node
						checktag mt,PostUserGroup
						Exit Sub
					End If	
				End If
		Next
		End If
		
		If instr(","& can_Post_Style &",",","& PostUserGroup &",") = 0 Then
			For Each Node in xml.documentElement.selectNodes("//@*")
				If LCase(Node.nodeName)="style" Then
					Stylestr=node.text
					Stylestr=split(Stylestr,";")
					newstyle=""
				 	For each style in Stylestr
				 		style1=split(style,":")
				 		If UBound(style1)>0 Then
				 			style_a=LCase(Trim(style1(0)))
					 		style_b=LCase(Trim(style1(1)))
					 		If UBound(style1)>1 Then
					 				For i =2 to UBound(style1)
					 				style_b=style_b& ":"& style1(i)
					 				Next
					 		End If
					 		'吃掉POSITION:,top,left几个属性
					 		If (style_a<>"top" and style_a<>"left" and style_a<>"bottom" and style_a<>"right" and style_a<>"" and style_a<> "position") Then
						 			'去掉过宽的属性
						 			If style_a="width" Then
						 				If InStr(style_b,"px")>0 Then
						 					style_b=replace(style_b,"px","")
						 					If IsNumeric(style_b) Then
						 						If CLng(style_b)>600 Then style_b=600
						 					End If
						 					style_b=style_b&"px"
						 				ElseIf InStr(style_b,"%")>0 Then
						 					style_b=replace(style_b,"%","")
						 					If IsNumeric(style_b) Then
						 						If CLng(style_b)>100 Then style_b=100
						 					End If
						 					style_b=style_b&"%"
						 			End If
					 				'去掉过大的字体
						 			If style_a = "font-size" Then
						 				If InStr(style_b,"px")>0 Then
						 					style_b=replace(style_b,"px","")
						 					If IsNumeric(style_b) Then
						 						If CLng(style_b)> 200 Then style_b=200
						 					End If
						 					style_b=style_b&"px"
						 				ElseIf InStr(style_b,"%")> 0 Then
						 					style_b=replace(style_b,"%","")
						 					If IsNumeric(style_b) Then
						 						If CLng(style_b)>100 Then style_b=100
						 					End If
						 					style_b=style_b&"%"
						 				End If
						 			End If
					 			End If 
					 			newstyle=newstyle&style_a&":"&style_b&";"
					 		End If
				 		End If
					Next
					node.text=newstyle
				End If
			Next
		End If
	End Sub
	Private Function checkXHTML(mt,PostUserGroup)
		checktag mt,PostUserGroup
		checkXHTML=Rexmlencode(Mid(xml.documentElement.xml,6,Len (xml.documentElement.xml)-11))
	End Function
	Function checkimg(textstr)
		Dim node,titletext,srctext,newnode
		If xml.loadxml("<div>" & xmlencode(textstr) &"</div>")Then
			For Each Node in xml.documentElement.getElementsByTagName("img")
				Set titletext=node.attributes.getNamedItem("title")
				If titletext is nothing Then
					titletext=""
				Else
					titletext=titletext.text
				End If 
				If titletext=UBB_TITLE Then
					Rem 是否开启滚轮改变图片大小的功能，如果不需要可以屏蔽
					Rem Node.attributes.setNamedItem(xml.createNode(2,"onmousewheel","")).text="return bbimg(this);"
					Node.attributes.setNamedItem(xml.createNode(2,"onload","")).text="imgresize(this);"
					Node.attributes.setNamedItem(xml.createNode(2,"alt","")).text="图片点击可在新窗口打开查看"	
				Else
					Rem 是否开启滚轮改变图片大小的功能，如果不需要可以屏蔽
					Rem Node.attributes.setNamedItem(xml.createNode(2,"onmousewheel","")).text="return bbimg(this);"
					Node.attributes.setNamedItem(xml.createNode(2,"onload","")).text="imgresize(this);"
					Node.attributes.setNamedItem(xml.createNode(2,"style","")).text="cursor: pointer;"
					Node.attributes.setNamedItem(xml.createNode(2,"alt","")).text="图片点击可在新窗口打开查看"
					Node.attributes.setNamedItem(xml.createNode(2,"onclick","")).text="javascript:window.open(this.src);"
					If Not node.parentNode is Nothing Then
						If node.parentNode.nodename = "a" Then
								node.attributes.removeNamedItem("onclick")
						End If
					End If
				End If
			Next
			checkimg=Rexmlencode(Mid(xml.documentElement.xml,6,Len (xml.documentElement.xml)-11))
		Else
			checkimg=textstr
		End If
	End Function
	Private Function bbimg(strText)
		Dim s
		s=strText
		re.Pattern="<img(.[^>]*)([/| ])>"
		s=re.replace(s,"<img$1/>")
		If InStr(Ubblists,",40,")=0 Then
			re.Pattern="<img(.[^>]*)/>"
			s=re.replace(s,"<img$1 onload=""imgresize(this);""/>")
		End If
		bbimg=s
	End Function
	'签名UBB转换
	Public Function Dv_SignUbbCode(s,PostUserGroup)
		Dim ii
		Dim po
		isxhtml=True
		If Dvbbs.forum_setting(66)="0" Then
			s= server.htmlEncode(s)
			re.Pattern="\[\/(img|dir|qt|mp|rm|sound|flash)\]"
			If re.Test(s) Then
				If Dv_FilterJS2(s)Then
					re.Pattern="\[(br)\]"
					s=re.Replace(s,"<$1>")
					re.Pattern = "(&nbsp;)"
					s = re.Replace(s,Chr(9))
					re.Pattern = "(<br/>)"
					s = re.Replace(s,vbNewLine)
					re.Pattern = "(<br />)"
					s = re.Replace(s,vbNewLine)
					re.Pattern = "(<p>)"
					s = re.Replace(s,"")
					re.Pattern = "(<\/p>)"
					s = re.Replace(s,vbNewLine)
					s=server.htmlencode(s)
					s="<form name=""scode"&replyid&""" method=""post"" action=""""><table class=""tableborder2"" cellspacing=""1"" cellpadding=""3"" width=""100%"" align=""center"" border=""0""><tr><th height=""22"">以下内容含脚本,或可能导致页面不正常的代码</th></tr><tr><td class=""tablebody1"" align=""middle"" width=""98%""><textarea id=""CodeText"" style=""BORDER-RIGHT: 1px dotted; BORDER-TOP: 1px dotted; OVERFLOW-Y: visible; OVERFLOW: visible; BORDER-LEFT: 1px dotted; WIDth: 98%; COLOR: #000000; BORDER-BOTTOM: 1px dotted"" rows=""20"" cols=""120"">"&s&"</textarea></td></tr><tr><td class=""tablebody2"" align=""middle"" width=""98%""><b>说明：</b>上面显示的是代码内容。您可以先检查过代码没问题，或修改之后再运行.</td></tr><tr><td class=""tablebody1"" align=""middle"" width=""98%""><input type=""button"" name=""run"" value=""运行代码"" onclick=""Dvbbs_ViewCode("&replyid&");""></td></tr></table></form>"
					s = Replace(s, vbNewLine, "")
					s = Replace(s, Chr(10), "")
					s = Replace(s, Chr(13), "")
					Dv_SignUbbCode=s
					Exit Function
				End If
			End If
			re.Pattern="([^"&Chr(13)&"])"& Chr(10)
			s= re.Replace(s,"$1<br />")
			re.Pattern=Chr(13)&Chr(10)&"(.*)" 
			s= re.Replace(s,"<p>$1</p>")
		Else
				If Dv_FilterJS(s) Then
					re.Pattern="\[(br)\]"
					s=re.Replace(s,"<$1>")
					re.Pattern = "(&nbsp;)"
					s = re.Replace(s,Chr(9))
					re.Pattern = "(<br/>)"
					s = re.Replace(s,vbNewLine)
					re.Pattern = "(<br>)"
					s = re.Replace(s,vbNewLine)
					re.Pattern = "(<p>)"
					s = re.Replace(s,"")
					re.Pattern = "(<\/p>)"
					s = re.Replace(s,vbNewLine)
					s=server.htmlencode(s)
					s="<form name=""scode"&replyid&""" method=""post"" action=""""><table class=""tableborder2"" cellspacing=""1"" cellpadding=""3"" width=""100%"" align=""center"" border=""0""><tr><th height=""22"">以下内容含脚本,或可能导致页面不正常的代码</th></tr><tr><td class=""tablebody1"" align=""middle"" width=""98%""><textarea id=""CodeText"" style=""BORDER-RIGHT: 1px dotted; BORDER-TOP: 1px dotted; OVERFLOW-Y: visible; OVERFLOW: visible; BORDER-LEFT: 1px dotted; WIDth: 98%; COLOR: #000000; BORDER-BOTTOM: 1px dotted"" rows=""20"" cols=""120"">"&s&"</textarea></td></tr><tr><td class=""tablebody2"" align=""middle"" width=""98%""><b>说明：</b>上面显示的是代码内容。您可以先检查过代码没问题，或修改之后再运行.</td></tr><tr><td class=""tablebody1"" align=""middle"" width=""98%""><input type=""button"" name=""run"" value=""运行代码"" onclick=""Dvbbs_ViewCode("&replyid&");""></td></tr></table></form>"
					Dv_SignUbbCode=s
					Exit Function
				End If
			re.Pattern="<((asp|\!|%))"
			s=re.Replace(s,"&lt;$1")
			re.Pattern="(>)("&vbNewLine&")(<)"
			s=re.Replace(s,"$1$3") 
			re.Pattern="(>)("&vbNewLine&vbNewLine&")(<)"
			s=re.Replace(s,"$1$3") 
		End If
		s = Replace(s, "  ", "&nbsp;&nbsp;")
		s = Replace(s, vbNewLine, "<br/>")
		s = Replace(s, Chr(13), "")
		'常规设置不支持UBB代码，则退出
		If Cint(Dvbbs.Forum_setting(65))=0 Then 
			Dv_SignUbbCode=s
			Exit Function
		End If 
		'img code
		If InStr(Lcase(s),"[/img]")>0 Then  s=Dv_UbbCode_iS2(s,"img","<img "& DV_UBB_TITLE &" src=""$1"" border=""0"" />","<img "& DV_UBB_TITLE &" src=""skins/default/filetype/gif.gif"" border=""0"" /><a href=""$1"" target=""_blank"">$1</a>",PostUserGroup,Cint(Dvbbs.forum_setting(67)),"")
		'media code
		If InStr(Lcase(s),"[/sound]")>0 Then s=Dv_UbbCode_iS2(s,"sound","<a href=""$1"" target=""_blank""><img "& DV_UBB_TITLE &" src=""skins/default/filetype/mid.gif""  border=""0"" alt=""背景音乐"" /></a><bgsound src=""$1"" loop=""-1"">","<a href=""$1"" target=""_blank"">$1</a>",PostUserGroup,Cint(Board_Setting(9) * mt),"")
		'flash code
		If InStr(Lcase(s),"[/flash]")>0 Then
			s=Dv_UbbCode_iS2(s,"flash",_
			"<a href=""$1"" target=""_blank""><img "& DV_UBB_TITLE &" src=""skins/default/filetype/swf.gif"" border=""0"" alt=""点击开新窗口欣赏该FLASH动画!"" height=""16"" width=""16"" />[全屏欣赏]</a><br/>"&_
			"<object "& DV_UBB_TITLE &" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0"" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" width=""500"" height=""400"">"&_
			"<param name=""movie"" value=""$1"" /><param name=""quality"" value=""high"" />"&_
			"<embed "& DV_UBB_TITLE &" src=""$1"" quality=""high"" pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" width=""500"" height=""400"">$1</embed></object>",_
			"<img "& DV_UBB_TITLE &" src=""skins/default/filetype/swf.gif"" border=""0"" alt=""""> <a href=""$1"" target=""_blank"">$1</a>（注意：Flash内容可能含有恶意代码）",_
			PostUserGroup,Cint(Dvbbs.forum_setting(71)),"")
			s=Dv_UbbCode_iS2(s,"flash",_
			"<a href=""$3"" target=""_blank""><img "& DV_UBB_TITLE &" src=""skins/default/filetype/swf.gif"" border=""0"" alt=""点击开新窗口欣赏该FLASH动画!"" height=""16"" width=""16"" />[全屏欣赏]</a><br/>"&_
			"<object "& DV_UBB_TITLE &" codeBase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0"" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" width=""$1"" height=""$2"">"&_
			"<param name=""movie"" value=""$3"" /><param name=""quality"" value=""high"" />"&_
			"<embed "& DV_UBB_TITLE &" src=""$3"" quality=""high"" pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" width=""$1"" height=""$2"">$3</embed></object>",_
			"<a href=""$3"" target=""_blank"">$3</a>（注意：Flash内容可能含有恶意代码）",_
			PostUserGroup,Cint(Dvbbs.forum_setting(71)),"=*([0-9]*),*([0-9]*)")
		End If
		'url code
		If InStr(Lcase(s),"[/url]")>0 Then
			s=Dv_UbbCode_S1(s,"url","<a href=""$1"" target=""_blank"">$1</a>")
			s=Dv_UbbCode_UF(s,"url","<a href=""$1"" target=""_blank"">$2</a>","0")
		End If
		'email code
		If InStr(Lcase(s),"[/email]")>0 Then
			s=Dv_UbbCode_S1(s,"email","<img "& DV_UBB_TITLE &" src=""skins/default/email1.gif"" alt="""" /><a href=""mailto:$1"">$1</a>")
			s=Dv_UbbCode_UF(s,"email","<img "& DV_UBB_TITLE &" src=""skins/default/email1.gif"" alt="""" /><a href=""mailto:$1"" target=""_blank"">$2</a>","0")
		End If
		If InStr(Lcase(s),"[/html]")>0 Then s=Dv_UbbCode_C(s,"html")
		If InStr(Lcase(s),"[/color]")>0 Then s=Dv_UbbCode_UF(s,"color","<font color=""$1"">$2</font>","1")
		If InStr(Lcase(s),"[/face]")>0 Then s=Dv_UbbCode_UF(s,"face","<font face=""$1"">$2</font>","1")
		If InStr(Lcase(s),"[/align]")>0 Then s=Dv_UbbCode_Align(s)

		If InStr(Lcase(s),"[/shadow]")>0 Then s=Dv_UbbCode_iS1(s,"shadow","<div style=""width:$1px;filter:shadow(color=$2, strength=$3)"">$4</div>")
		If InStr(Lcase(s),"[/glow]")>0 Then s=Dv_UbbCode_iS1(s,"glow","<div style=""width:$1px;filter:glow(color=$2, strength=$3)"">$4</div>")
		If InStr(Lcase(s),"[/i]")>0 Then s=Dv_UbbCode_S1(s,"i","<i>$1</i>")
		If InStr(Lcase(s),"[/b]")>0 Then s=Dv_UbbCode_S1(s,"b","<b>$1</b>")
		If InStr(Lcase(s),"[/u]")>0 Then s=Dv_UbbCode_S1(s,"u","<u>$1</u>")
		If InStr(Lcase(s),"[/size]")>0 Then
			s=Dv_UbbCode_UF(s,"size","<font size=""$1"">$2</font>","1-"&Maxsize&"")
		End If
		REM ：签名移动(如需使用则把以下屏蔽去掉)
		'If InStr(Lcase(s),"[/fly]")>0 Then s=Dv_UbbCode_S1(s,"fly","<marquee width=""90%"" behavior=""alternate"" scrollamount=""3"">$1</marquee>")
		'If InStr(Lcase(s),"[/move]")>0 Then s=Dv_UbbCode_S1(s,"move","<marquee scrollamount=""3"">$1</marquee>")
		'不开放HTML支持，不转换HREF
		REM 加上签名是否开放HTML判断 2004-5-6 Dvbbs.YangZheng
		If Board_Setting(5)="1" And Dvbbs.Forum_Setting(66) = "1" Then
			'自动识别网址
			If InStr(Lcase(s),"http://")>0 Then
				re.Pattern = "(^|[^<=""])(http:(\/\/|\\\\)(([\w\/\\\+\-~`@:%])+\.)+([\w\/\\\.\=\?\+\-~`@\':!%#]|(&amp;)|&)+)"
				s = re.Replace(s,"$1<a target=""_blank"" href=$2>$2</a>")
			End If
			'自动识别www等开头的网址
			If InStr(Lcase(s),"www.")>0 or InStr(Lcase(s),"bbs.")>0 Then
				re.Pattern = "(^|[^\/\\\w\=])((www|bbs)\.(\w)+\.([\w\/\\\.\=\?\+\-~`@\'!%#]|(&amp;))+)"
				s = re.Replace(s,"$1<a target=""_blank"" href=http://$2>$2</a>")
			End If
		End If
		If xml.loadxml("<div>" & xmlencode(s) &"</div>") Then
			s=checkimg(s)
			s=Rexmlencode(s)
			If showisxhtml=1 and Dvbbs.master Then
				s=s&"<p style=""color:green"" align=""right"">[符合XHML规范]</p>"
			End If
		Else
			s=bbimg(s)
			If showisxhtml=1 and Dvbbs.master Then
				s=s&"<p style=""color:red"" align=""right"">[不符合XHML规范]</p>"
			End If
		End If
		Dv_SignUbbCode=s
	End Function

	Private Function Dv_UbbCode_S1(strText,uCodeC,tCode)
		Dim s
		s=strText
		re.Pattern="\["&uCodeC&"\][\s\n]*\[\/"&uCodeC&"\]"
		s=re.Replace(s,"")
		re.Pattern="\[\/"&uCodeC&"\]"
		s=re.replace(s, Chr(1)&"/"&uCodeC&"]")
		re.Pattern="\["&uCodeC&"\]([^\x01]*)\x01\/"&uCodeC&"\]"
		s=re.Replace(s,tCode)
		re.Pattern="\x01\/"&uCodeC&"\]"
		s=re.replace(s,"[/"&uCodeC&"]")
		If isxhtml Then
			If xml.loadxml("<div>" & xmlencode(s) &"</div>") Then
				Dv_UbbCode_S1=s
			Else
				Dv_UbbCode_S1=strText
			End If
		Else
			Dv_UbbCode_S1=s
		End If
	End Function

	Private Function Dv_UbbCode_UF(strText,uCodeC,tCode,Flag)
		Dim s
		Dim LoopCount
		LoopCount=0
		s=strText
		re.Pattern="\["&uCodeC&"=([^\]]*)\][\s\n ]*\[\/"&uCodeC&"\]"
		s=re.Replace(s,"")
		re.Pattern="\[\/"&uCodeC&"\]"
		s=re.replace(s, chr(1)&"/"&uCodeC&"]")
		re.Pattern="\["&uCodeC&"=([^\]]*)\]([^\x01]*)\x01\/"&uCodeC&"\]"
		If Flag="1" Then 
			Do While Re.Test(s)
				s=re.Replace(s,tCode)
				LoopCount=LoopCount+1
				If LoopCount>MaxLoopCount Then Exit Do
			Loop
		ElseIf Flag="0" Then
			s=re.Replace(s,tCode)
		Else
			re.Pattern="\["&uCodeC&"=(["&Flag&"]*)\]([^\x01]*)\x01\/"&uCodeC&"\]"
			Do While Re.Test(s)
				s=re.Replace(s,tCode)
				LoopCount=LoopCount+1
				If LoopCount>MaxLoopCount Then Exit Do
			Loop
		End If
		re.Pattern="\x01\/"&uCodeC&"\]"
		s=re.replace(s,"[/"&uCodeC&"]")
		If isxhtml Then
			If xml.loadxml("<div>" & xmlencode(s) &"</div>") Then
				Dv_UbbCode_UF=s
			Else
				Dv_UbbCode_UF=strText
			End If
		Else
			Dv_UbbCode_UF=s
		End If
	End Function

	Private Function Dv_UbbCode_iS1(strText,uCodeC,tCode)
		Dim s
		s=strText
		re.Pattern="\["&uCodeC&"=[^\]]*\][\s\n]\[\/"&uCodeC&"\]"
		s=re.Replace(s,"")
		re.Pattern="\[\/"&uCodeC&"\]"
		s=re.replace(s, chr(1)&"/"&uCodeC&"]")
		re.Pattern="\["&uCodeC&"=([0-9]+),(#?[\w]+),([0-9]+)\]([^\x01]*)\x01\/"&uCodeC&"\]"
		s=re.Replace(s,tCode)
		re.Pattern="\x01\/"&uCodeC&"\]"
		s=re.replace(s, "[/"&uCodeC&"]")
		If isxhtml Then
			If xml.loadxml("<div>" & xmlencode(s) &"</div>") Then
				Dv_UbbCode_iS1=s
			Else
				Dv_UbbCode_iS1=strText
			End If
		Else
			Dv_UbbCode_iS1=s
		End If
	End Function

	Private Function Dv_UbbCode_iS2(strText,uCodeC,tCode1,tCode2,PostUserGroup,Flag,iCode)
		Dim s
		s=strText
		re.Pattern="\["&uCodeC&iCode&"\][\s\n]*\[\/"&uCodeC&"\]"
		s=re.replace(s,"")
		re.Pattern="\[\/"&uCodeC&"\]"
		s=re.replace(s, chr(1)&"/"&uCodeC&"]")
		If uCodeC<>"flash" Then
			re.Pattern="\["&uCodeC&"[^\]]*\](([^\x01\n]*)(\.swf|\.swi)([^\x01\n]*))\x01\/"&uCodeC&"\]"
			s=re.Replace(s,"非法"&uCodeC&"多媒体标签，文件地址：$1")
		End If
		If uCodeC="img" Then
			re.Pattern="\["&uCodeC&iCode&"\]([^""\x01\n]*)\x01\/"&uCodeC&"\]"
		Else
			re.Pattern="\["&uCodeC&iCode&"\]([^\x01\n]*)\x01\/"&uCodeC&"\]"
		End If
		If Flag = 1  Then
			s=re.Replace(s,tCode1)
		Else
			s=re.Replace(s,tCode2)
		End If 
		re.Pattern="\x01\/"&uCodeC&"\]"
		s=re.replace(s,"[/"&uCodeC&"]")
		If isxhtml Then
			If xml.loadxml("<div>" & xmlencode(s) &"</div>") Then
				Dv_UbbCode_iS2=s
			Else
				Dv_UbbCode_iS2=strText
			End If
		Else
			Dv_UbbCode_iS2=s
		End If
	End Function

	Private Function Dv_UbbCode_Align(strText)
		Dim s
		s=strText
		re.Pattern="\[align=(center|left|right)\][\s\n]*\[\/align\]"
		s=re.Replace(s,"")
		re.Pattern="\[\/align\]"
		s=re.replace(s,chr(1)&"/align]")
		re.Pattern="\[align=(center|left|right)\]([^\x01]*)\x01\/align\]"
		s=re.Replace(s,"<div align=""$1"">$2</div>")
		re.Pattern="\x01\/align\]"
		s=re.replace(s,"[/align]")
		If isxhtml Then
			If xml.loadxml("<div>" & xmlencode(s) &"</div>") Then
				Dv_UbbCode_Align=s
			Else
				Dv_UbbCode_Align=strText
			End If
		Else
			Dv_UbbCode_Align=s
		End If
	End Function

	Private Function Dv_UbbCode_U(strText,PostUserGroup,Flag)	'(帖子内容，用户组，是否开放图片标签)
		Dim s
		If Dvbbs.Forum_Setting(76)="" Or Dvbbs.Forum_Setting(76)="0" Then Dvbbs.Forum_Setting(76)="UploadFile/"
		If right(Dvbbs.Forum_Setting(76),1)<>"/" Then Dvbbs.Forum_Setting(76)=Dvbbs.Forum_Setting(76)&"/"
		s=strText
		re.Pattern="\[upload=([^\]\n]*)\][\s\n]\[\/UPLOAD\]"
		s=re.Replace(s,"")
		re.Pattern="\[\/UPLOAD\]"
		s=re.replace(s, chr(1)&"/upload]")
		re.Pattern="\[upload=(gif|jpg|jpeg|bmp|png)\]UploadFile/([^\x01\n]*)\x01\/UPLOAD\]"
		If Dvbbs.Forum_Setting(75)="0" Then 
			If Flag = 1 or PostUserGroup<4 Then
				s= re.Replace(s,"<br/><img "& DV_UBB_TITLE &" src=""skins/default/filetype/$1.gif"" border=""0"" />此主题相关图片如下：<br/><a href="""&Dvbbs.Forum_Setting(76)&"$2"" target=""_blank"" ><img "& DV_UBB_TITLE &" src="""&Dvbbs.Forum_Setting(76)&"$2"" border=""0"" alt=""按此在新窗口浏览图片"" /></a>")
			Else
				s= re.Replace(s,"<br/><img "& DV_UBB_TITLE &" src=""skins/default/filetype/$1.gif"" border=""0"" /><a href="""&Dvbbs.Forum_Setting(76)&"$2"" target=""_blank"">"&Dvbbs.Forum_Setting(76)&"$2</a>")
			End If 
		Else
			If Flag = 1 or PostUserGroup<4 Then
				s= re.Replace(s,"<br/><img "& DV_UBB_TITLE &" src=""skins/default/filetype/$1.gif"" border=""0"" />此主题相关图片如下：<br/><a href=""showimg.asp?BoardID="&Dvbbs.BoardID&"&filename=$2"" target=""_blank"" ><img "& DV_UBB_TITLE &" src=""showimg.asp?BoardID="&Dvbbs.BoardID&"&filename=$2"" border=""0"" /></a>")
			Else
				s= re.Replace(s,"<br/><img "& DV_UBB_TITLE &" src=""skins/default/filetype/$1.gif"" border=""0"" /><a href=""showimg.asp?BoardID="&Dvbbs.BoardID&"&filename=$2"" target=""_blank"">showimg.asp?BoardID="&Dvbbs.BoardID&"&filename=$2</a>")
			End If
		End If
		re.Pattern="\[upload=(swf|swi)\]UploadFile/([^\x01\n]*)\x01\/UPLOAD\]"
		If Dvbbs.Forum_Setting(75)="0" Then 
			If Board_Setting(44) = 1 or PostUserGroup<4 Then
				s= re.Replace(s,"<br/><img "& DV_UBB_TITLE &" src=""skins/default/filetype/swf.gif"" border=""0"" /><a href="""&Dvbbs.Forum_Setting(76)&"$2"" target=""_blank"">点击浏览该FLASH文件</a>：<br/>"&_
				"<embed "& DV_UBB_TITLE &" src="""&Dvbbs.Forum_Setting(76)&"$2"" quality=""high"" pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" width=""500"" height=""300""></embed>")
			Else
				s= re.Replace(s,"<br/><img "& DV_UBB_TITLE &" src=""skins/default/filetype/swf.gif"" border=""0"" /><a href="""&Dvbbs.Forum_Setting(76)&"$2"" target=""_blank"">"&Dvbbs.Forum_Setting(76)&"$2</a>")
			End If 
		Else
			s= re.Replace(s,"<br/><img "& DV_UBB_TITLE &" src=""skins/default/filetype/swf.gif"" border=""0"" /><a href=""showimg.asp?BoardID="&Dvbbs.BoardID&"&filename=$2"" target=""_blank"">论坛开启了防盗链，请点击浏览该FLASH文件</a>")
		End If
		re.Pattern="\[upload=([^\]\n]+)\]viewFile\.asp\?id=([0-9]*)\x01\/UPLOAD\]"
		s= re.Replace(s,"<img "& DV_UBB_TITLE &" src=""skins/default/filetype/$1.gif"" border=""0"" /><a href=""viewFile.asp?BoardID="&Dvbbs.Boardid&"&ID=$2"" target=""_blank"">点击浏览该文件</a>")
		re.Pattern="\x01\/upload]"
		re.Pattern="\[upload=([^\]\n]+)\]([^\x01]*)\x01\/UPLOAD\]"
		s= re.Replace(s,"<img "& DV_UBB_TITLE &" src=""skins/default/filetype/$1.gif"" border=""0"" /><a href=""$2"" target=""_blank"">点击浏览该文件</a>")
		re.Pattern="\x01\/upload]"
		s=re.replace(s,"[/upload]")
		If isxhtml Then
			If xml.loadxml("<div>" & xmlencode(s) &"</div>") Then
				Dv_UbbCode_U=s
			Else
				Dv_UbbCode_U=strText
			End If
		Else
				Dv_UbbCode_U=s
		End If
	End Function

	Private Function Dv_UbbCode_Q(strText)
		Dim s
		Dim LoopCount
		LoopCount=0
		s=strText
		re.Pattern="\[quote\]((.|\n)*?)\[\/quote\]"
		Do While re.Test(s)
			s=re.Replace(s,"<div class=""quote"">$1</div>")
			LoopCount=LoopCount+1
			If LoopCount>MaxLoopCount Then Exit Do
		Loop
		Dv_UbbCode_Q=s
	End Function

	Private Function Dv_UbbCode_name(strText)
		Dim s
		Dim po,match
		s=strText
		re.Pattern="\[\/username\]"
		s=re.Replace(s,Chr(1)&"/username]")
		re.Pattern="\[username=([^\]]+)]([^\x01]*)\x01\/username\]"
		If Cint(Board_Setting(56))=1 Then
			Set match = re.Execute(s)
			If match.count>0 Then
				po=re.Replace(match.item(0),",$1,")
				If  Dvbbs.Membername<>"" and (Dvbbs.Membername=UserName or InStr(po,","&Dvbbs.Membername&",")>0 or Dvbbs.master) Then
					s=re.Replace(s,"<hr /><font color=""red"">以下内容是专门发给<b>$1</b>浏览</font><br/>$2<hr />")
				Else
	
					s=re.Replace(s,"<hr /><font color=""gray"">以下内容是专门发给<b>$1</b>浏览</font><br/><hr />")
				End If
			End If
		Else
		s=re.Replace(s,"$2")
		End If
		re.Pattern="\x01\/username\]"
		s=re.Replace(s,"[/username]")
		Set match=Nothing
		If isxhtml Then
			If xml.loadxml("<div>" & xmlencode(s) &"</div>") Then
				Dv_UbbCode_name=s
			Else
				Dv_UbbCode_name=strText
			End If
		Else
			Dv_UbbCode_name=s
		End If
	End Function
	
	Private Function Dv_UbbCode_Get(strText,PostUserGroup,PostType,uCodeC,tCode1,tCode2,UsePoint,Flag)'帖子内容,发帖人组别,发帖类型,,,,,,用户积分,是否开放ubb标签
		Dim s,ii,match
		Dim LoopCount
		s=strText
		UsePoint=CLng(UsePoint)
		re.Pattern="\["&uCodeC&"= *[0-9]*\][\s\n]*\[\/"&uCodeC&"\]"
		s=re.replace(s,"")
		re.Pattern="\[\/"&uCodeC&"\]"
		s=re.replace(s,Chr(1)&"/"&uCodeC&"]")
		re.Pattern="\["&uCodeC&"= *([0-9]+)\]([^\x01]*)\x01\/"&uCodeC&"\]"
		If Issupport=1 Then
			Dim matches
			Set matches = re.Execute(s)
			re.Global=False
			For Each match In matches
				If (Flag=1 or PostUserGroup<4) and PostType=1 Then
					ii=int(match.SubMatches(0))
					If  Dvbbs.Membername<>"" and (Dvbbs.Membername=UserName or UsePoint>=ii or Dvbbs.master) Then
						s=re.Replace(s,tCode1)
					Else
						s=re.Replace(s,tCode2)
					End If
				Else
					s=re.Replace(s,"$2")
				End If
				LoopCount=LoopCount+1
				If LoopCount>MaxLoopCount Then Exit For
			Next
			Set matches=Nothing
		Else
			Dim Test
			re.Global=False
			Test=re.Test(s)
			Do While Test
				If (Flag=1 or PostUserGroup<4) and PostType=1 Then
					Set match = re.Execute(s)
					ii=int(re.Replace(match.item(0),"$1"))
					If  Dvbbs.Membername<>"" and (Dvbbs.Membername=UserName or UsePoint>=ii or Dvbbs.master) Then
						s=re.Replace(s,tCode1)
					Else
						s=re.Replace(s,tCode2)
					End If
				Else
					s=re.Replace(s,"$2")
				End If
				LoopCount=LoopCount+1
				If LoopCount>MaxLoopCount Then Exit Do
				Test=re.Test(s)
			Loop
			Set match=Nothing
		End If
		re.Global=true
		re.Pattern="\x01\/"&uCodeC&"\]"
		s=re.replace(s,"[/"&uCodeC&"]")
		If isxhtml Then
			If xml.loadxml("<div>" & xmlencode(s) &"</div>") Then
				Dv_UbbCode_Get=s
			Else
				Dv_UbbCode_Get=strText
			End If
		Else
			Dv_UbbCode_Get=s
		End If
	End Function

	Private Function UBB_REPLYVIEW(strText,PostUserGroup,PostType)
		Dim s
		Dim vrs
		s=strText
		re.Pattern="\[replyview\][\s\n]*\[\/replyview\]"
		s=re.Replace(s,"")
		re.Pattern="\[\/replyview\]"
		s=re.replace(s, chr(1)&"/replyview]")
		re.Pattern="\[replyview\]([^\x01]*)\x01\/replyview\]"
		If (Board_Setting(15)="1" or PostUserGroup<4) and PostType=1  Then
			If isgetreed<>1 Then 
				Set vrs=dvbbs.execute("select AnnounceID from "&TotalUsetable&" where rootid="&Announceid&" and PostUserID="&Dvbbs.UserID)
				isgetreed=1
				If Not vRs.eof Then
					reed=1 
				Else
					reed=0
				End If
				Set vrs=Nothing
			End If 
			If Dvbbs.Membername<>"" and (reed=1 or Dvbbs.master) Then
				s=re.Replace(s,"<hr noshade=""noshade"" size=""1"" /><font color=""gray"">以下内容只有<b>回复</b>后才可以浏览</font><br/>$1<hr noshade=""noshade"" size=""1"" />")
			Else
				s=re.Replace(s,"<hr noshade=""noshade"" size=""1"" /><font color="""&Dvbbs.Mainsetting(1)&""">以下内容只有<b>回复</b>后才可以浏览</font><hr noshade=""noshade"" size=""1"" />")
			End If 
		Else
			s=re.Replace(s,"$1")
		End If 
		re.Pattern="\x01\/replyview\]"
		s=re.replace(s, "[/replyview]")
		If isxhtml Then
			If xml.loadxml("<div>" & xmlencode(s) &"</div>") Then
				UBB_REPLYVIEW=s
			Else
				UBB_REPLYVIEW=strText
			End If
		Else
			UBB_REPLYVIEW=s
		End If
	End Function

	Private Function UBB_USEMONEY(strText,PostUserGroup,PostType)
		Dim s
		Dim Test
		Dim ii,iii,match,buied
		Dim SplitBuyUser,iPostBuyUser
		Dim LoopCount
		s=strText
		re.Global=False
		re.Pattern="\[USEMONEY=*([0-9]+)\]((.|\n)*)\[\/USEMONEY\]"
		Test=re.Test(s)
		If Test Then
			If T_GetMoneyType >0 Then
				s=re.Replace(s,"<hr size=""1"" /><font color=""gray"">由于使用了金币帖子设置，因此出售帖UBB模式失效，以下是帖子内容：</font>&nbsp;&nbsp;<br />$2<hr size=""1"" />")
			Else
				If (Cint(Board_Setting(23))=1 or PostUserGroup<4) and PostType=1 Then
					Set match = re.Execute(s)
					ii=int(re.Replace(match.item(0),"$1"))
					If  Dvbbs.Membername<>"" and (Dvbbs.Membername=UserName or Dvbbs.master) Then
						If (Not IsNull(PostBuyUser)) And PostBuyUser<>"" Then
							SplitBuyUser=split(PostBuyUser,"|")
							iPostBuyUser="<option value=""0"">已购买用户</option>"
							for iii=0 to ubound(SplitBuyUser)
								iPostBuyUser=iPostBuyUser & "<option value="""&iii&""">"&SplitBuyUser(iii)&"</option>"
							next
						Else
							iPostBuyUser="<option value=""0"">还没有用户购买</option>"
						End If
						s=re.Replace(s,"<hr noshade=""noshade"" size=""1"" /><font color=""gray"">以下内容需要花费现金<b>$1</b>才可以浏览</font>&nbsp;&nbsp;<select size=""1"" name=""buyuser"">"&iPostBuyUser&"</select><br/>$2<hr noshade=""noshade"" size=""1"" />")
						re.Global=true
						re.Pattern="\[\/?USEMONEY=*[0-9]*\]"
						s=re.Replace(s,"")
					Else
						buied=0
						If (Not IsNull(PostBuyUser)) and PostBuyUser<>"" Then
							If Instr("|"&PostBuyUser&"|","|"&Dvbbs.Membername&"|")>0 Then buied=1
						End If
						If buied=1 Then
							s=re.Replace(s,"<hr noshade=""noshade"" size=""1"" /><font color=""gray"">以下内容需要花费现金<b>$1</b>才可以浏览，您已经购买本帖</font><br/>$2<hr noshade=""noshade"" size=""1"" />")
							re.Global=true
							re.Pattern="\[\/?USEMONEY=*[0-9]*\]"
							s=re.Replace(s,"")
						Else
							If Clng(UserPointInfo(0))>=ii Then
								s=re.Replace(s,"<form action=""BuyPost.asp"" mothod=""post""><font color="""&Dvbbs.Mainsetting(1)&""">以下内容需要花费现金<b>$1</b>才可以浏览，您目前有现金<b>"&UserPointInfo(0)&"</b>。<br/><br/>　　　　<input type=""hidden"" name=""boardid"" value="""&Dvbbs.boardid&""" /><input type=""hidden"" value="""&replyid_a&""" name=""replyid"" /><input type=""hidden"" value="""&AnnounceID_a&""" name=""id"" /><input type=""hidden"" value="""&RootID_a&""" name=""rootid""/><input type=""hidden"" value="""&totalusetable&""" name=""posttable"" /><input type=""submit"" name=""submit"" value=""好黑啊…我…我买了！"" />&nbsp;&nbsp;</font></form>")
							Else
								s=re.Replace(s,"<hr noshade=""noshade"" size=""1"" /><font color="""&Dvbbs.Mainsetting(1)&""">以下内容需要花费现金<b>$1</b>才可以浏览，您只有现金<b>"&UserPointInfo(0)&"</b>，无法购买。</font><hr noshade=""noshade"" size=""1"" />")
							End If
						End If
					End If
				Else
					re.Global=true
					re.Pattern="\[\/?USEMONEY=*[0-9]*\]"
					s=re.Replace(s,"")
				End If
				Set match=Nothing
			End If
		End If
		re.Global=true
		If isxhtml Then
			If xml.loadxml("<div>" & xmlencode(s) &"</div>") Then
				UBB_USEMONEY=s
			Else
				UBB_USEMONEY=strText
			End If
		Else
			UBB_USEMONEY=s
		End If	
	End Function
	
	Public Function Dv_FilterJS(v)
		If Not Isnull(V) Then
			Dim t,test,Replacelist,t1
			t=v
			t1=v
			re.Pattern="&#36;"
			t1=re.Replace(t1,"$")
			re.Pattern="&#36"
			t1=re.Replace(t1,"$")
			re.Pattern="&#39;"
			t1=re.Replace(t1,"'")
			re.Pattern="&#39"
			t1=re.Replace(t1,"'")
			t1=entity2Str(t1)
			If InStr(Dvbbs.forum_setting(77),"|")=0 Then 
				Replacelist="(expression|xss:|function|meta|window\.|script|js:|about:|file:|Document\.|vbs:|frame|cookie|on(finish|mouse|Exit=|error|click|key|load|focus|Blur))"
			Else
				Replacelist="("&Dvbbs.forum_setting(77)&"expression|xss:|function|meta|window\.|script|js:|about:|file:|Document\.|vbs:|frame|cookie|on(finish|mouse|Exit|error|click|key|load|focus|Blur))"
			End If
			re.Pattern="<((.[^>]*"&Replacelist&"[^>]*)|"&Replacelist&")>"
			Test=re.Test(t1)
			If Test=False Then
				If IsNull(Ubblists)="" Then Dim Ubblists:Ubblists=",1,"
				re.Pattern=",[13-8],"
				If re.Test(Ubblists) Then
					If InStr(Dvbbs.forum_setting(77),"|")=0 Then 
						Replacelist="(expression|xss:|function|meta|window\.|script|js:|about:|file:|Document\.|vbs:|frame|cookie|on(finish|mouse|Exit=|error|click|key|load|focus|Blur))"
					Else
						Replacelist="("&Dvbbs.forum_setting(77)&"expression|xss:|function|meta|window\.|script|js:|about:|file:|Document\.|vbs:|frame|cookie|on(finish|mouse|Exit|error|click|key|load|focus|Blur))"
					End If
					re.Pattern="(\[(.[^\]]*)\])((.[^\]]*"&Replacelist&"[^\]]*)|"&Replacelist&")(\[\/(.[^\]]*)\])"
					Test=re.Test(t1)
				End If
			End If
			Dv_FilterJS=test
		End If
	End Function

	Public Function Dv_FilterJS2(v)
		If Not Isnull(V) Then
			Dim t,test,Replacelist,t1
			t=v
			t1=v
			re.Pattern="&#36;"
			t1=re.Replace(t1,"$")
			re.Pattern="&#36"
			t1=re.Replace(t1,"$")
			re.Pattern="&#39;"
			t1=re.Replace(t1,"'")
			re.Pattern="&#39"
			t1=re.Replace(t1,"'")
			t1=entity2Str(t1)
			If InStr(Dvbbs.forum_setting(77),"|")=0 Then 
				Replacelist="(expression|xss:|var |function|meta|window\.|script|js:|about:|file:|Document\.|vbs:|frame|cookie|on(finish|mouse|Exit=|error|click|key|load|focus|Blur))"	'|\[|\]
			Else
				Replacelist="("&Dvbbs.forum_setting(77)&"expression|xss:|var |function|meta|window\.|script|js:|about:|file:|Document\.|vbs:|frame|cookie|on(finish|mouse|Exit|error|click|key|load|focus|Blur))"
			End If
			re.Pattern="(\[(.[^\]]*)\])((.[^\]]*"&Replacelist&"[^\]]*)|"&Replacelist&")(\[\/(.[^\]]*)\])"
			Test=re.Test(t1)
			Dv_FilterJS2=test
		End If
	End Function

	Private Function Dv_UbbCode_C(strText,uCodeC)
		Dim s,matches,match,CodeStr,Floor
		Floor=1
		s=strText
		s=strText
		re.Pattern="\["&uCodeC&"\][\s\n]*\[\/"&uCodeC&"\]"
		s=re.replace(s,"")
		re.Pattern="\[\/"&uCodeC&"\]"
		s=re.replace(s,Chr(1)&"/"&uCodeC&"]")
		re.Pattern="\["&uCodeC&"\]([^\x01]*)\x01\/"&uCodeC&"\]"
		Set matches = re.Execute(s)
		re.Global=False
		For Each match In matches
			CodeStr=match.SubMatches(0)
			CodeStr = Replace(CodeStr,"&nbsp;",Chr(32),1,-1,1)
			CodeStr = Replace(CodeStr,"<p>","",1,-1,1)
			CodeStr = Replace(CodeStr,"</p>","&#13;&#10;",1,-1,1)
			CodeStr = Replace(CodeStr,"[br]","&#13;&#10;",1,-1,1)
			CodeStr = Replace(CodeStr,"<br/>","&#13;&#10;",1,-1,1)
			CodeStr = Replace(CodeStr,vbNewLine,"&#13;&#10;",1,-1,1)
			CodeStr = "<form name=""scode"& replyid &"_"& Floor &""" method=""post"" action=""""><table class=""tableborder1"" cellspacing=""1"" cellpadding=""3"" style=""width: 98%;"" align=""center""><tr><th height=""22"">以下是程序代码</th></tr><tr><td class=""tablebody"&(((Floor+1) Mod 2)+1)&""" align=""middle"" width=""98%""><textarea id=""CodeText"" style=""width: 100%;"" rows=""10"">"&CodeStr&"</textarea></td></tr><tr><td class=""tablebody"&(((Floor+1) Mod 2)+1)&""" align=""middle"" width=""98%""><b>说明：</b>上面显示的是代码内容。您可以先检查过代码没问题，或修改之后再运行.</td></tr><tr><td class=""tablebody"&(((Floor+1) Mod 2)+1)&""" align=""middle"" width=""98%""><input type=""button"" name=""run"" value=""运行代码"" onclick=""Dvbbs_ViewCode('"& replyid &"_"& Floor &"');"" />&nbsp;<input type=""button"" name=""copy"" value=""复制代码"" onclick=""Dvbbs_CopyCode('"& replyid &"_"& Floor &"');"" disabled=""disabled"" />&nbsp;<input type=""button"" name=""save"" value=""另存代码"" onclick=""Dvbbs_SaveCode('"& replyid &"_"& Floor &"');"" disabled=""disabled""/></td></tr></table></form>"
			s = re.Replace(s,CodeStr)
			Floor=Floor+1
		Next
		re.Global=true
		Set matches=Nothing
		re.Pattern="\x01\/"&uCodeC&"\]"
		s=re.replace(s,"[/"&uCodeC&"]")
		Dv_UbbCode_C=s
	End Function

		Private Function Dv_Alipay_PayTo(strText)
		If Not Isnull(strText) Then
			Dim s,ss
			Dim match,match2,urlStr,re2
			Dim t(2),temp,check,fee,i,encode8_tmp
			s=strText
			Set re2=new RegExp
			re2.IgnoreCase =true
			re2.Global=False
			t(0)="卖家承担运费"
			t(1)="买家承担运费"
			t(2)="虚拟物品不需邮递"
			s=strText
			re.Pattern="\[\/payto\]"
			s=re.replace(s, chr(1)&"/payto]")
			re.Pattern="\[payto\]([^\x01]+)\x01\/payto\]"
			Set match = re.Execute(s)
			re.Global=False
			For i=0 To match.count-1
				re2.Pattern="\(seller\)([^\n]+?)\(\/seller\)"
				If re2.Test(match.item(i)) Then
					Set match2 = re2.Execute(match.item(i))
					temp=re2.replace(match2.item(0),"$1")
					ss=""
					urlStr="http://pay.dvbbs.net/payto.asp?seller="&temp
					re2.Pattern="\(subject\)([^\n]+?)\(\/subject\)"
					If re2.Test(match.item(i)) Then
						Set match2 = re2.Execute(match.item(i))
						temp=re2.replace(match2.item(0),"$1")
						ss=ss&"<br/><b>商品名称</b>："&temp&"<br/><br/>"
						urlStr = urlStr & "&subject=" & Server.UrlEncode(temp)
						re2.Pattern="\(body\)((.|\n)*?)\(\/body\)"
						If re2.Test(match.item(i)) Then
							Set match2 = re2.Execute(match.item(i))
							temp=re2.replace(match2.item(0),"$1")
							ss=ss&"<b>商品说明</b>："&temp&"<br/><br/>"
							urlStr = urlStr & "&body=" & Server.UrlEncode(Left(Dvbbs.Replacehtml(temp), 200) & "...")
							re2.Pattern="\(price\)([\d\.]+?)\(\/price\)"
							If re2.Test(match.item(i)) Then
								Set match2 = re2.Execute(match.item(i))
								temp=re2.replace(match2.item(0),"$1")
								ss=ss&"<b>商品价格</b>："&temp&" 元<br/><br/>"
								urlStr=urlStr&"&price="&temp
								re2.Pattern="\(transport\)([1-3])\(\/transport\)"
								If re2.Test(match.item(i)) Then
									Set match2 = re2.Execute(match.item(i))
									temp=re2.replace(match2.item(0),"$1")
									check=true
									If int(temp)=2 Then
										re2.Pattern="\(express_fee\)([\d\.]+?)\(\/express_fee\)"
										If re2.Test(match.item(i)) Then
											Set match2 = re2.Execute(match.item(i))
											fee=re2.replace(match2.item(0),"$1")
											ss=ss&"<b>邮递信息</b>："&t(temp-1)&"，快递 "&fee&" 元<br/><br/>"
											urlStr=urlStr&"&transport="&temp&"&express_fee="&fee
										Else
											re2.Pattern="\(ordinary_fee\)([\d\.]+?)\(\/ordinary_fee\)"
											If re2.Test(match.item(i)) Then
												Set match2 = re2.Execute(match.item(i))
												fee=re2.replace(match2.item(0),"$1")
												ss=ss&"<b>邮递信息</b>："&t(temp-1)&"，平邮 "&fee&" 元<br/><br/>"
												urlStr=urlStr&"&transport="&temp&"&ordinary_fee="&fee
											Else
												check=False
											End If
										End If
									Else
										ss=ss&"<b>邮递信息</b>："&t(temp-1)&"<br/><br/>"
										urlStr=urlStr&"&transport="&temp
									End If
									If check=true Then
										check=False
										re2.Pattern="\(ww\)([^\n]+?)\(\/ww\)"
										If re2.Test(match.item(i)) Then
											Set match2 = re2.Execute(match.item(i))
											temp=re2.replace(match2.item(0),"$1")
											encode8_tmp=EncodeUtf8(temp)
											ss=ss&"<b>联系方法</b>：<a target=""_blank"" href=""http://amos1.taobao.com/msg.ww?v=2&amp;uid="&encode8_tmp&"&amp;s=1""><img border=""0"" src=""http://amos1.taobao.com/online.ww?v=2&amp;uid="&encode8_tmp&"&amp;s=1""/></a>"
											check=true
										End If
										re2.Pattern="\(qq\)(\d+?)\(\/qq\)"
										If re2.Test(match.item(i)) Then
											Set match2 = re2.Execute(match.item(i))
											temp=re2.replace(match2.item(0),"$1")
											If check=true Then
												ss=ss&"&nbsp;&nbsp;<a target=""_blank"" href=""http://wpa.qq.com/msgrd?V=1&Uin="&temp&"&Site=Dvbbs.Net&Menu=yes""><img border=""0"" src=""http://wpa.qq.com/pa?p=1:"&temp&":10"" alt=""联系我"" /></a><br/><br/>"
											Else
												ss=ss&"<b>联系方法</b>：<a target=""_blank"" href=""http://wpa.qq.com/msgrd?V=1&Uin="&temp&"&Site=Dvbbs.Net&Menu=yes""><img border=""0"" src=""http://wpa.qq.com/pa?p=1:"&temp&":10"" alt=""联系我"" /></a><br/><br/>"
											End If
										ElseIf check=true Then
											ss=ss&"<br/><br/>"
										End If
										re2.Pattern="\(demo\)([^\n]+?)\(\/demo\)"
										If re2.Test(match.item(i)) Then
											Set match2 = re2.Execute(match.item(i))
											temp=re2.replace(match2.item(0),"$1")
											ss=ss&"<b>演示地址</b>："&temp&"<br/><br/>"
											urlStr=urlStr&"&url="&temp
										End If
										ss=ss&"<a href="""&Server.HtmlEncode(urlStr&"&partner=2088002048522272&type=1&readonly=true")&""" target=""_blank""><img src=""images/alipay/alipay_7.gif"" border=""0"" alt=""通过支付宝交易，买卖都放心，免手续费、安全、快捷！"" /></a>&nbsp;&nbsp;<a href=""http://server.dvbbs.net/dvbbs/alipay/b.htm"" target=""_blank""><font color=""blue"">查看交易帮助，买卖放心</font></a><br/>"
										s=re.replace(s,ss)
									End If
								End If
							End If
						End If
					End If
				End If
			Next
			Set match=Nothing
			Set re2=Nothing
			Set match2=Nothing
			re.Global=true
			re.Pattern="\x01\/payto\]"
			s=re.replace(s,"[/payto]")
			If isxhtml Then
				If xml.loadxml("<div>" & xmlencode(s) &"</div>") Then
					Dv_Alipay_PayTo=s
				Else
					Dv_Alipay_PayTo=strText
				End If
			Else
					Dv_Alipay_PayTo=s
			End If
		End If
	End Function
	Function canusemt(GroupID)
		If Application(Dvbbs.CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& GroupID &"']/@groupsetting") Is Nothing Then GroupID=7
		canusemt = Split(Application(Dvbbs.CacheName &"_groupsetting").documentElement.selectSingleNode("usergroup[@usergroupid='"& GroupID &"']/@groupsetting").text,",")(72)
	End Function
End Class
</script>

<script type="text/javascript" runat="server" language=javascript>
 function EncodeUtf8(s1)
  {
      var s = escape(s1);
      var sa = s.split("%");
      var retV ="";
      if(sa[0] != "")
      {
         retV = sa[0];
      }
      for(var i = 1; i < sa.length; i ++)
      {
           if(sa[i].substring(0,1) == "u")
           {
               retV += Hex2Utf8(Str2Hex(sa[i].substring(1,5))) + sa[i].substring(5,sa[i].length);
               
           }
           else retV += "%" + sa[i];
      }
      
      return retV;
  }
  function Str2Hex(s)
  {
      var c = "";
      var n;
      var ss = "0123456789ABCDEF";
      var digS = "";
      for(var i = 0; i < s.length; i ++)
      {
         c = s.charAt(i);
         n = ss.indexOf(c);
         digS += Dec2Dig(eval(n));
           
      }
      //return value;
      return digS;
  }
  function Dec2Dig(n1)
  {
      var s = "";
      var n2 = 0;
      for(var i = 0; i < 4; i++)
      {
         n2 = Math.pow(2,3 - i);
         if(n1 >= n2)
         {
            s += '1';
            n1 = n1 - n2;
          }
         else
          s += '0';
          
      }
      return s;
      
  }
  function Dig2Dec(s)
  {
      var retV = 0;
      if(s.length == 4)
      {
          for(var i = 0; i < 4; i ++)
          {
              retV += eval(s.charAt(i)) * Math.pow(2, 3 - i);
          }
          return retV;
      }
      return -1;
  } 
  function Hex2Utf8(s)
  {
     var retS = "";
     var tempS = "";
     var ss = "";
     if(s.length == 16)
     {
         tempS = "1110" + s.substring(0, 4);
         tempS += "10" +  s.substring(4, 10); 
         tempS += "10" + s.substring(10,16); 
         var sss = "0123456789ABCDEF";
         for(var i = 0; i < 3; i ++)
         {
            retS += "%";
            ss = tempS.substring(i * 8, (eval(i)+1)*8);
            
            
            
            retS += sss.charAt(Dig2Dec(ss.substring(0,4)));
            retS += sss.charAt(Dig2Dec(ss.substring(4,8)));
         }
         return retS;
     }
     return "";
  } 
</script>