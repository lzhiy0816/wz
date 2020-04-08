<!--#include file="../Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="../inc/md5.asp" -->
<!-- #include file="../inc/chkinput.asp" -->
<%
Dim admin_flag
admin_flag=",2,"
CheckAdmin(admin_flag)
'Alimama管理
Dim Forum_Api,Alimama_Api
Dim Action
Dim Alimama
Dim AdsList

ChkForum_api()
Set Alimama = New Alimama_Apicls
Action = Request("act")
Head()
Top_nav()
Page_main()
Footer()
Set Alimama = Nothing
Set Forum_Api = Nothing
Set Alimama_Api = Nothing
Set AdsList = Nothing

Sub Page_main()
	Select Case Action
		Case "reg"
			Reg()
		Case "login"
			Login()
		Case "addads"
			Addads()
		Case "savereg"
			Savereg()
		Case "saveinfo"
			Saveinfo()
		Case "member_bind"
			member_bind()
		Case "info"
			Info()
		Case "saveads"
			SaveAds()
		Case "adslist"
			MyAdsList()
		Case "restore"
			restore()
		Case "editads"
			Editads()
		Case "saveeditads"
			SaveEditads()
		Case "sso_url"
			SSO_Url()
		Case Else
			If Alimama_Api.getAttribute("memberid")="" Then
				Reg()
			Else
				Info()
			End If
	End Select
End Sub

'测试用
Sub Restore()
	Dvbbs.Execute("update Dv_setup set Forum_apis=''")
	Response.redirect "plus_adsali.asp"
End Sub

Sub ChkForum_api()
	Dim Rs,XmlDoc
	Set Rs = Dvbbs.Execute("Select top 1 Forum_apis From Dv_Setup")
	XmlDoc = Rs(0)
	Rs.Close
	Set Rs = Nothing
	If IsNull(XmlDoc) or XmlDoc = "" Then
		Creat_Forum_Api()
	Else
		Set Forum_Api = Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		Forum_Api.LoadXml(XmlDoc)
		Set Alimama_Api = Forum_Api.documentElement.selectSingleNode("alimama")
	End If
End Sub

Sub Creat_Forum_Api()
	Set Forum_Api = Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	Forum_Api.LoadXml("<forum_api/>")
	Set Alimama_Api = Forum_Api.documentElement.appendChild(Forum_Api.createNode(1,"alimama",""))
	Alimama_Api.setAttribute "memberid",""
	Alimama_Api.setAttribute "email",""
	Alimama_Api.setAttribute "password",""
	Alimama_Api.setAttribute "nickname",""
	Alimama_Api.setAttribute "webname",""
	Update_Forum_Api()
End Sub

Sub Update_Forum_Api()
	Dvbbs.Execute("update Dv_setup set Forum_apis='"&Dvbbs.Checkstr(Forum_Api.xml)&"'")
	'Forum_Api.save Server.mappath("../Forum_Api.xml")
End Sub

Sub Top_nav()
%>
<table cellpadding="3" cellspacing="1" border="0" width="100%" align="center">
<tr><td height="25" class="td_title">站点营销推广说明</td></tr>
<%If Alimama_Api.getAttribute("memberid")<>"" Then%>
<tr><td class="td2">
您已成功成为阿里妈妈的会员，帐号是：<font class="font2"><%=Alimama_Api.getAttribute("email")%></font>
</td></tr>
<%End If%>
<tr><td>
<ul>
<li>在这里，您可以直接把论坛的广告位发布到阿里妈妈广告营销平台；</li>
<li>发布广告位的流程：1、注册成为阿里妈妈会员 --> 2、登记本论坛资料信息 --> 3、发布广告位。</li>
</ul>
</td></tr>

</table>
<table cellpadding="3" cellspacing="1" border="0" width="100%" align="center">
<tr><td class="td_title">
<%If Alimama_Api.getAttribute("memberid")="" Then%>
<a href="?act=reg" title="开通阿里妈妈用户帐号">申请帐号</a>
| <%End If%>
<%If Alimama_Api.getAttribute("webname")="" Then%>
<a href="?act=info">登记网站信息</a>
<%Else%>
<a href="?act=sso_url&react=0">查看网站信息</a>
<%End If%>
<!-- | <a href="?act=login">站点登陆</a>  -->
| <a href="?act=addads">发布广告位</a> 
| <a href="?act=adslist">我的广告位</a>
<%If Alimama_Api.getAttribute("memberid")<>"" Then%>
| <a href="?act=sso_url&react=1" target="_blank">交易管理</a>
| <a href="?act=sso_url&react=2" target="_blank">预期收入</a>
| <a href="?act=sso_url&react=3" target="_blank">绑定支付宝</a>
| <a href="?act=sso_url&react=4" target="_blank">提现</a>
| <a href="?act=sso_url&react=5" target="_blank">广告位效果</a>
<%End If%>
</td></tr>
</table>
<%
End Sub

Sub SSO_Url()
	If Alimama_Api.getAttribute("memberid")="" Then
		Errmsg=ErrMsg + "<BR/>请申请了阿里妈妈的帐号后，再进行发布广告位信息！"
		Dvbbs_error()
		Exit Sub
	End If
	If Alimama_Api.getAttribute("siteid")="" Then
		Errmsg=ErrMsg + "<BR/>您的网站信息还未登记，请登记后再进行发布广告位！"
		Dvbbs_error()
		Exit Sub
	End If
	Dim react
	Dim SsoUrl
	Dim Str
	react = Lcase(Request.QueryString("react"))
	Str = Alimama.ApiUrl & "?" & "act=1&signAccount=" & Alimama_Api.getAttribute("email")&"&tourl="&react
	'Response.Write Str
	Response.Redirect Str
End Sub


'显示Loading状态提示
Sub LoadingDiv()
%>
	<div id="loadingDiv">
	<br/>
	<table border="0" cellspacing="1" cellpadding="3"  align="center" width="90%">
	<tr><td class="td2" style="text-align:center;"><img src="../images/loading.gif" border="0" alt=""/> 数据提交外理中......</td></tr>
	</table>
	<br/>
	</div>
<%
Response.Flush
End Sub

'关闭Loading状态提示
Sub CloseLoadingDiv()
%>
<script language="JavaScript">
<!--
if (document.getElementById('loadingDiv')){
	document.getElementById('loadingDiv').style.display = "none";
}
//-->
</script>
<%
End Sub

Sub Info()
	If Alimama_Api.getAttribute("memberid")="" Then
		Errmsg=ErrMsg + "<BR/>请申请了阿里妈妈的帐号后，再进行发布网站信息！"
		Dvbbs_error()
		Exit Sub
	End If
	Dim ali_webname,ali_weburl,ali_webdesc,ali_sex,ali_age,occupationSelect,income,ali_weblike
	
	If Alimama_Api.getAttribute("webname")="" Then
		ali_webname = Server.HtmlEncode(Dvbbs.Forum_info(0))
		ali_weburl = Dvbbs.Forum_info(1)
		ali_webdesc = Server.HtmlEncode(Dvbbs.Forum_info(10))
	Else
		ali_webname = Alimama_Api.getAttribute("webname")
		ali_weburl = Alimama_Api.getAttribute("weburl")
		ali_webdesc = Alimama_Api.getAttribute("webdescription")
	End If

%>
<br/>
<table border="0" cellspacing="1" cellpadding="3"  align="center" width="100%">
<tr>
<th colspan="2" style="text-align:center;">
网站资料信息
</th></tr>
<tr><td class="td2 title" colspan="2">设置您的网站资料</td></tr>
<form method="post" action="?act=saveinfo" name="webinfo">
<tr>
<td align="right" width="20%">网站名称：</td>
<td width="80%"><input type="text" name="ali_webname" size="20" value="<%=ali_webname%>"/>  <font class="font1">网站名称请不要超过16个汉字或32个字符</font></td>
</tr>
<tr>
<td align="right">网站地址：</td>
<td><input type="text" name="ali_weburl" size="40" value="<%=ali_weburl%>"/>  <font class="font1">例：http://www.dvbbs.net 网站地址请不要超过255个字符</font></td>
</tr>
<tr>
<td align="right">网站描述：</td>
<td><textarea name="ali_webdesc" style="width:96%;height:80px;"><%=ali_webdesc%></textarea>  <br/><font class="font1">例：我的网站日独立IP××，PV××，PR：4，Alexa排名前1万。是国内专业的英语学习网站，在同行内排名前三。具有稳定的注册用户群，并且每日新增用户约××人。在抓虾上订阅我的网站的用户已经有3000人。</font></td>
</tr>
<tr><td class="td2 title" colspan="2">修改网站浏览人群</td></tr>

<tr>
<td align="right">性别：<font class="font2">*</font> </td>
<td><input type="radio" name="ali_sex" value="1" class="radio"/>多数是男的
<input type="radio" name="ali_sex" value="2" class="radio"/>男女基本一致
<input type="radio" name="ali_sex" value="3" class="radio"/>多数是女的
</td>
</tr>
<tr>
<td align="right">年龄：<font class="font2">*</font> </td>
<td>
<input type="radio" name="ali_age" value="1" class="radio"/>18岁以下
<input type="radio" name="ali_age" value="2" class="radio"/>18~25岁
<input type="radio" name="ali_age" value="3" class="radio"/>25~45岁
<input type="radio" name="ali_age" value="4" class="radio"/>45岁以上
</td>
</tr>
<tr>
<td align="right">职业：<font class="font2">*</font> </td>
<td>
<input type="checkbox" name="occupationSelect" value="1" class="checkbox"/>学生
<input type="checkbox" name="occupationSelect" value="2" class="checkbox"/>职员
<input type="checkbox" name="occupationSelect" value="3" class="checkbox"/>高级管理人员
<input type="checkbox" name="occupationSelect" value="4" class="checkbox"/>个体工商户
<input type="checkbox" name="occupationSelect" value="5" class="checkbox"/>企业主
<input type="checkbox" name="occupationSelect" value="6" class="checkbox"/>自由职业者
<input type="checkbox" name="occupationSelect" value="7" class="checkbox"/>待业者
<input type="checkbox" name="selall" value="1" class="checkbox" onclick="boxCheckAll(document.webinfo.occupationSelect,this);"/>所有
</td>
</tr>

<tr>
<td align="right">月收入：<font class="font2">*</font> </td>
<td>
<input type="radio" name="income" value="1" class="radio"/>1000以下
<input type="radio" name="income" value="2" class="radio"/>1000~3000
<input type="radio" name="income" value="3" class="radio"/>3000~5000
<input type="radio" name="income" value="4" class="radio"/>5000~8000
<input type="radio" name="income" value="5" class="radio"/>8000~10000
<input type="radio" name="income" value="6" class="radio"/>10000以上
</td>
</tr>
<tr>
<td align="right">兴趣爱好：</td>
<td><input type="text" name="ali_weblike" size="40" value="<%=ali_weblike%>"/>  <font class="font1">兴趣爱好请不要超过30个汉字或者60个字符</font></td>
</tr>
<%If Alimama_Api.getAttribute("webname")="" Then%>
<tr><td class="td2" colspan="2" align="center"><input type="submit" name="submit" value="确认提交"/></td></tr>
<%End If%>
</form>
</table>
<script language="JavaScript">
<!--
chkradio(document.webinfo.ali_sex,'<%=Alimama_Api.getAttribute("websex")%>');
chkradio(document.webinfo.ali_age,'<%=Alimama_Api.getAttribute("webage")%>');
chkradio(document.webinfo.income,'<%=Alimama_Api.getAttribute("webincome")%>');
chkcheckbox(document.webinfo.occupationSelect,'<%=Alimama_Api.getAttribute("webjob")%>');
//-->
</script>
<%

End Sub

'保存新增网站信息
Sub Saveinfo()
	If Alimama_Api.getAttribute("memberid")="" Then
		Errmsg=ErrMsg + "<BR/>请申请了阿里妈妈的帐号后，再进行发布网站信息！"
		Dvbbs_error()
		Exit Sub
	End If
	Dim ali_webname,ali_weburl,ali_webdesc,ali_sex,ali_age,occupationSelect,income,ali_weblike
	ali_webname = Trim(Request.Form("ali_webname"))
	ali_weburl = Trim(Request.Form("ali_weburl"))
	ali_webdesc = Trim(Request.Form("ali_webdesc"))
	ali_sex = Trim(Request.Form("ali_sex"))
	ali_age = Trim(Request.Form("ali_age"))
	occupationSelect = Replace(Trim(Request.Form("occupationSelect"))," ","")
	income = Trim(Request.Form("income"))
	ali_weblike = Trim(Request.Form("ali_weblike"))

	If ali_webname = "" or Len(ali_webname)>32 Then
		Errmsg=ErrMsg + "<BR/>网站的名称请不要超过16个汉字或32个字符！"
		Dvbbs_error()
		Exit Sub
	End If
	If Len(ali_weburl)<3 or Len(ali_weburl)>255 Then
		Errmsg=ErrMsg + "<BR/>网站的地址超出了字符限制！"
		Dvbbs_error()
		Exit Sub
	End If
	If ali_webdesc = "" Then
		Errmsg=ErrMsg + "<BR/>网站的描述不能为空！"
		Dvbbs_error()
		Exit Sub
	End If
	If ali_sex = "" or Not Isnumeric(ali_sex) Then
		Errmsg=ErrMsg + "<BR/>请选取网站浏览人群性别！"
		Dvbbs_error()
		Exit Sub
	End If
	If ali_age = "" or Not Isnumeric(ali_age) Then
		Errmsg=ErrMsg + "<BR/>请选取网站浏览人群年龄！"
		Dvbbs_error()
		Exit Sub
	End If
	If occupationSelect = "" or Not IsNumeric(Replace(occupationSelect,",","")) Then
		Errmsg=ErrMsg + "<BR/>请选取网站浏览人群职业！"
		Dvbbs_error()
		Exit Sub
	End If
	If income = "" or Not Isnumeric(income) Then
		Errmsg=ErrMsg + "<BR/>请选取网站浏览人群收入！"
		Dvbbs_error()
		Exit Sub
	End If
	If Len(ali_weblike)>60 Then
		Errmsg=ErrMsg + "<BR/>网站浏览人群的兴趣爱好不能超过60字符！"
		Dvbbs_error()
		Exit Sub
	End If


	'-------------------------------------------------------------------------
	'创建API接口参数
	'-------------------------------------------------------------------------
	Dim XmlDom
	Alimama.QueryStr "service","site_add"
	Alimama.QueryStr "memberid",Alimama_Api.getAttribute("memberid")
	Alimama.QueryStr "name",ali_webname
	Alimama.QueryStr "url",ali_weburl
	Alimama.QueryStr "type",1002		'网站类型，默认为1002:社区论坛
	Alimama.QueryStr "description",ali_webdesc
	Alimama.QueryStr "sex",ali_sex
	Alimama.QueryStr "age",ali_age
	Alimama.QueryStr "income",income
	Alimama.QueryStr "job",occupationSelect
	Alimama.QueryStr "hobby",ali_weblike

	'Response.Write Alimama.Posturl
	'显示Loading状态提示
	LoadingDiv()
	Set XmlDom = Alimama.Post
	If XmlDom is Nothing Then
		CloseLoadingDiv()
		Errmsg=ErrMsg + "<BR/>意外错误，获取提交数据中止，请稍候再试！"
		Dvbbs_error()
		Exit Sub
	End If
	CloseLoadingDiv()
	'Is_success,Error_code,Data_List
	If Not (Alimama.Is_success is Nothing) Then
		If Alimama.Is_success.text<>"true" Then
			'If Alimama.Error_code.text<"1999" Then
			'	Errmsg=ErrMsg + "<BR/>(Errcode:"&Alimama.Error_code.text&")系统错误，获取提交数据中止，请稍候再试！"
			'	Dvbbs_error()
			'	Exit Sub
			'End If
			Errmsg=ErrMsg + "<BR/>"&Alimama.Error_desc
			Dvbbs_error()
			Exit Sub
		Else
			If Not (Alimama.Data_List.selectSingleNode("site/siteid") is Nothing) Then
				Alimama_Api.setAttribute "siteid",Alimama.Data_List.selectSingleNode("site/siteid").text
				Alimama_Api.setAttribute "webname",ali_webname
				Alimama_Api.setAttribute "weburl",ali_weburl
				Alimama_Api.setAttribute "webdescription",ali_webdesc
				Alimama_Api.setAttribute "websex",ali_sex
				Alimama_Api.setAttribute "weburl",ali_weburl
				Alimama_Api.setAttribute "webage",ali_age
				Alimama_Api.setAttribute "webincome",income
				Alimama_Api.setAttribute "webjob",occupationSelect
				Alimama_Api.setAttribute "webhobby",ali_weblike
				Update_Forum_Api()
				Dv_suc("网站信息登记成功！")
			Else
				Errmsg=ErrMsg + "<BR/>site/siteid不存在，系统未知错误，请稍候再试！"
				Dvbbs_error()
				Exit Sub
			End If
		End If
	Else
		Errmsg=ErrMsg + "<BR/>返回数据错误，操作已中止，请稍候再试！"
		Dvbbs_error()
		Exit Sub
	End If
End Sub

'申请帐号
Sub Reg()
%>
<br/>
<table border="0" cellspacing="1" cellpadding="3"  align="center" width="100%">
<tr><th colspan="2" style="text-align:center;">阿里妈妈会员注册</th></tr>
<tr><td colspan="2">
<font class="font2">*</font>如果您已经在阿里妈妈申请了帐号，请点击<a href="?act=login" class="font2">帐户绑定</a>
</td></tr>
<tr><td class="td2 title" colspan="2">设置您的帐户信息</td></tr>
<form method="post" action="?act=savereg">
<tr>
<td align="right">昵称：</td>
<td><input type="text" name="ali_name" size="20" value="<%=Dvbbs.MemberName%>"/>  <font class="font1">这个昵称将用于阿里妈妈的论坛，请谨慎填写</font></td>
</tr>
<tr>
<td  align="right">EMAIL地址：</td>
<td>
<%
Dim userinfo,UserEmail
If TypeName(Dvbbs.UserSession)="DOMDocument" Then
	Set UserInfo=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo")
	UserEmail = UserInfo.getAttribute("useremail")
Else
	useremail=""
End If
%>
<input type="text" name="ali_email" size="20" value="<%=UserEmail%>"/>
&nbsp;&nbsp;&nbsp;<font class="font1">您将使用此邮件地址登录，请确保它是真实有效的</font>
</td>
</tr>
<tr>
<td align="right">密码：</td>
<td><input type="text" name="ali_password1" size="20" maxlength="16"/>  <font class="font1">密码由6-16个字符组成，请使用英文字母加数字的组合密码</font></td>
</tr>
<tr>
<td align="right">再输入一遍密码：</td>
<td><input type="text" name="ali_password2" size="20" maxlength="16"/>  <font class="font1">请再输入一遍您上面输入的密码</font></td>
</tr>
<tr><td class="td2 title" colspan="2">联系方式  (<font class="font2">*以下三种联系方式，手机号码为必填，其他为可选项</a>)</td></tr>

<tr>
<td  align="right">手机号码：</td>
<td><input type="text" name="ali_mobile" size="20"/>  <font class="font2">*</font><font class="font1">手机号码是阿里妈妈就付款、投放等重要问题和您保持联系的首选，请认真填写</font></td>
</tr>
<tr>
<td  align="right">联系电话：</td>
<td><input type="text" name="ali_phone" size="20"/>  <font class="font1">请输入您的联系电话</font></td>
</tr>
<tr>
<td  align="right">淘宝旺旺登录名：</td>
<td><input type="text" name="ali_wangwang" size="20"/>  <font class="font1">请输入您的淘宝旺旺名。为保证和广告买家、卖家的通畅沟通，请多多使用阿里旺旺（<a href="http://www.taobao.com/wangwang/index.php" target="_blank">淘宝版</a>）</font></td>
</tr>
<tr><td class="td2" colspan="2" align="center"><input type="submit" name="submit" value="确认提交"/></td></tr>
</form>
</table>
<%
End Sub

Sub	Savereg()
	If Alimama_Api.getAttribute("memberid")<>"" Then
		Errmsg=ErrMsg + "<BR/>你已经申请过帐号了！"
		Dvbbs_error()
		Exit Sub
	End If
	Dim Ali_Email,Ali_pass,Ali_name,Ali_mobile
	Dim Ali_phone,Ali_wangwang

	'-------------------------------------------------------------------------
	'获取表单数据
	'-------------------------------------------------------------------------
	Ali_name = Trim(Request.Form("ali_name"))
	Ali_Email = Trim(Request.Form("ali_email"))
	Ali_pass = Trim(Request.Form("ali_password1"))
	Ali_mobile = Trim(Request.Form("ali_mobile"))
	Ali_phone = Trim(Request.Form("ali_phone"))
	Ali_wangwang = Trim(Request.Form("ali_wangwang"))

	'-------------------------------------------------------------------------
	'验证提交数据开始
	'-------------------------------------------------------------------------
	If Ali_name=""or Len(Ali_name)<1 or Len(Ali_name)>32 Then
		Errmsg=ErrMsg + "<BR/>昵称不能为空或超过32个字符！"
		Dvbbs_error()
		Exit Sub
	End If
	If Len(Ali_Email)>128 or Ali_Email="" Then
		Errmsg=ErrMsg + "<BR/>EMAIL地址不能为空或超过128个字符，请重新填写！"
		Dvbbs_error()
		Exit Sub
	End If
	If IsValidEmail(Ali_Email) = False Then
		Errmsg=ErrMsg + "<BR/>EMAIL地址不符合要求，请重新填写！"
		Dvbbs_error()
		Exit Sub
	End If

	If Ali_pass="" or Len(Ali_pass)<6 or Len(Ali_pass)>16 Then
		Errmsg=ErrMsg + "<BR/>密码必须由6-16个字符组成，请重新填写！"
		Dvbbs_error()
		Exit Sub
	End If
	If ChkIsEnAndNum(Ali_pass) = False Then
		Errmsg=ErrMsg + "<BR/>请使用英文字母加数字的组合密码！"
		Dvbbs_error()
		Exit Sub
	End If
	If Ali_pass<>Trim(Request.Form("ali_password1")) Then
		Errmsg=ErrMsg + "<BR/>两次输入的密码不符，请重新填写！"
		Dvbbs_error()
		Exit Sub
	End If
	If Ali_mobile="" Or Len(Ali_mobile)>50 Then
		Errmsg=ErrMsg + "<BR/>手机不能为空或超过50个字符！"
		Dvbbs_error()
		Exit Sub
	End If
	If ChkIsPhone(Ali_mobile) = False Then
		Errmsg=ErrMsg + "<BR/>手机号码不符合要求，请检查一下！"
		Dvbbs_error()
		Exit Sub
	End If
	If Ali_phone<>"" and Len(Ali_phone)>50 Then
		Errmsg=ErrMsg + "<BR/>联系电话不能超过50个字符，请检查一下！"
		Dvbbs_error()
		Exit Sub
	End If
	
	If Ali_wangwang<>"" and Len(Ali_wangwang)>100 Then
		Errmsg=ErrMsg + "<BR/>淘宝旺旺登录名不能超过100个字符，请检查一下！"
		Dvbbs_error()
		Exit Sub
	End If
	'验证提交数据结束
	'-------------------------------------------------------------------------
	Ali_Email = Lcase(Ali_Email)
	Ali_pass = Md5(Ali_pass,32)

	'-------------------------------------------------------------------------
	'创建API接口参数
	'-------------------------------------------------------------------------
	Dim XmlDom
	Alimama.QueryStr "service","member_register"
	Alimama.QueryStr "logname",Ali_Email
	Alimama.QueryStr "logid",Alimama.FormatNickName(Ali_Email)
	Alimama.QueryStr "passwd",Ali_pass
	Alimama.QueryStr "nickname",Alimama.FormatNickName(Ali_name)
	Alimama.QueryStr "mobile",Ali_mobile
	Alimama.QueryStr "telephone",ali_phone
	Alimama.QueryStr "wangwang",ali_wangwang
	'Response.Write Alimama.Posturl
	'显示Loading状态提示
	LoadingDiv()
	Set XmlDom = Alimama.Post
	If XmlDom is Nothing Then
		CloseLoadingDiv()
		Errmsg=ErrMsg + "<BR/>意外错误，获取提交数据中止，请稍候再试！"
		Dvbbs_error()
		Exit Sub
	End If
	CloseLoadingDiv()
	'Is_success,Error_code,Data_List
	If Not (Alimama.Is_success is Nothing) Then
		If Alimama.Is_success.text<>"true" Then
			'If Alimama.Error_code.text<"1999" Then
			'	Errmsg=ErrMsg + "<BR/>(Errcode:"&Alimama.Error_code.text&")系统错误，获取提交数据中止，请稍候再试！"
			'	Dvbbs_error()
			'	Exit Sub
			'End If
			Errmsg=ErrMsg + "<BR/>"&Alimama.Error_desc
			Dvbbs_error()
			Exit Sub
		Else
			Alimama_Api.setAttribute "memberid",Alimama.Data_List.selectSingleNode("member/memberid").text
			Alimama_Api.setAttribute "email",Ali_Email
			Alimama_Api.setAttribute "password",Ali_pass
			Alimama_Api.setAttribute "nickname",Ali_name
			Update_Forum_Api()
			Dv_suc("申请帐号成功！")
		End If
	Else
		Errmsg=ErrMsg + "<BR/>返回数据错误，操作已中止，请稍候再试！"
		Dvbbs_error()
		Exit Sub
	End If
	
End Sub

'绑定用户
Sub Login()
%>
<br/>
<table border="0" cellspacing="1" cellpadding="3"  align="center" width="100%">
<form method="post" action="?act=member_bind">
<tr><th colspan="2" style="text-align:center;">绑定阿里妈妈会员帐号</th></tr>
<tr><td class="td2 title" colspan="2">填写您在阿里妈妈已开通的帐户信息</td></tr>
<tr>
<td  align="right">EMAIL地址：</td>
<td><input type="text" name="ali_email" size="20"/></td>
</tr>
<tr>
<td  align="right">密码：</td>
<td><input type="text" name="ali_password" size="20"/></td>
</tr>
<tr><td class="td2" colspan="2" align="center"><input type="submit" name="submit" value="现在绑定"/></td></tr>
</form>
</table>
<%
End Sub

'绑定用户
Sub member_bind()
	Dim Ali_Email,Ali_pass
	Ali_Email = Trim(Request.Form("ali_email"))
	Ali_pass = Trim(Request.Form("ali_password"))
	Ali_Email = Lcase(Ali_Email)
	Ali_pass = Md5(Ali_pass,32)

	'-------------------------------------------------------------------------
	'创建API接口参数
	'-------------------------------------------------------------------------
	Dim XmlDom
	Alimama.QueryStr "service","member_bind"
	Alimama.QueryStr "logname",Ali_Email
	Alimama.QueryStr "passwd",Ali_pass
	Alimama.QueryStr "logid",Alimama.FormatNickName(Ali_Email)
	
	'显示Loading状态提示
	LoadingDiv()
	Set XmlDom = Alimama.Post
	If XmlDom is Nothing Then
		CloseLoadingDiv()
		Errmsg=ErrMsg + "<BR/>意外错误，获取提交数据中止，请稍候再试！"
		Dvbbs_error()
		Exit Sub
	End If
	CloseLoadingDiv()
	'Is_success,Error_code,Data_List
	If Not (Alimama.Is_success is Nothing) Then
		If Alimama.Is_success.text<>"true" Then
			'If Alimama.Error_code.text<"1999" Then
			'	Errmsg=ErrMsg + "<BR/>(Errcode:"&Alimama.Error_code.text&")系统错误，获取提交数据中止，请稍候再试！"
			'	Dvbbs_error()
			'	Exit Sub
			'End If
			Errmsg=ErrMsg + "<BR/>"&Alimama.Error_desc
			Dvbbs_error()
			Exit Sub
		Else
			Alimama_Api.setAttribute "memberid",Alimama.Data_List.selectSingleNode("member/memberid").text
			Alimama_Api.setAttribute "email",Ali_Email
			Alimama_Api.setAttribute "password",Ali_pass
			Alimama_Api.setAttribute "nickname",""
			Update_Forum_Api()
			Dv_suc("帐号绑定成功！")
		End If
	Else
		Errmsg=ErrMsg + "<BR/>返回数据错误，操作已中止，请稍候再试！"
		Dvbbs_error()
		Exit Sub
	End If

End Sub

'添加/编辑广告位
Sub Addads()
	Dim Adzoneid,ActStats
	Adzoneid = Request.QueryString("adzoneid")
	If Adzoneid<>"" Then
		Set AdsList = Alimama_Api.selectSingleNode("ads[@adzoneid="&Adzoneid&"]")
		If AdsList is Nothing  Then
			Errmsg=ErrMsg + "<BR/>您需要编辑的广告位不存在！"
			Dvbbs_error()
			Exit Sub
		End If
		ActStats = "编辑"
	Else
		ActStats = "发布"
		Set AdsList = Alimama_Api.appendChild(Forum_Api.createNode(1,"ads",""))
	End If
	If Alimama_Api.getAttribute("memberid")="" Then
		Errmsg=ErrMsg + "<BR/>请申请了阿里妈妈的帐号后，再进行发布广告位信息！"
		Dvbbs_error()
		Exit Sub
	End If
	If Alimama_Api.getAttribute("siteid")="" Then
		Errmsg=ErrMsg + "<BR/>您的网站信息还未登记，请登记后再进行发布广告位！"
		Dvbbs_error()
		Exit Sub
	End If
%>
<style>
#adsize_label span{color:red;}
</style>
<script language="JavaScript">
function chkNum(oIpt){
	var v = oIpt.value;
	var reNum = /[^0-9\.]/;
	if(reNum.test(v)){
	av = v.split('');
	for(var i=0;i<av.length;i++){
	if(reNum.test(av[i])){
	av[i]='';
	}
	}
	oIpt.value = av.join('');
	}
	return true;
}

function _setdemo(v){
	var obj = document.getElementById('ads_setdemo');
	if (obj){
		obj.src = "skins/images/ads_set"+v+".gif";
	}
}
</script>
<br/>
<table border="0" cellspacing="1" cellpadding="3"  align="center" width="100%">
<form method="post" action="?act=saveads" name="adsform">
<tr><th colspan="2" style="text-align:center;"><%=ActStats%>广告位信息(<font class="font2">以下为必填项</font>)</th></tr>
<tr><td class="td2 title" colspan="2">广告位信息</td></tr>
<tr>
<td width="15%" align="right">广告位名称：</td>
<td width="85%">
<input type="text" name="zonename" size="20" value="<%=AdsList.getAttribute("name")%>"/>(例如：xxx网站顶部广告位) 
<br/>
<font class="font1">广告位名称将显示给买家，好的名称有助于买家快速理解广告位价值，请认真填写，并且长度不要超过16个汉字</font>
</td>
</tr>
<tr>
<td  align="right">时长计费设置：</td>
<td>
<input type="text" name="weekprice" size="5" onKeyup="return chkNum(this);" value="<%=AdsList.getAttribute("weekprice")%>"/> 元/周
</td>
</tr>
<tr>
<td  align="right">是否需要审核：</td>
<td>
<input type="radio" name="needAuditingSelect" class="radio" value="1"/>需要  <input type="radio" name="needAuditingSelect" checked="true" class="radio" value="0"/>不需要
<img style="cursor:help;" alt="当您选择需要审核后，有交易付款时，系统会发送邮件通知您审核买家的广告牌。如果您在交易创建后48小时内未审核。系统会自动为您审核通过。" id="sh_tips" src="skins/images/help.gif" align="absmiddle" />
</td>
</tr>
<tr>
<td  align="right">广告形式：</td>
<td>
<input type="radio" name="format" class="radio" value="1"/>文字广告  <input type="radio" name="format" checked="true" class="radio"  value="2"/>图片广告
</td>
</tr>
<tr>
<td  align="right">选择广告位尺寸：</td>
<td>
	<label id="adsize_label"> 
	<input type="radio" name="adsize" value="11" class="radio"/>横幅广告<span>760x90</span>&nbsp;
	<input type="radio" name="adsize" value="12" class="radio"/>横幅广告<span>468x60</span>&nbsp;
	<input type="radio" name="adsize" value="13" class="radio"/>横幅广告<span>250x60</span>&nbsp;
	<input type="radio" name="adsize" value="21" class="radio"/>垂直广告<span>120x600</span>
	<br/>
	<input type="radio" name="adsize" value="22" class="radio"/>垂直广告<span>120x240</span>
	<input type="radio" name="adsize" value="31" class="radio"/>巨幅广告<span>180x250</span>
	<input type="radio" name="adsize" value="32" class="radio"/>巨幅广告<span>250x300</span>
	<input type="radio" name="adsize" value="33" class="radio"/>巨幅广告<span>360x190</span>
	</label>
</td>
</tr>

<tr>
<td  align="right">广告显示位置：</td>
<td>
<div style="float:left;width:40%;">
<input type="checkbox" name="getskinid" value="1" class="checkbox"/>论坛默认广告<font class="font1">：（除具体版面内容以外的页面）</font>
<br/>
选择显示的版面<font class="font1">：（请按 CTRL 键进行多选）</font>
<select name="getboard" style="width:98%;" size="12" multiple="true">
<%
	Dim node,BoardNode,ii
	Set BoardNode=Application(Dvbbs.CacheName&"_boardlist").cloneNode(True)
	For each node in BoardNode.documentElement.getElementsByTagName("board")
		Response.Write "<option value="""&node.getAttribute("boardid")&""">"
		Select Case node.getAttribute("depth")
			Case 0
				Response.Write "╋"
			Case 1
				Response.Write "&nbsp;&nbsp;├"
		End Select
		If node.getAttribute("depth")>1 Then
			For ii=2 To node.getAttribute("depth")
				Response.Write "&nbsp;&nbsp;│"
			Next
			Response.Write "&nbsp;&nbsp;├"
		End If
		Response.Write node.getAttribute("boardtype")
		Response.Write "</option>"&vbNewline
	Next
%>
</select>
<br/>
</div>
<div style="float:left;width:20%;">
选择在页面中显示的位置
<br/>
页面顶部广告位：<input type="radio" name="adsset" value="1" class="radio" onclick="_setdemo(this.value);"/><br/>
页面底部广告位：<input type="radio" name="adsset" value="2" class="radio" onclick="_setdemo(this.value);"/><br/>
<!-- 页面文字广告位：<input type="radio" name="adsset" value="6" class="radio" onclick="_setdemo(this.value);"/><br/>
帖间随机文字广告位：<input type="radio" name="adsset" value="5" class="radio" onclick="_setdemo(this.value);"/><br/> -->
帖子楼主顶部广告位：<input type="radio" name="adsset" value="7" class="radio" onclick="_setdemo(this.value);"/><br/>
帖子楼主左边广告位：<input type="radio" name="adsset" value="8" class="radio" onclick="_setdemo(this.value);"/><br/>
帖子楼主右边广告位：<input type="radio" name="adsset" value="9" class="radio" onclick="_setdemo(this.value);"/><br/>
帖子楼主底部广告位：<input type="radio" name="adsset" value="10" class="radio" onclick="_setdemo(this.value);"/><br/>
</div>
<div style="float:left;width:36%;height:220px;">
<img id="ads_setdemo" src="skins/images/ads_set.gif" border="0" alt="广告位置预览"/>
</div>
</td>
</tr>

<tr>
<td width="15%" align="right">广告位关键字：</td>
<td width="85%">
<input type="text" name="keywords" size="20" value="<%=AdsList.getAttribute("keywords")%>"/>
</td>
</tr>

<tr>
<td  align="right">广告位描述：</td>
<td>
<textarea name="zonedesc" style="width:96%;height:80px;"><%=AdsList.getAttribute("zonedesc")%></textarea>
<br/><font class="font1">例如：广告位的位置，大小，以及该广告位可以在站内所有读帖页展现，访问人群是什么，适合投哪类广告！</font>
</td>
</tr>
<tr><td class="td2" colspan="2" align="center">
<input type="hidden" name="adzoneid" value="<%=AdsList.getAttribute("adzoneid")%>"/>
<input type="submit" name="submit" value="<%=ActStats%>我的广告位"/>
</td></tr>
</form>
</table>
<%If Adzoneid<>"" Then%>
<script language="JavaScript">
chkcheckbox(document.adsform.getskinid,'<%=AdsList.getAttribute("homepage")%>');
chkradio(document.adsform.needAuditingSelect,'<%=AdsList.getAttribute("needAuditingSelect")%>');
chkradio(document.adsform.format,'<%=AdsList.getAttribute("format")%>');
chkradio(document.adsform.adsize,'<%=AdsList.getAttribute("adsize")%>');
ChkSelected(document.adsform.getboard,'<%=AdsList.getAttribute("getboard")%>');
chkradio(document.adsform.adsset,'<%=AdsList.getAttribute("adsset")%>');
_setdemo('<%=AdsList.getAttribute("adsset")%>');
</script>
<%
End If
End Sub

'编辑广告位
Sub Editads()
	Dim Adzoneid
	Adzoneid = Request.QueryString("adzoneid")
	If Adzoneid<>"" Then
		Set AdsList = Alimama_Api.selectSingleNode("ads[@adzoneid="&Adzoneid&"]")
		If AdsList is Nothing  Then
			Errmsg=ErrMsg + "<BR/>您需要编辑的广告位不存在！"
			Dvbbs_error()
			Exit Sub
		End If
	End If
%>
<br/>
<table border="0" cellspacing="1" cellpadding="3"  align="center" width="100%">
<form method="post" action="?act=saveeditads" name="adsform">
<tr><th colspan="2" style="text-align:center;">广告位设置</th></tr>
<tr><td class="td2 title" colspan="2">广告位信息</td></tr>
<tr>
<td  align="right">广告位名称：</td>
<td>
<input type="text" name="zonename" size="20" value="<%=AdsList.getAttribute("name")%>" disabled/> 
</td>
</tr>
<tr>
<td  align="right">广告显示位置：</td>
<td>
<div style="float:left;width:40%;">
<input type="checkbox" name="getskinid" value="1" class="checkbox"/>论坛默认广告<font class="font1">：（除具体版面内容以外的页面）</font>
<br/>
选择显示的版面<font class="font1">：（请按 CTRL 键进行多选）</font>
<select name="getboard" size="12" style="width:98%;" multiple="true">
<%
	Dim node,BoardNode,ii
	Set BoardNode=Application(Dvbbs.CacheName&"_boardlist").cloneNode(True)
	For each node in BoardNode.documentElement.getElementsByTagName("board")
		Response.Write "<option value="""&node.getAttribute("boardid")&""">"
		Select Case node.getAttribute("depth")
			Case 0
				Response.Write "╋"
			Case 1
				Response.Write "&nbsp;&nbsp;├"
		End Select
		If node.getAttribute("depth")>1 Then
			For ii=2 To node.getAttribute("depth")
				Response.Write "&nbsp;&nbsp;│"
			Next
			Response.Write "&nbsp;&nbsp;├"
		End If
		Response.Write node.getAttribute("boardtype")
		Response.Write "</option>"&vbNewline
	Next
%>
</select>
<br/>
</div>
<div style="float:left;width:20%;">
选择在页面中显示的位置
<br/>
页面顶部广告位：<input type="radio" name="adsset" value="1" class="radio" onclick="_setdemo(this.value);"/><br/>
页面底部广告位：<input type="radio" name="adsset" value="2" class="radio" onclick="_setdemo(this.value);"/><br/>
<!-- 页面文字广告位：<input type="radio" name="adsset" value="6" class="radio" onclick="_setdemo(this.value);"/><br/>
帖间随机文字广告位：<input type="radio" name="adsset" value="5" class="radio" onclick="_setdemo(this.value);"/><br/> -->
帖子楼主顶部广告位：<input type="radio" name="adsset" value="7" class="radio" onclick="_setdemo(this.value);"/><br/>
帖子楼主左边广告位：<input type="radio" name="adsset" value="8" class="radio" onclick="_setdemo(this.value);"/><br/>
帖子楼主右边广告位：<input type="radio" name="adsset" value="9" class="radio" onclick="_setdemo(this.value);"/><br/>
帖子楼主底部广告位：<input type="radio" name="adsset" value="10" class="radio" onclick="_setdemo(this.value);"/><br/>
</div>
<div style="float:left;width:36%;height:220px;">
<img id="ads_setdemo" src="skins/images/ads_set.gif" border="0" alt="广告位置预览"/>
</div>
</td>
</tr>
<tr><td class="td2" colspan="2" align="center">
<input type="hidden" name="adzoneid" value="<%=AdsList.getAttribute("adzoneid")%>"/>
<input type="submit" name="submit" value="确认"/>
</td></tr>
</form>
</table>
<script language="JavaScript">
function _setdemo(v){
	var obj = document.getElementById('ads_setdemo');
	if (obj){
		obj.src = "skins/images/ads_set"+v+".gif";
	}
}
chkcheckbox(document.adsform.getskinid,'<%=AdsList.getAttribute("homepage")%>');
ChkSelected(document.adsform.getboard,'<%=AdsList.getAttribute("getboard")%>');
chkradio(document.adsform.adsset,'<%=AdsList.getAttribute("adsset")%>');
_setdemo('<%=AdsList.getAttribute("adsset")%>');
</script>
<%
End Sub

'保存广告位设置
Sub SaveEditads()
	If Alimama_Api.getAttribute("memberid")="" Then
		Errmsg=ErrMsg + "<BR/>请申请了阿里妈妈的帐号后，再进行发布广告位信息！"
		Dvbbs_error()
		Exit Sub
	End If
	If Alimama_Api.getAttribute("siteid")="" Then
		Errmsg=ErrMsg + "<BR/>您的网站信息还未登记，请登记后再进行发布广告位！"
		Dvbbs_error()
		Exit Sub
	End If
	Dim zonename,adsize,getboard,adsset
	Dim Adzoneid,ActStats
	Adzoneid = Trim(Request.Form("adzoneid"))
	getboard = Trim(Request.Form("getboard"))
	adsset = Trim(Request.Form("adsset"))
	If adsset = "" or Not Isnumeric(adsset) Then
		Errmsg=ErrMsg + "<BR/>请选取广告在页面中显示的位置！"
		Dvbbs_error()
		Exit Sub
	End If

	Set AdsList = Alimama_Api.selectSingleNode("ads[@adzoneid="&Adzoneid&"]")
	If AdsList is Nothing  Then
		Errmsg=ErrMsg + "<BR/>您需要编辑的广告位不存在！"
		Dvbbs_error()
		Exit Sub
	End If
	AdsList.setAttribute "adsset",adsset
	AdsList.setAttribute "getboard",getboard
	Update_Forum_Api()
	UpdateAdsSeting()
	Dv_suc("广告位置修改成功！")
End Sub

'设置广告位数据在页面显示的位置
Sub UpdateAdsSeting()
	Dim iSetting,i,Forum_ads
	Dim adsset,adsstr
	Dim Sql
	
	'广告代码字串
	'mm_alimama会员id_网站id_广告位id
	adsset = AdsList.getAttribute("adsset")
	Adsstr = "<script type=""text/JavaScript"">" &_ 
	"var alimama_pid=""mm_"&Alimama_Api.getAttribute("memberid")&"_"&Alimama_Api.getAttribute("siteid")&"_"&AdsList.getAttribute("adzoneid")&"""; " &_
	"var alimama_sizecode=""21""; " & _ 
	"var alimama_width="&AdsList.getAttribute("width")&"; " &_ 
	"var alimama_height="&AdsList.getAttribute("height")&"; " &_ 
	"var alimama_type="&AdsList.getAttribute("format")&"; " &_ 
	"</script> " &_ 
	"<script language=""javaScript"" type=""text/javascript"" src=""http://p.alimama.com/inf.js""></script>"

	Adsstr = Replace(Adsstr,"$","")
	If AdsList.getAttribute("homepage")="1" Then
		For i = 0 To 30
			iSetting = Trim(Dvbbs.Forum_ads(i))
			If (i = 2 or i = 7 or i = 13 or i=12 or i=15 or i = 17) and Dvbbs.Forum_ads(i)="" Then iSetting = 0

			If adsset = "1" and i=0 Then	'顶部
				iSetting = adsstr
			End If
			If adsset = "2" and i=1 Then	'底部
				iSetting = adsstr
			End If
			If adsset = "7" Then			'帖子楼主顶部广告位
				If i = 18 Then
					iSetting = 1			'开启帖子楼主顶部广告位
				End If
				If i = 19 Then
					iSetting = Adsstr
				End If
			End If

			If adsset = "8" Then			'帖子楼主左边广告位
				If i = 22 Then
					iSetting = 1			'开启帖子楼主左边广告位
				End If
				If i = 23 Then
					iSetting = Adsstr
				End If
			End If
			If adsset = "9" Then			'帖子楼主右边广告位
				If i = 22 Then
					iSetting = 2			'开启帖子楼主右边广告位
				End If
				If i = 23 Then
					iSetting = Adsstr
				End If
			End If
			If adsset = "10" Then			'帖子楼主底部广告位
				If i = 20 Then
					iSetting = 1			'开启帖子楼主底部广告位
				End If
				If i = 21 Then
					iSetting = Adsstr
				End If
			End If
			If i = 0 Then
				Forum_ads = iSetting
			Else
				Forum_ads = Forum_ads & "$" & iSetting
			End If
		Next
		Sql = "Update Dv_Setup Set Forum_ads='"&Replace(Forum_ads,"'","''")&"'"
		Dvbbs.Execute(sql)
	End If
	If AdsList.getAttribute("getboard")<>"" Then
		'查获更新版面数据，只更新设置投放广告项，避免清空原广告其他设置
		Dim Rs
		Set Rs = Dvbbs.Execute("select Boardid,Board_Ads from Dv_board where boardid in ("&Dvbbs.Checkstr(AdsList.getAttribute("getboard"))&")")
		do while not rs.eof
			Dvbbs.Forum_ads = Split(Rs(1),"$")
			For i = 0 To 30
				iSetting = Trim(Dvbbs.Forum_ads(i))
				If (i = 2 or i = 7 or i = 13 or i=12 or i=15 or i = 17) and Dvbbs.Forum_ads(i)="" Then iSetting = 0

				If adsset = "1" and i=0 Then	'顶部
					iSetting = adsstr
				End If
				If adsset = "2" and i=1 Then	'底部
					iSetting = adsstr
				End If
				If adsset = "7" Then			'帖子楼主顶部广告位
					If i = 18 Then
						iSetting = 1			'开启帖子楼主顶部广告位
					End If
					If i = 19 Then
						iSetting = Adsstr
					End If
				End If

				If adsset = "8" Then			'帖子楼主左边广告位
					If i = 22 Then
						iSetting = 1			'开启帖子楼主左边广告位
					End If
					If i = 23 Then
						iSetting = Adsstr
					End If
				End If
				If adsset = "9" Then			'帖子楼主右边广告位
					If i = 22 Then
						iSetting = 2			'开启帖子楼主右边广告位
					End If
					If i = 23 Then
						iSetting = Adsstr
					End If
				End If
				If adsset = "10" Then			'帖子楼主底部广告位
					If i = 20 Then
						iSetting = 1			'开启帖子楼主底部广告位
					End If
					If i = 21 Then
						iSetting = Adsstr
					End If
				End If
				If i = 0 Then
					Forum_ads = iSetting
				Else
					Forum_ads = Forum_ads & "$" & iSetting
				End If
			Next
			Sql = "Update Dv_Board Set Board_Ads='"&Replace(Forum_ads,"'","''")&"' Where BoardID ="&Rs(0)
			Dvbbs.Execute(Sql)
		Rs.movenext
		Loop
		Rs.close
		Set Rs = Nothing
	End If
	RestoreBoardCache()
	Dvbbs.loadSetup()
End Sub


'保存广告位信息
Sub SaveAds()
	If Alimama_Api.getAttribute("memberid")="" Then
		Errmsg=ErrMsg + "<BR/>请申请了阿里妈妈的帐号后，再进行发布广告位信息！"
		Dvbbs_error()
		Exit Sub
	End If
	If Alimama_Api.getAttribute("siteid")="" Then
		Errmsg=ErrMsg + "<BR/>您的网站信息还未登记，请登记后再进行发布广告位！"
		Dvbbs_error()
		Exit Sub
	End If
	Dim zonename,weekprice,needAuditingSelect,format,adsize,getboard,adsset,zonedesc,keywords
	Dim homepage,transtype
	Dim Adzoneid,ActStats
	Adzoneid = Request.Form("adzoneid")
	zonename = Trim(Request.Form("zonename"))
	weekprice = Trim(Request.Form("weekprice"))
	needAuditingSelect = Trim(Request.Form("needAuditingSelect"))
	format = Trim(Request.Form("format"))
	adsize = Trim(Request.Form("adsize"))
	getboard = Trim(Request.Form("getboard"))
	adsset = Trim(Request.Form("adsset"))
	zonedesc  = Trim(Request.Form("zonedesc"))
	keywords = Trim(Request.Form("keywords"))
	'提交信息验证
	If zonename=""or Len(zonename)<1 or Len(zonename)>32 Then
		Errmsg=ErrMsg + "<BR/>广告位名称不能为空或超过32个字符！"
		Dvbbs_error()
		Exit Sub
	End If
	If weekprice="" or Not IsNumeric(weekprice) Then
		Errmsg=ErrMsg + "<BR/>时长计费设置必须为数字字符！"
		Dvbbs_error()
		Exit Sub
	End If
	If Ccur(weekprice)<0.05 Then
		Errmsg=ErrMsg + "<BR/>时长计费设置必须大于等于0.05！"
		Dvbbs_error()
		Exit Sub
	End If
	If needAuditingSelect="" or Not IsNumeric(needAuditingSelect) Then
		Errmsg=ErrMsg + "<BR/>请正确选择是否需要审核！"
		Dvbbs_error()
		Exit Sub
	End If
	If format="" or Not IsNumeric(format) Then
		Errmsg=ErrMsg + "<BR/>请正确选择广告形式！"
		Dvbbs_error()
		Exit Sub
	End If
	If adsize="" or Not IsNumeric(adsize) Then
		Errmsg=ErrMsg + "<BR/>请正确选择广告位尺寸！"
		Dvbbs_error()
		Exit Sub
	End If
	If zonedesc="" Then
		Errmsg=ErrMsg + "<BR/>广告描述不能为空！"
		Dvbbs_error()
		Exit Sub
	End If
	If adsset = "" or Not Isnumeric(adsset) Then
		Errmsg=ErrMsg + "<BR/>请选取广告在页面中显示的位置！"
		Dvbbs_error()
		Exit Sub
	End If
	If keywords = "" or Len(keywords)>16 Then
		Errmsg=ErrMsg + "<BR/>请填写广告位关键字或不要超过16个字符！"
		Dvbbs_error()
		Exit Sub
	End If
	If format = "1" and (adsset="3" or adsset="4") Then
		Errmsg=ErrMsg + "<BR/>文字广告类型，不能设置在浮动或右下固定广告位！"
		dvbbs_error()
		Exit Sub
	End If

	'是否显示在首页类型
	homepage = Request.Form("getskinid") '在首页:1
	'？transtype  计费类型 1:cpc 2:cpt 3:cpc+cpt  4:cpm
	transtype = 2
	'-------------------------------------------------------------------------
	'广告位置检查
	'-------------------------------------------------------------------------
	'广告位名称同一个站点下名称不能重复
	If Adzoneid<>"" Then
		Set AdsList = Alimama_Api.selectSingleNode("ads[@adzoneid="&Adzoneid&"]")
		If AdsList is Nothing  Then
			Errmsg=ErrMsg + "<BR/>您需要编辑的广告位不存在！"
			Dvbbs_error()
			Exit Sub
		End If
		ActStats = "编辑"
		If Alimama_Api.selectNodes("ads[@name='"&zonename&"' and @adzoneid != "&Adzoneid&" ]").Length>0 Then
			Errmsg=ErrMsg + "<BR/>广告位名称不能有重复！"
			Dvbbs_error()
			Exit Sub
		End If
		AdsList.setAttribute "updatetime",Now()
		Alimama.QueryStr "service","zone_edit"
		Alimama.QueryStr "adzoneid",AdsList.getAttribute("adzoneid")
	Else
		ActStats = "发布"
		Set AdsList = Alimama_Api.selectNodes("ads[@name='"&zonename&"']")
		If AdsList.Length>0 Then
			Errmsg=ErrMsg + "<BR/>广告位名称不能有重复！"
			Dvbbs_error()
			Exit Sub
		End If
		'创建一个新的广告数据节点
		Set AdsList = Alimama_Api.appendChild(Forum_Api.createNode(1,"ads",""))
		AdsList.setAttribute "createtime",Now()
		AdsList.setAttribute "updatetime",Now()
		Alimama.QueryStr "service","zone_add"
	End If

	'根据选取类型，获取宽高
	Select Case adsize
		Case "11"
			AdsList.setAttribute "width",760
			AdsList.setAttribute "height",90
		Case "12"
			AdsList.setAttribute "width",468
			AdsList.setAttribute "height",60
		Case "13"
			AdsList.setAttribute "width",250
			AdsList.setAttribute "height",60
		Case "21"
			AdsList.setAttribute "width",120
			AdsList.setAttribute "height",600
		Case "22"
			AdsList.setAttribute "width",120
			AdsList.setAttribute "height",240
		Case "31"
			AdsList.setAttribute "width",180
			AdsList.setAttribute "height",250
		Case "32"
			AdsList.setAttribute "width",250
			AdsList.setAttribute "height",300
		Case "33"
			AdsList.setAttribute "width",360
			AdsList.setAttribute "height",190
		Case Else
			AdsList.setAttribute "width",0
			AdsList.setAttribute "height",0
	End Select

	AdsList.setAttribute "name",zonename
	AdsList.setAttribute "weekprice",weekprice
	AdsList.setAttribute "needAuditingSelect",needAuditingSelect
	AdsList.setAttribute "format",format
	AdsList.setAttribute "adsize",adsize
	AdsList.setAttribute "homepage",homepage
	AdsList.setAttribute "transtype",transtype
	AdsList.setAttribute "adsset",adsset
	AdsList.setAttribute "getboard",getboard
	AdsList.setAttribute "zonedesc",zonedesc
	AdsList.setAttribute "keywords",zonedesc
	AdsList.setAttribute "adzonecatids",1099

	'-------------------------------------------------------------------------
	'创建API接口参数
	'-------------------------------------------------------------------------
	Dim XmlDom
	Alimama.QueryStr "memberid",Alimama_Api.getAttribute("memberid")
	Alimama.QueryStr "siteid",Alimama_Api.getAttribute("siteid")
	Alimama.QueryStr "name",zonename
	Alimama.QueryStr "transtype",transtype
	Alimama.QueryStr "weekprice",weekprice
	Alimama.QueryStr "format",format
	Alimama.QueryStr "sizecode",adsize
	Alimama.QueryStr "needauditing",needAuditingSelect
	Alimama.QueryStr "homepage",homepage
	Alimama.QueryStr "adzonecatids",1099			'?广告位类目
	Alimama.QueryStr "keywords",keywords		'?关键字描述
	Alimama.QueryStr "zonedesc",zonedesc

	'显示Loading状态提示
	LoadingDiv()
	Set XmlDom = Alimama.Post
	If XmlDom is Nothing Then
		CloseLoadingDiv()
		Errmsg=ErrMsg + "<BR/>意外错误，获取提交数据中止，请稍候再试！"
		Dvbbs_error()
		Exit Sub
	End If
	CloseLoadingDiv()
	'Is_success,Error_code,Data_List
	If Not (Alimama.Is_success is Nothing) Then
		If Alimama.Is_success.text<>"true" Then
			'If Alimama.Error_code.text<"1999" Then
			'	Errmsg=ErrMsg + "<BR/>(Errcode:"&Alimama.Error_code.text&")系统错误，获取提交数据中止，请稍候再试！"
			'	Dvbbs_error()
			'	Exit Sub			
			'End If
			Errmsg=ErrMsg + "<BR/>"&Alimama.Error_desc
			Dvbbs_error()
			Exit Sub
		Else
			If Adzoneid="" Then
				AdsList.setAttribute "adzoneid",Alimama.Data_List.selectSingleNode("adzone/adzoneid").text
			End If
			Update_Forum_Api()
			UpdateAdsSeting()
			Dv_suc(ActStats&"广告位成功！")
		End If
	Else
		Errmsg=ErrMsg + "<BR/>返回数据错误，操作已中止，请稍候再试！"
		Dvbbs_error()
		Exit Sub
	End If
End Sub

'查看广告位列表
Sub MyAdsList()
	Set AdsList = Alimama_Api.selectNodes("ads")
	If AdsList.Length=0 Then
		Errmsg=ErrMsg + "<BR/>暂未有广告位数据！"
		Dvbbs_error()
		Exit Sub
	End If
	Dim Node
%>
<br/>
<table border="0" cellspacing="1" cellpadding="3"  align="center" width="100%">
<tr><th colspan="7" style="text-align:center;" colspan="4">我的广告位信息</th></tr>
<tr>
<td class="td2 title" >广告位名称</td>
<td class="td2 title" width="10%">广告位ID</td>
<td class="td2 title" width="10%">广告位形式</td>
<td class="td2 title" width="10%">价格(元/周)</td>
<td class="td2 title" width="15%">创建时间</td>
<td class="td2 title" width="15%">更新时间</td>
<td class="td2 title" >操作</td>
</tr>
<%
	For Each Node In AdsList
%>
<tr>
<td><%=Node.getAttribute("name")%></td>
<td><%=Node.getAttribute("adzoneid")%></td>
<td>cpt</td>
<td><%=Node.getAttribute("weekprice")%></td>
<td>
<%
If IsDate(Node.getAttribute("createtime")) Then
Response.Write Formatdatetime(Node.getAttribute("createtime"),0)
Else
Response.Write Node.getAttribute("createtime")
End If
%></td>
<td>
<%
If IsDate(Node.getAttribute("updatetime")) Then
Response.Write Formatdatetime(Node.getAttribute("updatetime"),0)
Else
Response.Write Node.getAttribute("updatetime")
End If
%></td>
<td><a href="?act=editads&adzoneid=<%=Node.getAttribute("adzoneid")%>">编辑</a></td>
</tr>
<%
	Next
%>
</table>
<%
Set AdsList = Nothing
End Sub

'更新版面广告缓存数据
Sub RestoreBoardCache()
	Dim Board,node
	Dvbbs.LoadBoardList()
	For Each node in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style/@id")
		Application.Contents.Remove(Dvbbs.CacheName & "_showtextads_"&node.text)
		For Each board in Application(Dvbbs.CacheName&"_boardlist").documentElement.selectNodes("board/@boardid")
			Dvbbs.LoadBoardData board.text
			Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_"& board.text &"_"&node.text)
			Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_"& board.text &"_"&node.text&"_-time")
		Next
		Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_0_"& node.text)
		Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_0_"& node.text&"_-time")
	Next
End Sub

'-------------------------------------------------------------------------
'阿里妈妈API接口类
'-------------------------------------------------------------------------

Class Alimama_Apicls
	Public Return_url	'可选设置
	Public ApiUrl,ApiKey,Service,Partner,Sign,Sign_type	'必选数据设置
	Private Query_Str
	Public Is_success,Error_code,Data_List,Error_desc

	Private Sub Class_Initialize()
		'ApiUrl = "http://192.168.1.103:89/aliads_api.asp"
		ApiUrl = "http://server.dvbbs.net/ads_api.asp"
		Sign_type = "MD5"
		Query_Str = ""
		'QueryStr "sign_type",Sign_type	'加密方式
		'QueryStr "return_url",Dvbbs.Get_ScriptNameUrl	'返回url
		Set Is_success = Nothing
		Set Error_code = Nothing
		Set Data_List = Nothing
		Error_desc = ""
	End Sub

	Private Sub Class_Terminate()
	End Sub
	
	'nickname 格式化
	Public Function FormatNickName(Str)
		FormatNickName = Str & "_cndw"
	End Function

	'设置URL参数（参数名，参数值）
	Public Function QueryStr(qName,qStr)
		If qStr="" or IsNull(qStr) Then Exit Function
		If Query_Str<>"" Then Query_Str = Query_Str & "&"
		Query_Str = Query_Str & qName & "=" & qStr
	End Function

	'获取发送URL地址
	Public Property Get PostUrl()
		Dim Str
		Str = SortByEn(Query_Str,"&")
		'Sign = Md5(Str & ApiKey,32)
		'Str = Str & "&sign=" & Sign
		'Str = Str & "&sign_type=" & Sign_type
		PostUrl = ApiUrl & "?" & SortByEn(Str,"&")
		'Response.Write PostUrl
		'Response.Write "<br>"
	End Property

	'将URL的参数按字母排序
	Public Function SortByEn(Str,S)
		Dim Sort,QStr
		QStr = Mid(Str,InStrRev(Str, "?")+1)
		Sort = JsSortByEn(QStr,S)
		SortByEn = Left(Str,Instr(Str,"?")) & Sort
	End Function

	Public Function Post()
		Dim XmlHttp,XmlDom
		Set XmlHttp = Server.CreateObject("Microsoft.XMLHTTP")
		XmlHttp.Open "post",PostUrl,false
		XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		'XmlHttp.SetRequestHeader "content-type", "text/xml"
		XmlHttp.send()

		Set XmlDom = Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		If XmlDom.Load(xmlhttp.responseXML) Then
			Set Post = XmlDom
			Set Is_success = XmlDom.documentElement.selectSingleNode("is_success")
			
			Set Error_code = XmlDom.documentElement.selectSingleNode("error_code")
			If Not (Error_code is Nothing) Then
				Error_desc = "(Errcode:" & Error_code.text & "):" & Error_code.getAttribute("desc")
			End If
			Set Data_List = XmlDom.documentElement.selectSingleNode("data/list")
		Else
			'Response.Write PostUrl
			Response.Write xmlhttp.responsetext
			Set Post = Nothing
		End If
		Set XmlHttp = Nothing
		Set XmlDom = Nothing
	End Function

End Class

'数组排序
Function VBSort(ary)
	Dim KeepChecking,I,FirstValue,SecondValue
	 KeepChecking = TRUE 
	Do Until KeepChecking = FALSE 
	 KeepChecking = FALSE 
	 For I = 0 to UBound(ary) 
	  If I = UBound(ary) Then Exit For 
	   If ary(I) > ary(I+1) Then 
		FirstValue = ary(I) 
		SecondValue = ary(I+1) 
		ary(I) = SecondValue 
		ary(I+1) = FirstValue 
		KeepChecking = TRUE 
	   End If 
	 Next 
	Loop 
	 VBSort = ary 
End Function

'获取当前时间
function get_nowtime()
	dim year_str , month_str , day_str , hour_str , minute_str , second_str
	year_str=cstr(year(now))
	month_str=cstr(month(now))
	day_str=cstr(day(now))
	hour_str=cstr(hour(now))
	minute_str=cstr(minute(now))
	second_str=cstr(second(now))
	if len(hour_str)=1 then hour_str="0"+hour_str
	if len(minute_str)=1 then minute_str="0"+minute_str
	if len(second_str)=1 then second_str="0"+second_str
	get_nowtime=year_str+"-"+month_str+"-"+day_str+" "+hour_str+":"+minute_str+":"+second_str 
end function


%>
<script type="text/javascript" runat="server" language="javascript">
//字符串按字母排序,S为分隔符
function JsSortByEn(Str,S){
   var a, l;
   a = Str.split(S);
   l = a.sort();                   // 排序数组。
   return(l.join(S));
}
</script>