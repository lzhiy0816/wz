<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
If Request("t")="1" Then
	Face_Main()
Else
	Main()
End If

Sub Main()
	Dvbbs.Loadtemplates("post")
	Dim star
	If Request("star")<>"" Then
		star=Request("star") 
	Else 
		star=1
	End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<meta HTTP-EQUIV=Expires CONTENT=0>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>Dvbbs多功能编辑器</TITLE>
<Link rel="stylesheet" type="text/css" href="images/post/pop.css">
<style type="text/css">
				.Disactive { border-right: 1px solid; border-top: 1px solid; border-left: 1px solid; width: 1%; cursor: hand; border-bottom: 1px solid; background-color: #dedbd6; text-align: center; }
				.Active { cursor: hand; background-color: #ffffcc; text-align: center; }
				.MainTable { border-right: #e8e8e8 5px solid; border-top: #e8e8e8 5px solid; border-left: #e8e8e8 5px solid; border-bottom: #e8e8e8 5px solid; }
				.Sample { border-right: 1px solid; border-top: 1px solid; font-size: 24px; border-left: 1px solid; border-bottom: 1px solid; background-color: #dedbd6; }
				.Empty { border-right: 1px solid; border-top: 1px solid; border-left: 1px solid; width: 1%; cursor: default; border-bottom: 1px solid; background-color: #dedbd6; }
</style>
<script language="javascript">
<!--
var oSample ;
var code;
var n
function insertsmilie(smilieface){
	if (self.opener.Dvbbs_Composition){
	self.opener.Dvbbs_Composition.document.body.innerHTML+=smilieface;
	}
	else if (self.opener.frmAnnounce.Body){
	self.opener.frmAnnounce.Body.value+=smilieface;
	}
}

function insertChar(td)
{
	window.returnValue = td.innerHTML+"*"+td.title || "" ;
	//window.close();
}

function over(td)
{
	oSample.innerHTML = td.innerHTML ;
	td.className = 'Active' ;
	code.innerHTML = td.title;
}
function out(td)
{
	oSample.innerHTML = "&nbsp;" ;
	td.className = 'Disactive' ;
	code.innerHTML = "";
}

function CloseWindow()
{
	window.returnValue = null ;
	window.close() ;
}

function pageopen(id){
//this.location.replace("smiley.asp?star="+id+"");
location.reload("smiley.asp?star="+id+"")
}
//-->
</script>
</head>
<body topmargin="0" rightmargin="0" bottommargin="0" leftmargin="0" bgcolor="menu" >
<table cellpadding="0" cellspacing="10" width="100%" height="100%">
<tr><td rowspan="2" width="90%"><INPUT name=SmileStar value=1 type=hidden>
<table class="MainTable" cellpadding="2" cellspacing="1" align="center" border="1" width="100%" height="100%">
<script language="javascript">
<!--
<%
Response.Write "var aImages='"&Dvbbs.Forum_Emot&"';"
Response.Write "var star="&star&";"
%>
var str=document.getElementById('SmileStar').value
var sBasePath = '' ;
aImages		  = aImages.split("|||");
var cols      = 8 ;
showmain(star)
function showmain(star){
star=Math.floor(star);
var ImgCount=aImages.length-2
var i = 24*star-24+1 ;
while (i <24*star)
{
	document.write("<TR>") ;
	for(var j = 0 ; j < cols ; j++) 
	{
		if (aImages[i])
		{
			if (i<10)
			{ii='0'+i}
			else
			{ii=i}
			//document.write("<TD title=[em"+ii+"] class='Disactive' align=center onclick='insertChar(this)' onmouseover='over(this);' onmouseout='out(this)'>") ;
			document.write("<TD title=[em"+ii+"] class='Disactive' align=center onclick=\"insertsmilie('[em" + ii + "]')\" style=\"CURSOR: hand\" onmouseover='over(this);' onmouseout='out(this)'>") ;
			insertsmilie
			document.write("<img src='" + aImages[0] + aImages[i] + "' border='0'>") ;
		}
		else
			document.write("<TD width='10%' class='Empty'>&nbsp;") ;
		document.write("</TD>") ;
		i++ ;
	}
	document.write("</TR>") ;
}
if (ImgCount>24){
document.write("<TR>") ;
document.write("<TD height=20 colspan="+cols+">") ;
pagecount('red',24,ImgCount,star)
}

//分页
function pagecount(alertcolor,PageSize,TopicNum,star){
PageSize=Math.floor(PageSize);
TopicNum=Math.floor(TopicNum);
star=Math.floor(star);
var n,p;
if ((star-1)%10==0) 
{
	p=(star-1) /10
}
else
{
	p=(((star-1)-(star-1)%10)/10)
}
if(TopicNum%PageSize==0) 
{
	n=TopicNum/PageSize;
}
else
{
	n=(TopicNum-TopicNum%PageSize)/PageSize+1;
}
document.write ('<table border="0" cellpadding="0" cellspacing="3" width="100%" >');
document.write ('<tr><td valign="middle" width="50%">');
document.write ('总文件数: <b>'+TopicNum+' </b>&nbsp;&nbsp;分页：');
if (star==1)
{
	document.write ('<font face=webdings color="'+alertcolor+'">9</font>');
}
else
{
	document.write ('<a href="#" onclick=pageopen('+star+') title="首页"><font face=webdings>9</font></a> ');
}
if (p*10 > 0)
{
	document.write ('<a onclick=pageopen('+p*10+') href="#" title="上十页"><font face=webdings>7</font></a> ');
}
document.write ('<b>');
for (var i=p*10+1;i<p*10+11;i++)
{
	if (i==star)
	{
		document.write (' <font color="'+alertcolor+'">'+i+'</font> ');
	}
	else
	{
		document.write (' <a href="#" onclick=pageopen('+i+') >'+i+'</a> ');
	}
	if (i==n) break;
}
document.write ('</b>');
if (i<n)
{
	document.write ('<a onclick=pageopen('+i+') href="#" title="下十页"><font face=webdings>8</font></a>   ');
}
if (star==n)
{
	document.write ('<Font face=webdings color="'+alertcolor+'">:</font>');
}
else
{
	document.write (' <a onclick=pageopen('+n+') href="#" title="尾页"><font face=webdings>:</font></a>  ');
}
document.write ('</td></tr></table>');
}

document.write("</TD>") ;
document.write("</TR>") ;
}
//-->
</script>
</table>
</td>
<td valign="top" width="10%" align="center">
<table class="MainTable" align="center" cellspacing="2">
<tr><td id="SampleTD" width="80" height="80" align="center" class="Sample">&nbsp;
</td></tr></table>
<div id=code></div>
</td>
</tr>
<tr><td align=center height=1><BUTTON onclick=window.close();>关闭</BUTTON></td></tr>
</table>
</body>
</html>
<script language="javascript">
<!--
oSample = document.getElementById("SampleTD") ;
code = document.getElementById("code") ;
//-->
</script>
<%
End Sub

Sub Face_Main()
	Dim n,m,h,s
	Dim y,z,show
	Dim stats,Forum_userfaceNum,Forum_userface,Forum_facepath
	Dim i
	stats="论坛头像列表"
	h=6 '每行显示个数
	s=30 '每页显示个数

	Dvbbs.LoadTemplates("usermanager")
	Dvbbs.Stats="用户头像列表"
	Dvbbs.head()
	'Dvbbs.Head_var 0,0,template.Strings(0),"usermanager.asp"
	Forum_userface=split(Dvbbs.Forum_userface,"|||")
	Forum_facepath=Forum_userface(0)
	Forum_userfacenum=ubound(Forum_userface)
	M=Forum_userfaceNum-1 '头象个数
	Dim TempStr,TempLateStr
	Dim Facebar,p
	TempStr=template.Strings(66)
	If m>s Then
		For Y=1 to (m+s-1)\s
		Facebar=Facebar&Replace(TempStr,"{$page}",y)
		Next
	End If
	TempStr=Split(template.html(19),"||")
	TempStr(0)=Replace(TempStr(0),"{$FaceBar}",Facebar)
	Response.Write TempStr(0)
	i=0
	p= trim(request("show"))
	if p="" or not IsNumeric(p) then
	p=1
	end if
	z=s*p
	if z/m>1 then
		z=m
	else
		z=s*p
	end if
	for n=s*p+1-s to z
	i=i+1
	TempLateStr=TempStr(1)
	TempLateStr=Replace(TempLateStr,"{$Forum_facepath}",Forum_facepath)
	TempLateStr=Replace(TempLateStr,"{$Forum_userface_n}",Forum_userface(n))
	Response.Write TempLateStr
	if i=h then
	response.write TempStr(2)
	i=0
	end if 
	next
	Response.Write TempStr(3)
	Dvbbs.Footer()
End Sub
%>