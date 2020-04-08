<!--#include file="../conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Dim captcha_chartype_chinese,captcha_chartype_english,captcha_chartype_number
Dim captcha_size
Dim captcha_width_lbound,captcha_width_ubound,captcha_height_lbound,captcha_height_ubound,captcha_spacing_lbound,captcha_spacing_ubound
Dim captcha_angle_lbound,captcha_angle_ubound,captcha_weight
Dim captcha_charshow
Dim captcha_charshow_stepbystep_r,captcha_charshow_stepbystep_g,captcha_charshow_stepbystep_b
Dim captcha_charshow_simple_r,captcha_charshow_simple_g,captcha_charshow_simple_b
Dim captcha_backshow
Dim captcha_backshow_stepbystep_r,captcha_backshow_stepbystep_g,captcha_backshow_stepbystep_b
Dim captcha_backshow_simple_r,captcha_backshow_simple_g,captcha_backshow_simple_b
Dim captcha_charshow_mix_percent
Dim captcha_backshow_mix_percent
Dim captcha_pic_width,captcha_pic_height

Dim admin_flag
admin_flag=",1,"
CheckAdmin(admin_flag)

If "save"=Request("action") Then
    Call SaveCaptcha()
ElseIf "savetotemp"=Request("action") Then
	Call PreviewCaptcha()
Else
    Call Head()
    Call SetForm()
End If

Dvbbs.PageEnd

Function SaveCaptchaFile(str, path)
	On Error Resume Next 
	DvStream.charset="gb2312"
	DvStream.Mode = 3
	DvStream.open()
	DvStream.WriteText(str)
	DvStream.SaveToFile Server.MapPath(path),2
	DvStream.close()
	If Err Then
		Err.clear
		SaveCaptchaFile = False 
	Else
		SaveCaptchaFile = True 
	End If 
End Function 

Sub PreviewCaptcha()
	Dim str
	str = GetScriptContent()
	If Not SaveCaptchaFile(str, "CaptchaPreview.asp") Then 
	Call Head()
	%>
		<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
		<tr> 
		<td height="23"><b><font color=red>更新临时文件(admin/CaptchaPreview.asp)失败！给临时文件加入可写权限才能预览。</font></b></td>
		</tr>
		</table>
	<%
		Dvbbs.PageEnd
		Response.End 
	Else
		Dvbbs.PageEnd
		Response.Redirect "CaptchaSet.asp?action=preview"
	End If 
End Sub 

Sub SaveCaptcha()
	Dim str
	str = GetScriptContent()
	If Not SaveCaptchaFile(str, "../"&DvCodeFile) Then 
	Call Head()
	%>
		<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
		<tr> 
		<td height="23"><b><font color=red>更新验证码文件(../<%=DvCodeFile%>)失败！请给验证码文件加入可写权限，或者直接复制以下内容到验证码文件，替换掉原来的内容。</font></b></td>
		</tr>
		<tr> 
		<td><textarea style="width:500px;height:200px;" onfocus="this.select()"><%=str%></textarea></td>
		</tr>
		</table>
	<%
		Dvbbs.PageEnd
		Response.End 
	Else
		Dvbbs.PageEnd
		Response.Redirect "CaptchaSet.asp"
	End If 
End Sub 

Function GetScriptContent()
    Dim str,temp
    captcha_chartype_chinese = CInt(Request("captcha_chartype_chinese"))
    captcha_chartype_english = CInt(Request("captcha_chartype_english"))
    captcha_chartype_number = CInt(Request("captcha_chartype_number"))
    captcha_size = CInt(Request("captcha_size"))
    captcha_width_lbound = CInt(Request("captcha_width_lbound"))
    captcha_width_ubound = CInt(Request("captcha_width_ubound"))
    captcha_height_lbound = CInt(Request("captcha_height_lbound"))
    captcha_height_ubound = CInt(Request("captcha_height_ubound"))
    captcha_spacing_lbound = CInt(Request("captcha_spacing_lbound"))
    captcha_spacing_ubound = CInt(Request("captcha_spacing_ubound"))
    captcha_angle_lbound = CInt(Request("captcha_angle_lbound"))
    captcha_angle_ubound = CInt(Request("captcha_angle_ubound"))
    captcha_weight = CInt(Request("captcha_weight"))
    captcha_charshow = CInt(Request("captcha_charshow"))
    captcha_charshow_simple_r = CInt(Request("captcha_charshow_simple_r"))
    captcha_charshow_simple_g = CInt(Request("captcha_charshow_simple_g"))
    captcha_charshow_simple_b = CInt(Request("captcha_charshow_simple_b"))
    captcha_charshow_stepbystep_r = CInt(Request("captcha_charshow_stepbystep_r"))
    captcha_charshow_stepbystep_g = CInt(Request("captcha_charshow_stepbystep_g"))
    captcha_charshow_stepbystep_b = CInt(Request("captcha_charshow_stepbystep_b"))
    captcha_charshow_mix_percent = CInt(Request("captcha_charshow_mix_percent"))
    captcha_backshow = CInt(Request("captcha_backshow"))
    captcha_backshow_simple_r = CInt(Request("captcha_backshow_simple_r"))
    captcha_backshow_simple_g = CInt(Request("captcha_backshow_simple_g"))
    captcha_backshow_simple_b = CInt(Request("captcha_backshow_simple_b"))
    captcha_backshow_stepbystep_r = CInt(Request("captcha_backshow_stepbystep_r"))
    captcha_backshow_stepbystep_g = CInt(Request("captcha_backshow_stepbystep_g"))
    captcha_backshow_stepbystep_b = CInt(Request("captcha_backshow_stepbystep_b"))
    captcha_backshow_mix_percent = CInt(Request("captcha_backshow_mix_percent"))
    captcha_pic_width = CInt(Request("captcha_pic_width"))
    captcha_pic_height = CInt(Request("captcha_pic_height"))
    str = Dvbbs.ReadTextFile("CaptchaTemplate.txt")
    str = Replace(str, "{$captcha_chartype_number}", captcha_chartype_number)
    str = Replace(str, "{$captcha_chartype_english}", captcha_chartype_english)
    str = Replace(str, "{$captcha_chartype_chinese}", captcha_chartype_chinese)
    str = Replace(str, "{$captcha_size}", captcha_size)
    str = Replace(str, "{$captcha_width_ubound}", captcha_width_ubound)
    str = Replace(str, "{$captcha_width_lbound}", captcha_width_lbound)
    str = Replace(str, "{$captcha_height_ubound}", captcha_height_ubound)
    str = Replace(str, "{$captcha_height_lbound}", captcha_height_lbound)
    str = Replace(str, "{$captcha_spacing_ubound}", captcha_spacing_ubound)
    str = Replace(str, "{$captcha_spacing_lbound}", captcha_spacing_lbound)
    str = Replace(str, "{$captcha_pic_width}", captcha_pic_width)
    str = Replace(str, "{$captcha_pic_height}", captcha_pic_height)
    str = Replace(str, "{$captcha_angle_lbound}", captcha_angle_lbound)
    str = Replace(str, "{$captcha_angle_ubound}", captcha_angle_ubound)
    str = Replace(str, "{$captcha_weight}", captcha_weight)
    If captcha_weight>0 Then
        str = Replace(str, "{$captcha_weight_real}", captcha_weight)
    Else
        str = Replace(str, "{$captcha_weight_real}", "GetRnd(1,2)")
    End If
    str = Replace(str, "{$captcha_charshow}", captcha_charshow)
    str = Replace(str, "{$captcha_charshow_mix_percent}", captcha_charshow_mix_percent)
	str = Replace(str, "{$captcha_charshow_simple_r}", captcha_charshow_simple_r)
	str = Replace(str, "{$captcha_charshow_simple_g}", captcha_charshow_simple_g)
	str = Replace(str, "{$captcha_charshow_simple_b}", captcha_charshow_simple_b)
	str = Replace(str, "{$captcha_charshow_stepbystep_r}", captcha_charshow_stepbystep_r)
	str = Replace(str, "{$captcha_charshow_stepbystep_g}", captcha_charshow_stepbystep_g)
	str = Replace(str, "{$captcha_charshow_stepbystep_b}", captcha_charshow_stepbystep_b)
    If 0=captcha_charshow Then
        temp = "dim ccr,ccg,ccb" & VBNewline
        If -1=captcha_charshow_simple_r Then
            temp = temp & "ccr=GetRnd(0,255)" & VBNewline
        Else
            temp = temp & "ccr=" & captcha_charshow_simple_r & VBNewline
        End If
        If -1=captcha_charshow_simple_g Then
            temp = temp & "ccg=GetRnd(0,255)" & VBNewline
        Else
            temp = temp & "ccg=" & captcha_charshow_simple_g & VBNewline
        End If
        If -1=captcha_charshow_simple_b Then
            temp = temp & "ccb=GetRnd(0,255)" & VBNewline
        Else
            temp = temp & "ccb=" & captcha_charshow_simple_b & VBNewline
        End If
        str = Replace(str, "{$pre_area1}", temp) & VBNewline
        str = Replace(str, "{$char_r}", "ccr")
        str = Replace(str, "{$char_g}", "ccg")
        str = Replace(str, "{$char_b}", "ccb")
    Else
        str = Replace(str, "{$pre_area1}", "")
        If -1=captcha_charshow_stepbystep_r Then
            str = Replace(str, "{$char_r}", "i")
        Else
            str = Replace(str, "{$char_r}", captcha_charshow_stepbystep_r)
        End If
        If -1=captcha_charshow_stepbystep_g Then
            str = Replace(str, "{$char_g}", "i")
        Else
            str = Replace(str, "{$char_g}", captcha_charshow_stepbystep_g)
        End If
        If -1=captcha_charshow_stepbystep_b Then
            str = Replace(str, "{$char_b}", "i")
        Else
            str = Replace(str, "{$char_b}", captcha_charshow_stepbystep_b)
        End If
    End If
    str = Replace(str, "{$captcha_backshow}", captcha_backshow)
    str = Replace(str, "{$captcha_backshow_mix_percent}", captcha_backshow_mix_percent)
	str = Replace(str, "{$captcha_backshow_simple_r}", captcha_backshow_simple_r)
	str = Replace(str, "{$captcha_backshow_simple_g}", captcha_backshow_simple_g)
	str = Replace(str, "{$captcha_backshow_simple_b}", captcha_backshow_simple_b)
	str = Replace(str, "{$captcha_backshow_stepbystep_r}", captcha_backshow_stepbystep_r)
	str = Replace(str, "{$captcha_backshow_stepbystep_g}", captcha_backshow_stepbystep_g)
	str = Replace(str, "{$captcha_backshow_stepbystep_b}", captcha_backshow_stepbystep_b)
    If 0=captcha_backshow Then
        temp = "dim bcr,bcg,bcb" & VBNewline
        If -1=captcha_backshow_simple_r Then
            temp = temp & "bcr=GetRnd(0,255)" & VBNewline
        Else
            temp = temp & "bcr=" & captcha_backshow_simple_r & VBNewline
        End If
        If -1=captcha_backshow_simple_g Then
            temp = temp & "bcg=GetRnd(0,255)" & VBNewline
        Else
            temp = temp & "bcg=" & captcha_backshow_simple_g & VBNewline
        End If
        If -1=captcha_backshow_simple_b Then
            temp = temp & "bcb=GetRnd(0,255)" & VBNewline
        Else
            temp = temp & "bcb=" & captcha_backshow_simple_b & VBNewline
        End If
        str = Replace(str, "{$pre_area2}", temp) & VBNewline
        str = Replace(str, "{$back_r}", "bcr")
        str = Replace(str, "{$back_g}", "bcg")
        str = Replace(str, "{$back_b}", "bcb")
    Else
        str = Replace(str, "{$pre_area2}", "")
        If -1=captcha_backshow_stepbystep_r Then
            str = Replace(str, "{$back_r}", "i")
        Else
            str = Replace(str, "{$back_r}", captcha_backshow_stepbystep_r)
        End If
        If -1=captcha_backshow_stepbystep_g Then
            str = Replace(str, "{$back_g}", "i")
        Else
            str = Replace(str, "{$back_g}", captcha_backshow_stepbystep_g)
        End If
        If -1=captcha_backshow_stepbystep_b Then
            str = Replace(str, "{$back_b}", "i")
        Else
            str = Replace(str, "{$back_b}", captcha_backshow_stepbystep_b)
        End If
    End If
    GetScriptContent = "<"&"%"&VBNewline&str&VBNewline&"%"&">"
End Function

Sub SetForm()
Dim str,re,match,xmlObj
If "preview"=request("action") Then
	str = Dvbbs.ReadTextFile("CaptchaPreview.asp")
Else
	str = Dvbbs.ReadTextFile("../"&DvCodeFile)
End If 
Set re=new RegExp
re.IgnoreCase =true
re.Global=True
re.Pattern ="<root>.+</root>"
If re.Test(str) Then
	Set match=re.Execute(str)
	str=match.item(0)
	Set match=Nothing 
Else
	str=""
End If 
Set re=Nothing
Set xmlObj=Dvbbs.CreateXmlDoc("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
If ""<>str And xmlObj.loadXml(str) Then
	captcha_chartype_chinese = xmlObj.documentElement.selectSingleNode("captcha_chartype_chinese").text
	captcha_chartype_english = xmlObj.documentElement.selectSingleNode("captcha_chartype_english").text
	captcha_chartype_number = xmlObj.documentElement.selectSingleNode("captcha_chartype_number").text
	captcha_size = xmlObj.documentElement.selectSingleNode("captcha_size").text
	captcha_width_ubound = xmlObj.documentElement.selectSingleNode("captcha_width_ubound").text
	captcha_width_lbound = xmlObj.documentElement.selectSingleNode("captcha_width_lbound").text
	captcha_height_ubound = xmlObj.documentElement.selectSingleNode("captcha_height_ubound").text
	captcha_height_lbound = xmlObj.documentElement.selectSingleNode("captcha_height_lbound").text
	captcha_spacing_lbound = xmlObj.documentElement.selectSingleNode("captcha_spacing_lbound").text
	captcha_spacing_ubound = xmlObj.documentElement.selectSingleNode("captcha_spacing_ubound").text
	captcha_angle_lbound = xmlObj.documentElement.selectSingleNode("captcha_angle_lbound").text
	captcha_angle_ubound = xmlObj.documentElement.selectSingleNode("captcha_angle_ubound").text
	captcha_weight = xmlObj.documentElement.selectSingleNode("captcha_weight").text
	captcha_charshow = xmlObj.documentElement.selectSingleNode("captcha_charshow").text
	captcha_charshow_simple_r = xmlObj.documentElement.selectSingleNode("captcha_charshow_simple_r").text
	captcha_charshow_simple_g = xmlObj.documentElement.selectSingleNode("captcha_charshow_simple_g").text
	captcha_charshow_simple_b = xmlObj.documentElement.selectSingleNode("captcha_charshow_simple_b").text
	captcha_charshow_stepbystep_r = xmlObj.documentElement.selectSingleNode("captcha_charshow_stepbystep_r").text
	captcha_charshow_stepbystep_g = xmlObj.documentElement.selectSingleNode("captcha_charshow_stepbystep_g").text
	captcha_charshow_stepbystep_b = xmlObj.documentElement.selectSingleNode("captcha_charshow_stepbystep_b").text
	captcha_charshow_mix_percent = xmlObj.documentElement.selectSingleNode("captcha_charshow_mix_percent").text
	captcha_backshow = xmlObj.documentElement.selectSingleNode("captcha_backshow").text
	captcha_backshow_simple_r = xmlObj.documentElement.selectSingleNode("captcha_backshow_simple_r").text
	captcha_backshow_simple_g = xmlObj.documentElement.selectSingleNode("captcha_backshow_simple_g").text
	captcha_backshow_simple_b = xmlObj.documentElement.selectSingleNode("captcha_backshow_simple_b").text
	captcha_backshow_stepbystep_r = xmlObj.documentElement.selectSingleNode("captcha_backshow_stepbystep_r").text
	captcha_backshow_stepbystep_g = xmlObj.documentElement.selectSingleNode("captcha_backshow_stepbystep_g").text
	captcha_backshow_stepbystep_b = xmlObj.documentElement.selectSingleNode("captcha_backshow_stepbystep_b").text
	captcha_backshow_mix_percent = xmlObj.documentElement.selectSingleNode("captcha_backshow_mix_percent").text
	captcha_pic_width = xmlObj.documentElement.selectSingleNode("captcha_pic_width").text
	captcha_pic_height = xmlObj.documentElement.selectSingleNode("captcha_pic_height").text
Else
	captcha_chartype_chinese = 1
	captcha_chartype_english = 0
	captcha_chartype_number = 0
	captcha_size = 2
	captcha_width_ubound = 30
	captcha_width_lbound = 25
	captcha_height_ubound = 30
	captcha_height_lbound = 25
	captcha_spacing_lbound = -6
	captcha_spacing_ubound = 0
	captcha_angle_lbound = -15
	captcha_angle_ubound = 15
	captcha_weight = 1
	captcha_charshow = 0
	captcha_charshow_simple_r = 0
	captcha_charshow_simple_g = 0
	captcha_charshow_simple_b = 0
	captcha_charshow_stepbystep_r = 0
	captcha_charshow_stepbystep_g = 0
	captcha_charshow_stepbystep_b = 0
	captcha_charshow_mix_percent = 30
	captcha_backshow = 0
	captcha_backshow_simple_r = 255
	captcha_backshow_simple_g = 255
	captcha_backshow_simple_b = 255
	captcha_backshow_stepbystep_r = 255
	captcha_backshow_stepbystep_g = 255
	captcha_backshow_stepbystep_b = 255
	captcha_backshow_mix_percent = 30
	captcha_pic_width = 60
	captcha_pic_height = 40
End If 
Set xmlObj=Nothing 
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
<tr>
<td height="23"><B>说明</B>：<br>①、使用此功能需要有修改文件的权限。<br>②、这里不能设置验证码显示的页面。设置页面是否显示验证码请到后台基本设置和各版面高级设置中设置。
</td>
</tr>
</table>
<script language="javascript">
<!--
function display_area(str,n,total){
	for (var i=total-1; i>=0; --i){
		document.getElementById(str+'_'+i+'_area').style.display='none';
	}
	document.getElementById(str+'_'+n+'_area').style.display='';
}
function init_select(selected){
    document.writeln('<option value="-1">变</option>');
    for (var i=0; i<=255; ++i){
        document.writeln('<option value="'+i+'"');
        if (selected==i) document.writeln(' selected');
        document.writeln('>'+i);
        document.writeln('</option>');
    }
}
function preview(){
    document.captcha_form.action="?action=savetotemp";
	check_form();
    document.captcha_form.submit();
}
function check_form(){
	var o = document.getElementsByTagName("input");
	for (var i=o.length-1; i>=0; --i){
		if('text'==o[i].getAttribute("type").toLowerCase())
		{
			o[i].value=o[i].value.replace(/[１２３４５６７８９０－]/gi,function(w){
				switch(w){
					case "１":return "1";
					case "２":return "2";
					case "３":return "3";
					case "４":return "4";
					case "５":return "5";
					case "６":return "6";
					case "７":return "7";
					case "８":return "8";
					case "９":return "9";
					case "０":return "0";
					case "－":return "-";
				}
			});
			o[i].value=o[i].value.replace(/[^0123456789-]/gi,'');
		}
	}
}
//-->
</script>
<form name="captcha_form" action="?action=save" method="post" style="margin:0px;padding:0px;" onsubmit="check_form()">
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr>
<th colspan="2">验证码显示设置</th>
</tr>
<tr>
<td width="30%" align="right"><b>显示效果：</b><br/>点预览按扭可以看到当前的设置效果</td>
<td>
	<img src="<%If "preview"=Request("action") Then Response.Write "CaptchaPreview.asp":Else Response.Write "../"&DvCodeFile:End If %>" id="captcha_view_img" />
</td>
</tr>
<tr>
<td width="20%" align="right"><b>字符类型：</b><br/>可以勾选多个类型，建议使用一种类型</td>
<td valign="top">
	<input type="checkbox" id="captcha_chartype_chinese" name="captcha_chartype_chinese" value="1" style="border:none" />
	<label for="captcha_chartype_chinese">汉字</label>
	<input type="checkbox" id="captcha_chartype_english" name="captcha_chartype_english" value="1" style="border:none" />
	<label for="captcha_chartype_english">字母</label>
	<input type="checkbox" id="captcha_chartype_number" name="captcha_chartype_number" value="1" style="border:none" />
	<label for="captcha_chartype_number">数字</label>
</td>
</tr>
<tr>
<td width="20%" align="right"><b>字符数：</b><br/>建议填写2-4个字符</td>
<td valign="top"><input type="text" name="captcha_size" value="<%=captcha_size%>" size="10" /></td>
</tr>
<tr>
<td width="20%" align="right"><b>字符宽度变化范围：</b><br/>由小到大变化，建议填写25至30</td>
<td valign="top"><input type="text" name="captcha_width_lbound" value="<%=captcha_width_lbound%>" size="10" /> -&gt; <input type="text" name="captcha_width_ubound" value="<%=captcha_width_ubound%>" size="10" /></td>
</tr>
<tr>
<td width="20%" align="right"><b>字符高度变化范围：</b><br/>由小到大变化，建议填写25至30</td>
<td valign="top"><input type="text" name="captcha_height_lbound" value="<%=captcha_height_lbound%>" size="10" /> -&gt; <input type="text" name="captcha_height_ubound" value="<%=captcha_height_ubound%>" size="10" /></td>
</tr>
<tr>
<td width="20%" align="right"><b>字符间距变化范围：</b><br/>由小到大变化，建议填写-6至0</td>
<td valign="top"><input type="text" name="captcha_spacing_lbound" value="<%=captcha_spacing_lbound%>" size="10" /> -&gt; <input type="text" name="captcha_spacing_ubound" value="<%=captcha_spacing_ubound%>" size="10" /></td>
</tr>
<tr>
<td width="20%" align="right"><b>字符旋转角度变化范围：</b><br/>建议填写-15至15，可调范围：-360~360</td>
<td valign="top"><input type="text" name="captcha_angle_lbound" value="<%=captcha_angle_lbound%>" size="10" /> -&gt; <input type="text" name="captcha_angle_ubound" value="<%=captcha_angle_ubound%>" size="10" /></td>
</tr>
<tr>
<td width="20%" align="right"><b>字符粗细：</b></td>
<td valign="top">
	<input type="radio" id="captcha_weight_rnd" name="captcha_weight"  value="0" style="border:none" />
	<label for="captcha_weight_rnd">随机</label>
	<input type="radio" id="captcha_weight_thin" name="captcha_weight" value="1" style="border:none" />
	<label for="captcha_weight_thin">细线</label>
	<input type="radio" id="captcha_weight_bold" name="captcha_weight" value="2" style="border:none" />
	<label for="captcha_weight_bold">粗线</label>
</td>
</tr>
<tr>
<td width="20%" align="right"><b>字符显示方式：</b></td>
<td valign="top">
	<input type="radio" id="captcha_charshow_simple" name="captcha_charshow" value="0" style="border:none" onclick="display_area('captcha_charshow',0,3)" />
	<label for="captcha_charshow_simple">纯色</label>
	<input type="radio" id="captcha_charshow_mix" name="captcha_charshow" value="1" style="border:none" onclick="display_area('captcha_charshow',1,3)" />
	<label for="captcha_charshow_mix">杂色</label>
	<input type="radio" id="captcha_chartype_stepbystep" name="captcha_charshow" value="2" style="border:none" onclick="display_area('captcha_charshow',2,3)" />
	<label for="captcha_chartype_stepbystep">渐变</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<span id="captcha_charshow_0_area">红 <select name="captcha_charshow_simple_r"><script language="JavaScript">init_select(<%=captcha_charshow_simple_r%>);</script></select>  绿 <select name="captcha_charshow_simple_g"><script language="JavaScript">init_select(<%=captcha_charshow_simple_g%>);</script></select>  蓝 <select name="captcha_charshow_simple_b"><script language="JavaScript">init_select(<%=captcha_charshow_simple_b%>);</script></select></span><span id="captcha_charshow_1_area">杂点率 <input type="text" name="captcha_charshow_mix_percent" value="<%=captcha_charshow_mix_percent%>" size="10" /> %</span><span id="captcha_charshow_2_area">红 <select name="captcha_charshow_stepbystep_r"><script language="JavaScript">init_select(<%=captcha_charshow_stepbystep_r%>);</script></select>  绿 <select name="captcha_charshow_stepbystep_g"><script language="JavaScript">init_select(<%=captcha_charshow_stepbystep_g%>);</script></select>  蓝 <select name="captcha_charshow_stepbystep_b"><script language="JavaScript">init_select(<%=captcha_charshow_stepbystep_b%>);</script></select></span>
</td>
</tr>
<tr>
<td width="20%" align="right"><b>背景显示方式：</b></td>
<td valign="top">
	<input type="radio" id="captcha_backshow_simple" name="captcha_backshow" value="0" style="border:none" onclick="display_area('captcha_backshow',0,3)" />
	<label for="captcha_backshow_simple">纯色</label>
	<input type="radio" id="captcha_backshow_mix" name="captcha_backshow" value="1" style="border:none" onclick="display_area('captcha_backshow',1,3)" />
	<label for="captcha_backshow_mix">杂色</label>
	<input type="radio" id="captcha_backshow_stepbystep" name="captcha_backshow" value="2" style="border:none" onclick="display_area('captcha_backshow',2,3)" />
	<label for="captcha_backshow_stepbystep">渐变</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<span id="captcha_backshow_0_area">红 <select name="captcha_backshow_simple_r"><script language="JavaScript">init_select(<%=captcha_backshow_simple_r%>);</script></select>  绿 <select name="captcha_backshow_simple_g"><script language="JavaScript">init_select(<%=captcha_backshow_simple_g%>);</script></select>  蓝 <select name="captcha_backshow_simple_b"><script language="JavaScript">init_select(<%=captcha_backshow_simple_b%>);</script></select></span><span id="captcha_backshow_1_area">杂点率 <input type="text" name="captcha_backshow_mix_percent" value="<%=captcha_backshow_mix_percent%>" size="10" /> %</span><span id="captcha_backshow_2_area">红 <select name="captcha_backshow_stepbystep_r"><script language="JavaScript">init_select(<%=captcha_backshow_stepbystep_r%>);</script></select>  绿 <select name="captcha_backshow_stepbystep_g"><script language="JavaScript">init_select(<%=captcha_backshow_stepbystep_g%>);</script></select>  蓝 <select name="captcha_backshow_stepbystep_b"><script language="JavaScript">init_select(<%=captcha_backshow_stepbystep_b%>);</script></select></span>
</td>
</tr>
<tr>
<td width="20%" align="right"><b>图片宽度：</b><br/>建议填写为字符数与字符宽度平均值的乘积</td>
<td valign="top"><input type="text" name="captcha_pic_width" value="<%=captcha_pic_width%>" size="10" /></td>
</tr>
<tr>
<td width="20%" align="right"><b>图片高度：</b><br/>建议填写为略大于字符最大高度</td>
<td valign="top"><input type="text" name="captcha_pic_height" value="<%=captcha_pic_height%>" size="10" /></td>
</tr>
<tr>
<td width="20%" align="right"></td>
<td>
<input type="submit" name="sub1" value=" 提交保存 " /> <input type="button" name="sub1" value=" 预 览 " onclick="preview()" /></td>
</tr>
</table>
</form>
<script language="javascript">
<!--
document.captcha_form.captcha_chartype_chinese.checked=1==<%=captcha_chartype_chinese%>;
document.captcha_form.captcha_chartype_english.checked=1==<%=captcha_chartype_english%>;
document.captcha_form.captcha_chartype_number.checked=1==<%=captcha_chartype_number%>;
document.captcha_form.captcha_weight[<%=captcha_weight%>].checked=true;
document.captcha_form.captcha_charshow[<%=captcha_charshow%>].checked=true;
document.captcha_form.captcha_backshow[<%=captcha_backshow%>].checked=true;
display_area('captcha_charshow',<%=captcha_charshow%>,3);
display_area('captcha_backshow',<%=captcha_backshow%>,3);
//-->
</script>
<%
End Sub
%>

