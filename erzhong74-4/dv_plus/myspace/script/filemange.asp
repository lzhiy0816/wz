<!--#include file="../../../Conn.asp"-->
<%
MyDbPath = "../../../"
%>
<!-- #include file="../../../inc/const.asp" -->
<!-- #include file="Dv_ClsSpace.asp" -->
<!-- #include File="../../../inc/Upload_Class.asp" -->
<%
Dvbbs.ScriptPath="../../../"
Dvbbs.LoadTemplates("")
Dim MySpace
Set MySpace = New Cls_Space
Dvbbs.Stats = "风格模板-文件管理"
Dvbbs.Head()
Page_Main()
Set MySpace = Nothing
Set Dvbbs = Nothing


Sub Page_Main()
	If Dvbbs.Userid=0 Then
		Response.write "你还未登陆！"
		Exit Sub
	End If
	If MySpace.Sid=0 Then
		Exit Sub
	End If
	If Cint(Dvbbs.GroupSetting(7))=0 then
		Response.write "您没有在本论坛上传文件的权限"
		Exit Sub
	End If
	Dim Fso,Folder,Skin_Path,Folderhrf
	Dim Files,File,GetFolder,FileName
	Dim CanEdit
	CanEdit = False
	Folder = MySpace.Space_Skinpath & MySpace.Space_Info.getAttribute("s_path")
	Folderhrf = Folder
	If Instr(Folder,"/userskins/skin_"&Dvbbs.Userid&"/") or Dvbbs.Master Then
		CanEdit = True
	Else
		CanEdit = False
	End If
	Folder = Server.MapPath(Folder)
	If Right(Folder,1)<>"\" Then Folder=Folder&"\"
	Dim PostRanNum
	If Request.QueryString("act")="" Then
		Randomize
		PostRanNum = Int(900*rnd)+1000
		Session("UploadCode") = Cstr(PostRanNum)
	End If
%>
<style>
a {text-decoration : none;color : #000000; } 
a:hover {text-decoration : underline; color : #559AE4; } 
body {
	text-align:center;margin-top:0;
} 
body,ol,ul,div,td{
	font-size : 12px; color : #000000; font-family : tahoma, 宋体, fantasy; 
}
form{margin:0;}
div.skindiv{
	margin:0;
	width:100%;
}
div.skindiv .readme{
	width:96%;
	background-color:#FFFFC1;
	border:1px dashed #569CFC;
	color:#000;
}

div.skindiv th{
	color:white;
	background-color:#569CFC;
	border-bottom:4px solid #9FC7FD;
}
div.skindiv table{
	border:1px solid #9FC7FD;
}
div.skindiv td{
	border-bottom:1px solid #9FC7FD;
}
input , select , textarea , option {
	font-family : tahoma, verdana, 宋体, fantasy; font-size : 12px;line-height : 15px; color : #000000;
} 
input { vertical-align:middle;border:1px solid #98BBD7; background:#EFF6FB; height:18px;}
select{margin-top:6px;}
.chkbox, .radio {border: 0px;background: none;vertical-align: middle; }
.button {border:1px solid #98BBD7; background:#EFF6FB; height:22px; color:#000000}
ol{
	margin-top:5px;
}
ul.skin{
	margin:0 auto;
	clear:both;
	width:98%;
	padding:0;
	list-style:none;
	
}
ul.skin li{
	margin:0px 5px 0px 0px;
	padding:0;
	float:left;
}
img.demogif{
	width:60px;
	height:60px;
	border:none;
	cursor : pointer;

}
hr {
height:0px;border :0px;border-top: gray 1px solid;width : 98%; 
} 

</style>
<script language="JavaScript" src="<%=MyDbPath%>inc/Pagination.js"></script>
<script language="JavaScript">
<!--
function viewfile(obj,putdiv)
{
	var fileext=obj.value.substring(obj.value.lastIndexOf("."),obj.value.length)
	fileext=fileext.toLowerCase()
        if ((fileext=='.jpg')||(fileext=='.gif')||(fileext=='.jpeg')||(fileext=='.png')||(fileext=='.bmp'))
        {
        //alert(''+document.form1.UpFile.value)//把这里改成预览图片的语句
		document.getElementById(putdiv).innerHTML="<img src='"+obj.value+"' width='40' height='40' border='0'/>"
		uploadframe(25)
        }
}
//-->
</script>
<div class="skindiv">
	<table cellpadding="3" cellspacing="3" class="readme">
	<tr>
	<td>
		<b>风格模板-文件管理</b>
		<ol>
		<li>可以添加gif,jpg,png,swf,jpeg,bmp到个人空间;</li>
		<li>单个文件不能超过100K;</li>
		<li>文件数限制为50个；</li>
		</ol>
	</td>
	</tr>
	<%If CanEdit = True Then%>
	<form method="post" name="skinform" action="filemange.asp?act=upload" enctype="multipart/form-data">
	<tr>
	<td>
	上传：<span id="preview_1"></span>
	&nbsp;&nbsp;<input type="file" name="file1" width="200" value="" size="40"  onchange="viewfile(this,'preview_1');"/>
	&nbsp;&nbsp;<input type="submit" name="submit" value="上传"/><INPUT TYPE="hidden" NAME="UploadCode" value="<%=PostRanNum%>"/>
	</td>
	</tr>
	</form>
	<%End If%>
	</table>
</div>
<div class="skindiv">
<table cellpadding="3" cellspacing="3" width="96%" align="center">
<tr>
<th width="15%">类型</th>
<th width="15%">图例</th>
<th width="20%">文件名称</th>
<th width="20%">修改时间</th>
<th width="10%">操作</th>
</tr>
<%
Dim FormName,Upload
'On error resume Next
Set FSO = Dvbbs.iCreateObject("Scripting.FileSystemObject")
If Err Then
	Response.Write "<tr><td colspan=""5"">该空间暂不支持FSO组件，因此不能对文件进行管理！</td></tr>"
	Err.Clear
Else
	If Fso.FolderExists(Folder)=False Then
		Skin_Path = MySpace.CreatStylePath
		MySpace.CopyFolder MySpace.Space_Info.getAttribute("s_path"),Skin_Path
	End If
	Set GetFolder = Fso.GetFolder(Folder)
	Set Files = GetFolder.Files
	'上传文件

	If Request.QueryString("act")="upload" and CanEdit Then
		If Files.Count<50 Then
		Set Upload = New UpFile_Cls
		Upload.UploadType			= Cint(Dvbbs.Forum_UploadSetting(2))	'设置上传组件类型
		Upload.UploadPath			= Folderhrf								'设置上传路径
		Upload.InceptFileType		= "gif,jpg,png,swf,jpeg,bmp"		'设置上传文件限制
		Upload.MaxSize				= 100								'单位 KB
		Upload.InceptMaxFile		= 5									'每次上传文件个数上限
		Upload.ReName				= 1									'以默认原来文件名保存
		'Upload.ChkSessionName		= "UploadCode"						'防止重复提交，SESSION名与提交的表单要一致
		'执行上传
		Upload.SaveUpFile
		If Upload.ErrCodes<>0 Then
			Response.Write "<tr><td colspan=""5"">"& Upload.Description & "[ <a href=""?"">重新上传</a> ]</td></tr>"
		Else
			If Upload.Count > 0 Then
				For Each FormName In Upload.UploadFiles
					Set File = Upload.UploadFiles(FormName)
					Response.Write "<tr><td colspan=""5"">文件："&File.FileName&"已上传！</td></tr>"
				Next
			Else
				Response.Write "<tr><td colspan=""5"">"& Upload.Description & "[ 请正确选择要上传的文件<a href=""?"">重新上传</a> ]</td></tr>"
			End If
		End If
		Set Upload = Nothing
		Else
			Response.Write "<tr><td colspan=""5"">文件已超出限制，不允许上传！</td></tr>"
		End if
	End If
	'删除文件
	If Request.QueryString("act")="del" and CanEdit Then
		FileName = Request.QueryString("fn")
		FileName = Replace(FileName,".asp","")
		FileName = Replace(FileName,"..","")
		FileName = Replace(FileName,Chr(0),"")
		FileName = Replace(FileName,"/","")
		FileName = Replace(FileName,"\","")
		If Fso.FileExists(Folder&FileName) Then
			Fso.DeleteFile(Folder&FileName)
			Response.Write "<tr><td colspan=""5"">文件："&FileName&"已删除！</td></tr>"
		End If
	End If
	
	
	'Files.Count 
	For Each File in Files
		Response.Write "<tr>"
		Response.Write "<td align=""center"">"&File.type&"</td>"
		
		If Lcase(Mid(File.name,InStrRev(File.name, ".")+1)) = "css" Then
			Response.Write "<td  align=""center""><img src=""../drag/txt.gif""/></td>"
		Else
			Response.Write "<td align=""center""><img src="""&Folderhrf&File.name&""" class=""demogif"" onclick=""window.open(this.src)""/></td>"
		End If
		Response.Write "<td>"&File.name&"</td>"
		Response.Write "<td>"&File.DateLastModified&"</td>"
		Response.Write "<td align=""center"">"
		If CanEdit Then
			Response.Write "<button class=""button"" onclick=""window.location='?act=del&fn="&File.name&"'"">删除</button>"
		End If
		Response.Write "</td>"
		Response.Write "</tr>"

	Next
	Set Files = Nothing
	Set GetFolder = Nothing
End If
Set Fso = Nothing
%>
</table>
</div>

<%
End Sub


%>