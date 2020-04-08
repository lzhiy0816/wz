<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<%
Head()
Dim TestConn
dim action
dim admin_flag
action=trim(request("action"))
dim dbpath,bkfolder,bkdbname,fso,fso1
Dim uploadpath
Dim okOS,okCpus,okCPU
If Dvbbs.Forum_Setting(76)="0" Or  Dvbbs.Forum_Setting(76)="" Then Dvbbs.Forum_Setting(76)="UploadFile/"
uploadpath="../"&Dvbbs.Forum_Setting(76)
select case action
case "CompressData"		'压缩数据
	admin_flag=",30,"
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=../admin_login.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		dvbbs_error()
	else
		dim tmprs
		dim allarticle
		dim Maxid
		dim topic,username,dateandtime,body
		call CompressData()
	end if	
case "BackupData"		'备份数据
	admin_flag=",31,"
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=../admin_login.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		dvbbs_error()
	else
		if request("act")="Backup" then
			call updata()
		else
			call BackupData()
		end if
	end if
case "RestoreData"		'恢复数据
	admin_flag=",32,"
	dim backpath
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=../admin_login.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		dvbbs_error()
	else
		if request("act")="Restore" then
			Dbpath=request.form("Dbpath")
			backpath=request.form("backpath")
			if dbpath="" then
			response.write "请输入您要恢复成的数据库全名"	
			else
			Dbpath=server.mappath(Dbpath)
			end if
			backpath=server.mappath(backpath)
			
			Set TestConn = Server.CreateObject("ADODB.Connection")
			On Error Resume Next
			TestConn.open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Dbpath
			If Err Then
				Response.Write Err.Description
				Err.Clear
				Set TestConn = Nothing
				Response.Write "备份的文件并非合法的数据库。"
				Response.End 
			Else
				Set TestConn = Nothing
			End If
			Set Fso=server.createobject("scripting.filesystemobject")
			if fso.fileexists(dbpath) then  					
			fso.copyfile Dbpath,Backpath
			response.write "成功恢复数据！"
			else
			response.write "备份目录下并无您的备份文件！"	
			end if
		else
			call RestoreData()
		end if
	end if
case "SpaceSize"		'系统空间占用
	admin_flag=",33,"
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li> 本页面为管理员专用，请<a href=../admin_login.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		dvbbs_error()
	else
		call SpaceSize()
	end if
case else
		Errmsg=ErrMsg + "<BR><li>选取相应的操作。"
		dvbbs_error()
end select

Footer()
response.write"</body></html>"


'====================系统空间占用=======================
sub SpaceSize()
On error resume next
GetSysInfo()
Dim t
't = GetAllSpace
Dim FoundFso
FoundFso = False
FoundFso = IsObjInstalled("Scripting.FileSystemObject")
%>
<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder"><tr>
<th height=25 colspan=5>
&nbsp;&nbsp;系统信息检测情况
</th>
</tr>
<tr>
<td class="forumrow" width="35%" height=23>
当前论坛版本
</td>
<td class="forumrow" width="15%">
<a href="http://www.dvbbs.net/download.asp" target=_blank>Dvbbs <%=Dvbbs.Forum_Version%></a>
</td>
<td width="8" class="forumrow"></td>
<td class="forumrow" width="35%">
数据库类型：
</td>
<td class="forumrow" width="15%">
<%
If IsSqlDataBase = 1 Then
	Response.Write "Sql Server"
Else
	Response.Write "Access"
End If
%>
</td>
</tr>
<tr>
<td class="forumRowHighlight" width="35%" height=23>
服务器名和IP
</td>
<td class="forumRowHighlight" width="15%">
<%=Request.ServerVariables("SERVER_NAME")%><BR><%=Request.ServerVariables("LOCAL_ADDR")%>
</td>
<td width="8" class="forumRowHighlight"></td>
<td class="forumRowHighlight" width="35%">
数据库占用空间
</td>
<td class="forumRowHighlight" width="15%">
<%
If IsSqlDataBase = 1 Then
	Set Rs=Dvbbs.Execute("Exec sp_spaceused")
	If Err <> 0 Then
		Err.Clear
		Response.Write "<font color=gray>未知</font>"
	Else
		Response.Write Rs(1)
	End If
Else
	If FoundFso Then
		Response.Write GetFileSize(MyDbPath & DB)
	Else
		Response.Write "<font color=gray>未知</font>"
	End If
End If
%>
</td>
</tr>
<tr>
<td class="forumrow" width="35%" height=23>
上传头像占用空间
</td>
<td class="forumrow" width="15%">
<%showSpaceinfo("../uploadface")%>
</td>
<td width="8" class="forumrow"></td>
<td class="forumrow" width="35%">
上传图片占用空间
</td>
<td class="forumrow" width="15%">
<%showSpaceinfo(uploadpath)%>
</td>
</tr>
<tr>
<td class="forumRowHighlight" width="100%" height=23 colspan=5>
<B>服务器相关信息</B>
</td>
</tr>
<tr>
<td class="forumrow" width="35%" height=23>
ASP脚本解释引擎
</td>
<td class="forumrow" width="15%">
<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %>
</td>
<td width="8" class="forumrow"></td>
<td class="forumrow" width="35%">
IIS 版本
</td>
<td class="forumrow" width="15%">
<%=Request.ServerVariables("SERVER_SOFTWARE")%>
</td>
</tr>
<tr>
<td class="forumRowHighlight" width="35%" height=23>
服务器操作系统
</td>
<td class="forumRowHighlight" width="15%">
<%=okos%>
</td>
<td width="8" class="forumRowHighlight"></td>
<td class="forumRowHighlight" width="35%">
服务器CPU数量
</td>
<td class="forumRowHighlight" width="15%">
<%=okcpus%> 个
</td>
</tr>
<tr>
<td class="forumrow" width="100%" height=23 colspan=5>
本文件路径：<%=Server.Mappath("data.asp")%>
</td>
</tr>
<tr>
<td class="forumRowHighlight" width="100%" colspan=5 height=23>
<B>主要组件信息</B>
</td>
</tr>
<tr>
<td class="forumrow" width="35%" height=23>
FSO文件读写
</td>
<td class="forumrow" width="15%">
<%
If FoundFso Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
</td>
<td width="8" class="forumrow"></td>
<td class="forumrow" width="35%">
Jmail发送邮件支持
</td>
<td class="forumrow" width="15%">
<%
If IsObjInstalled("JMail.SmtpMail") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
</td>
</tr>
<tr>
<td class="forumRowHighlight" width="35%" height=23>
CDONTS发送邮件支持
</td>
<td class="forumRowHighlight" width="15%">
<%
If IsObjInstalled("CDONTS.NewMail") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
</td>
<td width="8" class="forumRowHighlight"></td>
<td class="forumRowHighlight" width="35%">
AspEmail发送邮件支持
</td>
<td class="forumRowHighlight" width="15%">
<%
If IsObjInstalled("Persits.MailSender") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
</td>
</tr>
<tr>
<td class="forumrow" width="35%" height=23>
无组件上传支持
</td>
<td class="forumrow" width="15%">
<%
If IsObjInstalled("Adodb.Stream") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
</td>
<td width="8" class="forumrow"></td>
<td class="forumrow" width="35%">
AspUpload上传支持
</td>
<td class="forumrow" width="15%">
<%
If IsObjInstalled("Persits.Upload") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
</td>
</tr>
<tr>
<td class="forumRowHighlight" width="35%" height=23>
SA-FileUp上传支持
</td>
<td class="forumRowHighlight" width="15%">
<%
If IsObjInstalled("SoftArtisans.FileUp") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
</td>
<td width="8" class="forumRowHighlight"></td>
<td class="forumRowHighlight" width="35%">
DvFile-Up上传支持
</td>
<td class="forumRowHighlight" width="15%">
<%
If IsObjInstalled("DvFile.Upload") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
</td>
</tr>
<tr>
<td class="forumrow" width="35%" height=23>
CreatePreviewImage生成预览图片
</td>
<td class="forumrow" width="15%">
<%
If IsObjInstalled("CreatePreviewImage.cGvbox") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
</td>
<td width="8" class="forumrow"></td>
<td class="forumrow" width="35%">
AspJpeg生成预览图片
</td>
<td class="forumrow" width="15%">
<%
If IsObjInstalled("Persits.Jpeg") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
</td>
</tr>
<tr>
<td class="forumRowHighlight" width="35%" height=23>
SA-ImgWriter生成预览图片
</td>
<td class="forumRowHighlight" width="15%">
<%
If IsObjInstalled("SoftArtisans.ImageGen") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%>
</td>
<td width="8" class="forumRowHighlight"></td>
<td class="forumRowHighlight" width="35%">ADO(数据库访问)版本:<%=conn.Version%>
</td>
<td class="forumRowHighlight" width="15%"><font color=green><b>√</b></font>
</td>
</tr>
<FORM action="data.asp?action=SpaceSize" method=post id=form1 name=form1>
<tr>
<td class="forumrow" width="100%" height=23 colspan=5>
<%
If Request("classname")<>"" Then
	If IsObjInstalled(Request("classname")) Then
		Response.Write "<font color=green><b>恭喜，本服务器支持 "&Request("classname")&" 组件</b></font><BR>"
	Else
		Response.Write "<font color=red><b>抱歉，本服务器不支持 "&Request("classname")&" 组件</b></font><BR>"
	End If
End If
%>
其它组件支持情况查询：<input class=input type=text value="" name="classname" size=30>
<INPUT type=submit value="查 询" id=submit1 name=submit1>
输入组件的 ProgId 或 ClassId
</td>
</tr>
</form>
</table>
<%Response.Flush%>
<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
<tr>
<td class="forumRowHighlight" width="100%" colspan=5 height=23>
<B>磁盘文件操作速度测试</B>
</td>
</tr>
<tr>
<td class="forumrow" width="100%" colspan=5 height=23>
<%
	Response.Write "正在重复创建、写入和删除文本文件50次..."

	Dim thetime3,tempfile,iserr,t1,FsoObj,tempfileOBJ,t2
	Set FsoObj=Server.CreateObject("Scripting.FileSystemObject")

	iserr=False
	t1=timer
	tempfile=server.MapPath("./") & "\aspchecktest.txt"
	For i=1 To 50
		Err.Clear

		Set tempfileOBJ = FsoObj.CreateTextFile(tempfile,true)
		If Err <> 0 Then
			Response.Write "创建文件错误！"
			iserr=True
			Err.Clear
			Exit For
		End If
		tempfileOBJ.WriteLine "Only for test. Ajiang ASPcheck"
		If Err <> 0 Then
			Response.Write "写入文件错误！"
			iserr=True
			Err.Clear
			Exit For
		End If
		tempfileOBJ.close
		Set tempfileOBJ = FsoObj.GetFile(tempfile)
		tempfileOBJ.Delete 
		If Err <> 0 Then
			Response.Write "删除文件错误！"
			iserr=True
			Err.Clear
			Exit For
		end if
		Set tempfileOBJ=Nothing
	Next
	t2=timer
	If Not iserr Then
		thetime3=cstr(int(( (t2-t1)*10000 )+0.5)/10)
		Response.Write "...已完成！本服务器执行此操作共耗时 <font color=red>" & thetime3 & " 毫秒</font>"
	End If
%>
</td>
</tr>
<tr>
<td class="forumRowHighlight" width="100%" height=23 colspan=5>
<a href="http://www.aspsky.cn" target=_blank>动网科技虚拟主机 <font color=gray>双至强2.4,2GddrEcc,SCSI36.4G*2</font> 执行此操作需要 <font color=red>32～65</font> 毫秒</a>
</td>
</tr>
</table>
<%Response.Flush%>
<table border="0" cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
<tr>
<td class="forumrow" width="100%" colspan=5 height=23>
<B>ASP脚本解释和运算速度测试</B>
</td>
</tr>
<tr>
<td class="forumRowHighlight" width="100%" colspan=5 height=23>
<%

	Response.Write "整数运算测试，正在进行50万次加法运算..."
	dim lsabc,thetime,thetime2
	t1=timer
	for i=1 to 500000
		lsabc= 1 + 1
	next
	t2=timer
	thetime=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...已完成！共耗时 <font color=red>" & thetime & " 毫秒</font><br>"


	Response.Write "浮点运算测试，正在进行20万次开方运算..."
	t1=timer
	for i=1 to 200000
		lsabc= 2^0.5
	next
	t2=timer
	thetime2=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...已完成！共耗时 <font color=red>" & thetime2 & " 毫秒</font><br>"
%>
</td>
</tr>
<tr>
<td class="forumrow" width="100%" colspan=5 height=23>
<a href="http://www.aspsky.cn" target=_blank>动网科技虚拟主机 <font color=gray>双至强2.4,2GddrEcc,SCSI36.4G*2</font> 整数运算需要 <font color=red>171～203</font> 毫秒, 浮点运算需要 <font color=red>156～171</font> 毫秒</a>

</td>
</tr>
</table><BR>
<%
end sub

sub SQLUserReadme()
%>
		<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">			<tr>
  					<th height=25>
  					&nbsp;&nbsp;SQL数据库数据处理说明
  					</th>
  				</tr> 	
 				<tr>
 					<td class="forumrow"> 			
 			<blockquote>
<B>一、备份数据库</B>
<BR><BR>
1、打开SQL企业管理器，在控制台根目录中依次点开Microsoft SQL Server<BR>
2、SQL Server组-->双击打开你的服务器-->双击打开数据库目录<BR>
3、选择你的数据库名称（如论坛数据库Forum）-->然后点上面菜单中的工具-->选择备份数据库<BR>
4、备份选项选择完全备份，目的中的备份到如果原来有路径和名称则选中名称点删除，然后点添加，如果原来没有路径和名称则直接选择添加，接着指定路径和文件名，指定后点确定返回备份窗口，接着点确定进行备份
<BR><BR>
<B>二、还原数据库</B><BR><BR>
1、打开SQL企业管理器，在控制台根目录中依次点开Microsoft SQL Server<BR>
2、SQL Server组-->双击打开你的服务器-->点图标栏的新建数据库图标，新建数据库的名字自行取<BR>
3、点击新建好的数据库名称（如论坛数据库Forum）-->然后点上面菜单中的工具-->选择恢复数据库<BR>
4、在弹出来的窗口中的还原选项中选择从设备-->点选择设备-->点添加-->然后选择你的备份文件名-->添加后点确定返回，这时候设备栏应该出现您刚才选择的数据库备份文件名，备份号默认为1（如果您对同一个文件做过多次备份，可以点击备份号旁边的查看内容，在复选框中选择最新的一次备份后点确定）-->然后点击上方常规旁边的选项按钮<BR>
5、在出现的窗口中选择在现有数据库上强制还原，以及在恢复完成状态中选择使数据库可以继续运行但无法还原其它事务日志的选项。在窗口的中间部位的将数据库文件还原为这里要按照你SQL的安装进行设置（也可以指定自己的目录），逻辑文件名不需要改动，移至物理文件名要根据你所恢复的机器情况做改动，如您的SQL数据库装在D:\Program Files\Microsoft SQL Server\MSSQL\Data，那么就按照您恢复机器的目录进行相关改动改动，并且最后的文件名最好改成您当前的数据库名（如原来是bbs_data.mdf，现在的数据库是forum，就改成forum_data.mdf），日志和数据文件都要按照这样的方式做相关的改动（日志的文件名是*_log.ldf结尾的），这里的恢复目录您可以自由设置，前提是该目录必须存在（如您可以指定d:\sqldata\bbs_data.mdf或者d:\sqldata\bbs_log.ldf），否则恢复将报错<BR>
6、修改完成后，点击下面的确定进行恢复，这时会出现一个进度条，提示恢复的进度，恢复完成后系统会自动提示成功，如中间提示报错，请记录下相关的错误内容并询问对SQL操作比较熟悉的人员，一般的错误无非是目录错误或者文件名重复或者文件名错误或者空间不够或者数据库正在使用中的错误，数据库正在使用的错误您可以尝试关闭所有关于SQL窗口然后重新打开进行恢复操作，如果还提示正在使用的错误可以将SQL服务停止然后重起看看，至于上述其它的错误一般都能按照错误内容做相应改动后即可恢复<BR><BR>

<B>三、收缩数据库</B><BR><BR>
一般情况下，SQL数据库的收缩并不能很大程度上减小数据库大小，其主要作用是收缩日志大小，应当定期进行此操作以免数据库日志过大<BR>
1、设置数据库模式为简单模式：打开SQL企业管理器，在控制台根目录中依次点开Microsoft SQL Server-->SQL Server组-->双击打开你的服务器-->双击打开数据库目录-->选择你的数据库名称（如论坛数据库Forum）-->然后点击右键选择属性-->选择选项-->在故障还原的模式中选择“简单”，然后按确定保存<BR>
2、在当前数据库上点右键，看所有任务中的收缩数据库，一般里面的默认设置不用调整，直接点确定<BR>
3、<font color=blue>收缩数据库完成后，建议将您的数据库属性重新设置为标准模式，操作方法同第一点，因为日志在一些异常情况下往往是恢复数据库的重要依据</font>
<BR><BR>

<B>四、设定每日自动备份数据库</B><BR><BR>
<font color=red>强烈建议有条件的用户进行此操作！</font><BR>
1、打开企业管理器，在控制台根目录中依次点开Microsoft SQL Server-->SQL Server组-->双击打开你的服务器<BR>
2、然后点上面菜单中的工具-->选择数据库维护计划器<BR>
3、下一步选择要进行自动备份的数据-->下一步更新数据优化信息，这里一般不用做选择-->下一步检查数据完整性，也一般不选择<BR>
4、下一步指定数据库维护计划，默认的是1周备份一次，点击更改选择每天备份后点确定<BR>
5、下一步指定备份的磁盘目录，选择指定目录，如您可以在D盘新建一个目录如：d:\databak，然后在这里选择使用此目录，如果您的数据库比较多最好选择为每个数据库建立子目录，然后选择删除早于多少天前的备份，一般设定4－7天，这看您的具体备份要求，备份文件扩展名一般都是bak就用默认的<BR>
6、下一步指定事务日志备份计划，看您的需要做选择-->下一步要生成的报表，一般不做选择-->下一步维护计划历史记录，最好用默认的选项-->下一步完成<BR>
7、完成后系统很可能会提示Sql Server Agent服务未启动，先点确定完成计划设定，然后找到桌面最右边状态栏中的SQL绿色图标，双击点开，在服务中选择Sql Server Agent，然后点击运行箭头，选上下方的当启动OS时自动启动服务<BR>
8、这个时候数据库计划已经成功的运行了，他将按照您上面的设置进行自动备份
<BR><BR>
修改计划：<BR>
1、打开企业管理器，在控制台根目录中依次点开Microsoft SQL Server-->SQL Server组-->双击打开你的服务器-->管理-->数据库维护计划-->打开后可看到你设定的计划，可以进行修改或者删除操作
<BR><BR>
<B>五、数据的转移（新建数据库或转移服务器）</B><BR><BR>
一般情况下，最好使用备份和还原操作来进行转移数据，在特殊情况下，可以用导入导出的方式进行转移，这里介绍的就是导入导出方式，导入导出方式转移数据一个作用就是可以在收缩数据库无效的情况下用来减小（收缩）数据库的大小，本操作默认为您对SQL的操作有一定的了解，如果对其中的部分操作不理解，可以咨询动网相关人员或者查询网上资料<BR>
1、将原数据库的所有表、存储过程导出成一个SQL文件，导出的时候注意在选项中选择编写索引脚本和编写主键、外键、默认值和检查约束脚本选项<BR>
2、新建数据库，对新建数据库执行第一步中所建立的SQL文件<BR>
3、用SQL的导入导出方式，对新数据库导入原数据库中的所有表内容<BR>
 			</blockquote> 	
 					</td>
 				</tr>
 			</table>
<%
end sub

'====================恢复数据库=========================
sub RestoreData()
If IsSqlDataBase = 1 Then
	SQLUserReadme()
	Exit Sub
End If
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder"		<tr>
	<th height=25 >
   					&nbsp;&nbsp;<B>恢复论坛数据</B>( 需要FSO支持，FSO相关帮助请看微软网站 )
  					</th>
  				</tr>
				<form method="post" action="data.asp?action=RestoreData&act=Restore">
  				
  				<tr>
  					<td height=100 class="forumrow">
  						&nbsp;&nbsp;备份数据库路径(相对)：<input type=text size=30 name=DBpath value="../DataBackup/dvbbs7_Backup.MDB">&nbsp;&nbsp;<BR>
  						&nbsp;&nbsp;目标数据库路径(相对)：<input type=text size=30 name=backpath value="../<%=db%>"><BR>&nbsp;&nbsp;填写您当前使用的数据库路径，如不想覆盖当前文件，可自行命名（注意路径是否正确），然后修改conn.asp文件，如果目标文件名和当前使用数据库名一致的话，不需修改conn.asp文件<BR>
						&nbsp;&nbsp;<input type=submit value="恢复数据"> <br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;在上面填写本程序的数据库路径全名，本程序的默认备份数据库文件为DataBackup\dvbbs_Backup.MDB，请按照您的备份文件自行修改。<br>
  						&nbsp;&nbsp;您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>
  						&nbsp;&nbsp;注意：所有路径都是相对与程序空间根目录的相对路径</font>
  					</td>
  				</tr>	
  				</form>
  			</table>
<%
end sub

'====================备份数据库=========================
sub BackupData()
If IsSqlDataBase = 1 Then
	SQLUserReadme()
	Exit Sub
End If
%>
	<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  				<tr>
  					<th height=25 >
  					&nbsp;&nbsp;<B>备份论坛数据</B>( 需要FSO支持，FSO相关帮助请看微软网站 )
  					</th>
  				</tr>
  				<form method="post" action="data.asp?action=BackupData&act=Backup">
  				<tr>
  					<td height=100 class="forumrow">
  						&nbsp;&nbsp;
						当前数据库路径(相对路径)：<input type=text size=15 name=DBpath value="../<%=db%>"><BR>&nbsp;&nbsp;
						备份数据库目录(相对路径)：<input type=text size=15 name=bkfolder value="../Databackup">&nbsp;如目录不存在，程序将自动创建<BR>&nbsp;&nbsp;
						备份数据库名称(填写名称)：<input type=text size=15 name=bkDBname value="dvbbs7_Backup.mdb">&nbsp;如备份目录有该文件，将覆盖，如没有，将自动创建<BR>
						&nbsp;&nbsp;<input type=submit value="确定"><br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;在上面填写本程序的数据库路径全名，本程序的默认数据库文件为Data\dvbbs7.MDB，<B>请一定不能用默认名称命名备份数据库</B><br>
  						&nbsp;&nbsp;您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>
  						&nbsp;&nbsp;注意：所有路径都是相对与程序空间根目录的相对路径				</font>
  					</td>
  				</tr>	
  				</form>
  			</table>
<%
end sub

sub updata()
		On error resume next
		Dim FileConnStr,Fileconn
		Dbpath=request.Form("Dbpath")
		Dbpath=server.mappath(Dbpath)
		bkfolder=request.Form("bkfolder")
		bkdbname=request.Form("bkdbname")

		FileConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Dbpath
		Set Fileconn = Server.CreateObject("ADODB.Connection")
		Fileconn.open FileConnStr
		If Err Then
			Response.Write Err.Description
			Err.Clear
			Set Fileconn = Nothing
			Response.Write "备份的文件并非合法的数据库。"
			Exit Sub
		Else
			Set Fileconn = Nothing
		End If
		Set Fso=server.createobject("scripting.filesystemobject")
		If Fso.fileexists(dbpath) then
			If CheckDir(bkfolder) = True Then
				Fso.copyfile dbpath,bkfolder& "\"& bkdbname
			else
				MakeNewsDir bkfolder
				Fso.copyfile dbpath,bkfolder& "\"& bkdbname
			end if
			response.write "备份数据库成功，您备份的数据库路径为" &bkfolder& "\"& bkdbname
		Else
			response.write "找不到您所需要备份的文件。"
		End if
end sub
'------------------检查某一目录是否存在-------------------
Function CheckDir(FolderPath)
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '存在
       CheckDir = True
    Else
       '不存在
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------根据指定名称生成目录-----------------------
Function MakeNewsDir(foldername)
	dim f
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function


'====================压缩数据库 =========================
sub CompressData()

If IsSqlDataBase = 1 Then
	SQLUserReadme()
	Exit Sub
End If
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
<form action="data.asp?action=CompressData" method="post">
<tr>
<td class="forumrow" height=25><b>注意：</b><br>输入数据库所在相对路径,并且输入数据库名称（正在使用中数据库不能压缩，请选择备份数据库进行压缩操作） </td>
</tr>
<tr>
<td class="forumrow">压缩数据库：<input type="text" name="dbpath" value="../Databackup/dvbbs7_Backup.MDB">&nbsp;
<input type="submit" value="开始压缩"></td>
</tr>
<tr>
<td class="forumrow"><input type="checkbox" name="boolIs97" value="True">如果使用 Access 97 数据库请选择
(默认为 Access 2000 数据库)<br><br></td>
</tr>
<form>
</table>
<%
dim dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")

If dbpath <> "" Then
dbpath = server.mappath(dbpath)
	response.write(CompactDB(dbpath,boolIs97))
End If

end sub

'=====================压缩参数=========================
Function CompactDB(dbPath, boolIs97)
Dim fso, Engine, strDBPath,JET_3X
strDBPath = left(dbPath,instrrev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(dbPath) Then
	fso.CopyFile dbpath,strDBPath & "temp.mdb"
	Set Engine = CreateObject("JRO.JetEngine")

	If boolIs97 = "True" Then
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb;" _
		& "Jet OLEDB:Engine Type=" & JET_3X
	Else
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb"
	End If

fso.CopyFile strDBPath & "temp1.mdb",dbpath
fso.DeleteFile(strDBPath & "temp.mdb")
fso.DeleteFile(strDBPath & "temp1.mdb")
Set fso = nothing
Set Engine = nothing

	CompactDB = "你的数据库, " & dbpath & ", 已经压缩成功!" & vbCrLf

Else
	CompactDB = "数据库名称或路径不正确. 请重试!" & vbCrLf
End If

End Function


'=====================系统空间参数=========================
Sub ShowSpaceInfo(drvpath)
	dim fso,d,size,showsize
	set fso=server.createobject("scripting.filesystemobject") 		
	drvpath=server.mappath(drvpath) 		 		
	set d=fso.getfolder(drvpath) 		
	size=d.size
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write "<font face=verdana>" & showsize & "</font>"
End Sub	
 	
Sub Showspecialspaceinfo(method)
	dim fso,d,fc,f1,size,showsize,drvpath 		
	set fso=server.createobject("scripting.filesystemobject")
	drvpath=server.mappath("../index.asp")
	drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
	set d=fso.getfolder(drvpath)
	if method="All" then 		
		size=d.size
	elseif method="Program" then
		set fc=d.Files
		for each f1 in fc
			size=size+f1.size
		next	
	end if
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write "<font face=verdana>" & showsize & "</font>"
end sub 	 	 	
	
Function Drawbar(drvpath)
	dim fso,drvpathroot,d,size,totalsize,barsize
	set fso=server.createobject("scripting.filesystemobject")
	drvpathroot=server.mappath("../index.asp")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	drvpath=server.mappath(drvpath)
	if fso.FolderExists(drvpath) then		
		set d=fso.getfolder(drvpath)
		size=d.size
	End If
	barsize=cint((size/totalsize)*400)
	Drawbar=barsize
End Function 	
 	
Function Drawspecialbar()
	dim fso,drvpathroot,d,fc,f1,size,totalsize,barsize
	set fso=server.createobject("scripting.filesystemobject")
	drvpathroot=server.mappath("../index.asp")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	set fc=d.files
	for each f1 in fc
		size=size+f1.size
	next
	barsize=cint((size/totalsize)*400)
	Drawspecialbar=barsize
End Function
	
Function GetAllSpace()
	Dim fso,drvpath,d,size
	set fso=server.createobject("scripting.filesystemobject")
	drvpath=server.mappath("../index.asp")
	drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
	set d=fso.getfolder(drvpath)	
	size=d.size
	set fso=nothing
	GetAllSpace = size
End Function

Function GetFileSize(FileName)
	Dim fso,drvpath,d,size,showsize
	set fso=server.createobject("scripting.filesystemobject")
	drvpath=server.mappath(FileName)
	set d=fso.getfile(drvpath)	
	size=d.size
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	set fso=nothing
	GetFileSize = showsize
End Function

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

Sub GetSysInfo()
	On Error Resume Next
	Dim WshShell,WshSysEnv
	Set WshShell = Server.CreateObject("WScript.Shell")
	Set WshSysEnv = WshShell.Environment("SYSTEM")
	okOS = Cstr(WshSysEnv("OS"))
	okCPUS = Cstr(WshSysEnv("NUMBER_OF_PROCESSORS"))
	okCPU = Cstr(WshSysEnv("PROCESSOR_IDENTIFIER"))
	If IsNull(okCPUS) Then
		okCPUS = Request.ServerVariables("NUMBER_OF_PROCESSORS")
	ElseIf okCPUS="" Then
		okCPUS = Request.ServerVariables("NUMBER_OF_PROCESSORS")
	End If
	If Request.ServerVariables("OS")="" Then okOS=okOS & "(可能是 Windows Server 2003)"
End Sub
%>