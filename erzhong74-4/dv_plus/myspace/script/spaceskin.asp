<!--#include file="../../../Conn.asp"-->
<%
MyDbPath = "../../../"
%>
<!-- #include file="../../../inc/const.asp" -->
<!-- #include file="Dv_ClsSpace.asp" -->
<%
Dvbbs.ScriptPath="../../../"
Dvbbs.LoadTemplates("")
Dvbbs.Stats = "风格模板"
Dvbbs.Head()
If Dvbbs.UserID = 0 Then
	Response.Write "请开通个人首页再访问！"
	Response.End
End If
Page_Main()
Set Dvbbs = Nothing

Sub Page_Main()
%>
<style>
a {text-decoration : none;color : #000000; } 
a:hover {text-decoration : underline; color : #559AE4; } 
body {
	text-align:center;margin-top:0;
} 
body,ol,ul,div{
	font-size : 12px; color : #000000; font-family : tahoma, 宋体, fantasy; 
}
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
	margin:0px 10px 0px 0px;
	padding:0;
	float:left;
}
img.demogif{
	width:100px;
	height:80px;
	border:none;
}
hr {
height:0px;border :0px;border-top: gray 1px solid;width : 98%; 
} 

</style>
<script language="JavaScript" src="<%=MyDbPath%>inc/Pagination.js"></script>
<form method="post" name="skinform">
<div class="skindiv">
	<table cellpadding="3" cellspacing="3" class="readme">
	<tr>
	<td width="100%">
		<ol>
			<li>点击图片可预览风格效果；</li>
			<li>选取喜欢的模板后，点击右边的确认按钮；</li>
		</ol>
	</td>
	<td width="20%">
		<input type="hidden" name="skinid" value=""/>
		<input type="hidden" name="skinpath" value=""/>
		<input type="button" value="确认" name="selskins" style="height:60px;width:50px;" onclick="parent.changskin(document.skinform.skinid.value,document.skinform.skinpath.value);"/>
	</td>
	</tr>
	</table>
</div>
<div class="skindiv">
<%
	Dim Rs,Sql
	Dim Page,MaxRows,Endpage,CountNum,PageSearch,i,j
	Sql = "Select id,s_name,s_username,s_userid,s_style,s_path,s_lock,s_addtime From Dv_Space_skin where s_lock=1 order by id"

	PageSearch = ""
	Endpage = 0
	MaxRows = 6
	CountNum = 0
	Page = Request("Page")
	If IsNumeric(Page) = 0 or Page="" Then Page=1
	Page = Clng(Page)
	Set Rs = Dvbbs.iCreateObject ("adodb.recordset")
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open Sql,Conn,1,1
	If Not Rs.eof Then
			CountNum = Rs.RecordCount
			If CountNum Mod MaxRows=0 Then
				Endpage = CountNum \ MaxRows
			Else
				Endpage = CountNum \ MaxRows+1
			End If
			Rs.MoveFirst
			If Page > Endpage Then Page = Endpage
			If Page < 1 Then Page = 1
			If Page >1 Then 				
				Rs.Move (Page-1) * MaxRows
			End if
			SQL=Rs.GetRows(MaxRows)
	Else
			Response.Write "暂时没有任何风格数据！"

	End If
	Rs.close:Set Rs = Nothing 
	
	If IsArray(Sql) Then
		j = 0
		For i=0 To Ubound(SQL,2)
			If j=0 Then Response.Write "<br/><ul class=skin>"
			Response.Write "<li>"
			Response.Write "<br/><input class=""radio"" type=""radio"" name=""skin_id"" value="""&Sql(0,i)&""" onclick=""document.skinform.skinid.value=this.value;document.skinform.skinpath.value='"&Sql(5,i)&"';""/>"&Sql(1,i)
			Response.Write "<br/><a href="""&MyDbPath&"skins/myspace/"&Sql(5,i)&"demo.htm"" title=""预览风格效果"" target=""_blank""><img src="""&MyDbPath&"skins/myspace/"&Sql(5,i)&"demo.gif"" class=""demogif""/></a>"
			Response.Write "<br/>提供者："&Sql(2,i)
			Response.Write "</li>"
			If j=2 or i=Ubound(SQL,2) Then
				Response.Write "</ul><div style=""clear:both""></div><hr/>"
				j = 0
			Else
				j = j+1
			End If
		Next

	End If
%>
</div>
</form>
<div class="skindiv" style="height:25px;margin-top:8px;" aligh="right">
<script language="JavaScript">
<!--
PageList(<%=Page%>,3,<%=MaxRows%>,<%=CountNum%>,"<%=PageSearch%>",0);
//-->
</script>
</div>
<%
End Sub


%>