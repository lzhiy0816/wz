<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/ubblist.asp"-->
<!-- #include file="inc/dv_clsother.asp" -->
<%
Dim XMLDom,paramnode,TableList,idlist
If Not Dvbbs.Master Then Response.redirect "showerr.asp?ErrCodes=<li>您没有权限。</li>&action=OtherErr"
LoadTableList()
 Select Case  Request("action")
 	Case "dotransform"
 		dotransform
 	Case "get"
 		GetPost()
 	Case "transform"
 		transform()
 	Case Else
 		Dvbbs.LoadTemplates("query")
		LoadCount()
		Set paramnode=XMLDom.documentElement.appendChild(XMLDom.createNode(1,"param",""))
		paramnode.attributes.setNamedItem(XMLDom.createNode(2,"action","")).text=Request("action")
		Response.Write Dvbbs.mainhtml(18)
		Dvbbs.stats="批量更新帖子到xhtml格式"
		Dvbbs.nav()
		Dvbbs.Head_var 2,0,"",""
 		 ShowHTML()
		Dvbbs.Footer
 End Select


Sub ShowHTML()
	Dim xslt,proc,XMLStyle
	Set XMLStyle=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	'XMLStyle.loadxml template.html(2)
	XMLStyle.load Server.MapPath("inc/Admin_transformxhml.xslt")
	Set XSLT=Server.CreateObject("Msxml2.XSLTemplate" & MsxmlVersion)
	XSLT.stylesheet=XMLStyle
	Set proc = XSLT.createProcessor()
	proc.input = XMLDom
  proc.transform()
  Response.Write  proc.output
	Set XMLDOM=Nothing
	Set XSLt=Nothing
	Set proc=Nothing	
End Sub
Sub LoadTableList()
	Dim Rs
	Set Rs=Dvbbs.Execute("select * from [Dv_TableList]")
	TableList=Rs.GetRows(-1)
	Set XMLDom=Dvbbs.ArrayToxml(TableList,Rs,"posttable","xml")
End Sub
Sub LoadCount()
	Dim Node,SQL
	For Each Node In XMLDom.documentElement.selectNodes("posttable")
		Node.attributes.setNamedItem(XMLDom.createNode(2,"count","")).text=Dvbbs.Execute("select Count(*) From "& Node.selectSingleNode("@tablename").text )(0)
	Next
End Sub
Sub transform()
	Response.Write "<script language=""Javascript"" type=""text/javascript"">"& vbnewline
	If Request.form("submit")="开 始" Then
		GetIDList()
		Response.Write "parent.document.Dvform1.submit.value='暂 停';"
		Response.Write "parent.document.getElementById('title1').innerHTML='正在获取贴子信息';"
		Response.Write "location.href='Admin_transformxhml.asp?action=get&tableid="&Request.form("tableid")&"'"
	ElseIf Request.form("submit")="暂 停" Then
		Response.Write "parent.document.Dvform1.submit.value='继 续';"
	ElseIf Request.form("submit")="继 续" Then
		Response.Write "parent.document.Dvform1.submit.value='暂 停';"
		Response.Write "location.href='Admin_transformxhml.asp?action=get&tableid="&Request.form("tableid")&"'"
	End If
	Response.Write vbnewline&"</script>"
End Sub
Sub dotransform()
	Dim SQL,AnnounceID,body,tablename,i
	AnnounceID=CLng(Request.form("id"))
	body=Request.form("body")
	tablename=XMLDom.documentElement.selectSingleNode("posttable[@id="& Request("tableid")&"]/@tablename").nodeValue
	SQL="Update ["& tablename &"] Set Body='"& Dvbbs.Checkstr(body)&"',Ubblist='"& Ubblist(body) &"'  Where AnnounceID="& AnnounceID
	Dvbbs.Execute SQL 
	'Response.Write "<div id=""body"">"
	'Response.Write SQL
	'Response.Write "</div>"
	Response.Write "<script language=""Javascript"" type=""text/javascript"">"& vbnewline
	'Response.Write "alert(document.getElementById('body').innerHTML);"
	Response.Write "location.href='Admin_transformxhml.asp?action=get&tableid="&Request("tableid")&"'"
	Response.Write vbnewline&"</script>"
End Sub
Sub GetPost()
	Dim Rs,SQL,AnnounceID,node,xml,body,tablename,tmp,tmp1,i
	Set Xml=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	xml.Loadxml Session("IDlist")
	Set Node=xml.documentElement.selectSingleNode("row")
	If Node is Nothing Then
		Response.Write "<script language=""Javascript"" type=""text/javascript"">"& vbnewline
		Response.Write "parent.document.Dvform1.submit.value='开 始';"
		Response.Write "parent.document.getElementById('title1').innerHTML='上一批贴子更新完成';"
		Response.Write "parent.document.getElementById('title2').innerHTML='';"
		Response.Write vbnewline&"</script>"
	Else
		AnnounceID=Node.selectSingleNode("@announceid").nodeValue
		xml.documentElement.removeChild(node)
		Session("IDlist")=xml.xml
		tablename=XMLDom.documentElement.selectSingleNode("posttable[@id="& Request("tableid")&"]/@tablename").nodeValue
		SQL="Select Body From "& tablename &" where AnnounceID="& AnnounceID
		Set Rs=Dvbbs.Execute(SQL)
		body=Rs(0)&""
		'Body=Replace(Body,"&lt;br /&gt;","<br />")
		'Body=Replace(body,"<br />",Chr(10))
		Set Rs=Nothing
		Response.Write "<textarea id=""body1"" name=""body1"">"
		Response.Write body
		Response.Write "</textarea>"
		Response.Write "<script language=""Javascript"" type=""text/javascript"">"& vbnewline
		Response.Write "parent.document.getElementById('title2').innerHTML='正在处理编号为："&AnnounceID&" 的贴子';"& vbnewline
		Response.Write "parent.document.Dvform1.sid.value='"&AnnounceID+1&"';"& vbnewline
		Response.Write "parent.document.Dvform.id.value='"&AnnounceID&"';"& vbnewline
		Response.Write "parent.document.Dvform.tableid.value='"&Request("tableid")&"';"& vbnewline
		Response.Write "parent.copyto(document.getElementById('body1').value);"& vbnewline
		Response.Write vbnewline&"</script>"
	End If

End Sub
Sub GetIDList()
	Dim sid,transcount,SQL,tablename,Rs
	transcount=CLng(Request.form("transcount"))
	If transcount< 1 Then transcount=1000
	sid=CLng(Request.form("sid"))
	tablename=XMLDom.documentElement.selectSingleNode("posttable[@id="& Request.form("tableid")&"]/@tablename").nodeValue
	SQL="Select Top "& transcount &" AnnounceID From "&tablename &" where AnnounceID  >="&Sid& " Order By AnnounceID"
	'Response.Write "parent.document.getElementById('title2').innerHTML='"&SQL&"';"
	Set Rs=Dvbbs.Execute(SQL)
	Set IDlist=Dvbbs.RecordsetToxml(Rs,"row","xml")
	Set Rs=Nothing
	Session("IDlist")=IDlist.xml
End Sub
%>