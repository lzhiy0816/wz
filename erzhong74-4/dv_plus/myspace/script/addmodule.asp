<!--#include file="../../../Conn.asp"-->
<%
MyDbPath = "../../../"
%>
<!-- #include file="../../../inc/const.asp" -->
<!-- #include file="Dv_ClsSpace.asp" -->
<%
Dim MySpace
Set MySpace = New Cls_Space
Dvbbs.ScriptPath="../../../"
Dvbbs.LoadTemplates("")
Dvbbs.Stats = "�����չƵ��ģ��"
Dvbbs.Head()
Page_Main()
Set MySpace = Nothing
Set Dvbbs = Nothing

Sub Page_Main()
	Dim Url : Url = Request.QueryString("url")
	Dim ID : ID = Request.QueryString("id")
	Dim Path : Path = Request.QueryString("path")
	Dim XmlDom
	If Dvbbs.UserID = 0 Then
		Response.Write "�뿪ͨ������ҳ�ٷ��ʣ�"
		Response.End
	End If
	If Not Isnumeric(Replace(ID,"mod_","")) Then
		Response.Write "�����Ƿ���"
		Response.End
	End If
	If Url="" or ID="" or Path="" Then
		Response.Write "ȱ�ٲ�����"
		Response.End
	End If
	Set XmlDom = GetXmldoc(Url)
	If XmlDom is Nothing Then
		Response.Write "��Դ�����Ƿ���"
		Response.End
	End If

	'�����ĵ�
	Dim ModNode,ConNode
	Dim Rs,Sql
	Dim ModulesNode,Node
	'ModulePrefs,Content
	Set ModNode = XmlDom.DocumentElement.selectSingleNode("ModulePrefs")
	Set ConNode = XmlDom.DocumentElement.selectSingleNode("Content")

	If (ModNode is nothing) or (ConNode  is nothing) THen
		Response.Write "��Դ���ݲ����ϱ�׼!"
		Exit Sub
	End If
	ModNode.setAttribute "id",ID
	Set ModulesNode = MySpace.Space_Info.selectSingleNode("modules")
	Set Node = ModulesNode.selectSingleNode("//Module/ModulePrefs[@id='"&id&"']")
	If Not (Node is Nothing) Then
		ModulesNode.RemoveChild(Node.parentNode)
	End If
	ModulesNode.appendChild(XmlDom.DocumentElement.cloneNode(True))

	If Not IsObject(Conn) Then ConnectionDatabase
	Sql = "Select Top 1 id,userid,plusdb,cachedb from [Dv_Space_user] where Userid=" & MySpace.Sid
	Set Rs=Dvbbs.iCreateObject("Adodb.RecordSet")
	Rs.Open Sql,Conn,1,3
	If Not Rs.Eof Then
		Rs("cachedb") = MySpace.XmlDoc.xml
		Rs("plusdb") = ModulesNode.xml
		Rs.Update
	Else
		Response.Write "�뿪ͨ���˿ռ���ִ�в���!"
		Exit Sub
	End If
	Rs.Close
	Set Rs = Nothing

	%>
<div id="tempdiv" style="display:none">
	<%=ConNode.text%>
</div>
<script language="JavaScript">
//alert(parent.Drag)
parent.Drag.addmodule('<%=id%>','<%=path%>','<%=Replace(ModNode.getAttribute("title"),"'","/'")%>','<%=ConNode.getAttribute("type")%>',document.getElementById("tempdiv").innerHTML);
alert("[<%=id%>] <%=Replace(ModNode.getAttribute("title"),"'","/'")%> ��ӳɹ���");
window.location.replace(document.referrer);
</script>
	
<%

End Sub

Function GetXmldoc(GetUrl)
	Dim XmlHttp,XmlDom
	Set XmlHttp = Dvbbs.iCreateObject("Microsoft.XMLHTTP")
	XmlHttp.Open "get",GetUrl,false
	XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	'XmlHttp.SetRequestHeader "content-type", "text/xml"
	XmlHttp.send()

	Set XmlDom = Dvbbs.CreateXmlDoc("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	If XmlDom.Load(xmlhttp.responseXML) Then
		Set GetXmldoc = XmlDom
	Else
		Set GetXmldoc = Nothing
	End If
	Set XmlHttp = Nothing
	Set XmlDom = Nothing
End Function
%>