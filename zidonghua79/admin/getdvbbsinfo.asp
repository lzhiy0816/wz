<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="../inc/md5.asp" -->
<!-- #include file="../inc/myadmin.asp" -->
<%
	Rem 本文件用于论坛和动网官方主服务器通讯验证用途
	Rem 所有信息经过<双重不可解密>的加密算法加密,加密信息为当前操作的论坛管理员最后登陆IP
	Rem 查询用户为前台的管理员用户名,均为公开或<双重不可解密>的信息
	Rem 非法的对该文件请求只能猜得前台管理员帐户(这是前台公开信息)
	Rem 非法的对该文件请求并不能获知后台登陆账号,也只能获取无用的IP不可解密加密信息
	Dim UserName
	UserName = Request("Name")
	If UserName = "" Then
		Response.Write "非法的参数！"
		Response.End
	End If
	UserName = Replace(UserName,"'","")
	Set Rs = Dvbbs.Execute("Select LastLoginIP From "&AdminTable&" Where AddUser = '"&UserName&"'")
	If Rs.Eof And Rs.Bof Then
		Response.Write "Null"
	Else
		Response.Write Md5(Md5(Rs(0),32),32)
	End If
	Rs.Close
	Set Rs=Nothing
	Conn.Close
	Set Conn=Nothing
%>