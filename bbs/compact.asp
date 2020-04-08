<%
option explicit
if instr(Request.ServerVariables("http_referer"),"http://"&Request.ServerVariables("server_name")&"") = 0 then
response.write "<script>alert('来源错误！');history.back();</script>"
response.end
end if
Const JET_3X = 4
Function CompactDB(dbPath, boolIs97)
Dim fso, Engine, strDBPath
strDBPath = left(dbPath,instrrev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(dbPath) Then
Set Engine = CreateObject("JRO.JetEngine")
If boolIs97 = "True" Then
Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb;" _
& "Jet OLEDB:Engine Type=" & JET_3X
Else
Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
End If
fso.CopyFile strDBPath & "temp.mdb",dbpath
fso.DeleteFile(strDBPath & "temp.mdb")
Set fso = nothing
Set Engine = nothing
CompactDB = "<script>alert('压缩成功！');history.back();</script>"
Else
CompactDB = "<script>alert('找不到数据库！\n请检查数据库路径是否输入错误！');history.back();</script>"
End If
End Function
Dim dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")
If dbpath <> "" Then
dbpath = server.mappath(dbpath)
response.write(CompactDB(dbpath,boolIs97))
End If
%>