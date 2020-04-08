<%
Rem ==========论坛通用函数=========

Rem Check for valid syntax in an email address.
Function IsValidEmail(email)
	Dim names, name, i, c
	IsValidEmail = True
	names = Split(email, "@")
	If UBound(names) <> 1 Then
	   IsValidEmail = False
	   Exit Function
	End If
	For Each name In names
	   If Len(name) <= 0 Then
		 IsValidEmail = False
		 Exit Function
	   End If
	   For i = 1 To Len(name)
		 c = Lcase(Mid(name, i, 1))
		 If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) Then
		   IsValidEmail = False
		   Exit Function
		 End If
	   Next
	   If Left(name, 1) = "." or Right(name, 1) = "." Then
		  IsValidEmail = False
		  Exit Function
	   End If
	Next
	If InStr(names(1), ".") <= 0 Then
	   IsValidEmail = False
	   Exit Function
	End If
	i = Len(names(1)) - InStrRev(names(1), ".")
	If i <> 2 and i <> 3 Then
	   IsValidEmail = False
	   Exit Function
	End If
	If InStr(email, "..") > 0 Then
	   IsValidEmail = False
	End If
End Function

Function StrLength(str)
	ON ERROR RESUME NEXT
	Dim WINNT_CHINESE
	WINNT_CHINESE    = (len("论坛")=2)
	If WINNT_CHINESE Then
		Dim l,t,c
		Dim i
		l=len(str)
		t=l
		For i=1 To l
			c=asc(mid(str,i,1))
			If c<0 Then c=c+65536
			If c>255 Then
				t=t+1
			End If
		Next
		strLength=t
	Else 
		strLength=len(str)
	End If
	If err.number<>0 Then err.clear
End Function

Function CutStr(str,strlen)
	Dim l,t,c
	l=len(str)
	t=0
	For i=1 To l
		c=Abs(Asc(Mid(str,i,1)))
		If c>255 Then
			t=t+2
		Else
			t=t+1
		End If
		If t>=strlen Then
			cutStr=left(str,i)&"..."
			Exit for
		Else
			cutStr=str
		End If
	Next
	CutStr=replace(cutStr,chr(10),"")
End Function
%>
