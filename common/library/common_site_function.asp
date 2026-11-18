<%
'========================================
' 로그인 체크
' - /pages/login/ 을 제외한 모든 페이지는 로그인 세션(SNT_ADMIN_ID)이 있어야 접근 가능
' - 2026-11-18 복구: 기존 코드는 다른 프로젝트의 세션 키(FMS_ADMIN_ID)를 검사해
'   항상 로그인 화면으로 튕겨서 주석 처리돼 있었음 → 로그인 시 저장하는 SNT_ADMIN_ID 로 수정
'========================================
If InStr(1, Request.ServerVariables("SCRIPT_NAME") , "/pages/login/", vbTextCompare) = 0  Then
	If Session("SNT_ADMIN_ID") = "" Then
		Call errorParentMove("세션이종료되어 로그인 페이지로 이동합니다!","/pages/login/")
		Response.End
	End If
End If

Sub isCurrentChk(dbForder , curForder )
	'Response.write dbForder & "=" & curForder & "<br/>"
	If curForder = dbForder Then 
		Response.write " current "
	End If 
End Sub 

Sub getClassChk(chk1 , chk2 , className)
	'rw("<br/>" & chk1 & chk2)
	If Trim(chk1) = Trim(chk2) Then 
		Response.write className
	End If 
End Sub 
'========================================
' 인젝션 처리
'========================================
Public Function checkRequest(byref str)
Dim tempStr
	tempStr = Empty
'	tempStr = instr(str,"'") + instr(str,"--") +  instr(str,"varchar") + instr(Lcase(str),"exec ") + instr(Lcase(str),"exec(") + instr(Lcase(str),"declare") + instr(Lcase(str),".js<") + instr(Lcase(str),"<script") + instr(str,";") +  instr(Lcase(str),"select") +  instr(Lcase(str),"insert ") + instr(Lcase(str),"update ") + instr(Lcase(str),"delete ") + instr(Lcase(str),"union ") + instr(Lcase(str)," or ") +  instr(Lcase(str)," and ")
	If tempStr = 0 and Len(str) < 200 then
		tempStr = str
	Else
		tempStr = Null
	End If
	checkRequest = tempStr
End Function

'========================================
' html 제거 
'========================================
Function noHtml(contents)
		contents = Trim(contents) '앞, 뒤 공백제거
		contents = Replace(contents,chr(13)&chr(10)," ") 'Enter 처리
		contents = Replace(contents,chr(9)," ") 'Tab 처리
		contents = Replace(contents,"'","''") '작은따옴표(') 처리
		noHtml = contents
End Function 

Function strAddClsss(str1 , str2 , className)
	If str1 = str2 Then 
		strAddClsss = className
	End If 
End Function 

'========================================
' 문자가 null일 경우 restr 에 입력된 값으로 변환
'========================================
public function replaceNull(str, restr)
    dim str_len
    
    If str = "" Then
      str_len = restr
    elseif IsNull(str) Then
      str_len = restr
    elseif IsEmpty(str) Then
      str_len = restr
    else 
    	str_len = str
    end if
    replaceNull =  str_len
End Function

Function getThDayColor(t , d)
	If WeekDay(d) = 1 Then 
		If t = "num" Then getThDayColor = "<th  class=""primary sunday"">" & day(d) & "</th>"		
		If t = "str" Then getThDayColor = "<th  class=""primary sunday"">" & getWeekDayStr(d) & "</th>"		
	ElseIf WeekDay(d) = 7 Then 
		If t = "num" Then getThDayColor = "<th  class=""primary saturday"">" & day(d) & "</th>"
		If t = "str" Then getThDayColor = "<th  class=""primary saturday"">" & getWeekDayStr(d) & "</th>"
	Else 
		If t = "num" Then getThDayColor = "<th class=""primary"">" & day(d) & "</th>"
		If t = "str" Then getThDayColor = "<th class=""primary"">" & getWeekDayStr(d) & "</th>"
	End If 
End Function 

Function getWeekDayStr(n)
	Dim strWeek

	Select Case WeekDay(n)
		Case "1"
			strWeek = "일"
		Case "2"
			strWeek = "월"
		Case "3"
			strWeek = "화"
		Case "4"
			strWeek = "수"
		Case "5"
			strWeek = "목"
		Case "6"
			strWeek = "금"
		Case "7"
			strWeek = "토"
	End Select 
	getWeekDayStr  = strWeek
End Function 
%>