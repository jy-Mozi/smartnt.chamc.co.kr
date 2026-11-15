<%@CodePage=65001%>
<%
'==============================
' 작 성 자 : 
' 작 성 일 : 
' 파일내용 :  
'==============================
Option Explicit


%>

<!-- #include virtual = "/common/library/common.asp"-->
<%

Dim sCommand, sParCodeName, sParCodeUse, sParCode, sParCodeEtc1, sParCodeEtc3 
Dim sComCode, sComCodeName, sComCodeUse, sComCodeEtc1, sComCodeEtc2, sComCodeEtc3
sCommand		= nullFillWith(Request("command"), "", False)
sParCodeName	= nullFillWith(Request("parCodeName"), "", False)
sParCodeUse		= nullFillWith(Request("parCodeUse"), "", False)
sParCode		= nullFillWith(Request("parCode"), "", False)
sParCodeEtc1	= nullFillWith(Request("parCodeEtc1"), "", False)
sParCodeEtc3	= nullFillWith(Request("parCodeEtc3"), "", False)
sComCode		= nullFillWith(Request("comCode"), "", False)
sComCodeName	= nullFillWith(Request("comCodeName"), "", False)
sComCodeUse		= nullFillWith(Request("comCodeUse"), "", False)
sComCodeEtc1	= nullFillWith(Request("comCodeEtc1"), "", False)
sComCodeEtc2	= nullFillWith(Request("comCodeEtc2"), "", False)
sComCodeEtc3	= nullFillWith(Request("comCodeEtc3"), "", False)


Execute("Call "& sCommand &"_Process()")    '함수 호츨

'====================================
' 부모코드 등록
'====================================
Sub Par_Add_Process()
'	rw("<br/>" & sParCode)
'	rw("<br/>" & sParCodeName)
'	rw("<br/>" & sParCodeUse)
'	rw("<br/>" & sParCodeEtc1)
'	rw("<br/>" & sParCodeEtc2)
'	rw("<br/>" & sParCodeEtc3)
	Dim sql , rs , OP_PAR_CODE
	sql = "EXEC USP_CODE_ADD_PAR '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}';"
	sql = Replace(sql , "{P1}" , sParCodeName , vbTextCompare)
	sql = Replace(sql , "{P2}" , sParCodeUse , vbTextCompare)
	sql = Replace(sql , "{P3}" , sParCodeEtc1 , vbTextCompare)
	sql = Replace(sql , "{P4}" , sParCodeEtc2 , vbTextCompare)
	sql = Replace(sql , "{P5}" , sParCodeEtc3 , vbTextCompare)
	sql = Replace(sql , "{P6}" , "REG_ID" , vbTextCompare)
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		OP_PAR_CODE = rs("PAR_VAL")
	End If 

	session("Param") = "?parCode=" & OP_PAR_CODE
	Call gotoAlert("success" , "처리완료되었습니다.","m" , "./code.asp" & session("Param"))
End Sub 

'====================================
' 부모코드 수정
'====================================
Sub Par_Mod_Process()
'	rw("<br/>" & sParCode)
'	rw("<br/>" & sParCodeName)
'	rw("<br/>" & sParCodeUse)
'	rw("<br/>" & sParCodeEtc1)
'	rw("<br/>" & sParCodeEtc2)
'	rw("<br/>" & sParCodeEtc3)
	Dim sql , rs 
	sql = "EXEC USP_CODE_MOD_PAR '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}';"
	sql = Replace(sql , "{P1}" , sParCode , vbTextCompare)
	sql = Replace(sql , "{P2}" , sParCodeName , vbTextCompare)
	sql = Replace(sql , "{P3}" , sParCodeUse , vbTextCompare)
	sql = Replace(sql , "{P4}" , sParCodeEtc1 , vbTextCompare)
	sql = Replace(sql , "{P5}" , sParCodeEtc2 , vbTextCompare)
	sql = Replace(sql , "{P6}" , sParCodeEtc3 , vbTextCompare)
	sql = Replace(sql , "{P7}" , "REG_ID" , vbTextCompare)

    Call setQueryDB(sql)

	session("Param") = "?parCode=" & sParCode
	Call gotoAlert("success" , "처리완료되었습니다.","m" , "./code.asp" & session("Param"))
End Sub 
'====================================
' 자식코드 등록
'====================================
Sub SON_ADD_Process()
	'rw("<br/>" & sParCode)
	'rw("<br/>" & sComCode)	
	'rw("<br/>" & sComCodeName)
	'rw("<br/>" & sComCodeUse)	
	'rw("<br/>" & sComCodeEtc1)
	'rw("<br/>" & sComCodeEtc2)
	'rw("<br/>" & sComCodeEtc3)	
	Dim sql , rs
	sql = "EXEC USP_CODE_ADD_COMMON '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}';"
	sql = Replace(sql , "{P1}" , sParCode , vbTextCompare)
	sql = Replace(sql , "{P2}" , sComCodeName , vbTextCompare)
	sql = Replace(sql , "{P3}" , sComCodeUse , vbTextCompare)
	sql = Replace(sql , "{P4}" , sComCodeEtc1 , vbTextCompare)
	sql = Replace(sql , "{P5}" , sComCodeEtc2 , vbTextCompare)
	sql = Replace(sql , "{P6}" , sComCodeEtc3 , vbTextCompare)
	sql = Replace(sql , "{P7}" , "REG_ID" , vbTextCompare)
	Call setQueryDB(sql)
	session("Param") = "?parCode=" & sParCode
	'Call errorParentMove("처리완료되었습니다.","./" & session("Param"))
	Call gotoAlert("success" , "처리완료되었습니다.","m" , "./code.asp" & session("Param"))
End Sub 

Sub Com_Mod_Process()
	'rw("<br/>" & sParCode)
	'rw("<br/>" & sComCode)	
	'rw("<br/>" & sComCodeName)
	'rw("<br/>" & sComCodeUse)	
	'rw("<br/>" & sComCodeEtc1)
	'rw("<br/>" & sComCodeEtc2)
	'rw("<br/>" & sComCodeEtc3)
	Dim sql , rs
	sql = "EXEC USP_CODE_MOD_COMMON '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}';"
	sql = Replace(sql , "{P1}" , sComCode , vbTextCompare)
	sql = Replace(sql , "{P2}" , sComCodeName , vbTextCompare)
	sql = Replace(sql , "{P3}" , sComCodeUse , vbTextCompare)
	sql = Replace(sql , "{P4}" , sComCodeEtc1 , vbTextCompare)
	sql = Replace(sql , "{P5}" , sComCodeEtc2 , vbTextCompare)
	sql = Replace(sql , "{P6}" , sComCodeEtc3 , vbTextCompare)
	sql = Replace(sql , "{P7}" , "REG_ID" , vbTextCompare)
	Call setQueryDB(sql)
	session("Param") = "?parCode=" & sParCode
	'Call errorParentMove("처리완료되었습니다.","./" & session("Param"))
	Call gotoAlert("success" , "처리완료되었습니다.","m" , "./code.asp" & session("Param"))
End Sub 

Sub ComCodeList_Process()	
	Dim sql , rs  
	Dim OP_IDX , OP_CODE_NM , OP_CODE_VAL , OP_PAR_CODE_VAL , OP_SUB_CODE_NM1 , OP_SUB_CODE_NM2 , OP_SUB_CODE_NM3 , OP_USE_YN
	Dim vBody , dataBody , i
	sql = "EXEC USP_CODE_COMMON_LIST '{P1}';"
	sql = Replace(sql , "{P1}" , sParCode , vbTextCompare)
	
	Set rs = getQueryDB(sql)

	i = 0 
	If Not(rs.Eof Or rs.Bof) Then 
		vBody = "[{'header':{'resultCode':'S'},'result':"		
		
		Do Until rs.Eof 
			OP_IDX			= rs("IDX")
			OP_CODE_NM		= rs("CODE_NM")
			OP_CODE_VAL		= rs("CODE_VAL")
			OP_PAR_CODE_VAL	= rs("PAR_CODE_VAL")
			OP_SUB_CODE_NM1	= rs("SUB_CODE_NM1")
			OP_SUB_CODE_NM2	= rs("SUB_CODE_NM2")
			OP_SUB_CODE_NM3	= rs("SUB_CODE_NM3")
			OP_USE_YN		= rs("USE_YN")
			
			If i = 0 Then 
				vBody = vBody & "["
			End If 
			dataBody = dataBody &",{'idx':'"& OP_IDX &"', 'codeNm':'"& OP_CODE_NM &"', 'codeVal':'"& OP_CODE_VAL &"', 'codeEtc1':'"& OP_SUB_CODE_NM1 &"',  'codeEtc2':'"& OP_SUB_CODE_NM2 &"','codeEtc3':'"& OP_SUB_CODE_NM3 &"','useYn':'"& OP_USE_YN &"' }"
			
			i = i + 1
			rs.MoveNext
		Loop 

		vBody = vBody & Mid(dataBody,2) & "]}]"
	Else 
		vBody = "[{'header':{'resultCode':'E'}}]"
	End If 
	rsClose(rs)
	rw(vBody)
End Sub 


Sub ComCodeDetailView_Process()
	Dim sql , rs  
	Dim OP_IDX , OP_CODE_NM , OP_CODE_VAL , OP_PAR_CODE_VAL , OP_SUB_CODE_NM1 , OP_SUB_CODE_NM2 , OP_SUB_CODE_NM3 , OP_USE_YN	
	Dim vBody , dataBody , i

	sql = "EXEC USP_CODE_COMMON_DETAIL_VIEW '{P1}';"
	sql = Replace(sql , "{P1}" , sComCode , vbTextCompare)
	
	Set rs = getQueryDB(sql)

	i = 0 
	If Not(rs.Eof Or rs.Bof) Then 
		vBody = "[{'header':{'resultCode':'S'},'result':"		
		
		Do Until rs.Eof 
			OP_IDX			= rs("IDX")
			OP_CODE_NM		= rs("CODE_NM")
			OP_CODE_VAL		= rs("CODE_VAL")
			OP_PAR_CODE_VAL	= rs("PAR_CODE_VAL")
			OP_SUB_CODE_NM1	= rs("SUB_CODE_NM1")
			OP_SUB_CODE_NM2	= rs("SUB_CODE_NM2")
			OP_SUB_CODE_NM3	= rs("SUB_CODE_NM3")
			OP_USE_YN		= rs("USE_YN")
			
			If i = 0 Then 
				vBody = vBody & "["
			End If 
			dataBody = dataBody &",{'idx':'"& OP_IDX &"', 'codeNm':'"& OP_CODE_NM &"', 'codeVal':'"& OP_CODE_VAL &"', 'codeEtc1':'"& OP_SUB_CODE_NM1 &"',  'codeEtc2':'"& OP_SUB_CODE_NM2 &"','codeEtc3':'"& OP_SUB_CODE_NM3 &"','useYn':'"& OP_USE_YN &"' }"
			
			i = i + 1
			rs.MoveNext
		Loop 

		vBody = vBody & Mid(dataBody,2) & "]}]"
	Else 
		vBody = "[{'header':{'resultCode':'E'}}]"
	End If 
	rsClose(rs)
	rw(vBody)	

End Sub 
%>