<%@CodePage=65001%>
<%
'==============================
' 작 성 자 : 노광옹
' 작 성 일 : 2024-10-29
' 파일내용 : 점검일지 등록 > 배선카트
'==============================
Option Explicit
%>
<!-- #include virtual = "/common/library/common.asp"-->
<!-- #include virtual = "/common/include/requestValue.asp"-->
<%
Dim sCommand, sDinner, sPage, sCurrDate, sCartCode, sIdx, sItemCode, sBreakfast, sLunch, sItemIdx, sConfirm
Dim sComplaint, sEtc, sManagerId
Dim arrComplaint, arrEtc , arrManagerId
sCommand = nullFillWith(Request("command"), "", False)
sDinner = nullFillWith(Request("dinner"), "", False)
sPage = nullFillWith(Request("page"), "", False)
sCurrDate = nullFillWith(Request("currDate"), "", False)
sCartCode = nullFillWith(Request("cartCode"), "", False)
sIdx = nullFillWith(Request("idx"), "", False)
sItemCode = nullFillWith(Request("itemCode"), "", False)
sBreakfast = nullFillWith(Request("breakfast"), "", False)
sLunch = nullFillWith(Request("lunch"), "", False)
sComplaint = nullFillWith(Request("complaint"), "", False)
sEtc = nullFillWith(Request("etc"), "", False)
sManagerId = nullFillWith(Request("managerId"), "", False)

sItemIdx = nullFillWith(Request("itemIdx"), "", False)	
sConfirm  = nullFillWith(Request("confirm"), "", False)	

arrComplaint	= Split(sComplaint, ",")
arrEtc			= Split(sEtc, ",")
arrManagerId	= Split(sManagerId, ",")

Execute("Call "& sCommand &"_Process()")    '함수 호츨

Sub Meal_Cart_Data_Process()
	Dim arrItemIdx, arrItemCode, arrBreakFast, arrLunch, arrDinner, i
	Dim sql , rs
	arrItemIdx		= Split(sItemIdx, ",")
	arrItemCode		= Split(sItemCode, ",")
	arrBreakFast	= Split(sBreakfast, ",")
	arrLunch		= Split(sLunch, ",")
	arrDinner		= Split(sDinner, ",")

	For i = 0 To Ubound(arrItemIdx)
		'조식
		sql = "EXEC USP_CART_RECORD_DATA_MEAL_PROC '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}','{P8}','{P9}','{P10}';"
		sql = Replace(sql , "{P1}" , Trim(arrItemIdx(i)) , vbTextCompare)
		sql = Replace(sql , "{P2}" , sCurrDate , vbTextCompare)
		sql = Replace(sql , "{P3}" , sCartCode , vbTextCompare)
		sql = Replace(sql , "{P4}" , Trim(arrItemCode(i)) , vbTextCompare)
		sql = Replace(sql , "{P5}" , "1" , vbTextCompare)
		sql = Replace(sql , "{P6}" , Trim(arrBreakFast(i)) , vbTextCompare)
		sql = Replace(sql , "{P7}" , Trim(arrComplaint(0)) , vbTextCompare)
		sql = Replace(sql , "{P8}" , Trim(arrEtc(0)) , vbTextCompare)
		sql = Replace(sql , "{P9}" , Trim(arrManagerId(0)) , vbTextCompare)
		sql = Replace(sql , "{P10}" , Session("SNT_ADMIN_ID") , vbTextCompare)
		Call setQueryDB(sql)

		'중식
		sql = "EXEC USP_CART_RECORD_DATA_MEAL_PROC '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}','{P8}','{P9}','{P10}';"
		sql = Replace(sql , "{P1}" , Trim(arrItemIdx(i)) , vbTextCompare)
		sql = Replace(sql , "{P2}" , sCurrDate , vbTextCompare)
		sql = Replace(sql , "{P3}" , sCartCode , vbTextCompare)
		sql = Replace(sql , "{P4}" , Trim(arrItemCode(i)) , vbTextCompare)
		sql = Replace(sql , "{P5}" , "2" , vbTextCompare)
		sql = Replace(sql , "{P6}" , Trim(arrLunch(i)) , vbTextCompare)
		sql = Replace(sql , "{P7}" , Trim(arrComplaint(1)) , vbTextCompare)
		sql = Replace(sql , "{P8}" , Trim(arrEtc(1)) , vbTextCompare)
		sql = Replace(sql , "{P9}" , Trim(arrManagerId(1)) , vbTextCompare)
		sql = Replace(sql , "{P10}" , Session("SNT_ADMIN_ID") , vbTextCompare)
		Call setQueryDB(sql)

		'석식
		sql = "EXEC USP_CART_RECORD_DATA_MEAL_PROC '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}','{P8}','{P9}','{P10}';"
		sql = Replace(sql , "{P1}" , Trim(arrItemIdx(i)) , vbTextCompare)
		sql = Replace(sql , "{P2}" , sCurrDate , vbTextCompare)
		sql = Replace(sql , "{P3}" , sCartCode , vbTextCompare)
		sql = Replace(sql , "{P4}" , Trim(arrItemCode(i)) , vbTextCompare)
		sql = Replace(sql , "{P5}" , "3" , vbTextCompare)
		sql = Replace(sql , "{P6}" , Trim(arrDinner(i)) , vbTextCompare)
		sql = Replace(sql , "{P7}" , Trim(arrComplaint(2)) , vbTextCompare)
		sql = Replace(sql , "{P8}" , Trim(arrEtc(2)) , vbTextCompare)
		sql = Replace(sql , "{P9}" , Trim(arrManagerId(2)) , vbTextCompare)
		sql = Replace(sql , "{P10}" , Session("SNT_ADMIN_ID") , vbTextCompare)
		Call setQueryDB(sql)

	Next 
	Call onlyAlert("success","처리완료 되었습니다.")
End Sub 

Sub Meal_Cart_Approve_Process()
	Dim sql 
	sql = "EXEC AMD_APPROVE_PROC '{P1}','{P2}','{P3}','{P4}';" 
	sql = Replace(sql , "{P1}" , Replace(sCurrDate,"-","") , vbTextCompare)
	sql = Replace(sql , "{P2}" , "CART" , vbTextCompare)
	sql = Replace(sql , "{P3}" , Session("SNT_ADMIN_ID") , vbTextCompare)
	sql = Replace(sql , "{P4}" , sConfirm , vbTextCompare)
	Call setQueryDB(sql)
	Call gotoAlert("success","처리완료 되었습니다.","r","")
End Sub 

Sub Meal_Cart_Confirm_Process()
	Dim sql
	sql = "EXEC USP_CART_RECORD_DATA_COMFIRM '{P1}','{P2}','{P3}';"
	sql = Replace(sql , "{P1}" , sCurrDate , vbTextCompare)
	sql = Replace(sql , "{P2}" , sCartCode , vbTextCompare)
	sql = Replace(sql , "{P3}" , sConfirm , vbTextCompare)
	Call setQueryDB(sql)
	Call gotoAlert("success","처리완료 되었습니다.","r","")
End Sub 

Sub Meal_Cart_Remove_Process()
	Dim sql
	sql = "EXEC USP_CART_RECORD_DATA_DEL '{P1}','{P2}';"
	sql = Replace(sql,"{P1}",sCurrDate,vbTextCompare)
	sql = Replace(sql,"{P2}",sCartCode,vbTextCompare)
	Call setQueryDB(sql)
	Call gotoAlert("success","처리완료 되었습니다.","m","./meal_cart.asp")
End Sub 
%>