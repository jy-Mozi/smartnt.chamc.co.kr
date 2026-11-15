<%@CodePage=65001%>
<%
'==============================
' 작 성 자 : 노 광 옹
' 작 성 일 : 2024-10-24
' 파일내용 : 점검항목 등록 프로세스
'==============================
Option Explicit
%>
<!-- #include virtual = "/common/library/common.asp"-->
<%

Dim vCommand
	vCommand = nullFillWith(Request("command"), "" , False)

Execute("Call "& vCommand &"_Process()")    '함수 호츨



' 배선카트 등록/수정
Sub mealCartRegister_Process()

	Dim vIdx, vCartcode, vCartname, vItem, vEtc , vHospitalType
		vIdx			= nullFillWith(Request("idx"), "0" , False)
		vCartcode		= nullFillWith(Request("cartcode"), "" , False)
		vCartname		= nullFillWith(Request("cartname"), "" , False)
		vItem			= nullFillWith(Request("itemlist"), "" , False)
		vEtc			= nullFillWith(Request("etc"), "" , False)
		vHospitalType	= nullFillWith(Request("hospitalType"), "" , False)
			

	Dim vIn(6), vRs, vResult, vBody
		vIn(0) = vIdx
		vIn(1) = vCartcode
		vIn(2) = vCartname
		vIn(3) = vItem
		vIn(4) = vEtc
		vIn(5) = ""
		vIn(6) = vHospitalType
	Set vRs = getSP("USP_MEAL_CART_PROC", vIn)
	vResult = vRs("RESULT")

	vBody = "[{'resultCode':'" & vResult & "'}]"
	rw(vBody)

End Sub 

' 일일배식 등록/수정
Sub dailyMealRegister_Process()

	Dim vIdx, vUser, vItem, vEtc
		vIdx	= nullFillWith(Request("idx"), "0" , False)
		vUser	= nullFillWith(Request("user"), "" , False)
		vItem	= nullFillWith(Request("itemlist"), "" , False)
		vEtc	= nullFillWith(Request("etc"), "" , False)

	Dim vIn(4), vRs, vResult, vBody
		vIn(0) = vIdx
		vIn(1) = vUser
		vIn(2) = vItem
		vIn(3) = vEtc
		vIn(4) = ""

	Set vRs = getSP("USP_FOOD_LOCATION_PROC", vIn)
	vResult = vRs("RESULT")

	vBody = "[{'resultCode':'" & vResult & "'}]"
	rw(vBody)

End Sub 


' 기기설비 온도 등록/수정
Sub cart_temperatureRegister_Process()

	Dim vIdx, vAreacode, vAreaname, vItem, vEtc , vEtc2
		vIdx		= nullFillWith(Request("idx"), "0" , False)
		vAreacode	= nullFillWith(Request("areacode"), "" , False)
		vAreaname	= nullFillWith(Request("areaname"), "" , False)
		vItem		= nullFillWith(Request("itemlist"), "" , False)
		vEtc		= nullFillWith(Request("etc"), "" , False)
		vEtc2		= nullFillWith(Request("etc2"), "" , False)

	Dim vIn(6), vRs, vResult, vBody
		vIn(0) = vIdx
		vIn(1) = vAreacode
		vIn(2) = vAreaname
		vIn(3) = vItem
		vIn(4) = vEtc
		vIn(5) = vEtc2
		vIn(6) = ""
	'rw("USP_MACHINE_CELSIUS_PROC " & Join(vIn,"','"))
	Set vRs = getSP("USP_MACHINE_CELSIUS_PROC", vIn)
	vResult = vRs("RESULT")

	vBody = "[{'resultCode':'" & vResult & "'}]"
	rw(vBody)

End Sub 


' 기기설비 등록/수정
Sub cart_equipmentRegister_Process()

	Dim vIdx, vAreacode, vAreaname, vItem, vEtc, vHospitalType
		vIdx		= nullFillWith(Request("idx"), "0" , False)
		vAreacode	= nullFillWith(Request("areacode"), "" , False)
		vAreaname	= nullFillWith(Request("areaname"), "" , False)
		vItem		= nullFillWith(Request("itemlist"), "" , False)
		vEtc		= nullFillWith(Request("etc"), "" , False)
		vHospitalType = nullFillWith(Request("hospitalType"), "" , False)

	Dim vIn(6), vRs, vResult, vBody
		vIn(0) = vIdx
		vIn(1) = vAreacode
		vIn(2) = vAreaname
		vIn(3) = vItem
		vIn(4) = vEtc
		vIn(5) = ""
		vIn(6) = vHospitalType
	Set vRs = getSP("ADM_MACHINE_FACILITY_PROC", vIn)
	vResult = vRs("RESULT")

	vBody = "[{'resultCode':'" & vResult & "'}]"
	rw(vBody)

End Sub 


Sub meal_room_register_process()
	Dim sCommand, sIdx, sHospitalType, sHospitalRoom, sHospitalRoomCode
	Dim sql 
	sCommand		= nullFillWith(Request("command"), "", False)
	sIdx			= nullFillWith(Request("idx"), "", False)
	sHospitalType	= nullFillWith(Request("hospitalType"), "", False)
	sHospitalRoom	= nullFillWith(Request("hospitalRoom"), "", False)
	sHospitalRoomCode = nullFillWith(Request("hospitalRoomCode"), "", False)
	sql = "EXEC USP_HOSPITAL_ROOM_PROC '{P1}','{P2}','{P3}','{P4}';"
	sql = Replace(sql , "{P1}" , sIdx , vbTextCompare)
	sql = Replace(sql , "{P2}" , sHospitalRoom , vbTextCompare)
	sql = Replace(sql , "{P3}" , sHospitalRoomCode , vbTextCompare)
	sql = Replace(sql , "{P4}" , sHospitalType , vbTextCompare)
	Call setQueryDB(sql)
	Call gotoAlert("success","처리완료 되었습니다.","r","")	
End Sub 
%>