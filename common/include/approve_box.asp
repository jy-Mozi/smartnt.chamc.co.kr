<%
'==============================
' 파일내용 : 점검일지 출력 화면 공통 결재 박스
'   - loadApproveInfo : 결재자 정보 조회 (ADM_APPROVE_DETAIL_VIEW)
'   - drawApproveBox  : 결재 표 출력 (현대그린푸드 담당/점장, 본원 담당/팀장)
'   사용 화면 : pages/checklist/*_print.asp
'   승인/취소 버튼은 각 화면의 btnApprove(), btnApproveCancel() 을 호출한다.
'   vAdminCompany, vAdminPosition 은 layout_header.asp 에서 설정된다.
'==============================
Dim AP_CHA_MANAGER_NM , AP_CHA_STAFF_NM , AP_HD_MANAGER_NM , AP_HD_STAFF_NM
Dim AP_CHA_MANAGER_ID , AP_CHA_STAFF_ID , AP_HD_MANAGER_ID , AP_HD_STAFF_ID

'------------------------------
' 결재자 정보 조회
' fDate         : 일지 날짜 (YYYYMMDD)
' fFlag         : CART / MEAL / TEMPERATURE / MACHINE
' fHospitalType : 병원구분 코드
'------------------------------
Sub loadApproveInfo(fDate, fFlag, fHospitalType)
	Dim apSql , apRs
	apSql = "EXEC ADM_APPROVE_DETAIL_VIEW '{P1}','{P2}','{P3}','{P4}';"
	apSql = Replace(apSql , "{P1}" , fDate			, vbTextCompare)
	apSql = Replace(apSql , "{P2}" , fFlag			, vbTextCompare)
	apSql = Replace(apSql , "{P3}" , fHospitalType	, vbTextCompare)
	apSql = Replace(apSql , "{P4}" , ""				, vbTextCompare)

	AP_CHA_MANAGER_NM	= ""
	AP_CHA_STAFF_NM		= ""
	AP_HD_MANAGER_NM	= ""
	AP_HD_STAFF_NM		= ""
	AP_CHA_MANAGER_ID	= ""
	AP_CHA_STAFF_ID		= ""
	AP_HD_MANAGER_ID	= ""
	AP_HD_STAFF_ID		= ""

	Set apRs = getQueryDB(apSql)
	If Not(apRs.Eof Or apRs.Bof) Then
		AP_CHA_MANAGER_NM	= nullFillWith(apRs("CHA_MANAGER_NAME"),"",False)
		AP_CHA_STAFF_NM		= nullFillWith(apRs("CHA_STAFF_NAME"),"",False)
		AP_HD_MANAGER_NM	= nullFillWith(apRs("HD_MANAGER_NAME"),"",False)
		AP_HD_STAFF_NM		= nullFillWith(apRs("HD_STAFF_NAME"),"",False)
		AP_CHA_MANAGER_ID	= nullFillWith(apRs("CHA_MANAGER"),"",False)
		AP_CHA_STAFF_ID		= nullFillWith(apRs("CHA_STAFF"),"",False)
		AP_HD_MANAGER_ID	= nullFillWith(apRs("HD_MANAGER"),"",False)
		AP_HD_STAFF_ID		= nullFillWith(apRs("HD_STAFF"),"",False)
	End If
	rsClose(apRs)
End Sub

'------------------------------
' 결재 칸 하나 출력
' fGrade     : 1~4 (btnApprove 에 넘기는 결재 단계)
' fName, fId : 이미 결재한 사람 이름/아이디
' fCompany, fPosition : 이 칸을 승인할 수 있는 소속(HD/CHA)/직급(STAFF/MANAGER)
' fColor     : 승인 버튼 색 (green/orange)
'------------------------------
Sub drawApproveCell(fGrade, fName, fId, fCompany, fPosition, fColor)
	If fName = "" And vAdminCompany = fCompany And vAdminPosition = fPosition Then
		rw("<small class=""badge " & fColor & " noprint"" onClick=""javascript:btnApprove('" & fGrade & "');void(0);"" style=""cursor:pointer;"">승인</small>")
	End If

	If fName <> "" Then
		rw(fName)
		If Session("SNT_ADMIN_ID") = fId Then
			rw("<BR/><small class=""badge red noprint"" onClick=""javascript:btnApproveCancel('" & fGrade & "');void(0);"" style=""cursor:pointer;"">취소</small>")
		End If
	End If
End Sub

'------------------------------
' 결재 표 출력 (loadApproveInfo 호출 후 사용)
'------------------------------
Sub drawApproveBox()
%>
				<table>
					<tbody>
						<tr>
							<td rowspan="3">결재</td>
							<td colspan="2">현대그린푸드</td>
							<td colspan="2">본원</td>
						</tr>
						<tr>
							<td>담당</td>
							<td>점장</td>
							<td>담당</td>
							<td>팀장</td>
						</tr>
						<tr>
							<td style="height:30px;"><% Call drawApproveCell("1", AP_HD_STAFF_NM, AP_HD_STAFF_ID, "HD", "STAFF", "green") %></td>
							<td><% Call drawApproveCell("2", AP_HD_MANAGER_NM, AP_HD_MANAGER_ID, "HD", "MANAGER", "orange") %></td>
							<td><% Call drawApproveCell("3", AP_CHA_STAFF_NM, AP_CHA_STAFF_ID, "CHA", "STAFF", "orange") %></td>
							<td><% Call drawApproveCell("4", AP_CHA_MANAGER_NM, AP_CHA_MANAGER_ID, "CHA", "MANAGER", "orange") %></td>
						</tr>
					</tbody>
				</table>
<%
End Sub
%>
