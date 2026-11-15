<%@CodePage=65001%>
<% 
'=============================================
' 작 성 자 : 노 광 옹
' 작 성 일 : 2025-03-21
' 파일내용 : 일일배식 출력
'=============================================
Option Explicit 
%>
<!-- #include virtual = "/common/library/common.asp"-->
<!doctype html>
<html lang="ko">
  <head>
	<!-- #include virtual = "/common/include/layout_header.asp"-->
	<style>
		@media print {
			.noprint {display:none;}
			.print {display:block;}
		}

		.section {
			border-radius: 0;
			box-shadow: none;
			h2 {
				font-weight: 800;
				font-size: 22px;
			}
			
			.table {
				overflow: hidden;
				display: flex;
				justify-content: end;

				& table {
					& th, & td {
						border: 1px solid var(--color-gray-01);
					}

					td {
						padding: 5px;
						text-wrap: initial;
					}
				}
			}
			
			.head_container {
				.table {
					table {
						width: 40%;
					}
				}
			}

			.body_container {
				.table {
					margin-top: 30px;
					table {
						border: 3px solid var(--color-gray-01);
						width: 800px;
					}
				}
			}
			    
			
			.info {
				margin-top: 20px;
				display: flex;
				justify-content: space-between;
				.input {
					&.date {
						width: 180px;
					}
				}

				ul li {
					width: 300px;
					display: flex;
					:nth-child(odd) {
						flex: .3;
					}
					:nth-child(even) {
						flex: .7;
						&::before {
							content: ":";
							margin-right: 10px;
						}
					}
				}
			}
			
			.text_blue {
				color: var(--color-sub-02);
			}
			.text_red {
				color: var(--color-red);
			}
			.left {
				text-align: left;
			}
			.left_ti2 {
				text-indent: 2%;
			}
			
			.hide {
				display: none;
			}
		}

		.bx-calendar {
			position: absolute;
			right: 10px;
			pointer-events: none;
		}

	</style>
 <%
	Dim selDate, sHospitalType, selHoType
	selDate			= nullFillWith(Request("selDate"),Date(),False)
	sHospitalType	= nullFillWith(Request("hospitalType"),Date(),False)

	Dim printerYn	:	printerYn = "Y"
	Dim sql , rs , i , arrData
	sql = "EXEC USP_FOOD_LOTATION_RECORD_DETAIL_VIEW  '{P1}','{P2}';"
	sql = Replace(sql , "{P1}" , selDate , vbTextCompare)
	sql = Replace(sql , "{P2}" , sHospitalType , vbTextCompare)
	'rw(sql)
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrData = rs.GetRows()
	End If 
	rsClose(rs)


	Dim HP_IDX, HP_ROOM_CODE, HP_ROOM_NAME , arrRoom
	sql = "EXEC ADM_HOSPITAL_ROOM_LIST '{P1}','{P2}';"
	sql = Replace(sql , "{P1}" , sHospitalType , vbTextCompare)
	sql = Replace(sql , "{P2}" , "" , vbTextCompare)
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrRoom	= rs.GetRows()
	End If 
	rsClose(rs)


	Dim OP_CURR_YMD, OP_STAFF, OP_HOSPITAL_ROOM, OP_FM_TIME, OP_TO_TIME, OP_REG_ID , OP_ROOM_NAME , OP_REG_NM  
	Dim OP_HOSPITAL_TYPE , OP_HOSPITAL_TYPE_NAME 
	IF isArray(arrData) Then
		OP_CURR_YMD			= nullFillWith(arrData(1,0),"", False)
		OP_STAFF			= nullFillWith(arrData(2,0),"", False)
		OP_HOSPITAL_ROOM	= nullFillWith(arrData(4,0),"", False)
		OP_FM_TIME			= nullFillWith(arrData(5,0),"", False)
		OP_TO_TIME			= nullFillWith(arrData(6,0),"", False)
		OP_REG_ID			= nullFillWith(arrData(12,0),"", False)
		OP_ROOM_NAME		= nullFillWith(arrData(13,0),"", False)
		OP_REG_NM			= nullFillWith(arrData(15,0),"", False)
		OP_HOSPITAL_TYPE		= nullFillWith(arrData(18,0),"", False)
		OP_HOSPITAL_TYPE_NAME	= nullFillWith(arrData(17,0),"", False)
		If OP_FM_TIME <> "" Then OP_FM_TIME			= Left(OP_FM_TIME,2) & ":" & Right(OP_FM_TIME,2)
		If OP_TO_TIME <> "" Then OP_TO_TIME			= Left(OP_TO_TIME,2) & ":" & Right(OP_TO_TIME,2)
	End If 

	Dim AP_CHA_MANAGER_NM , AP_CHA_STAFF_NM , AP_HD_MANAGER_NM , AP_HD_STAFF_NM
	Dim AP_CHA_MANAGER_ID , AP_CHA_STAFF_ID , AP_HD_MANAGER_ID , AP_HD_STAFF_ID
	sql = "EXEC ADM_APPROVE_DETAIL_VIEW '{P1}','{P2}','{P3}','{P4}';"
	sql = Replace(sql , "{P1}" , selDate , vbTextCompare)
	sql = Replace(sql , "{P2}" , "MEAL"    , vbTextCompare)
	sql = Replace(sql , "{P3}" , sHospitalType    , vbTextCompare)
	sql = Replace(sql , "{P4}" , ""    , vbTextCompare)

	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		AP_CHA_MANAGER_NM	 = nullFillWith(rs("CHA_MANAGER_NAME"),"",False)
		AP_CHA_STAFF_NM	 = nullFillWith(rs("CHA_STAFF_NAME"),"",False)
		AP_HD_MANAGER_NM	 = nullFillWith(rs("HD_MANAGER_NAME"),"",False)
		AP_HD_STAFF_NM		 = nullFillWith(rs("HD_STAFF_NAME"),"",False)
		AP_CHA_MANAGER_ID	= nullFillWith(rs("CHA_MANAGER"),"",False)
		AP_CHA_STAFF_ID		= nullFillWith(rs("CHA_STAFF"),"",False)
		AP_HD_MANAGER_ID	= nullFillWith(rs("HD_MANAGER"),"",False)
		AP_HD_STAFF_ID		= nullFillWith(rs("HD_STAFF"),"",False)
	Else 
		AP_CHA_MANAGER_NM	= ""
		AP_CHA_STAFF_NM		= ""
		AP_HD_MANAGER_NM	= ""
		AP_HD_STAFF_NM		= ""
		AP_CHA_MANAGER_ID	= ""
		AP_CHA_STAFF_ID		= ""
		AP_HD_MANAGER_ID	= ""
		AP_HD_STAFF_ID		= ""
	End If 
	rsClose(rs)
 %>
  </head>
  <body>
    <form class="form" name="dataForm" id="dataForm">
	<input type="hidden" name="command" id="command" value="Approve" title="명령" style="background-color:lime;">
	<input type="hidden" name="FLAG" id="FLAG" value="MEAL" title="플래그값" style="background-color:lime;">
	<input type="hidden" name="approveGrade" id="approveGrade" value="" title="결재" style="background-color:lime;">
	<section class="section">
		<div class="head_container">
			<div style="display: flex;justify-content: space-between;">
				<h2>배식서비스 체크리스트</h2>
				<%
				If AP_CHA_MANAGER_NM = "" Or AP_CHA_STAFF_NM = "" Or AP_HD_MANAGER_NM = "" Or AP_HD_STAFF_NM = "" Then 
					printerYn = "N"
				End If 

				printerYn = "Y"
				If printerYn = "Y" Then %>
				<button class="button noprint" id="btnPrint" onClick="javascript:reportPrint();void(0);" style="width:100px;padding: 0;height: 100%;">출력하기</button>
				<% End If %>
			</div>
			<div class="table">
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
							<td style="height:30px;">
							<%
							If AP_HD_STAFF_NM = "" And vAdminCompany = "HD" And vAdminPosition = "STAFF" Then 
								%><small class="badge green noprint" onClick="javascript:btnApprove('1');void(0);" style="cursor:pointer;">승인</small><%
							End If 

							If AP_HD_STAFF_NM <> "" Then 
								rw(AP_HD_STAFF_NM)
								If Session("SNT_ADMIN_ID") = AP_HD_STAFF_ID Then 
									rw("<BR/><small class=""badge red noprint"" onClick=""javascript:btnApproveCancel('1');void(0);"" style=""cursor:pointer;"">취소</small>")
								End If 
							End If 
							%>
							</td>
							<td>
							<%
							If AP_HD_MANAGER_NM = "" And vAdminCompany = "HD" And vAdminPosition = "MANAGER" Then 
								%><small class="badge orange" onClick="javascript:btnApprove('2');void(0);" style="cursor:pointer;">승인</small><%
							End If 

							If AP_HD_MANAGER_NM <> "" Then 
								rw(AP_HD_MANAGER_NM)
								If Session("SNT_ADMIN_ID") = AP_HD_MANAGER_ID Then 
									rw("<BR/><small class=""badge red noprint"" onClick=""javascript:btnApproveCancel('2');void(0);"" style=""cursor:pointer;"">취소</small>")
								End If
							End If 
							%>
							</td>
							<td>
							<%
							If AP_CHA_STAFF_NM = "" And vAdminCompany = "CHA" And vAdminPosition = "STAFF"Then 
								%><small class="badge orange" onClick="javascript:btnApprove('3');void(0);" style="cursor:pointer;">승인</small><%
							End If 
							If AP_CHA_STAFF_NM <> "" Then 
								rw(AP_CHA_STAFF_NM)
								If Session("SNT_ADMIN_ID") = AP_CHA_STAFF_ID Then 
									rw("<BR/><small class=""badge red noprint"" onClick=""javascript:btnApproveCancel('3');void(0);"" style=""cursor:pointer;"">취소</small>")
								End If
							End If 
							%>
							</td>
							<td>
							<%
							If AP_CHA_MANAGER_NM = "" And vAdminCompany = "CHA" And vAdminPosition = "MANAGER" Then 
							%><small class="badge orange" onClick="javascript:btnApprove('4');void(0);"" style="cursor:pointer;">승인</small><% 
							End If 
							If AP_CHA_MANAGER_NM <> "" Then 
								rw(AP_CHA_MANAGER_NM)
								If Session("SNT_ADMIN_ID") = AP_CHA_MANAGER_ID Then 
									rw("<BR/><small class=""badge red noprint"" onClick=""javascript:btnApproveCancel('4');void(0);"" style=""cursor:pointer;"">취소</small>")
								End If
							End If 
							%>
							</td>
						</tr>
					</tbody>
				</table>
			</div>
			<div>
			</div>
			<div style="display:flex;gap:10px;">
				<div class="select noprint"  style="width:200px;height:100%;">
					<select id="hospitalType" name="hospitalType">
						<option value="">전체</option>
						<%=getCodeDB("select" , "P11" , sHospitalType)%>			
					</select>
				</div>
				<div class="date-picker noprint" style="margin-bottom:10px;">
					<div class="input date" style="width:150px;">
						<input type="text" name="selDate" id="selDate" placeholder="YYYY-MM-DD" readonly value="<%=selDate%>"/>
						<i class="bx bx-calendar"></i>
					</div>
				</div>
			</div>
			<div class="info">
				<div>						
					<ul>
						<li>
							<div>일시</div>
							<div><%=convertDate(OP_CURR_YMD, "-")%></div>
						</li>
						<li>
							<div>병동명</div>
							<div><%=OP_HOSPITAL_ROOM%></div>
						</li>
						<li>
							<div>담당배선</div>
							<div><%=OP_STAFF%></div>
						</li>
						<li>
							<div>점검자</div>
							<div><%=OP_REG_NM%> (<%=OP_REG_ID%>)</div>
						</li>
					</ul>
				</div>
				<div style="padding-right: 50px;">
					<span>라운딩 시작/종료시간 (<%=OP_FM_TIME%> ~ <%=OP_TO_TIME%>)</span>
				</div>

			</div>
		</div>

		<div class="body_container">
			<div class="table">
				<table>
					<colgroup>
						<col width="10%">
						<col width="70%">
						<col width="20%">
					</colgroup>
					<thead>
						<tr>
							<th>순서</th>
							<th>평가내용</th>
							<th>평가</th>
						</tr>
					</thead>
					<tbody>
						<%
						Dim strEtc
						strEtc = ""
						Dim OP_ITEM_CODE, OP_ITEM_NAME, OP_ITEM_RESULT, OP_ETC_CODE, OP_ETC
						IF isArray(arrData) Then
							For i = 0 To UBound(arrData, 2)
								OP_ITEM_CODE		= nullFillWith(arrData(7,i),"", False)
								OP_ITEM_NAME		= nullFillWith(arrData(8,i),"", False)
								OP_ITEM_RESULT		= nullFillWith(arrData(9,i),"", False)
								OP_ETC_CODE			= nullFillWith(arrData(10,i),"", False)
								OP_ETC				= nullFillWith(arrData(11,i),"", False)			
								
								
								If i = 0 Then 
									strEtc = OP_ETC_CODE & "&nbsp;" & OP_ETC
									strEtc = Replace(strEtc , "없음" , "")
								End If 
						%>
						<tr>
							<td><%=i+1%></td>
							<td class="left"><%=OP_ITEM_NAME%></td>
							<td <%=txtClass(OP_ITEM_RESULT)%>><%=getCode("name","CHECK_FLAG",OP_ITEM_RESULT)%></td>
						</tr>						
						<%
							Next 
						End If 
						%>
						<tr>
							<td>지적사항</td>
							<td class="left left_ti2" colspan="2"><%=strEtc%></td>
						</tr>
					</tbody>
				</table>
			</div>
		</div>
	</section>
	</form>
  </body>
</html>
<%
Function txtClass(t) 
	Dim strTxt
	Select Case t
		Case "1" 
			strTxt = "class=""text_blue"""
		Case "2"
			strTxt = ""
		Case "3" 
			strTxt = "class=""text_red"""
		Case Else 
			strTxt = ""
	End Select 
	txtClass = strTxt
End Function 
%>
<script type="text/javascript">
<!--
$(document).ready(function(){
	const elem = document.getElementById('selDate');
	const rangepicker = new DateRangePicker(elem, {});

	elem.addEventListener('changeDate', function(e) {
	  //console.log(e.detail.date); // new Date() 찍어봤을때랑 똑같은 형태로 나옴
		var selDate = e.detail.date;
		var year = selDate.getFullYear();
		var month = (selDate.getMonth() + 1).toString().padStart(2,'0');
		var day = selDate.getDate();
		//console.log(year + "-" + month + "-" + day);
		btnSubmit();
	});

	$("#hospitalRoom, #hospitalType").change(function(){
		btnSubmit();
	});
});

function btnSubmit() {
	var frm = document.dataForm;
	
	frm.action = "./daily_meal_print.asp";
	frm.target = "_self";
	frm.method = "POST";
	frm.submit();		
}

function btnApprove(v) {
	var frm = document.dataForm;
	if(!setInputAlert(frm.selDate, "날짜를")) return false;
	if(!setInputAlert(frm.hospitalType, "병원구분을")) return false;

	frm.approveGrade.value = v;
	frm.command.value = "Approve";
	frm.action = "./approve_proc.asp";			
	frm.target = "procFrame";
	frm.method = "POST";
	frm.submit();	
}

function btnApproveCancel(v) {
	var frm = document.dataForm;
	frm.approveGrade.value = v;
	frm.command.value = "ApproveCancel";
	frm.action = "./approve_proc.asp";			
	frm.target = "procFrame";
	frm.method = "POST";
	frm.submit();	
}

function reportPrint() {
	window.print();
}

//-->
</script>
<iframe  name="procFrame" id="procFrame" style="width:80%; height:500px;display:none;" class="LEDB"></iframe>