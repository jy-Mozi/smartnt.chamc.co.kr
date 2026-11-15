<%@CodePage=65001%>
<% 
'=============================================
' 작 성 자 : 노 광 옹
' 작 성 일 : 2024-10-17
' 파일내용 : 점검일지 등록 > 배선카트 리스트 출력
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
					justify-content: center;
					table {
						border: 3px solid var(--color-gray-01);
						width: 800px;
					}
				}
			}
			    
			
			.date-picker {
				margin-top: 20px;
				.input {
					&.date {
						width: 180px;
					}
				}
			}
			
			.text_blue {
				color: var(--color-sub-02);
			}
			.text_red {
				color: var(--color-red);
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
	Dim selDate, selHoType , seqNum
	selDate		= nullFillWith(Request("selDate"),Date(),False)
	selHoType	= nullFillWith(Request("selHoType"),"",False)
	Dim sql , rs , i , j , arrData 
	Dim OP_CART_CODE , OP_CART_NAME , arrRecord , currRowSpan 
	Dim OP_BREAKFAST_VAL , OP_LUNCH_VAL , OP_DINNER_VAL, OP_ITEM_NAME, OP_ETC , printerYn,OP_ETC_desc
	printerYn = "Y"

	
	Dim AP_CHA_MANAGER_NM , AP_CHA_STAFF_NM , AP_HD_MANAGER_NM , AP_HD_STAFF_NM
	Dim AP_CHA_MANAGER_ID , AP_CHA_STAFF_ID , AP_HD_MANAGER_ID , AP_HD_STAFF_ID
	sql = "EXEC ADM_APPROVE_DETAIL_VIEW '{P1}','{P2}','{P3}','{P4}';"
	sql = Replace(sql , "{P1}" , selDate , vbTextCompare)
	sql = Replace(sql , "{P2}" , "CART"    , vbTextCompare)
	sql = Replace(sql , "{P3}" , selHoType    , vbTextCompare)
	sql = Replace(sql , "{P4}" , "" , vbTextCompare)

	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		AP_CHA_MANAGER_NM	= nullFillWith(rs("CHA_MANAGER_NAME"),"",False)
		AP_CHA_STAFF_NM		= nullFillWith(rs("CHA_STAFF_NAME"),"",False)
		AP_HD_MANAGER_NM	= nullFillWith(rs("HD_MANAGER_NAME"),"",False)
		AP_HD_STAFF_NM		= nullFillWith(rs("HD_STAFF_NAME"),"",False)
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
	<input type="hidden" name="FLAG" id="FLAG" value="CART" title="플래그값" style="background-color:lime;">
	<input type="hidden" name="approveGrade" id="approveGrade" value="CART" title="결재" style="background-color:lime;">
	<input type="hidden" name="hospitalType" id="hospitalType" title="병원구분" style="background-color:lime;">
	<section class="section">
		<div class="head_container">
			<div style="display: flex;justify-content: space-between;">
				<h2>배선카트 일지</h2>
				<%
				If AP_CHA_MANAGER_NM = "" Or AP_CHA_STAFF_NM = "" Or AP_HD_MANAGER_NM = "" Or AP_HD_STAFF_NM = "" Then 
					printerYn = "N"
				End If 
				
				printerYn = "Y"
				If printerYn = "Y" Then %>
				<button class="button noprint" id="btnPrint" onClick="javascript:reportPrint();void(0);" style="width:100px;padding:0;height: 100%;">출력하기</button>
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
								%><small class="badge orange noprint" onClick="javascript:btnApprove('2');void(0);" style="cursor:pointer;">승인</small><%
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
								%><small class="badge orange noprint" onClick="javascript:btnApprove('3');void(0);" style="cursor:pointer;">승인</small><%
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
							%><small class="badge orange noprint" onClick="javascript:btnApprove('4');void(0);"" style="cursor:pointer;">승인</small><% 
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
			<div class="select noprint"  style="width:200px;margin-top:20px;height:100%;">
				<select id="selHoType" name="selHoType">
					<option value="">전체</option>
					<%=getCodeDB("select" , "P11" , selHoType)%>			
				</select>
			</div>
			<div class="date-picker">
				<div class="input date noprint" id="cal">
					<input type="text" name="selDate" id="selDate" placeholder="YYYY-MM-DD" readonly value="<%=selDate%>"/>
					<i class="bx bx-calendar"></i>
				</div>
				<span class="dateTxt print">일자 : <%=selDate%></span>
			</div>
		</div>
		</div>
		<div class="body_container">
			<div class="table">
				<table>
					<colgroup>
						<col width="8%">
						<col width="12%">
						<col width="20%">
						<col width="12%">
						<col width="12%">
						<col width="12%">
					</colgroup>
					<thead>
						<tr>
							<th>순서</th>
							<th>배선카트명</th>
							<th>구분</th>
							<th>조식</th>
							<th>중식</th>
							<th>석식</th>
						</tr>
					</thead>
					<tbody>
						<%
						Dim rowSpan , chkCartCode , CP_CHKIDX , CP_CNT , CP_CART_CODE , CP_CART_NAME , CP_CODE_NM
						Dim CP_BREAKFAST , CP_LUNCH , CP_DINNER
						seqNum = 1
						sql = "EXEC ADM_CART_RECORD_LIST_PRINT '{P1}','{P2}';"
						sql = Replace(sql , "{P1}" , selHoType , vbTextCompare)
						sql = Replace(sql , "{P2}" , selDate , vbTextCompare)
					    'rw(sql)
						Set rs = getQueryDB(sql)
						If Not(rs.Eof Or rs.Bof) Then 
							Do Until rs.Eof 							
								CP_CHKIDX		= rs("CHKIDX") 
								CP_CNT			= rs("CNT")
								CP_CART_CODE	= rs("CART_CODE")
								CP_CART_NAME	= rs("CART_NAME")	
								CP_CODE_NM		= rs("CODE_NM")
								CP_BREAKFAST	= rs("BREAKFAST")
								CP_LUNCH		= rs("LUNCH")
								CP_DINNER		= rs("DINNER")
						%>
						<tr>
							<td><%=seqNum%></td>
							<% If CP_CHKIDX = "1" Then %>
							<td rowspan="<%=CP_CNT%>"><%=rs("CART_NAME")%><BR/>(<%=rs("CART_CODE")%>)</td>
							<% End If %>
							<td><%=CP_CODE_NM%></td>
							<td <%=txtClass(CP_BREAKFAST)%>><%=getCode("name","CHECK_FLAG",nullFillWith(CP_BREAKFAST, "", False))%></td>
							<td <%=txtClass(CP_LUNCH)%>><%=getCode("name","CHECK_FLAG",nullFillWith(CP_LUNCH, "", False))%></td>
							<td <%=txtClass(CP_DINNER)%>><%=getCode("name","CHECK_FLAG",nullFillWith(CP_DINNER, "", False))%></td>
						</tr>
<%
		seqNum = seqNum + 1
		rs.MoveNext
	Loop 
End If 
rsClose(rs)

sql = "EXEC ADM_CART_RECORD_LIST_PRINT_ETC '{P1}','{P2}';"
sql = Replace(sql , "{P1}" , selHoType , vbTextCompare)
sql = Replace(sql , "{P2}" , selDate , vbTextCompare)
OP_ETC_DESC = ""
Set rs = getQueryDB(sql)
If Not(rs.Eof Or rs.Bof) Then 
	Do Until rs.Eof 
		If nullFillWith(rs("CODE_NM") , "" , False) <> "" Then 
			OP_ETC_DESC = OP_ETC_DESC & rs("CODE_NM") & " "
		End If 
		If nullFillWith(rs("ETC") , "" , False) <> "" Then 
			OP_ETC_DESC = OP_ETC_DESC & rs("ETC") & " " 
		End If 		
		rs.MoveNext
	Loop 

End If 
rsclose(rs)
OP_ETC_desc = Replace(OP_ETC_desc , "없음" , "")
%>
							<th colspan="2">개선 및 조치사항</th>
							<td colspan="4"><%=OP_ETC_desc%></td>
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

	$("#selHoType").change(function(){
		btnSubmit();
	});
});	//(e)_ready

function btnSubmit() {
	var frm = document.dataForm;
	
	frm.action = "./meal_cart_print.asp";
	frm.target = "_self";
	frm.method = "POST";
	frm.submit();		
}

function btnApprove(v) {
	var frm = document.dataForm;
	if(!setInputAlert(frm.selDate, "날짜를")) return false;
	if(!setInputAlert(frm.selHoType, "병원구분을")) return false;
	frm.hospitalType.value = $("#selHoType").val();
	frm.approveGrade.value = v;
	frm.command.value = "Approve";
	frm.action = "./approve_proc.asp";			
	frm.target = "procFrame";
	frm.method = "POST";
	frm.submit();	
}

function btnApproveCancel(v) {
	var frm = document.dataForm;
	if(!setInputAlert(frm.selDate, "날짜를")) return false;
	if(!setInputAlert(frm.selHoType, "병원구분을")) return false;
	frm.hospitalType.value = $("#selHoType").val();
	frm.approveGrade.value = v;
	frm.command.value = "ApproveCancel";
	frm.action = "./approve_proc.asp";			
	frm.target = "procFrame";
	frm.method = "POST";
	frm.submit();	
}

function reportPrint() {
//	$(".dateTxt").removeClass("hide");
//	$("#cal").addClass("hide");
//	$("#btnPrint").addClass("hide");
	window.print();
}

//-->
</script>
<iframe  name="procFrame" id="procFrame" style="width:0%; height:500px;display:none;" class="LEDB"></iframe>