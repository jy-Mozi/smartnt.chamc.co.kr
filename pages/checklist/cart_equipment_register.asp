<%@CodePage=65001%>
<% 
'=============================================
' 작 성 자 : 노 광 옹
' 작 성 일 : 2024-10-17
' 파일내용 : 
'=============================================
Option Explicit 
%>
<!-- #include virtual = "/common/library/common.asp"-->
<!doctype html>
<html lang="ko">
  <head>
    <!-- #include virtual = "/common/include/layout_header.asp"-->
	<%
	Dim sCommand, sIdx, sFmDate, sPage, sCurrDate, sAreaCode, sToDate
	sCommand = nullFillWith(Request("command"), "", False)
	sIdx = nullFillWith(Request("idx"), "0", False)
	sFmDate = nullFillWith(Request("fmDate"), "", False)
	sPage = nullFillWith(Request("page"), "", False)
	sCurrDate = nullFillWith(Request("currDate"), Date(), False)
	sAreaCode = nullFillWith(Request("areaCode"), "", False)
	sToDate = nullFillWith(Request("toDate"), "", False)

	Dim sql , rs , i , j , arrData , arrArea , arrComment , arrManager
	Dim AP_IDX, AP_AREA_CODE, AP_AREA_NAME, AP_ITEM_LIST , AP_ETC
	sql = "EXEC USP_MACHINE_FACILITY_LIST_PAGEING '{P1}','{P2}','{P3}','{P4}';"
	sql = Replace(sql , "{P1}", "1" , vbTextCompare)
	sql = Replace(sql , "{P2}", "1000" , vbTextCompare)
	sql = Replace(sql , "{P3}", "" , vbTextCompare)
	sql = Replace(sql , "{P4}", "" , vbTextCompare)
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) then 
		arrArea = rs.GetRows()
	End If 
	rsClose(rs)


	Dim VP_IDX, VP_CODE_NM,  VP_CART_VAL
	sql = "EXEC USP_CODE_COMMON_LIST 'P08';"
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrComment = rs.GetRows()
	End If 
	rsClose(rs)

	sql = "EXEC USP_ADMIN_LIST_PAGEING '{P1}','{P2}','{P3}','{P4}','{P5}','','';"
	sql = Replace(sql , "{P1}", "1" , vbTextCompare)
	sql = Replace(sql , "{P2}", "1000" , vbTextCompare)
	sql = Replace(sql , "{P3}", "" , vbTextCompare)
	sql = Replace(sql , "{P4}", "" , vbTextCompare)
	sql = Replace(sql , "{P5}", "" , vbTextCompare)
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrManager = rs.GetRows()
	End If 
	rsClose(rs)

	Dim OP_IDX, OP_CURR_YMD, OP_AREA_CODE, OP_ITEM_CODE, OP_ITEM_CODE_NM
	Dim OP_ITEM_CODE_RESULT, OP_ETC_CODE, OP_ETC, OP_CONFIRM_FLAG, OP_REG_ID	
	Dim arrEtc(2) , arrManagerId(2) , arrTakeCd(2)

	sql = "EXEC USP_MACHINE_FACILITY_RECORD_DETAIL_VIEW_JOIN  '{P1}','{P2}';"
	sql = Replace(sql , "{P1}" , sCurrDate , vbTextCompare)
	sql = Replace(sql , "{P2}" , sAreaCode , vbTextCompare)
	
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrData = rs.GetRows()
	End If 
	rsClose(rs)

	%>

	<style type="text/css">
		.bx-calendar {
			position: absolute;
			right: 10px;
			pointer-events: none;
		}
	</style>
  </head>
  <body>
    <div id="root">
      <div id="app">
        <header id="header">
			<!-- #include virtual = "/common/include/layout_left.asp"-->
        </header>
        <div id="wrap">
          <div class="scroll-area">
            <main id="main">
              <div class="main-container">
                <!-- #include virtual = "/common/include/layout_top.asp"-->
                <div class="breadcrumb">
                  <a href=""><i class="bx bx-home-alt-2"></i></a>
                  <i class="bx bx-chevron-right"></i>
                  <a href="">점검일지 등록</a>
                  <i class="bx bx-chevron-right"></i>
                  <a class="current" href="">기기설비 서비스 일지 등록/수정</a>
                </div>
                <div class="main-content">
                  <section class="section">
                    <div class="section-container">
                      <div class="section-title">
                        <h2>기기설비 서비스 일지 등록/수정</h2>
                      </div>
                      <div class="section-content">
                        <hr class="divider" />
                        <form class="form" name="dataForm" id="dataForm">
						<input type="hidden" name="command" id="command" value="machine_equipment_Data" title="명령어:일일배식데이타" class="LEB">
						<input type="hidden" name="idx" id="idx" value="<%=sIdx%>" title="페이지번호"  class="LEB">
						<input type="hidden" name="page" id="page" value="<%=sPage%>" title="페이지번호"  class="LEB">
						<input type="hidden" name="confirm" id="confirm" title="승인여부" class="LEB">
                          <fieldset class="input-group flex-box">
                            <div class="input-field item1-5">
                              <label class="label">구역번호</label>
								<div class="select">
									<select name="areaCode" id="areaCode">
										<option value="">구역번호(전체)</option>
										<%
										IF isArray(arrArea) Then
											For i = 0 To UBound(arrArea, 2)
												AP_IDX			= arrArea(2 , i)
												AP_AREA_CODE	= arrArea(3 , i)
												AP_AREA_NAME	= arrArea(4 , i)
												AP_ITEM_LIST	= arrArea(5 , i)
												AP_ETC			= arrArea(6 , i)
												%><option value="<%=AP_AREA_CODE%>" <%=selectValue(AP_AREA_CODE,sAreaCode,"select")%>><%=AP_AREA_NAME%>&nbsp;(<%=AP_AREA_CODE%>)</option><%	
											Next 
										End If 
										%>
									</select>
								</div>
                            </div>
                          </fieldset>
                          <hr class="divider" />	


						  <fieldset class="input-group flex-box">
                            <div class="input-field item1-6">
                              <label class="label">날짜</label>
							<div class="date-picker">
							  <div class="input date">
								<input type="text" name="currDate" id="currDate" placeholder="YYYY-MM-DD" readonly value="<%=convertDate(sCurrDate,"-")%>"/>
								<i class="bx bx-calendar"></i>
							  </div>
							</div>
							</div>
                          </fieldset>
						  <hr class="divider" />
					  
						  <fieldset class="input-group flex-box" <%If sAreaCode = "" Then rw("style=""display:none;"")")  End If %>>
							<div class="input-field">
							  <label class="label">구분 리스트</label>
							  <div class="table">
								  <table>
									<thead>
									  <tr>
										<th class="fit-cell">구역</th>
										<th class="fit-cell">구분</th>
										<th class="fit-cell">점검</th>
									  </tr>
									</thead>
									<tbody>
									<%
									IF isArray(arrData) Then
										For i = 0 To UBound(arrData, 2)
											OP_IDX				= nullFillWith(arrData(0,i),"", False)
											OP_CURR_YMD			= nullFillWith(arrData(1,i),"", False)
											OP_AREA_CODE		= nullFillWith(arrData(2,i),"", False)
											OP_ITEM_CODE		= nullFillWith(arrData(3,i),"", False)
											OP_ITEM_CODE_NM		= nullFillWith(arrData(4,i),"", False)
											OP_ITEM_CODE_RESULT	= nullFillWith(arrData(5,i),"", False)
											OP_ETC_CODE			= nullFillWith(arrData(6,i),"", False)
											OP_ETC				= nullFillWith(arrData(7,i),"", False)
											OP_CONFIRM_FLAG		= nullFillWith(arrData(8,i),"", False)
											OP_REG_ID			= nullFillWith(arrData(9,i),"", False)

									%>
										<input type="hidden" name="itemIdx" id="itemIdx" value="<%=OP_IDX%>" title="인덱스" class="LEDB">
										<input type="hidden" name="itemCode" id="itemCode" value="<%=OP_ITEM_CODE%>" title="아이템코드" class="LEDB">
									  <tr>
										<td class="fit-cell"><%=OP_AREA_CODE%></td>
										<td class="fit-cell"><%=OP_ITEM_CODE_NM%></td>
										<td class="fit-cell">
												<div class="select">
													<select id="selResult" name="selResult">
														<option value="">선택</option>
														<%=getCode("select","CHECK_FLAG",OP_ITEM_CODE_RESULT)%>
													</select>
												</div>	
										</td>
										
									  </tr>
									<%
										Next 
									Else 
										rw("<tr><td colspan=""5"">등록된 데이타가 없습니다</td></tr>")
									End If 
									%>		
									
									</tbody>
								  </table>
							  </div>
							</div>
                          </fieldset>
						  <br/><br/>

						  <br/><br/>
						  <fieldset class="input-group flex-box">
							<div class="input-field">
							  <label class="label">특이사항</label>
							  <div class="table">
								  <table>
									<thead>
									  <tr>
										<th class="fit-cell">조치사항</th>
										<th class="fit-cell">조치내용</th>
										<th class="fit-cell">점검자</th>
									  </tr>
									</thead>
									<tbody>
									<tr>
										<td class="fit-cell">                             
											<div class="select">
											  <select name="selEtc" id="selEtc">
												<option value="">선택</option>
												<%												
												IF isArray(arrComment) Then
													For i = 0 To UBound(arrComment, 2)
														VP_IDX			= arrComment(0 , i)
														VP_CODE_NM	= arrComment(1 , i)
														VP_CART_VAL	= arrComment(2 , i)
												%><option value="<%=VP_CART_VAL%>" <%=selectValue(VP_CART_VAL,OP_ETC_CODE,"select")%>><%=VP_CODE_NM%></option><%	
													Next 
												End If 
												%>
											  </select>
											</div>
						                </td>
										<td class="fit-cell">                             
											<div class="input">
												<input id="etc" name="etc" type="text" placeholder="조치내용" value="<%=OP_ETC%>" />
											</div>
						                </td>
										<td class="fit-cell">                             
											<div class="select">
											  <select name="selRegId" id="selRegId">
											  <option value="">선택</option>
												<%												
												IF isArray(arrManager) Then
													For i = 0 To UBound(arrManager, 2)
												%><option value="<%=arrManager(3 , i)%>" <%=selectValue(arrManager(3 , i),OP_REG_ID,"select")%>><%=arrManager(4 , i)%>&nbsp;(<%=arrManager(3 , i)%>)</option><%	
													Next 
												End If 
												%>
											  </select>
											</div>
						                </td>
									</tr>										
									</tbody>
								  </table>
							  </div>
							</div>
                          </fieldset>		


						  <%
							Dim AP_CHA_MANAGER , AP_CHA_STAFF , AP_HD_MANAGER , AP_HD_STAFF
							sql = "EXEC ADM_APPROVE_DETAIL_VIEW '{P1}','{P2}','','';"
							sql = Replace(sql , "{P1}" , Replace(sCurrDate,"-","") , vbTextCompare)
							sql = Replace(sql , "{P2}" , "MACHINE" , vbTextCompare)

							Set rs = getQueryDB(sql)
							If Not(rs.Eof Or rs.Bof) Then 
								AP_CHA_MANAGER	 = nullFillWith(rs("CHA_MANAGER"),"",False)
								AP_CHA_STAFF	 = nullFillWith(rs("CHA_STAFF"),"",False)
								AP_HD_MANAGER	 = nullFillWith(rs("HD_MANAGER"),"",False)
								AP_HD_STAFF		 = nullFillWith(rs("HD_STAFF"),"",False)
							Else 
								AP_CHA_MANAGER	= ""
								AP_CHA_STAFF	= ""
								AP_HD_MANAGER	= ""
								AP_HD_STAFF		= ""
							End If 
							rsClose(rs)
						  %>
						  <fieldset class="input-group flex-box" style="margin-top:40px;display:none;">
                            <div class="input-field">
                              <label class="label">결재</label>
								<div class="table">
									<table>
										<thead>
											<tr>
												<th class="fit-cell" colspan="2">분당차병원</th>
												<th class="fit-cell" colspan="2" style="border-left:1px solid #cdd0d4;">현대그린푸드</th>
											</tr>
											<tr>
											<td class="fit-cell">팀장</td>
											<td class="fit-cell" style="border-left:1px solid #cdd0d4;">팀원</td>
											<td class="fit-cell" style="border-left:1px solid #cdd0d4;">팀장</td>
											<td class="fit-cell" style="border-left:1px solid #cdd0d4;">팀원</td>
											</tr>
										</thead>
									<tbody>
										<tr>
											<td class="fit-cell">
												<%
												If AP_CHA_MANAGER = "" And vAdminCompany = "CHA" And vAdminPosition = "MANAGER" Then 
												%>
													<button class="button orange" onClick="javascript:btnApprove('4');void(0);">승인하기</button>
												<% 
												End If 
												If AP_CHA_MANAGER <> "" Then 
													rw("<button class=""button dark"" >승인완료 : "&AP_CHA_MANAGER&"</button>")
												End If 
												%>
											</td>
											<td class="fit-cell" style="border-left:1px solid #cdd0d4;">
												<%
												If AP_CHA_STAFF = "" And vAdminCompany = "CHA" And vAdminPosition = "STAFF"Then 
													rw("<button class=""button orange"" onClick=""javascript:btnApprove('3');void(0);"">승인하기</button>")
												End If 

												If AP_CHA_STAFF <> "" Then 
													rw("<button class=""button dark"" >승인완료 : "&AP_CHA_STAFF&"</button>")
												End If 
												%>
											</td>
											<td class="fit-cell" style="border-left:1px solid #cdd0d4;">
												<%
												If AP_HD_MANAGER = "" And vAdminCompany = "HD" And vAdminPosition = "MANAGER" Then 
													rw("<button class=""button sub"" onClick=""javascript:btnApprove('2');void(0);"">승인하기</button>")
												End If 

												If AP_HD_MANAGER <> "" Then 
													rw("<button class=""button dark"" >승인완료 : "&AP_HD_MANAGER&"</button>")
												End If 
												%>
											</td>
											<td class="fit-cell" style="border-left:1px solid #cdd0d4;">
												<%
												If AP_HD_STAFF = "" And vAdminCompany = "HD" And vAdminPosition = "STAFF" Then 
													rw("<button class=""button sub"" onClick=""javascript:btnApprove('1');void(0);"">승인하기</button>")
												End If 

												If AP_HD_STAFF <> "" Then 
													rw("<button class=""button dark"" >승인완료 : "&AP_HD_STAFF&"</button>")
												End If 
												%>
											</td>
										</tr>
									</tbody>
									</table>
								</div>

                            </div>
                          </fieldset>
                
                          <hr class="divider" />
                          <div class="form-control">
                            <div class="right-box">
                              <a href="cart_equipment.asp" class="button outline def">돌아가기</a>
                              <button type="button" class="button" id="btnSubmit" style="display:block;">등록하기
							  <% If sIdx <> "0" Then %>
                              <button type="button" class="button orange" id="btnRemove" style="display:block;">삭제하기</button>
							  <% End If %>							  </button>
							  <%If OP_CONFIRM_FLAG = "2" Then %>
							  <!-- <button class="button red" onClick="javascript:chgConfirm('1');void(0);">승인하기</button> -->
							  <% Else %>
							  <!-- <button class="button dark" onClick="javascript:chgConfirm('2');void(0);">승인취소</button> -->
							  <% End If %>
                            </div>
                          </div>
                        </form>
                      </div>
                    </div>
                  </section>
                </div>
              </div>
            </main>
            <!-- #include virtual = "/common/include/layout_footer.asp"-->
          </div>
        </div>
      </div>
    </div>
  </body>
</html>
<script type="text/javascript">
<!--
$(document).ready(function(){
	const elem = document.getElementById('currDate');
	const rangepicker = new DateRangePicker(elem, {});

	elem.addEventListener('changeDate', function(e) {
	  //console.log(e.detail.date); // new Date() 찍어봤을때랑 똑같은 형태로 나옴
		var selDate = e.detail.date;
		var year = selDate.getFullYear();
		var month = (selDate.getMonth() + 1).toString().padStart(2,'0');
		var day = selDate.getDate();
		//console.log(year + "-" + month + "-" + day);
		$("#areaCode").change();
	});

	$("#btnSubmit").click(function(){
		var frm = document.dataForm;

		if(!setInputAlert(frm.currDate, "날짜를")) return false;
		if(!setSelectAlert(frm.areaCode, "구역번호를")) return false;

		frm.page.value = "1";
		frm.action = "./cart_equipment_proc.asp";
		frm.target = "procFrame";
		frm.method = "POST";
		frm.submit();	
	});

	$("#areaCode").change(function(){
		var frm = document.dataForm;

		if(!setInputAlert(frm.currDate, "날짜를")) return false;
		if(!setSelectAlert(frm.areaCode, "구역번호를")) return false;
		frm.idx.value = "0";
		frm.page.value = "1";
		frm.action = "./cart_equipment_register.asp";
		frm.target = "_self";
		frm.method = "POST";
		frm.submit();	
	});

	// 삭제하기
	$("#btnRemove").click(function(){
		Swal.fire({
		  title: "삭제하시겠습니까?",
		  text: "삭제된 데이타는 복구가 불가능합니다!",
		  icon: "question",
		  showCancelButton: true,
		  confirmButtonColor: "#3085d6",
		  cancelButtonColor: "#d33",
		  confirmButtonText: "삭제",
  		  cancelButtonText: "취소"
		}).then((result) => {
		  if (result.isConfirmed) {
			var frm = document.dataForm;
			if(!setSelectAlert(frm.areaCode, "구역번호를")) return false;
			frm.command.value = "cart_equipment_Remove";
			frm.action = "./cart_equipment_proc.asp";			
			frm.target = "procFrame";
			frm.method = "POST";
			frm.submit();	
		  }
		});	
	});
});		
// End Ready

function btnApprove(v) {
	if ($("#idx").val() == "0") {
		sAlert("error","등록 후 승인처리 가능합니다.")
		return false;
	}
	var frm = document.dataForm;
	frm.confirm.value = v;
	frm.command.value = "Machine_Equipment_Approve";
	frm.action = "./cart_equipment_proc.asp";			
	frm.target = "procFrame";
	frm.method = "POST";
	frm.submit();	
}


function chgConfirm(v) {
	var frm = document.dataForm;
	frm.confirm.value = v;
	frm.command.value = "Machine_Equipment_Confirm";
	frm.action = "./cart_equipment_proc.asp";			
	frm.target = "procFrame";
	frm.method = "POST";
	frm.submit();	
}

//-->
</script>
<iframe  name="procFrame" id="procFrame" style="width:900px; height:500px;display:none;background-color:lime;"></iframe>