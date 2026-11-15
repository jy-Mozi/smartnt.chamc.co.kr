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
	Dim sUserId, sFmDate, sToDate, page, sHospitalType
	sUserId			= nullFillWith(Request("userId"),"",False)
	sFmDate			= nullFillWith(request("fmDate"),Date()-7,False)
	sToDate			= nullFillWith(request("toDate"),Date(),False)
	page			= nullFillWith(Request("page"),"1",false)
	sHospitalType	= nullFillWith(Request("hospitalType"),"",False)

	Dim listSize , pageSize , totCnt , totPage
	Dim vTemp , vNumber 
	listSize	= 25
	pageSize    = 10
	Session("Param") = "&fmDate=" & sFmDate & "&toDate=" & sToDate& "&hospitalType=" & sHospitalType


	Dim sql , rs , i , arrUser, arrData

	Dim UP_USER_ID, UP_USER_NM , UP_EMPLOYEE_NUM
	sql = "EXEC USP_FOOD_LOTATION_USER_LIST;"
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrUser = rs.GetRows()
	End If
	rsClose(rs)

	Dim HP_IDX, HP_ROOM_CODE, HP_ROOM_NAME , arrRoom
	sql = "EXEC USP_CODE_COMMON_LIST 'P11';"
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrRoom	= rs.GetRows()
	End If 
	rsClose(rs)

	sql = "EXEC ADM_FOOD_LOTATION_RECORD_LIST_PAGEING_V2 '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}';"
	sql = Replace(sql , "{P1}", page , vbTextCompare)
	sql = Replace(sql , "{P2}", listSize , vbTextCompare)
	sql = Replace(sql , "{P3}", "ID" , vbTextCompare)
	sql = Replace(sql , "{P4}", sUserId , vbTextCompare)
	sql = Replace(sql , "{P5}", sFmDate , vbTextCompare)
	sql = Replace(sql , "{P6}", sToDate , vbTextCompare)
	sql = Replace(sql , "{P7}", sHospitalType , vbTextCompare)

	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrData	= rs.GetRows()
		totCnt		= arrData(0,0)			'총 계시물
		totPage		= Ceil(totCnt , listSize)	'총 페이지
	Else 
		totCnt		= 0
		totPage		= 1
	End If 

	%>
	<style type="text/css">
		.table-search {
		 & .form {
			.datepicker-controls {
				.button {
					width: auto;
					padding-inline: 0px;
					}		
				}
			}
		}

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
                  <a class="current" href="">배식서비스 체크리스트</a>
                </div>
                <div class="main-content">
                  <section class="section">
                    <div class="section-container">
                      <div class="section-title">
                        <h2>배식서비스 체크리스트 일지</h2>
                      </div>
                      <div class="section-content">
                        <div class="table-search">
                          <form class="form" name="dataForm" id="dataForm">
							<input type="hidden" name="command" id="command" value="FOOD_LOTATION" title="명령어:기본 부모정보코드등록" style="background-color:lime;">
							<input type="hidden" name="idx" id="idx" value="" title="인덱스" style="background-color:lime;">
							<input type="hidden" name="page" id="page" value="<%=page%>" title="페이지번호" style="background-color:lime;">
							<input type="hidden" name="currYmd" id="currYmd" title="선택 날짜">
                            <div class="select">
							<select id="hospitalType" name="hospitalType">
								<option value="">병동선택</option>										
								<%
								IF isArray(arrRoom) Then
									For i = 0 To UBound(arrRoom, 2)
								%>
								<option value="<%=arrRoom(2,i)%>" <%=selectvalue(arrRoom(2,i),sHospitalType,"select")%>><%=arrRoom(1,i)%></option>
								<%
									Next
								End If 
								%>
							</select>
                            </div>
							<div class="date-picker">
							  <div class="input date">
								<input type="text" name="fmDate" id="fmDate" placeholder="YYYY-MM-DD" readonly value="<%=sFmDate%>"/>
								<i class="bx bx-calendar"></i>
							  </div>
							</div>
							<div class="date-picker"">
							  <div class="input date">
								<input type="text" name="toDate" id="toDate"placeholder="YYYY-MM-DD" readonly value="<%=sToDate%>"/>
								<i class="bx bx-calendar"></i>
							  </div>
							</div>
                            <button type="button" class="button" id="btnSearch">검색</button>
                            <button type="button" class="button outline def">엑셀 다운로드</button>
							<button class="button red" onClick="javascript:reportPrint();void(0);">승인•출력하기</button>
                          </form>
                        </div>
                        <div class="table">
                          <table>
                            <thead>
                              <tr>
                                <th class="fit-cell">번호</th>
								<th class="fit-cell">점검날짜</th>
                                <th class="fit-cell">병원구분</th>
                                <th class="fit-cell">병동</th>
								<th class="fit-cell">담당자</th>
                              </tr>
                            </thead>
                            <tbody>
							<%

							Dim OP_CURR_YMD, OP_STAFF, OP_STAFF_NM, OP_REG_DATE, OP_REG_ID 
							Dim OP_REG_NM, OP_HOSPITAL_ROOM, OP_HOSPITAL_TYPE_NAME, OP_HOSPITAL_TYPE
							IF isArray(arrData) Then
							vTemp = 1
							vNumber = totCnt - listSize * (page - 1)
								For i = 0 To UBound(arrData, 2)
									OP_CURR_YMD				= arrData(2 , i)
									OP_STAFF				= arrData(6 , i)
									OP_STAFF_NM				= arrData(6 , i)
									'OP_REG_DATE				= arrData(5 , i)
									'OP_REG_ID				= arrData(6 , i)
									'OP_REG_NM				= arrData(7 , i)
									OP_HOSPITAL_ROOM		= arrData(5 , i)
									OP_HOSPITAL_TYPE		= arrData(3 , i)
									OP_HOSPITAL_TYPE_NAME   = arrData(4 , i)
							%>
                              <tr onClick="btnDetail('<%=OP_CURR_YMD%>','<%=OP_HOSPITAL_TYPE%>');" style="cursor:pointer;">
                                <td class="fit-cell"><%=vNumber%></td>
								<td class="fit-cell"><%=convertDate(OP_CURR_YMD,"-")%></td>
                                <td class="fit-cell"><%=OP_HOSPITAL_TYPE_NAME%></td>
                                <td class="fit-cell"><%=OP_HOSPITAL_ROOM%></td>
								<td class="fit-cell"><%=OP_STAFF%></td>
                              </tr>
                              </tr>
							<%
									vNumber = vNumber - 1
									vTemp = vTemp + 1
								Next 
							Else 
								rw("<tr><td colspan=""8"">등록된 데이타가 없습니다</td></tr>")
							End If 
							%>

                            </tbody>
                          </table>
                        </div>
                        <div class="table-control">
                          <div class="left-box">
							<%=getListPaging(page , Session("Param") , totCnt, totPage)%>
                          </div>
                          <div class="right-box">
                            <a  class="button" id="btnSubmit">등록하기</a>
                          </div>
                        </div>
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
		$("#btnSearch").click(function(){
			var frm = document.dataForm;			
			frm.action = "./daily_meal.asp";
			frm.target = "_self";
			frm.method = "POST";
			frm.submit();				
		});

		$("#btnSubmit").click(function(){
			var frm = document.dataForm;
			frm.page.value = "1";
			frm.idx.value = "0";
			frm.action = "./daily_meal_register.asp";			
			frm.target = "_self";
			frm.method = "POST";
			frm.submit();	
		});

		$("#selHoType").change(function(){
			$("#btnSearch").trigger("click");		
		});
	});		
	// End Ready

	function btnDetail(d,i) {
		var frm = document.dataForm;
		$("#currYmd").val(d);
		$("#hospitalType").val(i);
		frm.idx.value = "1";
		frm.page.value = "1";
		frm.action = "./daily_meal_register.asp";			
		frm.target = "_self";
		frm.method = "POST";
		frm.submit();	
	}

	//출력하기
	function reportPrint() {
		 openPopup('daily_meal_print.asp?selDate='+$("#fmDate").val() + "&hospitalType=" + $("#hospitalType").val() ,'DAILY_MEAL',800,700,'scrollbars=yes,width=800,height=700');
	}

//-->
</script>
<iframe  name="procFrame" id="procFrame" style="width:900px; height:500px;display:none;background-color:lime;"></iframe>