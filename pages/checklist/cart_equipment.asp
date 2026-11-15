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
	Dim sAreaCode , sFmDate , sToDate, page, selHoType
	sAreaCode	= nullFillWith(Request("areaCode"),"",False)
	sFmDate		= nullFillWith(request("fmDate"),Date() - weekDay(Date())+2, false)
	sToDate		= nullFillWith(request("toDate"),Date() , false)
	page		= nullFillWith(Request("page"),"1",false)
	selHoType		= nullFillWith(Request("selHoType"),"",False)

	Dim listSize , pageSize , totCnt , totPage
	Dim vTemp , vNumber 
	listSize	= 25
	pageSize    = 10
	Session("Param") = "&fmDate=" & sFmDate & "&toDate=" & sToDate& "&areaCode=" & sAreaCode

	Dim sql , rs , i , j , arrData , arrArea
	Dim AP_IDX, AP_AREA_CODE, AP_AREA_NAME, AP_ITEM_LIST , AP_ETC
	sql = "EXEC ADM_MACHINE_FACILITY_LIST_PAGEING '{P1}','{P2}','{P3}','{P4}','{P5}';"
	sql = Replace(sql , "{P1}", "1" , vbTextCompare)
	sql = Replace(sql , "{P2}", "1000" , vbTextCompare)
	sql = Replace(sql , "{P3}", "" , vbTextCompare)
	sql = Replace(sql , "{P4}", "" , vbTextCompare)
	sql = Replace(sql , "{P5}", selHoType , vbTextCompare)
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) then 
		arrArea = rs.GetRows()
	End If 
	rsClose(rs)


	Dim OP_CURR_YMD , OP_AREA_CODE , OP_AREA_NAME
	sql = "EXEC ADM_MACHINE_FACILITY_RECORD_LIST_PAGEING_V2 '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}';"
	sql = Replace(sql , "{P1}", page		, vbTextCompare)
	sql = Replace(sql , "{P2}", listSize	, vbTextCompare)
	sql = Replace(sql , "{P3}", "CD"		, vbTextCompare)
	sql = Replace(sql , "{P4}", sAreaCode	, vbTextCompare)
	sql = Replace(sql , "{P5}", sFmDate		, vbTextCompare)
	sql = Replace(sql , "{P6}", sToDate		, vbTextCompare)
	sql = Replace(sql , "{P7}", selHoType	, vbTextCompare)
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrData = rs.GetRows()
		totCnt		= arrData(0,0)			'총 계시물
		totPage		= Ceil(totCnt , listSize)	'총 페이지
	Else 
		totCnt		= 0
		totPage		= 1
	End If 
	rsClose(rs)
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
                  <a class="current" href="">기기설비</a>
                </div>
                <div class="main-content">
                  <section class="section">
                    <div class="section-container">
                      <div class="section-title">
                        <h2>기기설비 서비스 일지</h2>
                      </div>
                      <div class="section-content">
                        <div class="table-search">
                          <form class="form" name="dataForm" id="dataForm">
							<input type="hidden" name="command" id="command" value="Machine_Facility" title="명령어" class="LEB">
							<input type="hidden" name="idx" id="idx" value="" title="인덱스" class="LEB">
							<input type="hidden" name="page" id="page" value="<%=page%>" title="페이지번호" class="LEB">
							<input type="hidden" name="currDate" id="currDate" title="선택 날짜" class="LEB">
                            <div class="select">
							<select id="selHoType" name="selHoType">
								<option value="">전체</option>
								<%=getCodeDB("select" , "P11" , selHoType)%>			
							</select>
                            </div>
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
							<div class="date-picker">
							  <div class="input date">
								<input type="text" name="fmDate" id="fmDate" value="<%=sFmDate%>" placeholder="YYYY-MM-DD" readonly />
								<i class="bx bx-calendar"></i>
							  </div>
							</div>
							<div class="input"  style="flex-grow:0;display:none;">
                              <input type="text" name="toDate" id="toDate" placeholder="YYYY-MM-DD" value="<%=sToDate%>" readonly/>
							  <i class="bx bx-calendar"></i>
                            </div>
							<!-- <div class="date-picker"> -->
							  <!-- <div class="input noprint" style="width:100px;"> -->
								<!-- <input type="text" name="toDate" id="toDate" value="<%=sToDate%>" placeholder="YYYY-MM-DD" readonly /> -->
								<!-- <i class="bx bx-calendar"></i> -->
							  <!-- </div> -->
							<!-- </div> -->
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
								<th class="fit-cell">날짜</th>
                                <th class="fit-cell">구역번호</th>
                                <th class="fit-cell">구역명</th>
								<th class="fit-cell">구분</th>
                                <th class="fit-cell">점검</th>
								<th class="fit-cell">담당자</th>
                              </tr>
                            </thead>
                            <tbody>
							<%
							Dim OP_ITEM_CODE , OP_ITEM_NAME , OP_ITEM_RESULT
							Dim OP_CONFIRM_FLAG , OP_REG_ID , OP_REG_NM
							IF isArray(arrData) Then			
								vTemp = 1
								vNumber = totCnt - listSize * (page - 1)							
								For i = 0 To UBound(arrData, 2)
									
									OP_CURR_YMD		= arrData(2,i)
									OP_AREA_CODE	= arrData(3,i)
									OP_AREA_NAME	= arrData(4,i)
									OP_ITEM_CODE	= arrData(5,i)
									OP_ITEM_NAME	= arrData(6,i)
									OP_ITEM_RESULT	= arrData(7,i)
									OP_REG_ID		= arrData(8,i)
									OP_REG_NM		= arrData(9,i)
									OP_CONFIRM_FLAG	= arrData(10,i)
							
							%>
                              <tr onClick="btnDetail('<%=OP_CURR_YMD%>','<%=OP_AREA_CODE%>');" style="cursor:pointer;">
                                <td class="fit-cell" ><%=vNumber%></td>
                                <td class="fit-cell" ><%=convertDate(OP_CURR_YMD,"-")%></td>
                                <td class="fit-cell" ><%=OP_AREA_CODE%></td>
                                <td class="fit-cell" ><%=OP_AREA_NAME%></td>
                                <td class="fit-cell" ><%=OP_ITEM_NAME%></td>
                                <td class="fit-cell" ><%=getCode("name","CHECK_FLAG",OP_ITEM_RESULT)%></td>
                                <td class="fit-cell" ><%=OP_REG_NM%></td>
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
                            <button type="button" class="button" id="btnSubmit">등록하기</a>
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
	    var earDay = $("#fmDate").val();
		const elem = document.getElementById('fmDate');

		elem.addEventListener('changeDate', function(e) {
		  //console.log(e.detail.date); // new Date() 찍어봤을때랑 똑같은 형태로 나옴					
			var selDate = e.detail.date;
			var year = selDate.getFullYear();
			var month = selDate.getMonth() + 1;
				month = (month.toString().padStart(2, "0"));
			var day = selDate.getDate();
			var curDay = year + "-" + month + "-" + day;
			//console.log(year + "-" + month + "-" + day);
			//alert(year + "-" + month + "-" + day);
			$("#toDate").val(dateAdd(curDay , 6 , "d"));
		});


		$("#btnSearch").click(function(){
			var frm = document.dataForm;
			
			frm.action = "./cart_equipment.asp";
			frm.target = "_self";
			frm.method = "POST";
			frm.submit();				
		});

		$("#btnSubmit").click(function(){
			var frm = document.dataForm;
			frm.page.value = "1";
			frm.idx.value = "0";
			frm.action = "./cart_equipment_register.asp";			
			frm.target = "_self";
			frm.method = "POST";
			frm.submit();	
		});

		$("#selHoType , #areaCode").change(function(){
			$("#btnSearch").trigger("click");
		});
	});		
	// End Ready

	function btnDetail(d,i) {
		var frm = document.dataForm;
		$("#currDate").val(d);
		$("#areaCode").val(i);
		frm.idx.value = "1";
		frm.page.value = "1";
		frm.action = "./cart_equipment_register.asp";			
		frm.target = "_self";
		frm.method = "POST";
		frm.submit();	
	}
	
	//출력하기
	function reportPrint() {

		 openPopup('cart_equipment_print.asp?fmDate='+$("#fmDate").val()+'&toDate='+$("#toDate").val()+'&selHoType=' + $("#selHoType").val(),'CART_EQUIPMENT',1600,700,'scrollbars=yes,width=1600,height=700');
	}

	//ex) getDayOfWeek('2023-04-04') 의 형태로 사용
	function getDayOfWeek(yyyyMMdd){
		const dayOfWeek = new Date(yyyyMMdd).getDay();    
		//0:일, 1:월, 2:화, 3:수, 4:목, 5:금, 6:토
		return dayOfWeek;
	}


	window.onload = function() {		
		const datePickerElement = document.querySelector('.datepicker-input');
				
		// 기존 Datepicker 인스턴스가 있는지 확인하고 제거
		if (datePickerElement.datepicker) {
			datePickerElement.datepicker.destroy();
		}

		// datepicker 초기화
		const datepicker = new Datepicker(datePickerElement, {
			language: 'ko',
			autohide: true,
			orientation: 'left',
//			format: 'yyyy-mm-dd', // 입력 필드에 표시될 날짜 형식
			beforeShowDay: (date) => {
				const day = date.getDay();
				return {
					enabled: day === 1, // 월요일만 활성화
					classes: day === 1 ? 'monday-date' : 'disabled-date'
				};
			}
		});

		datePickerElement.datepicker = datepicker;
	}
//-->
</script>
<iframe  name="procFrame" id="procFrame" style="width:900px; height:500px;display:none;background-color:lime;"></iframe>