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
<%
	Dim vIdx
		vIdx	= nullFillWith(Request("idx"), "" , False)
	
	Dim vSql, vRs, vI, vTxt
	
	' 담당자 리스트
	Dim vUserData
	vSql = "SELECT ADMIN_ID, ISNULL(ADMIN_NM,'') AS ADMIN_NM  FROM TBL_ADMIN WHERE USE_YN = 'Y' AND ADMIN_LEVEL = 'P01002'"
	Set vRs = getQueryDB(vSql)
	If Not(vRs.Eof Or vRs.Bof) Then
		vUserData = vRs.GetRows()
	End If 
	rsClose(vRs)


	' 구분항목 리스트
	Dim vItemData, vIn(0)
	vIn(0) = "P03"
	Set vRs = getSP("USP_CODE_COMMON_LIST",vIn)
	If Not(vRs.Eof Or vRs.Bof) Then
		vItemData = vRs.GetRows()
	End If 
	rsClose(vRs)

	
	' 수정일 경우 불러올 데이터
	Dim vUserId, vItemList, vEtc, arr_itemlist, arr_itemlistCode, vItemListNm
	If vIdx <> "" Then 
		vTxt = "수정"
		vSql = "SELECT USER_ID, ITEM_LIST, ETC, dbo.FUNC_STRING_COMMON_CODE(ITEM_LIST,',') AS ITEM_LIST_NM FROM TBL_FOOD_LOTATION WHERE USE_YN = 'Y' AND IDX = " & vIdx
		Set vRs = getQueryDB(vSql)
		If Not(vRs.Eof Or vRs.Bof) Then
			vUserId		= vRs("USER_ID")
			vItemList	= vRs("ITEM_LIST")
			vEtc		= vRs("ETC")
			vItemListNm = vRs("ITEM_LIST_NM")
		End If 
		rsClose(vRs)
	Else
		vTxt = "등록"
	End If 

%>
<!doctype html>
<html lang="ko">
  <head>
    <!-- #include virtual = "/common/include/layout_header.asp"-->
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
                  <a href="">점검항목 등록</a>
                  <i class="bx bx-chevron-right"></i>
                  <a class="current" href="">점검일지 등록</a>
                </div>
                <div class="main-content">
                  <section class="section">
                    <div class="section-container">
                      <div class="section-title">
                        <h2>배식서비스 체크리스트 등록/수정</h2>
                      </div>
                      <div class="section-content">
                        <hr class="divider" />
                        <form class="form">
						  <input type="hidden" id="idx" name="idx" value="<%=vIdx%>">
                          <fieldset class="input-group flex-box">
                            <div class="input-field item1-6" style="display:none;">
                              <label class="label">담당자</label>
                              <div class="select">							  
                                <select id="user" name="user">
								<option value="">선택</option>
								<% 
									If isArray(vUserData) Then 
										For vI = 0 To UBound(vUserData,2)
								%>
									<option value="<%=vUserData(0,vI)%>" <% If vUserId = vUserData(0,vI) Then rw("selected") End If %> ><%=vUserData(1,vI)%></option>
								<%
										Next
									End If 
								%>
                                </select>
                              </div>
                            </div>
                          </fieldset>
						 
						  <fieldset class="input-group flex-box">
							<div class="input-field">
							  <label class="label">구분항목</label>
							  <div class="flex-box x-small">
								<div class="select w-350">
								  <select id="item" name="item">
								  <option value="">선택</option>
								  <% 
									If isArray(vItemData) Then 
										For vI = 0 To UBound(vItemData,2)
											rw("<option value='" & vItemData(2,vI) & "'>" & vItemData(1,vI) & "</option>")
										Next
									End If 
								  %>
								  </select>
								</div>
								<div class="mw-100">
									<button type="button" class="button" onclick="itemAdd()">추가</button>
								</div>								
							  </div>
							  <div id="itemList" class="w-500 mt-30">
							  <%
								
								vSql = "EXEC USP_CODE_TABLE_LIST '{P1}','{P2}';"
								vSql = Replace(vSql , "{P1}" , vItemList , vbTextCompare)
								vSql = Replace(vSql , "{P2}" , ","			, vbTextCompare)
								Set vRs = getQueryDB(vSql)
								If Not(vRs.Eof Or vRs.Bof) Then 
									Do Until vRs.Eof 
									arr_itemlist = vRs(4)
									arr_itemlistCode = vRs(3)

							  %>
								<div class="area">
									<div class="flex-box between center my-10 px-20">
										<p><%=arr_itemlist%></p>
										<button type="button" class="button outline def w-auto" style="margin-left: 30px;" onclick="del(this,'<%=arr_itemlistCode%>')">삭제하기</button>
									</div>
									<hr class="divider my-10" alt="<%=vI%>"/>
								</div>
							  <%
										vI = vI + 1
										vRs.MoveNext
									Loop 
								End If 
								rsClose(vRs)
							  %>
							  </div>
							</div>
                          </fieldset>

                          <hr class="divider" />
						  <fieldset class="input-group flex-box">
                            <div class="input-field item1-3">
                              <label class="label">비고</label>
                              <div class="input">
                                <input id="etc" name="etc" type="text" placeholder="비고" value="<%=vEtc%>" />
                              </div>
                            </div>
                          </fieldset>
                          <hr class="divider" />
                          <div class="form-control">
                            <div class="right-box">
                              <a href="daily_meal.asp" class="button outline def">돌아가기</a>
                              <button class="button" onclick="register()"><%=vTxt%>하기</button>
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

  <script>
	
	var arr = new Array();
	$(document).ready(function() {
		// 아이템리스트 배열로 넣기
		var itemlist = "<%=vItemList%>";
		if (itemlist != "") {			
			var split_list = itemlist.split(",");
			for (var i=0;i<split_list.length;i++ ) {
				arr.push(split_list[i]);
			}
		}

		$(document).on("click", ".newdel", function() {
			if (arr.indexOf($(this).attr("alt")) >= 0) {
				arr.splice(arr.indexOf($(this).attr("alt")), 1);
				$(this).closest(".area").remove();
			}
		});
	});

	

	function itemAdd() {
		var itemVal = $("#item").val();
		var itemNm = $("#item option:checked").text();

		if (itemVal == "") {
			sAlert("warning","항목을 선택해주세요");
		} else {
			if (arr.indexOf(itemVal) >= 0) {
				sAlert("warning","이미 등록된 항목입니다");
			} else {
				var html = "";
				html += "<div class=\"area\">";
				html += "	<div class=\"flex-box between center my-10 px-20\">";
				html += "	<p>" + itemNm + "</p>";
				html += "	<button type=\"button\" class=\"button outline def w-auto newdel\" style=\"margin-left: 30px;\" alt=\"" + itemVal + "\">삭제하기</button>";
				html += "	</div>";
				html += "	<hr class=\"divider my-10\" />";
				html += "</div>";
									
				arr.push(itemVal);
				$("#itemList").append(html);
			}
		}		
	}

	function register() {
		if (arr.length == 0) {
			sAlert("warning","구분항목을 추가해주세요");
		} else {
			$.ajax({
				type : "GET",
				url : "proc.asp",
				data	: {
					command		:	"dailyMealRegister",
					idx			:	$("#idx").val(),
					user     	:	$("#user").val(),
					itemlist	:	arr.join(),
					etc			:	$("#etc").val()
				},
				success : function(json){
					var jsonData	= eval(json);
					var resultCd	= jsonData[0].resultCode;
					
					if (resultCd == "S") {
						sAlertMove("success","요청 성공!","m","daily_meal.asp");
					} else if (resultCd == "E") {
						sAlert("warning","이미 등록된 담당자입니다");
					}
				},
				error : function(request, status, error){
					sAlert("error","처리도중 에러가 발생했습니다");
					return false;
				}
			});		//END-AJAX
		}
	}
	
	// 항목삭제
	function del(e,code) {
		if (arr.indexOf(code) >= 0) {
			arr.splice(arr.indexOf(code), 1);
			$(e).closest(".area").remove();
		}
	}

  </script>
  </body>
</html>
