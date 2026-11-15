<%@CodePage=65001%>
<% 
'=============================================
' 작 성 자 : 송 쟁 영
' 작 성 일 : 2025-04-07
' 파일내용 : 배식서비스 병동리스트 상세보기
'=============================================
Option Explicit 
%>
<!-- #include virtual = "/common/library/common.asp"-->
<%
	Dim vIdx, sql , rs 
	Dim OP_ROOM_CODE , OP_ROOM_NAME , OP_HOSPITAL_TYPE , OP_USE_YN
	vIdx	= nullFillWith(Request("idx"), "" , False)
	sql = "EXEC USP_HOSPITAL_ROOM_DETAIL_VIEW '{P1}';"
	sql = Replace(sql , "{P1}" , vIdx , vbTextCompare)
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		OP_ROOM_CODE		= rs("ROOM_CODE")
		OP_ROOM_NAME		= rs("ROOM_NAME")
		OP_HOSPITAL_TYPE	= rs("HOSPITAL_TYPE")
		OP_USE_YN			= rs("USE_YN")
	End If 
	rsClose(rs)
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
                  <a class="current" href="">배식서비스 병동리스트</a>
                </div>
                <div class="main-content">
                  <section class="section">
                    <div class="section-container">
                      <div class="section-title">
                        <h2>배식서비스 병동리스트 등록/수정</h2>
                      </div>
                      <div class="section-content">
                        <hr class="divider" />
                        <form class="form" name="dataForm" id="dataForm">
						  <input type="hidden" id="command" name="command" value="meal_room_register">
						  <input type="hidden" id="idx" name="idx" value="<%=vIdx%>">
						  <input type="hidden" id="hospitalRoomCode" name="hospitalRoomCode" value="<%=OP_ROOM_CODE%>">
						  
						  <fieldset class="input-group flex-box">
                            <div class="input-field item1-5">
                              <label class="label">병원구분</label>
								<div class="select w-350">
								  <select id="hospitalType" name="hospitalType">
								  <option value="">선택</option>
								  <%=getCodeDB("select" , "P11" , OP_HOSPITAL_TYPE)%>
								  </select>
								</div>
                            </div>
                          </fieldset>
                          <hr class="divider" />
						  <fieldset class="input-group flex-box">
                            <div class="input-field item1-3">
                              <label class="label">병동명</label>
                              <div class="input">
                                <input id="hospitalRoom" name="hospitalRoom" type="text" placeholder="병동명" value="<%=OP_ROOM_NAME%>" readonly/>
                              </div>
                            </div>
                          </fieldset>						 

                          <hr class="divider" />
                          <div class="form-control">
                            <div class="right-box">
                              <a href="daily_meal_room.asp" class="button outline def">돌아가기</a>
                              <button class="button" onclick="register()">수정하기</button>
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

	$(document).ready(function() {
	
		
	});

	//수정하기
	function register() {
		var frm = document.dataForm;
		if(!setInputAlert(frm.hospitalRoom, "병동명을")) return false;
		if(!setSelectAlert(frm.hospitalType, "병원구분을")) return false;
		frm.action = "./proc.asp";
		frm.target = "procFrame";
		frm.method = "POST";
		frm.submit();	
	}
	
  </script>
  </body>
</html>
<iframe  name="procFrame" id="procFrame" style="width:900px; height:500px;display:none;background-color:lime;"></iframe>