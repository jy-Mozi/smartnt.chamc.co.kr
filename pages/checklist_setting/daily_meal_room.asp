<%@CodePage=65001%>
<% 
'=============================================
' 작 성 자 : 송 쟁 영
' 작 성 일 : 2025-04-07 오후 1:44
' 파일내용 : 배식서비스 병동리스트
'=============================================
Option Explicit 
%>
<!-- #include virtual = "/common/library/common.asp"-->
<%
	Dim sCommand, sIdx, sFilter, sKeyword, page
	sCommand = nullFillWith(Request("command"), "", False)
	sIdx = nullFillWith(Request("idx"), "", False)
	sFilter = nullFillWith(Request("filter"), "", False)
	sKeyword = nullFillWith(Request("keyword"), "", False)
	page	= nullFillWith(Request("page"),"1",false)

	Dim sql , rs, arrData , i 

	Dim listSize , pageSize , totCnt , totPage
	Dim vTemp , vNumber 
	listSize	= 25
	pageSize    = 10
	Session("Param") = "&filter=" & sFilter & "&keyword=" & sKeyword

	sql = "EXEC USP_HOSPITAL_ROOM_LIST_PAGEING '{P1}','{P2}','{P3}','{P4}';"
	sql = Replace(sql , "{P1}", page , vbTextCompare)
	sql = Replace(sql , "{P2}", listSize , vbTextCompare)
	sql = Replace(sql , "{P3}", sFilter , vbTextCompare)
	sql = Replace(sql , "{P4}", sKeyword , vbTextCompare)
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
                        <h2>배식서비스 병동리스트</h2>
                      </div>
                      <div class="section-content">
                        <div class="table-search">
                          <form class="form" id="dataForm" name="dataForm">
						  <input type="hidden" name="idx" id="idx" value="" title="인덱스" style="background-color:lime;">
						  <input type="hidden" name="page" id="page" value="<%=page%>" title="인덱스" style="background-color:lime;">
                          </form>
                        </div>
                        <div class="table">
                          <table>
							<colgroup>
								<col width="10%"/>
								<col width="30%"/>
								<col width="30%"/>
								<col width="20%"/>
							</colgroup>
                            <thead>
                              <tr>
                                <th class="fit-cell">번호</th>
                                <th class="fit-cell">병동 코드</th>
                                <th class="fit-cell">병동 명</th>
                                <th class="fit-cell">병원구분</th>
                              </tr>
                            </thead>
                            <tbody>						
							<%
							Dim OP_IDX , OP_ROOM_CODE , OP_ROOM_NAME , OP_HOSPITAL_TYPE
							IF isArray(arrData) Then
								vTemp = 1
								vNumber = totCnt - listSize * (page - 1)								
								For i = 0 To UBound(arrData, 2)
									OP_IDX				= arrData(2 , i)
									OP_ROOM_CODE		= arrData(3 , i)
									OP_ROOM_NAME		= arrData(4 , i)
									OP_HOSPITAL_TYPE	= arrData(5 , i)
								
							%>
                              <tr onClick="move('<%=OP_IDX%>');" style="cursor:pointer;">
                                <td class="fit-cell"><%=vNumber%></td>
                                <td class="fit-cell"><%=OP_ROOM_CODE%></td>
                                <td class="fit-cell"><%=OP_ROOM_NAME%></td>
                                <td class="fit-cell"><%=getCodeDB("name" , "P11" , OP_HOSPITAL_TYPE)%></td>
                              </tr>	
							<%
									vNumber = vNumber - 1
									vTemp = vTemp + 1
								Next 
							End If 
							%>						
                            </tbody>
                          </table>
                        </div>
                        <div class="table-control">
                          <div class="left-box">
							<%=getListPaging(page , Session("Param") , totCnt, totPage)%>
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

  <script>
	function search() {
		var frm = document.dataForm;
		frm.action = "./daily_meal_room.asp";
		frm.target = "_self";
		frm.method = "POST";
		frm.submit();
	}

	function move(v) {
		$("#idx").val(v);
		var frm = document.dataForm;
		frm.action = "./daily_meal_room_register.asp";
		frm.target = "_self";
		frm.method = "POST";
		frm.submit();
	}

  </script>
  </body>
</html>
