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
	Dim sCommand, sIdx, sFilter, sKeyword, page
	sCommand = nullFillWith(Request("command"), "", False)
	sIdx = nullFillWith(Request("idx"), "", False)
	sFilter = nullFillWith(Request("filter"), "", False)
	sKeyword = nullFillWith(Request("keyword"), "", False)
	page	= nullFillWith(Request("page"),"1",false)

	Dim sql , rs, arrData , i 

	Dim listSize , pageSize , totCnt , totPage
	Dim vTemp , vNumber 
	listSize	= 10
	pageSize    = 10
	Session("Param") = "&filter=" & sFilter & "&keyword=" & sKeyword

	sql = "EXEC USP_DAILYMEAL_LIST_PAGEING '{P1}','{P2}','{P3}','{P4}';"
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
                  <a class="current" href="">배식서비스 체크리스트</a>
                </div>
                <div class="main-content">
                  <section class="section">
                    <div class="section-container">
                      <div class="section-title">
                        <h2>배식서비스 체크리스트</h2>
                      </div>
                      <div class="section-content">
                        <div class="table-search">
                          <form class="form" id="dataForm" name="dataForm">
						  <input type="hidden" name="idx" id="idx" value="" title="인덱스" style="background-color:lime;">
						  <input type="hidden" name="page" id="page" value="<%=page%>" title="인덱스" style="background-color:lime;">
						  <!--
                            <div class="select">
                              <select id="filter" name="filter">
                                <option value="NM" <%=selectValue("NM" , sFilter,"select" )%> >이름</option>
                                <option value="ID" <%=selectValue("ID" , sFilter,"select" )%> >아이디</option>
                              </select>
                            </div>
                            <div class="input" style="flex-grow: initial;">
                              <input type="search" id="keyword" name="keyword" placeholder="검색어를 입력하세요." style="width:300px;"/>
                            </div>

                            <button class="button outline" onclick="search()">검색</button>

							<button class="button" onclick="location.href='daily_meal_register.asp'" style="margin-left:auto;">등록하기</button>
							-->
                          </form>
                        </div>
                        <div class="table">
                          <table>
							<colgroup>
								<col width="10%"/>
								<!-- <col width="15%"/> -->
								<!-- <col width="15%"/> -->
								<col width="40%"/>
								<col width="20%"/>
							</colgroup>
                            <thead>
                              <tr>
                                <th class="fit-cell">번호</th>
                                <!-- <th class="fit-cell">담당자 아이디</th> -->
                                <!-- <th class="fit-cell">담당자 명</th> -->
                                <th class="fit-cell">구분항목</th>
                                <th class="fit-cell">비고</th>
                              </tr>
                            </thead>
                            <tbody>
							<%
								Dim OP_IDX, OP_ADMIN_ID , OP_ADMIN_NM, OP_ITEM_LIST, OP_ETC, ARR_ITEM_LIST, vI, vItemList, OP_ITEM_LIST_CODE
								IF isArray(arrData) Then
									vTemp = 1
									vNumber = totCnt - listSize * (page - 1)
									
									For i = 0 To UBound(arrData, 2)
										OP_IDX				= arrData(2 , i)
										OP_ADMIN_ID			= arrData(3 , i)
										OP_ADMIN_NM			= arrData(4 , i)
										OP_ITEM_LIST		= arrData(5 , i)
										OP_ETC				= arrData(6 , i)
										OP_ITEM_LIST_CODE	= arrData(7 , i)

										vItemList			= ""
										sql = "EXEC USP_CODE_TABLE_LIST '{P1}','{P2}';"
										sql = Replace(sql , "{P1}" , OP_ITEM_LIST_CODE , vbTextCompare)
										sql = Replace(sql , "{P2}" , ","			, vbTextCompare)

										Set rs = getQueryDB(sql)
										If Not(rs.Eof Or rs.Bof) Then 
											Do Until rs.Eof 
													vItemList = vItemList & "<br>" & rs(4)
												rs.MoveNext
											Loop 
										End If 
										rsClose(rs)
										
							%>
                              <tr onClick="move('<%=OP_IDX%>');" style="cursor:pointer;">
                                <td class="fit-cell"><%=vNumber%></td>
                                <td class="fit-cell" style="text-align:left;padding-left:20px;"><%=vItemList%></td>
                                <td class="fit-cell"><%=OP_ETC%></td>
                              </tr>							
							<%
										vNumber = vNumber - 1
										vTemp = vTemp + 1
									Next 
								Else 
									rw("<tr><td colspan=""5"">등록된 데이타가 없습니다</td></tr>")
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
		frm.action = "./daily_meal.asp";
		frm.target = "_self";
		frm.method = "POST";
		frm.submit();
	}

	function move(v) {
		$("#idx").val(v);
		var frm = document.dataForm;
		frm.action = "./daily_meal_register.asp";
		frm.target = "_self";
		frm.method = "POST";
		frm.submit();
	}

  </script>
  </body>
</html>
