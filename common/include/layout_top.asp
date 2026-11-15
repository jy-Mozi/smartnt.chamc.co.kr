<div class="top-bar">
	<div class="menu-group">
		<ul>
			<li><a href="../../pages/setting/user_list.asp">설정</a></li>
			<li><a href="../../pages/checklist_setting/meal_cart.asp">점검항목 등록</a></li>
			<li><a href="../../pages/checklist/meal_cart.asp">일지 등록</a></li>
			<li>통계</li>
		</ul>
	</div>
	<nav class="top-nav">		
		<div class="info-group">
			<div id="setting" class="setting">
				<div class="button" style="cursor:auto;"><%=getCodeDB("name","P09",Session("SNT_ADMIN_COMPANY"))%>&nbsp;&nbsp;
				<%=getCodeDB("name" , "P11",Session("SNT_HOSPITAL_TYPE"))%>&nbsp;&nbsp;<%=Session("SNT_ADMIN_POSITION_STR")%>&nbsp;&nbsp;<%=Session("SNT_ADMIN_ID")%>(<%=Session("SNT_ADMIN_NM")%>)</div>
				<ul class="setting-list">
				  <li class="setting-item"><button type="button" class="button">비밀번호 변경</button></li>
				</ul>
			</div>
			<a href="/pages/login/proc.asp?command=Logout_Chk" class="button">로그아웃<i class="bx bx-log-in"></i></a>
		</div>
	</nav>
</div>