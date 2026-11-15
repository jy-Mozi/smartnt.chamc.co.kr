<%
'//=================================================================
'작  성 자 : @전륜기
'작  성 일 : @2007-11-15
'파일 내용 : @리스트페이징 처리 함수
'=================================================================//
'* 기능 : List에서 페이징 처리를 반환한다.

Function getListPaging(fPage, fParameter, fTotalPage, fListSize)
	Dim fBody, fBlockPage, fI
	fBlockPage = int((fPage - 1) / fListSize) * fListSize + 1

	fBody = ""
	fBody = fBody &"<div class=""pagination"">"
	fBody = fBody &"<ul class=""pagination-list"">"

	If fBlockPage = 1 then  
		fBody = fBody &"<li class=""pagination-item"">"
		fBody = fBody &"  <button class=""control-button prev"">"
		fBody = fBody &"	<i class=""bx bx-chevron-left""></i>"
		fBody = fBody &"  </button>"
		fBody = fBody &"</li>"
	Else 
		fBody = fBody &"<li class=""pagination-item"">"
		fBody = fBody &"  <button class=""control-button prev""  onClick=""top.location.href='?page="& fBlockPage & fParameter &"'"">"
		fBody = fBody &"	<i class=""bx bx-chevron-left""></i>"
		fBody = fBody &"  </button>"
		fBody = fBody &"</li>"
	End if 

	If fTotalPage = 0 Then 
		fBody = fBody &"<li class=""pagination-item"">"
		fBody = fBody &"	<button class=""page-button current"">1</button>"
		fBody = fBody &"</li>"
	Else 
		fI = 1
		Do Until fI > fListSize or fBlockPage > fTotalPage
			IF fBlockPage = Int(fPage) Then
				fBody = fBody &"<li class=""pagination-item"">"
				fBody = fBody &"	<button class=""page-button current"">"& fBlockPage &"</button>"
				fBody = fBody &"</li>"
			Else
				fBody = fBody &"<li class=""pagination-item"">"
				fBody = fBody &"	<button class=""page-button"" onClick=""top.location.href='?page="& fBlockPage & fParameter &"'"">"& fBlockPage &"</button>"
				fBody = fBody &"</li>"
			End if

			fBlockPage = fBlockPage + 1
			fI = fI + 1
		Loop
	End if 

	IF fBlockPage > fTotalPage Then

		fBody = fBody &"<li class=""pagination-item"">"
		fBody = fBody &"<button class=""control-button next"">"
		fBody = fBody &"<i class=""bx bx-chevron-right""></i>"
		fBody = fBody &"</button>"
		fBody = fBody &"</li>"
	Else
		fBody = fBody &"<li class=""pagination-item"">"
		fBody = fBody &"<button class=""control-button next"" onClick=""top.location.href='?page="& fBlockPage-1 & fParameter &"'"">"
		fBody = fBody &"<i class=""bx bx-chevron-right""></i>"
		fBody = fBody &"</button>"
		fBody = fBody &"</li>"
	End if

		fBody = fBody &"</ul>" & vbCrLf
		fBody = fBody &"</div>" & vbCrLf
	getListPaging = fBody
End Function

%>
