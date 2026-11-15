<%
'//=================================================================
'작  성 자 : @전윤기
'작  성 일 : @2007-11-15
'파일 내용 : @코드 모음
'=================================================================//
'* 기   능 : Dictionary객체에서 코드상 데이터를 원하는 형식으로 리턴
Function getCode(fMode, fItem, fVal)
	Dim fDic,fKeys,fItems,fBody,fData() ,fI
	Set fDic = Server.CreateObject("Scripting.Dictionary")
		Select Case fItem
			Case "CONFIRM_FLAG"
				fDic.Add "1", "승인"
				fDic.Add "2", "반려"
				fDic.Add "0", "미승인"

			Case "CHECK_FLAG" 
				fDic.Add "1", "양호"
				fDic.Add "2", "미흡"			
				fDic.Add "3", "불량"

			Case "CHECK_FLAG_2" 
				fDic.Add "1", "양호"
				fDic.Add "2", "미흡"
				
			Case "MONTH"
				fDic.Add "01", "01"						
				fDic.Add "02", "02"
				fDic.Add "03", "03"
				fDic.Add "04", "04"
				fDic.Add "05", "05"
				fDic.Add "06", "06"
				fDic.Add "07", "07"
				fDic.Add "08", "08"
				fDic.Add "09", "09"
				fDic.Add "10", "10"
				fDic.Add "11", "11"
				fDic.Add "12", "12"

			Case "DAY"
				fDic.Add "01", "01"						
				fDic.Add "02", "02"						
				fDic.Add "03", "03"						
				fDic.Add "04", "04"						
				fDic.Add "05", "05"						
				fDic.Add "06", "06"						
				fDic.Add "07", "07"						
				fDic.Add "08", "08"						
				fDic.Add "09", "09"
				fDic.Add "10", "10"
				fDic.Add "11", "11"
				fDic.Add "12", "12"
				fDic.Add "13", "13"
				fDic.Add "14", "14"
				fDic.Add "15", "15"
				fDic.Add "16", "16"
				fDic.Add "17", "17"
				fDic.Add "18", "18"
				fDic.Add "19", "19"
				fDic.Add "20", "20"
				fDic.Add "21", "21"
				fDic.Add "22", "22"
				fDic.Add "23", "23"
				fDic.Add "24", "24"
				fDic.Add "25", "25"
				fDic.Add "26", "26"
				fDic.Add "27", "27"
				fDic.Add "28", "28"
				fDic.Add "29", "29"
				fDic.Add "30", "30"
				fDic.Add "31", "31"
			Case "WEEK"
				fDic.Add "1", "1"
				fDic.Add "2", "2"
				fDic.Add "3", "3"
				fDic.Add "4", "4"
				fDic.Add "5", "5"

         
           
		End Select

		fKeys	= fDic.Keys
		fItems	= fDic.Items

		fVal = Cstr(fVal)	'자료형을 문자로
		Select Case fMode
			Case "name"
				getCode = fDic.Item(fVal)
			Case "select"
				For fI = 0 To UBound(fKeys)
					fBody = fBody &"<option value="""& fKeys(fI) &""" "& isSelected(fKeys(fI),fVal) &">"& fItems(fI) &"</option>"
				Next
				getCode = fBody
			Case "checkbox", "radio"
				For fI = 0 To UBound(fKeys)
					fBody = fBody &"<label><input type="""& fMode &""" name="""& fItem &""" value="""& fKeys(fI) &""" "& isChecked(fKeys(fI), fVal) &" class=""radio"">"& fItems(fI) &"</label>&nbsp;"
				Next
				getCode = fBody
			Case "array"	'2차원 배열로 리턴한다.
				ReDim fData(1,UBound(fKeys))
				For fI = 0 To UBound(fKeys)
					fData(0,fI) = fItems(fI)
					fData(1,fI) = fKeys(fI)
				Next
				getCode = fData
			Case "li"
				For fI = 0 To UBound(fKeys)
					fBody = fBody &"<li class=""option"" data-value=""" & fKeys(fI) & """"& isCurrentData(fKeys(fI), fVal) &" >" & fItems(fI) & "</li>"
				Next
				getCode = fBody
			Case Else
			getCode = "코드모드 오류!!"
		End Select
	Set fDic = Nothing
End Function

'* 기   능 : 코드테이블(DAS_CODE)에서 코드상 데이터를 원하는 형식으로 리턴
Function getCodeDB(fMode, fItem, fVal)
	IF Not isNone(fMode) and Not isNone(fItem) Then
		Dim fRs, fData, fI, fIn(0)
		fIn(0) = fItem
		Set fRs = getSP("USP_CODEBOOK_GROUP", fIn)
		IF Not(fRs.Bof or fRs.Eof) Then
			fData = fRs.getRows()
			IF isArray(fData) Then
				Dim fCode, fValue, fBody, fSelected , fGroupNm , fViewFlag
				Select Case fMode
					Case "name"
						For fI = 0 To UBound(fData, 2)
							fCode	= fData(0, fI)
							fValue	= fData(1, fI)
							IF fVal = fCode Then
								getCodeDB = fValue
								Exit For
							End if
						Next
					Case "rename"
						For fI = 0 To UBound(fData, 2)
							fCode	= fData(0, fI)
							fValue	= fData(1, fI)
							IF fVal = fValue Then
								getCodeDB = fCode
								Exit For
							End if
						Next
					Case "select"
						fBody = ""
						For fI = 0 To UBound(fData, 2)
							fCode	    = fData(0, fI)
							fValue	    = fData(1, fI)
                            fSelected   = fData(2, fI)
                            fViewFlag   = fData(3 , fI)

                            'If fViewFlag > 0  Then 
                                IF isNone(fVal) Then
                                    IF CStr(Trim(fVal)) = CStr(Trim(fCode)) Then
                                        fBody = fBody &"<option value="""& fCode &""" selected>"& fValue &"</option>"
                                    Else
                                        fBody = fBody &"<option value="""& fCode &""">"& fValue &"</option>"
                                    End if
                                Else
                                    fBody = fBody &"<option value="""& fCode &""" "& isSelected(fCode, fVal) &">"& fValue &"</option>"
                                End If
                            'End If     
						Next
						getCodeDB = fBody
					Case "reselect"
						fBody = ""
						For fI = 0 To UBound(fData, 2)
							fCode	    = fData(0, fI)
							fValue	    = fData(1, fI)
                            fSelected   = fData(2, fI)
                            fViewFlag   = fData(3 , fI)

                            If fViewFlag > 0 Then 
                                IF isNone(fVal) Then
                                    IF CStr(fSelected) = "1" Then
                                        fBody = fBody &"<option value="""& fCode &""" selected>"& fValue &"</option>"
                                    Else
                                        fBody = fBody &"<option value="""& fCode &""">"& fValue &"</option>"
                                    End If 
                                Else
                                    fBody = fBody &"<option value="""& fCode &""" "& isSelected(fValue, fVal) &">"& fValue &"</option>"
                                End If
                            End If 
						Next
						getCodeDB = fBody
					Case "checkbox", "radio"
						fBody = ""
						For fI = 0 To UBound(fData, 2)
							fCode	= fData(0, fI)
							fValue	= fData(1, fI)
							fGroupNm = fData(3, fI)
                            fViewFlag   = fData(3 , fI)

                            If fViewFlag > 0 Then     
    							fBody = fBody &"<label><input type="""& fMode &""" name=""" & fGroupNm & """  id=""" & fGroupNm & """ value="""& fCode &""" "& isChecked(fCode, fVal) &">"& fValue &"</label>&nbsp;&nbsp;"
                            End If 

						Next
						getCodeDB = fBody
					Case "array"
						getCodeDB = fData
					Case Else
						getCodeDB = "코드모드 오류!!"
				End Select
			Else
				getCodeDB = "코드없음!!"
			End if
		Else
			getCodeDB = "코드없음!!"
		End if
	End if
End Function

'* 기   능 : 코드테이블에서 코드상 데이터를 원하는 형식으로 리턴
'* 기   능 :
Function getCodeManager(fMode, fItem, fVal)

End Function

'* 파일에서 코드정보를 가져온다.
Function getCodeInfo(fMode, fCodeID, fVal)
	Dim fFile, fInfo, fI, fDic, fKeys, fItems, fData(), fBody, fCode
	fFile = getReadFile(gCodeFilePath)	'코드파일 읽어오기

	'읽어온 파일을 코드그룹을 찾아서 Dictionary객체에 저장
	Set fDic = Server.CreateObject("Scripting.Dictionary")
	fInfo = Split(fFile, Chr(13)&Chr(10))
	For fI = 0 To UBound(fInfo)
		IF Not isNone(fInfo(fI)) Then
			fCode = Split(fInfo(fI), ",")
			IF fCodeID = Trim(fCode(0)) Then
				fDic.Add Trim(fCode(1)), Trim(fCode(2)	)	' Dictionary객체에 저장
			End if
		End if
	Next

	fKeys	= fDic.Keys
	fItems	= fDic.Items

	fVal = Cstr(fVal)	'자료형을 문자로
	Select Case fMode
		Case "name"
			getCodeInfo = fDic.Item(fVal)
		Case "select"
			For fI = 0 To UBound(fKeys)
				fBody = fBody &"<option value="""& fKeys(fI) &""" "& isSelected(fKeys(fI), fVal) &">"& fItems(fI) &"</option>"
			Next
			getCodeInfo = fBody
		Case "array"	'2차원 배열로 리턴한다.
			ReDim fData(1,UBound(fKeys))
			For fI = 0 To UBound(fKeys)
				fData(0,fI) = fItems(fI)
				fData(1,fI) = fKeys(fI)
			Next
			getCodeInfo = fData
		Case Else
		getCodeInfo = "코드모드 오류!!"
	End Select
	Set fDic = Nothing
End Function

'* 월 정보를 가져올 때 10월 미만일 경우 앞에 '0'을 붙여 2자로 반환
Function getMonth2Char(month)
	If Len(month) = 1 Then
		getMonth2Char = "0" & month
	Else
		getMonth2Char = month
	End If
End Function

'* 백분율 표기시 소숫점 첫째자리까지만 반환
Function getDecimal(num)
	Dim arrNum

	If InStr(num , ".") Then
		arrNum = Split(num, ".")
		If isArray(arrNum) Then 
			If Len(arrNum(1)) > 1 Then
				getDecimal = arrNum(0) & "." & Left(arrNum(1), 1)
			Else
				getDecimal = num
			End If
		End If 
	Else
		getDecimal = num
	End If
End Function
%>