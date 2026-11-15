<%
'//=================================================================
'작  성 자 : @전륜기
'작  성 일 : @2007-11-15
'파일 내용 : @각종 함수모음
'=================================================================//

'//==================================================================(문자열 관련)
'* 기능 : fValue가없으면 fData로 치환(배열도 처리)
'fValue	: 원본 데이터
'fData	: 원본 데이터 부제시 대체할 문자
'fNull	: 필수 여부 False => 미필수, True => 필수
Function nullFillWith(fValue, fData, fNull)
    IF isArray(fValue) Then
        Dim fI
        For fI = 0 To UBound(fValue)
            fValue(fI) = setFillter(fValue(fI))
        Next
        nullFillWith = Join(fValue,",")
    Else
        IF isNone(fValue) Then
			IF fNull Then
				Call errorBack("필수 항목이 누락되었습니다.")
			Else
				nullFillWith = fData
			End if
        Else
            IF IsNumeric(fValue) then
                nullFillWith = fValue
            Else
                nullFillWith = setFillter(fValue)
            End if
        End if
    End if
End Function

'* 기능 : SQL injection 공격대비, 문자열 치환
Function setFillter(fString)
    fString = Replace(fString, "'", "&#39;")
    fString = Replace(fString, """", "&#34;")
'    fString = Replace(fString, ";", "&#41;")
'    fString = Replace(fString, "--", "&#45;")
    fString = Replace(fString, "sp_", "")
    fString = Replace(fString, "xp_", "")
'    fString = Replace(fString, "union", "")
'	fString = Replace(fString, "exec", "")
'    fString = Replace(fString, "select", "")
'	fString = Replace(fString, "update", "")
'	fString = Replace(fString, "delete", "")
'	fString = Replace(fString, "drop", "")
'	fString = Replace(fString, "truncate", "")
'	fString = Replace(fString, "union", "")
'	fString = Replace(fString, "join", "")
'	fString = Replace(fString, "sys", "")
'	fString = Replace(fString, "declare", "")
'	fString = Replace(fString, "sysobject", "")
'	fString = Replace(fString, "sysadmin", "")
	fString = Replace(fString, "db_", "")
	fString = Replace(fString, "0x6", "")
	fString = Replace(fString, "0x7", "")
	fString = Replace(fString, "0x3", "")
'	fString = Replace(fString, "chr(", "")
    setFillter = fString
End Function

'* 기능 : 치환된 문자를 원 문자로 돌려준다.
Function getRStr(fValue, fMode)
    IF Not isNone(fValue) Then
        fValue = Replace(fValue, "&#39;", "'")
        fValue = Replace(fValue, "&#34;", """")
        Select Case fMode
            Case "txt"   '텍스트
                fValue = Replace(fValue, "<","&lt")
                fValue = Replace(fValue, ">","&gt")
                fValue = Replace(fValue, Chr(13)&Chr(10),"<br>")
				fValue = Replace(fValue, Chr(10),"<br>")
            Case "xml"  'xml
                fValue = Replace(fValue, "<","&lt;")
                fValue = Replace(fValue, ">","&gt;")
                fValue = Replace(fValue, Chr(13)&Chr(10),"&lt;br&gt;")
            Case "edit"  '에디터(아무것도 안함)
            Case ""
                fValue = Replace(fValue, Chr(13)&Chr(10),"<br>")
        End Select
        getRStr = fValue
    End if
End Function

'* 기능 : 문자열속에 fKey값이 몇개 있는지 카운트
Function getStrCnt(fStr, fKey)
	IF Not isNone(fStr) Then
		Dim fI, fCnt
		fCnt = 0
		For fI = 1 To Len(fStr)
			IF CStr(Mid(fStr, fI, 1)) = CStr(fKey) Then
				fCnt = fCnt + 1
			End if
		Next
		getStrCnt = fCnt
	Else
		getStrCnt = 0
	End if
End Function

'* 기능 : 문자을 넘주는 길이보다 크면 잘라준다.
Function getCutStr(fStr, fLen)
	IF Len(fStr) > fLen Then
		Dim fI, fByte, fChar
		fByte = 1	'모든문자를 Byte로 셋팅
		For fI = 1 To fLen
			fChar = Asc(Mid(fStr, fI, 1))
			IF fChar >= 48 and fChar <= 57 Then			'숫자일때
				fByte = fByte + 1
			Elseif fChar >= 65 and fChar <= 90 Then	'영문대문자
				fByte = fByte + 1
			Elseif fChar >= 97 and fChar <= 122 Then	'영문소문자
				fByte = fByte + 1
			Else
				'fByte = fByte + 0.5
				fByte = fByte + 0.5
			End if
		Next

		getCutStr = Left(fStr, fByte) &"..."
	Else
		getCutStr = fStr
	End if
End Function

'* 기능 : 문자열중 HTML태그를 삭제한다.
Function getRMHTML(fHtmlStr)
	Dim fRex
	Set fRex = new Regexp
	fRex.Pattern= "<[^>]+>"
	fRex.Global=True
	getRMHTML = fRex.Replace(fHtmlStr,"")
	Set fRex = Nothing
End Function

'* 기능 : 설정한 자리수만큰 앞자리를 채움
'* 사용법 : setFillStr("원본","설정 자리수","채울 문자")
Function setFillStr(fStr, fLength, fFillStr)
	IF Not isNone(fStr) Then
		IF Len(fStr) < fLength Then
			setFillStr = String(fLength - Len(fStr), fFillStr) & fStr
		Else
			setFillStr = fStr
		End if
	End if
End Function

'* 기능 : 조건절에 in을 사용할 수 있도록 홀따움표로 감싸준다.
Function addMark(fStr, fSplit)
	IF Not isNone(fStr) Then
		Dim fArr, fArrTemp(), fI, fG
		fArr = Split(fStr, fSplit)
		fG = 0
		For fI = 0 To UBound(fArr)
			IF Not isNone(fArr(fI)) Then
				ReDim Preserve fArrTemp(fG)
				fArrTemp(fG) = Trim(fArr(fI))
				fG = fG + 1
			End if
		Next
		addMark = "'"& Join(fArrTemp, "','") &"'"
	Else
		addMark = ""
	End if
End Function
'======================================================================//

'//======================================================================(이메일 관련)
'* 기능 : 메일폼을 선택하여 이메일보낸다.
Sub setSendEmail(fEmailPath, fEmail, fSubject, fContent)
	Dim fFso, fFile, fMailHTML
	Set fFso=CreateObject("Scripting.FileSystemObject")
	IF Not fFso.FileExists(fEmailPath) Then
		Set fFile = fFso.OpenTextFile(fEmailPath)
		Do While Not fFile.AtEndOfStream
			fMailHTML = fMailHTML & fFile.ReadLine & chr(13) & chr(10)
		Loop
		Set fFile = Nothing
		fMailHTML = Replace(fMailHTML,"[SUBJECT]", fSubject)
		fMailHTML = Replace(fMailHTML,"[CONTENT]", fContent)

		Call setEmail(fEmail, fSubject, fMailHTML)
	End if
	Set fFso = Nothing
End Sub

'* 기능 : 이메일보낸다.
Sub setEmail(fEmail, fSubject, fContent)
	Dim fObjMail
	Set fObjMail = Server.CreateObject("CDO.Message")
	With fObjMail
		.From = gMasterEmail '보내는 사람 메일주소
		.To = fEmail
		.Subject = fSubject
		.HTMLBody = fContent
		.Send
	End With
	Set fObjMail = nothing
End Sub
'======================================================================//

'//======================================================================(문자열 인코딩 & 디코딩 & 암호화)
'*기능 : 문자열 인코딩
Function URLEncode(fString)
	IF isNone(fString) Then
		URLEncode = Server.URLEncode(fString)
	End if
End Function

'*기능 :  문자열 디코딩
Function URLDecode(fString)
	Dim fStrSource, fStrTemp, fStrResult, fStrChr
	Dim fLngPos, fAddNum, fIFKor
	fStrSource = Replace(fString, "+", " ")
	For fLngPos = 1 To Len(fStrSource)
		fAddNum = 2
		fStrTemp = Mid(fStrSource, fLngPos, 1)
		If fStrTemp = "%" Then
			If fLngPos + fAddNum < Len(fStrSource) + 1 Then
				fStrChr = CInt("&H" & Mid(fStrSource, fLngPos + 1, fAddNum))
				If fStrChr > 130 Then
					fAddNum = 5
					fIFKor = Mid(fStrSource, fLngPos + 1, fAddNum)
					fIFKor = Replace(fIFKor, "%", "")
					fStrChr = CInt("&H" & fIFKor )
				End If
				fStrResult = fStrResult & Chr(fStrChr)
				fLngPos = fLngPos + fAddNum
			End If
		Else
			fStrResult = fStrResult & fStrTemp
		End If
	Next
	URLDecode = fStrResult
End Function

'* 기능 : 중요한 데이터를 인코딩한다. (암호화)
Function base64Encode(fStr)
	IF Not isNone(fStr) Then
		Dim fBase64
		Set fBase64 = Server.CreateObject("Base64Lib.Base64")
			base64Encode = fBase64.Encode(CStr(fStr))
		Set fBase64 = Nothing
	Else
		base64Encode = ""
	End if
End Function

'* 기능 : 인코딩한 데이터를 다시 디코딩한다.(복화)
Function base64Decode(fStr)
	IF Not isNone(fStr) Then
		Dim fBase64
		Set fBase64 = Server.CreateObject("Base64Lib.Base64")
			base64Decode = fBase64.Decode(CStr(fStr))
		Set fBase64 = Nothing
	Else
		base64Decode = ""
	End if
End Function
'======================================================================//

'//======================================================================(파일 관련)
'기능 : 파일을 다운 받을수 있는 링크를 생성한다.
Function getDownFile(fFolder, fFile)
	IF Not isNone(fFolder) and Not isNone(fFile) Then
		getDownFile = "<a href=""/common/library/common_download.asp?folder="& fFolder &"&file="& Server.URLEncode(fFile) &""">"& fFile &"</a>&nbsp;&nbsp;"
	Else
		getDownFile = ""
	End if
End Function

'기능 : 수정화면에서 기존 파일 삭제 체크 박스를 생성한다.
Function getAttFile(fFolder, fAttFile)
	Dim fBody
	IF Not isNone(fAttFile) Then
		getAttFile = "<br><span style=""padding:3 0 3 0"">기존파일 : "& getDownFile(fFolder, fAttFile) &"</span>"
	End if
End Function

'기능 : 수정화면에서 기존 파일 삭제 체크 박스를 생성한다. 첨부파일이 여러게 있을경우사용한다.
Function getAttFiles(fFolder, fAttFile, fName)
	Dim fBody
	IF Not isNone(fAttFile) Then
		fBody = "<table width=""100%"" border=""0""><tr>"& vbCrLf
		fBody = fBody & "<td width=""60"" style=""padding-top:2px;"">첨부파일 : </td>"& vbCrLf
		fBody = fBody & "<td width=""10""><input type=""checkbox"" name="""& fName &""" value="""& fAttFile &""" style=""height:16px;"" /></td>"& vbCrLf
		fBody = fBody & "<td style=""padding-top:2px;"">"& getDownFile(fFolder, fAttFile) &"</td>"& vbCrLf
		fBody = fBody & "</tr></table>"& vbCrLf
		getAttFiles = fBody
	End if
End Function

'기능 : 파일의 확장자를 가져온다.
Function getExt(fAttFile)
	IF Not isNone(fAttFile)Then
		getExt = Mid(fAttFile,InStrRev(fAttFile,".")+1)
	End if
End Function

'기능 : 파일을 생성하고 내용을 작성함
Function setWriteFile(fFilePath, fContent)
	Dim fFso, fFileObj
	Set fFso = Server.CreateObject("Scripting.FileSystemObject")
	Set fFileObj = fFso.OpenTextFile(fFilePath,2,True)
		fFileObj.WriteLine(fContent)
	fFileObj.Close()
	Set fFileObj = Nothing
	Set fFso = Nothing
End Function

'* 기능 : 파일의 내용을 읽어 드린다.
Function getReadFile(fPath)
	Dim fFso, fFile, fBody
	Set fFso=CreateObject("Scripting.FileSystemObject")
	IF fFso.FileExists(fPath) Then
		Set fFile = fFso.OpenTextFile(fPath,1)
		getReadFile = fFile.ReadAll()
		'라인으로 읽어들일때
		'Do While Not fFile.AtEndOfStream
		'	getReadFile = getReadFile & fFile.ReadLine & chr(13) & chr(10)
		'Loop
		fFile.Close()
		Set fFile = Nothing
	End if
	Set fFso = Nothing
End Function

'기능 : 텍스트 파일 생성 설장한 시간동안은 텍스파일을 보여줌
Function setMakeTextFile(fFileName, fContent, fInterval, fTime)
	Dim fFso, fFilePath, fFileObj, fFile, fFileDate
	fFilePath = gTxtFilePath & fFileName &".txt"
	Set fFso = Server.CreateObject("Scripting.FileSystemObject")
		IF Not fFso.FileExists(fFilePath) Then	'파일이 존재하지 않는다면
			Call setWriteFile(fFilePath, fContent)
			setMakeTextFile =  Replace(getReadFile(fFilePath),Chr(13)&Chr(10),"<br>")
		Else	'파일이 존재한다면
			Set fFile = fFso.GetFile(fFilePath)
				 fFileDate = fFile.DateLastModified
				 '파일을 마지막으로 수정한 날짜에 설정한 단위의 시간을 더 해준다.
				 fFileDate = DateAdd(fInterval, fTime, fFileDate)
				 IF Datediff("s",Now(), fFileDate) >= 0 Then
					setMakeTextFile = Replace(getReadFile(fFilePath), Chr(13)&Chr(10), "<br>")
				 Else
					Call setWriteFile(fFilePath, fContent)
					setMakeTextFile = Replace(getReadFile(fFilePath), Chr(13)&Chr(10), "<br>")
				 End if
			Set fFile = Nothing
		End if
	Set fFso = Nothing
End Function

'기능 fFile형식이 fkey와 같다면 True 다르다면 False반환
Function getFileType(fFile, fKey)
	Dim fExt
	fExt = LCase(getExt(fFile))
	Select Case fKey
		Case "image"
			IF fExt = "jpg" or fExt = "gif" or fExt = "jpeg" or fExt = "png" Then
				getFileType = True
			Else
				getFileType = False
			End if
		Case "attach"
			IF fExt = "aspx" or fExt = "asp"  or fExt = "jsp" or fExt = "php" or fExt = "exe" or fExt = "js" Then
				getFileType = True
			Else
				getFileType = False
			End if
		Case Else
			getFileType = "함수에 추가요망"
	End Select
End Function

'* 기능 : 파일을 삭제한다.
'* 사용법 : setDelFile "경로","파일이름"
Sub setDelFile(fFolder, fFile)
	Dim fFso, fFilePath
	Set fFso = Server.CreateObject("Scripting.FileSystemObject")
	fFilePath = gFileUpPath & fFolder &"\"& fFile
  	IF fFso.FileExists(fFilePath) Then
    	fFso.DeleteFile fFilePath, false
  	End if
  	set fFso = nothing
End Sub
'======================================================================//

'//======================================================================(이미지 관련)
'*기능 : 이미지 사이즈를 설정하고 그 사이즈보다 크면 설정값으로 변환한다.
Function getImage(fFolder, fFile, fSize)
	IF Not isNone(fFile) Then
		Dim fImgWidth
		fImgWidth = getImgWSize(fFolder, fFile)
		IF fImgWidth > fSize then
			fImgWidth = fSize
		End if
		getImage = "<div><img src="""& USER_PATH & fFolder &"/"& Server.URLencode(fFile) &""" width="""& fImgWidth &"""></div>"
	End if
End Function

'이미지의 가로 사이즈
Function getImgWSize(fFolder, fFile)
'	Dim fImgPath, fImgFile, fFso
'	IF getFileType(fFile,"image") Then
'		fImgPath = gFileUpPath & fFolder &"\"& fFile
'		Set fFso = Server.CreateObject("Scripting.FileSystemObject")
'		IF Not fFso.FileExists(fImgPath) Then
'			getImgWSize = "이미지 찾을 수 없음!!"
'		Else
'			Dim fObjImg
'			Set fObjImg = LoadPicture(fImgPath)
'				getImgWSize = CLng(fObjImg.Width) * 24 / 635 '픽셀단위
'			Set fObjImg = Nothing
'		End if
'		Set fFso = Nothing
'	Else
'		getImgWSize = "이미지가 아님!!"
'	End if
End Function

'이미지의 세로 사이즈
Function getImgHSize(fFolder, fFile)
'	Dim fFso, fImgPath, fImgFile
'	IF getFileType(fFile,"image") Then
'		fImgPath = gFileUpPath & fFolder &"\"& fFile
'		Set fFso = Server.CreateObject("Scripting.FileSystemObject")
'		IF Not fFso.FileExists(fImgPath) Then
'			getImgHSize = "이미지 찾을 수 없음!!"
'		Else
'			Dim fObjImg
'			Set fObjImg = LoadPicture(fImgPath)
'				getImgHSize = CLng(fObjImg.Height) * 24 / 635 '픽셀단위
'			Set fObjImg = Nothing
'		End if
'		Set fFso = Nothing
'	Else
'		getImgHSize = "이미지가 아님!!"
'	End if
End Function

'*설정된 값을 기준으로 가로와 세로를 리사이즈 한다.
Function getThumbImage(fFolder, fFile, fSize)
	Dim fImgWidth, fImgHeight
	fImgWidth = getImgWSize(fFolder, fFile)
	fImgHeight = getImgHSize(fFolder, fFile)

'	IF fImgWidth >= fImgHeight Then
		IF fImgWidth > fSize then
			fImgHeight = Round(fImgHeight / (Round(fImgWidth/fSize)))
			fImgWidth = fSize
		End if
'	Else
'		IF fImgHeight > fSize then
'			fImgWidth = Round(fImgWidth / (Round(fImgHeight/fSize)))
'			fImgHeight = fSize
'		End if
'	End if
	getThumbImage = "<img src="""& USER_PATH & fFolder &"/"& Server.URLencode(fFile) &""" width="""& fImgWidth &""" height="""& fImgHeight &""" style=""border:1px solid #000000;"">"
End Function
'======================================================================//

'//======================================================================(날짜관련)
'* 기능 : 사이트 날짜 포맷 설정
Function getDateFormat(fDate, fKey)
	IF Not isNone(fDate) Then
		getDateFormat = Replace(FormatDateTime(fDate,2), "-", fKey)
	End if
End Function

'해당날짜의 요일은 가져온다.
Function getDateName(fDate)
	Select case Weekday(fDate, 1)
		case 1	getDateName = "일요일"
		case 2	getDateName = "월요일"
		case 3	getDateName = "화요일"
		case 4	getDateName = "수요일"
		case 5	getDateName = "목요일"
		case 6	getDateName = "금요일"
		case 7	getDateName = "토요일"
	End select
End Function

'해당날짜가 속한 주의 시작와 끝
Function getWeekDay(fGubun, fToday)
	Dim fTodayNum, fX, fY, fStart, fEnd
	fTodayNum	= Weekday(CDate(fToday), 1)
	Select Case fTodayNum
		Case 1
			fX = 1
			fY = 7
		Case 2
			fX = 0
			fY = 6
		Case 3
			fX = -1
			fY = 5
		Case 4
			fX = -2
			fY = 4
		Case 5
			fX = -3
			fY = 3
		Case 6
			fX = -4
			fY = 2
		Case 7
			fX = -5
			fY = 1
	End Select

	fStart	= Dateadd("d", fX, fToday)
	fEnd	= Dateadd("d", fY, fToday)
	IF fGubun = "start" Then
		getWeekDay = fStart
	Else
		getWeekDay = fEnd
	End if
End Function

'* 기능 : 해당월의 마지막 일
Function getLastOfMonth(fYear, fMonth)
	getLastOfMonth = Day(Dateadd("d",-1, Dateadd("m",1, fYear &  "-" & fMonth & "-1")))
End Function

'* 기능 : 현재시간을 MSSQL의 GETDATE()와 같은 포맷으로 반환
Function getDateTime(fDate)
	IF CDate(fDate) Then
		getDateTime = FormatDateTime(fDate,2) &" "& setFillStr(Hour(fDate), 2, "0") &":"& setFillStr(Minute(fDate), 2, "0") &":"& setFillStr(Second(fDate), 2, "0")
	Else
		getDateTime = ""
	End if
End Function

Function getConvertDate(fDate, fMode)
	IF Not isNone(fDate) Then
		IF fMode = "date" Then
			getConvertDate = Mid(fDate, 1, 4) &"-"& Mid(fDate, 5, 2) &"-"& Mid(fDate, 7, 2) &" "& Mid(fDate, 9, 2) &":"& Mid(fDate, 11, 2) &":"& Mid(fDate, 13, 2)
		Elseif fMode = "timestemp" Then
			fDate = Replace(fDate, " ", "")
			fDate = Replace(fDate, ":", "")
			fDate = Replace(fDate, "-", "")
			getConvertDate = fDate
		Else
			getConvertDate = ""
		End if
	Else
		getConvertDate = ""
	End if
End Function
'======================================================================//

'//======================================================================(ASP내장함수)
'* 기능 : 화면에 값을 찍어준다.
Function RW(fStr)
	Response.Write fStr
End Function

Function RWB(fStr)
	Response.Write fStr & "</BR>"
End Function



'* 기능 : 이 함수에서 모든 실행을 중지한다.
Function RE()
	Response.End
End Function

'* 기능 : 화면에 값을 찍어주고 실행을 중지한다.
Function RWE(fStr)
	Response.Write fStr
	Response.End
End Function

'* 기능 : 화면을 이동한다.
Function RR(fURL)
	Response.Redirect fURL
End Function
'======================================================================//

'//======================================================================(루트관련)
'* 기능 : 현재 파일의 루트를 반환한다.
Function ScriptName(fRoot)
	Dim fScript, fScriptName(10)
	fScript = Split(Request.ServerVariables("SCRIPT_NAME"),"/")

	Dim fI
	For fI = 1 To UBound(fScript)
		fScriptName(0) = Lcase(Request.ServerVariables("SERVER_NAME"))
		fScriptName(fI) = Lcase(fScript(fI))
	Next

	ScriptName = fScriptName(fRoot)
End Function
'======================================================================//

'//======================================================================(값비교)
'* 기능 : 값이 값으면 selected를 반환한다.
Function isSelected(fData, fValue)
	IF CStr(fData) = CStr(fValue) Then
		IsSelected = " selected"
	End if
End Function

'* 기능 : 값이 값으면 checked를 반환한다.
Function isChecked(fData, fValue)
	IF CStr(fData) = CStr(fValue) Then
		IsChecked = " checked"
	End if
End Function



'기 능 : 2차원배열에서 fKey값이 있다면 "checked"반환
Function isArrChecked(fData, fKey)
	IF isArray(fData) Then
		Dim fI, fJ, fRtnVal
		For fI = 0 To UBound(fData,2)
			For fJ = 0 To UBound(fData,1)
				IF fKey = fData(fJ, fI) Then
					fRtnVal = True
					Exit For
				End if
			Next
		Next

		IF fRtnVal = True Then
			isArrChecked = "checked"
		End if
	End if
End Function

'* 기능 : 공백이면 False 공백이 아니면 True Boolean형으로 반환
Function isNone(fString)
	fString = Trim(fString)
	IF isNumeric(fString) Then
		isNone = False
	Else
		IF isNull(fString) or isEmpty(fString) or fString = "" Then
			isNone = True
		Else
			isNone = False
		End if
	End if
End Function

'* 기능 : fItem의 문자열에 fKey가 있다면  True, 없다면 False Boolean형으로 반환
Function isInStr(fItem, fKey)
	IF Not (isNone(fItem) and isNone(fkey)) Then
		IF InStr(1, fItem, fKey,1) >= 1 Then	'찾는 문자열이 있다면
			isInStr = True
		Else
			isInStr = False
		End if
	End if
End Function
'======================================================================//

'//======================================================================(쿠키)
'* 기능 : 쿠키 저장
Function setCookies(fGroup, fID, fValue, fExpire)
	Response.Cookies(fGroup)(fID) = fValue
	Response.Cookies(fGroup).Domain = gScripDomain
	Response.Cookies(fGroup).Path = "/"
	Response.Cookies(fGroup).Expires = fExpire
End Function

'* 기능 : 쿠키 가져오기
Function getCookies(fGroup, fID)
	getCookies = Request.Cookies(fGroup)(fID)
End Function
'======================================================================//

'//======================================================================(기타)
Function Ceil(fCnt, fPage)
	fCnt	= Int(fCnt)
	fPage	= Int(fPage)

	Dim fTPage, fMod
	fTPage	= Int(fCnt / fPage)
	fMod	= fCnt Mod fPage

	IF fMod > 0 Then
		fTPage = fTPage + 1
	End if
	Ceil = fTPage
End Function

'Querystring 파라메터에서 필요한 파라메터값을 가져온다.
Function getParameter(fStr, fParam)
	IF Not isNone(fStr) Then
		Dim fArr, fI, fArrStr
		fArr = Split(fStr, "&")
		For fI = 0 To UBound(fArr)
			fArrStr = Split(fArr(fI), "=")

			IF Trim(fArrStr(0)) = Trim(fParam) Then
				getParameter = Trim(fArrStr(1))
				Exit For
			End if
		Next
	End if
End Function

Function getRefer()
	getRefer = LCase(Request.ServerVariables("SCRIPT_NAME")) &"?"& Request.ServerVariables("QUERY_STRING")
End Function

'리플 숫자를 오렌지 색으로 표기한다.
Function getReplyNum(fItem)
	IF Cint(fItem) > 0 Then
		getReplyNum = "<span style=""color:#F7941C"">("& fItem &")</span>"
	Else
		getReplyNum = ""
	End if
End Function

'답변형 깊만큼 들여쓰기한다.
Function getDepthNum(fItem, fICon)
	Dim fBody, fI
	IF Cint(fItem) > 0 Then
		For fI = 0 To fItem
			fBody = fBody & "&nbsp;"
		Next

		fBody = fBody & fICon
	Else
		fBody = ""
	End if
	getDepthNum = fBody
End Function

'등록날짜보 24시간 이내이면 아이콘 생성
Function getNewICon(fDate)
	IF DateDiff("h", fDate, now()) =< 24 Then
		getNewICon = " <img src=""/images/icon_12.jpg"" style=""vertial-align:middle"" />"
	Else
		getNewICon = ""
	End if
End Function

'* 테이블에 접근 레코드셋을 가져오는 문자열 만들고, Insert구문을 만든다!(Dictionary 객체 사용버전으로 MSSQL 2000 에서 사용)
Function getF(fTbl, fMode)
	Dim fRs, fWhere
	fWhere		= "where a.id = b.id and a.name = '"& fTbl &"' and b.xtype = c.xtype order by colid"
	Set fRs = getSelectQuery("dbo.sysobjects a, dbo.syscolumns b, dbo.systypes c","b.name", fWhere)
	IF fRs.Bof or fRs.Eof Then
		RW("해당하는 테이블 정보가 없습니다.")
		RE()
	Else
		Dim fData
		fData = fRs.GetRows
	End if
	RsClose(fRs)

	'필드명을 동적배열로 만든다.
	IF isArray(fData) Then
		Dim fI, fJ, fName, fFieldNM, fBody, fParamNM
		Dim fParam(), fVariable(), fField()
		For fI = 0 To UBound(fData,2)
			fName = fData(0, fI)
			ReDim Preserve fVariable(fI)	'변수명
			ReDim Preserve fParam(fI)		'파라매타용 변수
			ReDim Preserve fField(fI)			'컬럼 명
			fField(fI) = fName
			IF InStr(fName,"_") <> 0 Then
				Dim fStr
				fStr = Split(fName,"_")
				For fJ = 0 To UBound(fStr)
					Dim fTemp1, fTemp2
					fTemp1 = Ucase(Left(fStr(fJ),1))
					fTemp2 = Lcase(Right(fStr(fJ),Len(fStr(fJ)) - 1))
					IF fJ = 0 Then
						fFieldNM		= "v"& fTemp1 & fTemp2
						fParamNM		= "s"& fTemp1 & fTemp2
					Else
						fFieldNM		= fFieldNM & fTemp1 & fTemp2
						fParamNM		= fParamNM & fTemp1 & fTemp2
					End if
				Next

				fVariable(fI)	= fFieldNM
				fParam(fI)		= fParamNM
			Else
				fVariable(fI)	= "v"& Ucase(Left(fName,1)) & Lcase(Right(fName,Len(fName) - 1))
				fParam(fI)		= "s"& Ucase(Left(fName,1)) & Lcase(Right(fName,Len(fName) - 1))
			End if
		Next

		Select Case fMode
			Case "R"
				fBody = "Dim "& Join(fVariable,",") &"<br>"
				For fI = 0 To UBound(fField)
					fBody = fBody & fVariable(fI) &" = vRs("""& fField(fI) &""")<br>"
				Next
			Case "D"
				fBody = "Dim vDic,vKeys,vItems<br>Set vDic = Server.CreateObject(""Scripting.Dictionary"")<br>"
				For fI = 0 To UBound(fField)
					fBody = fBody &"vDic.Add """& fField(fI) &""","& fParam(fI) &"<br>"
				Next
				fBody = fBody &"<br>vKeys = vDic.Keys<br>vItems = vDic.Items<br>Set vDic = Nothing<br>Call setInsertQuery("& fTbl &",vKeys,vItems)"
		End Select
		getF = fBody
	Else
		fBody = "해당하는 테이블의 컬럼 정보가 없습니다."
		getF = vBody
	End if
End Function

'* 2차원 배열에서 값을 꺼내고 동적 변수 생성
Function getV(fField, fArrNM)
	IF Not isNone(fField) Then
		fField = Split(fField,",")
		Dim fI, fVariable()
		For fI = 0 To UBound(fField)
			ReDim Preserve fVariable(fI)
			IF isInStr(fField(fI),"_") Then
				Dim fFieldVal, fJ, fFieldNM
				fFieldVal = Split(fField(fI),"_")
				For fJ = 0 To UBound(fFieldVal)
					IF fJ = 0 Then
						fFieldNM = "v"& Ucase(Left(fFieldVal(fJ),1)) & Lcase(Right(fFieldVal(fJ),Len(fFieldVal(fJ)) - 1))
					Else
						fFieldNM = fFieldNM & Ucase(Left(fFieldVal(fJ),1)) & Lcase(Right(fFieldVal(fJ),Len(fFieldVal(fJ)) - 1))
					End if
				Next
			Else
				fFieldNM = "v"& Ucase(Left(fField(fI),1)) & Lcase(Right(fField(fI),Len(fField(fI)) - 1))
			End if
			fVariable(fI) = fFieldNM
		Next

		Dim fBody
		fBody = "Dim "& Join(fVariable,", ") &"<br>"
		For fI = 0 To UBound(fVariable)
			fBody = fBody & fVariable(fI) &" = "& fArrNM &"("& fI &", vI) <br>"
		Next

		getV = fBody
	End if
End Function

' 배열로 넘어온 데이터를 json데이터로 만들어 준다.
' 자바스크립트에서 위 함수를 사용한 json데이터를 사용할때는 json객체.jdata1 ~ json객체.jdata...n으로 사용해야된다.
Function getJSON(fData)
	Dim fI, fG
	IF isArray(fData) Then
		Dim fStr
		fStr = "["
		For fI = 0 To UBound(fData, 2)
			fStr = fStr &"{"
			For fG = 0 To UBound(fData, 1)
				fStr = fStr &"""jdata"& (fG + 1) &""" : """& Trim(fData(fG, fI)) &""""
				IF UBound(fData, 1) <> fG Then
					fStr = fStr &", "
				End if
			Next
			fStr = fStr &"}"
			IF UBound(fData, 2) <> fI Then
				fStr = fStr &", "
			End if
		Next
		fStr = fStr &"]"
		getJSON = fStr
	Else
		getJSON = "[]"
	End if
End Function
'======================================================================//

'//======================================================================(Javascript경고창)
'* 기능 : 경고창을 호출한다.
'Sub errorAlert(btn , fStr)
'	Dim fBody
'	fBody = "<script type=""text/javascript"">"& vbCrLf
'	fBody = fBody & "//<![CDATA["& vbCrLf
'	fBody = fBody & "window.parent.sAlert("""& btn &""",""" & fStr & """);"& vbCrLf
'	fBody = fBody & "//]]>"& vbCrLf
'	fBody = fBody & "</script>"& vbCrLf
'	Response.Write fBody
'	Response.End
'End sub

'/* 기능 : 단순 경고창만 호출한다
Sub onlyAlert(btn , fStr)
	Dim fBody
	fBody = "<script type=""text/javascript"">"& vbCrLf
	fBody = fBody & "//<![CDATA["& vbCrLf
	fBody = fBody & "window.parent.sAlert("""& btn &""",""" & fStr & """);"& vbCrLf
	fBody = fBody & "//]]>"& vbCrLf
	fBody = fBody & "</script>"& vbCrLf
	Response.Write fBody
	Response.End
End sub

'* 기능 : 페이지를 이동한다.
Sub gotoAlert(btn , fStr , flag , url)
	Dim fBody
	fBody = "<script type=""text/javascript"">"& vbCrLf
	fBody = fBody & "//<![CDATA["& vbCrLf
	fBody = fBody & "window.parent.sAlertMove("""& btn &""",""" & fStr & """,""" & flag & """,""" & url & """);"& vbCrLf
	fBody = fBody & "//]]>"& vbCrLf
	fBody = fBody & "</script>"& vbCrLf
	Response.Write fBody
	Response.End
End Sub

Sub errorGotoParent(fGotoURL)
	Dim fBody
	fBody = "<script type=""text/javascript"">"& vbCrLf
	fBody = fBody & "//<![CDATA["& vbCrLf
	fBody = fBody & "parent.location.href="""& fGotoURL &""";"& vbCrLf
	fBody = fBody & "//]]>"& vbCrLf
	fBody = fBody & "</script>"& vbCrLf
	Response.Write fBody
	Response.End
End sub

'* 기능 : 경고창을 호출하고 창을 닫는다.
Sub errorClose(fStr)
	Dim fBody
	fBody = "<script type=""text/javascript"">"& vbCrLf
	fBody = fBody & "//<![CDATA["& vbCrLf
	fBody = fBody & "alert("""& fStr &""");"& vbCrLf
	fBody = fBody & "self.close();"& vbCrLf
	fBody = fBody & "//]]>"& vbCrLf
	fBody = fBody & "</script>"& vbCrLf
	Response.Write fBody
	Response.End
End sub

'* 기능 : 경고창을 호출하고 페이지를 이동시킨다.
'Sub errorMove(fStr, fGotoURL)
'	Dim fBody
'	Dim fBody
'	fBody = "<script type=""text/javascript"">"& vbCrLf
'	fBody = fBody & "//<![CDATA["& vbCrLf
'	fBody = fBody & "window.parent.sAlert("""& btn &""",""" & fStr & """);"& vbCrLf
'	fBody = fBody & "//]]>"& vbCrLf
'	fBody = fBody & "</script>"& vbCrLf
'	Response.Write fBody
'	Response.End
'End sub


'* 기능 : 경고창을 호출하고 페이지를 이동시킨다.
Sub errorParentMove(fStr, fGotoURL)
	Dim fBody
	fBody = "<script type=""text/javascript"">"& vbCrLf
	fBody = fBody & "//<![CDATA["& vbCrLf
	fBody = fBody & "alert("""& fStr &""");"& vbCrLf
	fBody = fBody & "parent.location.href="""& fGotoURL &""";"& vbCrLf
	fBody = fBody & "//]]>"& vbCrLf
	fBody = fBody & "</script>"& vbCrLf
	Response.Write fBody
	Response.End
End sub

'* 기능 : 경고창을 호출하고 페이지를 이동시킨다.
Sub errorConfirm(fStr, fGotoURL, fUrl)
	Dim fBody
	fBody = "<script type=""text/javascript"">"& vbCrLf
	fBody = fBody & "//<![CDATA["& vbCrLf
	fBody = fBody & "if(confirm("""& fStr &""")){"& vbCrLf
	fBody = fBody & "	 location.href="""& fGotoURL &""";"& vbCrLf
	fBody = fBody & "}else{"& vbCrLf
	fBody = fBody & "	 location.href="""& fUrl &""";"& vbCrLf
	fBody = fBody & "}"& vbCrLf
	fBody = fBody & "//]]>"& vbCrLf
	fBody = fBody & "</script>"& vbCrLf
	Response.Write fBody
	Response.End
End sub

'* 기능 : 경고창을 호출하고 이전페이지로 이동한다.
Sub errorBack(fStr)
	Dim fBody
	fBody = "<script type=""text/javascript"">"& vbCrLf
	fBody = fBody & "//<![CDATA["& vbCrLf
	fBody = fBody & "alert("""& fStr &""");"& vbCrLf
	fBody = fBody & "history.back();"& vbCrLf
	fBody = fBody & "//]]>"& vbCrLf
	fBody = fBody & "</script>"& vbCrLf
	Response.Write fBody
	Response.End
End sub

'* 기능 : 팝업창을 닫고 부모창을 새로고침한다.
Sub errorReload(fGotoURL)
	Dim fBody
	fBody = "<script type=""text/javascript"">"& vbCrLf
	fBody = fBody & "//<![CDATA["& vbCrLf
	fBody = fBody & "self.close();"& vbCrLf
	fBody = fBody & "opener.location.replace('"& fGotoURL &"');"& vbCrLf
	fBody = fBody & "//]]>"& vbCrLf
	fBody = fBody & "</script>"& vbCrLf
	Response.Write fBody
	Response.End
End sub

'* 기능 : 팝업창을 닫고 부모창을 새로고침한다.
Sub setScript(fScript)
	Dim fBody
	fBody = "<script type=""text/javascript"">"& vbCrLf
	fBody = fBody & "//<![CDATA["& vbCrLf
	fBody = fBody & fScript & vbCrLf
	fBody = fBody & "//]]>"& vbCrLf
	fBody = fBody & "</script>"& vbCrLf
	Response.Write fBody
	'Response.End
End sub
'======================================================================//
%>