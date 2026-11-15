<%
'//=================================================================
'작  성 자 : @전륜기
'작  성 일 : @2007-11-15
'파일 내용 : @Request객체로부터 받은 Element를 동적변수로 생성, 파일 첨부도 처리
'단, 파일첨부시 dir 변수가 넘어와야함!, 해당파일이 저장될 하위경로의 폴더 이름만!
'=================================================================//
'* 기능 : Request나 컴포넌트 개체로 받을 필요없이 동적 변수로 사용가능하다.
'* 'save.asp'파일에서만 작동한다. 이유는 알아서 생각해요-_-;;
Dim gLScriptName
gLScriptName = LCase(ScriptName(UBound(Split(Request.ServerVariables("SCRIPT_NAME"),"/"))))
'IF gLScriptName = "save.asp" Then
    Dim fElem, fElemName, fObjStr, fTabs, fG
    Dim fArrFile(),fI
    'multipart/form-data방식
    IF Lcase(Left(Request.ServerVariables("HTTP_CONTENT_TYPE"),19)) = "multipart/form-data" Then
        'DEXT사용 기본 조건
        Set fTabs = Server.CreateObject("TABSUpload4.Upload")
        fTabs.CodePage = 65001    'UTF-8 언어셋 셋팅
        fTabs.Start gFileUpPath

        For Each fElem In fTabs.Form
            '인자이름
            fElemName = "s"& Ucase(Left(fElem.Name,1)) & Mid(fElem.Name,2,Len(fElem.Name))
            IF fElem.IsFile Then
                Execute "Dim "& fElemName &" : "& fElemName &" = setFileUpload(fTabs, fElem.Name)"
            Else
                Execute "Dim "& fElemName &" : "& fElemName &" = nullFillWith(fTabs("""& fElem.Name &"""),"""", False)"
            End if
        Next
        Set fTabs = Nothing
    Else
        'POST방식
        IF Lcase(Request.ServerVariables("REQUEST_METHOD")) = "post" Then
            For Each fElem In Request.Form

                fElem = Replace(fElem,"&","")
                fElemName = "s"& Ucase(Left(fElem,1)) & Mid(fElem,2,Len(fElem))

                Execute "Dim "& fElemName &" : "& fElemName &" = nullFillWith(Request.Form("""& fElem &"""), """", False)"
            Next
        'GET방식
        Else
            For Each fElem In Request.QueryString
                fElemName = "s"& Ucase(Left(fElem,1)) & Mid(fElem,2,Len(fElem))

                
                Execute "Dim "& fElemName &" : "& fElemName &" = nullFillWith(Request.QueryString("""& fElem &"""), """", False)"
            Next
        End if
    End if
'End if

'첨부파일 업로드
Function setFileUpload(fUp, fFile)
    fUp.MaxBytesToAbort = MAX_SIZE    '업로드 용량(5M)
    Dim fNum, fFileName, fExt, fFilePath, fMType, fRealName

    If fUp("dir") <> "CURRPAGE_UP" Then 
        IF Not isNone(fUp(fFile)) Then
            IF fUp(fFile).FileSize > MAX_SIZE Then
                Call errorBack("파일 첨부는 "& (MAX_SIZE / 1024 / 1024) &"MB가 이상할 수 없습니다.")
            Else
                fFileName    = fUp(fFile).FileName        '파일이름
                fExt        = fUp(fFile).FileType        '파일 확장자
                fMType      = fUp(fFile).ContentType    '파일 mimetype

                'IF fMType = "image/jpeg" or fMType = "image/gif" Then
                    '파일 만든기 만들기
                    Dim vRNum, vRandom
                    Randomize
                    vRNum    = Int(900 * Rnd) + 100
                    vRandom    = Year(Date()) & setFillStr(Month(Now()), 2, "0") & setFillStr(Day(Now()), 2, "0") & setFillStr(Hour(Now()), 2, "0")
                    vRandom    = vRandom & setFillStr(Minute(Now()), 2, "0") & setFillStr(Second(Now()), 2, "0") & vRNum
                    
                    'fFilePath = gFileUpPath & fUp("dir") &"\"& fFileName '기존파일이름, 저장할 때 중복파일이 있으면 앞에 [숫자]+ 파일이름의 형식으로 저장
                    'fFilePath = gFileUpPath & fUp("dir") &"\"& vRandom &"."& fExt

                    fRealName = Mid(fFileName, 1, Instr(fFileName, ".") - 1) & "_" & vRandom '기존 파일이름에 난수를 붙여준다
                    fFilePath = gFileUpPath & fUp("dir") &"\"& fRealName &"."& fExt
                    fUp(fFile).SaveAs fFilePath, False    '파일 저장
                    setFileUpload = fUp(fFile).ShortSaveName
                'Else
                '    Call errorBack("이미지만 첩부할 수 있습니다.\n파일 형식을 확인 후 업로드 해주세요.")
                'End if
            End if
        Else
            setFileUpload = ""
        End If
    Else 
            setFileUpload = ""
    End If 
End Function

''* 기능 : 해당페이지로 넘어온 파라메타 변수를 만들어 준다.
Sub R(fArrN)
    Dim vI
    IF LCase(Left(Request.ServerVariables("HTTP_CONTENT_TYPE"),19)) <> "multipart/form-data" Then
        IF LCase(Request.ServerVariables("REQUEST_METHOD")) = "post" Then
            vI = 0
            RW("Dim "& fArrN &"("& Request.Form.Count - 1 &")<br>")
            For Each fElem In Request.Form
                'fElemName = "s"& Ucase(Left(fElem,1)) & Mid(fElem,2,Len(fElem))
                RW(fArrN &"("& vI &") = nullFillWith(Request("""& fElem &"""), """", False)<br>")
                vI = vI + 1
            Next
        Else
            vI = 0
            RW("Dim "& fArrN &"("& Request.QueryString.Count - 1 &")<br>")
            For Each fElem In Request.QueryString
                fElemName = "s"& Ucase(Left(fElem,1)) & Mid(fElem,2,Len(fElem))
                RW(fArrN &"("& vI &") = "& fElemName &"<br>")
                vI = vI + 1
            Next
        End if
    End if
    RE()
End Sub

'* 기능 :     Request.Form으로 넘어온 모든 Element이름과  Element Value값을 찍어준다.
Sub RF()
    Dim fItem,fStr, fI, fComa
    fStr = "Dim "
    fI = 1
    For Each fItem In Request.Form
        IF Request.Form.Count <> fI Then
            fComa = ","
        Else
            fComa = ""
        End if
        fStr = fStr &"s"& Ucase(Left(fItem,1)) & Mid(fItem,2,Len(fItem)) &", "
        fI = fI + 1
    Next
    Response.Write fStr &"<br>"

    For Each fItem In Request.Form
        fStr = "s"& Ucase(Left(fItem,1)) & Mid(fItem,2,Len(fItem)) & " = "
        fStr = fStr & "nullFillWith(Request(" & chr(34) & fItem & chr(34) & "), """", False)"
        Response.Write fStr & "<br>"
    Next
    Response.End
End Sub


'* 기능 :     Request.Form으로 넘어온 모든 Element이름과  Element Value값을 찍어준다.
Sub S_RF()
    Dim fItem,fStr, fI, fComa
    fStr = "Dim "
    fI = 1
    For Each fItem In Request.Form
        IF Request.Form.Count <> fI Then
            fComa = ","
        Else
            fComa = ""
        End if
        fStr = fStr &"s_"& Lcase(Left(fItem,1)) & Mid(fItem,2,Len(fItem)) &", "
        fI = fI + 1
    Next
    Response.Write fStr &"<br>"

    For Each fItem In Request.Form
        fStr = "s_"& Lcase(Left(fItem,1)) & Mid(fItem,2,Len(fItem)) & " = "
        fStr = fStr & "nullFillWith(Request(" & chr(34) & fItem & chr(34) & "), """", False)"
        Response.Write fStr & "<br>"
    Next
    Response.End
End Sub

'* 기능 :     Request.QueryString으로 넘어온 모든 Element이름과 Element Value값을 찍어준다.
Sub RQ()
    Dim fItem, fStr, fI, fComa
    fStr = "Dim "
    fI = 1
    For Each fItem In Request.QueryString
        IF Request.QueryString.Count <> fI Then
            fComa = ","
        Else
            fComa = ""
        End if
        fStr = fStr &"s"& Ucase(Left(fItem,1)) & Mid(fItem,2,Len(fItem)) & fComa
        fI = fI + 1
    Next
    Response.Write fStr &"<br>"

    For Each fItem In Request.QueryString
        fStr = "s"& Ucase(Left(fItem,1)) & Mid(fItem,2,Len(fItem)) & " =    "
        fStr = fStr & "nullFillWith(Request(" & chr(34) & fItem & chr(34) & "),"""", False)"
        Response.Write fStr & "<br>"
    Next
    Response.End
End Sub


'* 기능 :     Request.QueryString으로 넘어온 모든 Element이름과 Element Value값을 찍어준다.
Sub S_RQ()
    Dim fItem, fStr, fI, fComa
    fStr = "Dim "
    fI = 1
    For Each fItem In Request.QueryString
        IF Request.QueryString.Count <> fI Then
            fComa = ","
        Else
            fComa = ""
        End if
        fStr = fStr &"s_"& Lcase(Left(fItem,1)) & Mid(fItem,2,Len(fItem)) & fComa
        fI = fI + 1
    Next
    Response.Write fStr &"<br>"

    For Each fItem In Request.QueryString
        fStr = "s_"& Lcase(Left(fItem,1)) & Mid(fItem,2,Len(fItem)) & " =    "
        fStr = fStr & "nullFillWith(Request(" & chr(34) & fItem & chr(34) & "),"""", False)"
        Response.Write fStr & "<br>"
    Next
    Response.End
End Sub


'서버정보,스크립트정보,쿠키정보,세션정보
Sub getServerInfo()
    RW("<font color=""red"">ASP "& ScriptEngine &" Version "& ScriptEngineMajorVersion &"."& ScriptEngineMinorVersion &"</font><br>")
    RW(String(120, "="))

    Dim fStrVar,fNum,fCookie,fKey,fSession
    RW("<br><font color=""red"">Server 정보 (Request.ServerVariables)</font><br>")
    'Server 정보
    For Each fStrVar In Request.ServerVariables
        RW(fStrVar &" : "& Request.ServerVariables(fStrVar) &"<br>")
    Next

    RW(String(120, "="))
    RW("<br><font color=""red"">쿠키정보</font><br>")

    '쿠키정보
    For Each fCookie In Request.Cookies
      IF Not Request.Cookies(fCookie).HasKeys Then
        RW(fCookie &" : "& Request.Cookies(fCookie) &"<br>")
      Else
        For Each fKey In Request.Cookies(fCookie)
          RW(fCookie &"("& fKey &") : "& Request.Cookies(fCookie)(fKey) &"<br>")
        Next
      End if
    Next

    RW(String(120, "="))
    RW("<br><font color=""red"">세션정보</font><br>")
    '세션정보
    For Each fSession In Session.Contents
        RW(fSession &" : "& Session(fSession) & "<br>")
    Next

    RW(String(120, "="))
    RW("<br><font color=""red"">Other Variables</font><br>")

    RW("Session.Sessionid : "& Session.Sessionid &"<br>")
    RW("Server.MapPath : "& Server.MapPath ("/"))
    RE()
End Sub


'* 기능 : 파라메터를 받아서 post로 전송
'* furl : 전송할 액션 URL
Sub setActionPost(fUrl)
    Dim fBody
    fBody = ""
    fBody = fBody &"<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"& vbCrLf
    fBody = fBody &"<html xmlns=""http://www.w3.org/1999/xhtml"">"& vbCrLf
    fBody = fBody &"<head>"& vbCrLf
    fBody = fBody &"    <meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />"& vbCrLf
    fBody = fBody &"    <title>document</title>"& vbCrLf
    fBody = fBody &"</head>"& vbCrLf
    fBody = fBody &"<body>"& vbCrLf
    fBody = fBody &"<form name=""dataForm"" action="""& fUrl &""" method=""post"">"& vbCrLf
    'POST방식
    IF Lcase(Request.ServerVariables("REQUEST_METHOD")) = "post" Then
        For Each fElem In Request.Form
            fBody = fBody &"    <input type=""hidden"" name="""& fElem &""" value="""& Request(fElem) &""" />"& vbCrLf
        Next
    'GET방식
    Else
        For Each fElem In Request.QueryString
            fBody = fBody &"    <input type=""hidden"" name="""& fElem &""" value="""& Request(fElem) &""" />"& vbCrLf
        Next
    End if
    fBody = fBody &"</form>"& vbCrLf
    fBody = fBody &"<script type=""text/javascript"">"& vbCrLf
    fBody = fBody &"    document.dataForm.submit();"& vbCrLf
    fBody = fBody &"</script>"& vbCrLf
    fBody = fBody &"</body>"& vbCrLf
    fBody = fBody &"</html>"& vbCrLf
    RWE(fBody)
End Sub
%>