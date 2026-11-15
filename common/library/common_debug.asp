<%
'//=================================================================
'작  성 자 : nangido
'작  성 일 : 20210331
'파일 내용 : 웹 debuging을 위한 log 저장
'=================================================================//
Function Debug_addWebLog(path, debugtype, logtext,appcode)


    Dim rs , arrParam(4) , rsRetVal

    
    arrParam(0) =   nullFillWith(path ,                 "" , False)
    arrParam(1) =   nullFillWith(debugtype ,            "" , False)
    arrParam(2) =   nullFillWith(logtext ,              "" , False)
    arrParam(3) =   nullFillWith(appcode,               "" , False)
    arrParam(4) =   nullFillWith(session("USER_ID"),    "" , False)

    rsResult = ""
    Set rs = getSp("USP_TBL_DEBUG_WEBLOG_INS" , arrParam)

    If Not(rs.Eof Or rs.Bof) Then 
        rsRetVal  =   rs("RetVal")
    End If 
    rsClose(rs) 
    Debug_addWebLog  =   rsRetVal 

End Function
%>
