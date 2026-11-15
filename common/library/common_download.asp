<%@CodePage=65001%>
<%

'//=================================================================
'작  성 자 : @전윤기
'작  성 일 : @2010-04-26
'파일 내용 : @파일다운로드
'=================================================================//
Option Explicit
Response.CharSet = "utf-8"

%><!--#include virtual="/common/library/common.asp"--><!--#include virtual="/common/library/secret_config.asp"--><%
Dim sFile, sFolder
sFile	= nullFillWith(Request("file"), "", True)
sFolder	= nullFillWith(Request("folder"), "", True)

IF Not isNone(sFile) and Not isNone(sFolder) Then
	Dim vDown

	Set vDown = Server.CreateObject("DEXT.FileDownload")
	vDown.AuthKey = gloDextAuthKey
	

	vDown.Download gFileUpPath & sFolder & "\" & sFile, sFile, true, false
	Set vDown = Nothing
Else
	Call errorBack("존재하지 않는 파일입니다.")
End if

%>
