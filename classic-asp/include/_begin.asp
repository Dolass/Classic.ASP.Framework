<%@ CodePage=65001 Language="VBScript"%> 
<%	
Option Explicit

' Response 기본 설정
With Response
	.Buffer = True
	.Expires = -1
	.AddHeader "pragma", "no-cache"
	.AddHeader "cache-control", "no-store"
	.ContentType = "text/html"
	.CharSet = "UTF-8"
End With

' Session 기본 설정
With Session
	.CodePage = "65001"
End With

' 사이트 타이틀
Const SITE_TITLE = "Classic ASP"
Const SITE_COPYRIGHT = "2014 &copy; [회사명]. All rights reserved."
Const SITE_DESCRIPTION = ""
Const SITE_AUTHOR = "zzzzz@znoz.com"
%>
