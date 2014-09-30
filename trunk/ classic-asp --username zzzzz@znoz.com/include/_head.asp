<!--#include file="_begin.asp"-->
<%
' ASP 페이지의 서버 처리 속도 측정
Dim ASP_Speed_BeginTime : ASP_Speed_BeginTime = Timer()

' UI 출력과 관련된 공통 변수


%><!DOCTYPE html>
<!--[if IE 8]> <html lang="en" class="ie8 no-js"> <![endif]-->
<!--[if IE 9]> <html lang="en" class="ie9 no-js"> <![endif]-->
<!--[if !IE]><!-->
<html lang="ko" class="no-js">
<!--<![endif]-->
<!-- BEGIN HEAD -->
<head>
<meta charset="utf-8"/>
<title><%=SITE_TITLE%></title>
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta content="width=device-width, initial-scale=1" name="viewport"/>
<meta content="<%=SITE_DESCRIPTION%>" name="description"/>
<meta content="<%=SITE_AUTHOR%>" name="author"/>

<!-- BEGIN CSS STYLES -->

<!-- END CSS STYLES -->
<link rel="shortcut icon" href="favicon.ico"/>
</head>

<!-- BEGIN BODY -->
<body>