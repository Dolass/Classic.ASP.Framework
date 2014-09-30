<%
' 공통 함수

' -------------- 에러 체크해서 메시지 출력 -------------
' strErrorMsg : 출력할 메시지
' ------------------------------------------------------
Public Function checkError(strErrorMsg)
     if Err.number <> 0 then
          if strErrorMsg = "" then
               Response.write "<BR>" & Request.ServerVariables("SCRIPT_NAME") & " : " & Err.Description & "<BR>"
          else
               Response.write strErrorMsg
          end if

          checkError = true
     Else
          checkError = false
     End if
End Function

' ----------------- 파일 내용 출력 -----------------
' filename : 출력할 파일명
' --------------------------------------------------
Public Function printFile(filename)
     Dim fso, path, data

     Set fso = Server.CreateObject("Scripting.FileSystemObject")
   
     path = Server.MapPath(".") & "\" & filename
     data = fso.OpenTextFile(path).ReadAll
   
     Response.write data

     Set fso = Nothing
End Function

' ----------------- 문자열에서 파일명 가져오기 -----------------
' text : 원본 텍스트
' -----------------------------------------------------------
Public Function getFilename(text)
     Dim regEx, result

     Set regEx = New RegExp

     regEx.Pattern = "[^\x00-\x1f\\?*:\"";|/]+$"
     regEx.IgnoreCase = True
     regEx.Global = True

     Set result = regEx.Execute(text)

     If 0 < result.count Then getFilename = result(0).Value
End Function

' ----------------- 문자열에서 파일 확장자 가져오기 ----------------
' text : 원본 텍스트
' --------------------------------------------------------------
Function getExt(fileName)
     Dim pos, ext

     pos = InstrRev(fileName, ".")

     GetExt = Mid(fileName, pos+1)
End Function

' --------- 해당 디렉토리가 있는지 검사하여 없으면 생성  ----------
' dir : 생성할 디렉토리
' 리턴 : 생성한 디렉토리, 실패할 경우 ""
' -----------------------------------------------------------------
Public Function makeDir(dir)
     Dim fso, f

     Set fso = Server.CreateObject("Scripting.FileSystemObject")
   
     If (fso.FolderExists(dir)) Then
          makeDir = ""
     else
          Set f = fso.CreateFolder(dir)             
          makeDir = f
     end If
    
     Set fso = Nothing
End Function

' --------------- 해당 파일 있는지 검사  -----------------
' path : 파일 경로
' 리턴 : 파일이 있으면 true
' -----------------------------------------------------
Public Function fileExists(path)
     Dim fso

     Set fso = Server.CreateObject("Scripting.FileSystemObject")
   
     fileExists = fso.FileExists(path)

     Set fso = Nothing
End Function

' --------------- 해당 파일 있는지 검사  -----------------
' url : 웹 상에서의 파일 경로
' 리턴 : 파일이 있으면 true
' -----------------------------------------------------
Public Function urlExists(url)
     urlExists = fileExists( Server.MapPath(url) )
End Function

' --------------- 파일 삭제  ----------------------------
' path : 파일 경로
' 리턴 : 파일이 있으면 true
' -----------------------------------------------------
Public Function fileDelete(path)
     Dim fso

     Set fso = Server.CreateObject("Scripting.FileSystemObject")
   
     fileDelete = fso.FileExists(path)

     If fileDelete Then fso.DeleteFile(path)
    
     Set fso = Nothing
End Function

' --------------- 폴더 삭제  ----------------------------
' path : 파일 경로
' 리턴 : 폴더가 있으면 true
' -----------------------------------------------------
Public Function folderDelete(path)
     Dim fso

     Set fso = Server.CreateObject("Scripting.FileSystemObject")
   
     folderDelete = fso.FolderExists(path)

     If folderDelete Then fso.DeleteFolder(path)
    
     Set fso = Nothing
End Function

' --------------- 파일 & 폴더 삭제  -----------------------
' path : 파일 경로
' 리턴 : 폴더가 있으면 true
' ------------------------------------------------------
Public Function fileFolderDelete(path)
     Call fileDelete(path)

     Dim fileName : fileName = getFilename(path)

     Call folderDelete( Replace(path, "\" & fileName, "") )
End Function
 
' ------------ 문자열의 확장자를 보고 이미지인지 검사 -----------
' text : 원본 텍스트
' -----------------------------------------------------------
Public Function isImage(text)
     Dim regEx, result

     Set regEx = New RegExp
    
     regEx.Pattern = ".(jpg|png|gif|jpeg)$"
     regEx.IgnoreCase = True
     regEx.Global = True
    
     isImage = regEx.Test(text)
End Function

' ---------- 배열에서 최저값을 찾아 해당 배열 인덱스 리턴 -------------
' array : 배열
' arrCount : 배열 갯수
' 리턴 : 최저값을 가지고 있는 배열의 인덱스, 최저값이 0일 경우 -1 리턴
' ----------------------------------------------------------------
Public Function getLowIndex(array, arrCount)
  Dim i, index
 
  index = 0
for i = 1 to arrCount - 1
  if 0 < array(i) And array(i) < array(index) then index = i
next

if array(index) <= 0 then index = -1

getLowIndex = index
end Function

' ---------- 배열에서 최고값을 찾아 해당 배열 인덱스 리턴 -------------
' array : 배열
' arrCount : 배열 갯수
' 리턴 : 최고값을 가지고 있는 배열의 인덱스, 최저값이 0일 경우 -1 리턴
' ----------------------------------------------------------------
Public Function getHighIndex(array, arrCount)
  Dim i, index
 
  index = 0
for i = 1 to arrCount - 1
  if 0 < array(i) And array(index) < array(i) then index = i
next

if array(index) <= 0 then index = -1

getHighIndex = index
end Function


' ---------- 배열에 특정 값이 있는지 체크 ---------------------------
' arr : 배열
' search : 찾을 값
' 리턴 : 있는 경우 true, 없으면 false
' ----------------------------------------------------------------
Function inArray(arr, search)
     Dim i

     inArray = False

     For i=0 To Ubound(arr)
          If Trim(arr(i)) = Trim(search) Then
               inArray = True
               Exit Function     
          End If
     Next
End Function


' ---------- HTML 문자열을 Decode  ------------------------------
' encodedstring : Server.HTMLEncode 로 인코딩된 문자열
' --------------------------------------------------------------
Public Function HTMLDecode(byVal encodedstring)
Dim tmp, i

tmp = encodedstring
tmp = Replace( tmp, "&quot;", chr(34) )
tmp = Replace( tmp, "&lt;"  , chr(60) )
tmp = Replace( tmp, "&gt;"  , chr(62) )
tmp = Replace( tmp, "&amp;" , chr(38) )
tmp = Replace( tmp, "&nbsp;", chr(32) )

For i = 1 to 255
     tmp = Replace( tmp, "&#" & i & ";", chr( i ) )
Next

HTMLDecode = tmp
End Function

' ---------- 자바 스크립트 alert 출력하고 target 위치로 이동 ---------------
' msg : 출력 메시지
' target : 이동할 위치, ""=history.back
' ------------------------------------------------------------------
Public Function Alert(msg, target)
Dim output

output = "<Script language='JavaScript' type='text/JavaScript'>" & vbcrlf
output = output & "alert('" & msg & "');" & vbcrlf

if target = "" then
  output = output & "history.back();" & vbcrlf
else
  output = output & "location.replace('" & target & "');" & vbcrlf
end if

output = output & "</Script>" & vbcrlf

Response.write output
End Function

' ---------- 설정 시간을 select 로 출력 --------------------------------
' defaultTime : 설정 시간(datetime)
' hourName : 시 select 이름
' MinName : 분 select 이름
' ------------------------------------------------------------------
Public Function printTimeSelect(defaultTime, hourName, MinName)
     Dim output, h, m, i
    
     h = Hour(defaultTime)
     m = Minute(defaultTime)
    
     output = "<Select name='" & hourName & "'>"
     for i = 0 to 23
          output = output & "<option value='" & i & "'"
          if i = h then output = output & " Selected"
          output = output & ">" & i & "</option>"
     next
     output = output & "</select>시 " & vbCrlf
    
     output = output & "<Select name='" & MinName & "'>"
     ' 5분 간격으로 출력
     for i = 0 to 59 step 5
          output = output & "<option value='" & i & "'"
          if i = m then output = output & " Selected"
          output = output & ">" & i & "</option>"
     next
     output = output & "</select>분 "
    
     Response.write output
End Function

' ---------- 설정 년월을 select 로 출력 -----------------------------
' defaultDate : 설정 년월(datetime)
' yearName : 년 select 이름
' monthName : 월 select 이름
' ---------------------------------------------------------------
Public Function printYMSelect(defaultDate, yearName, monthName)
     Dim output, y, m, i
    
     y = Year(defaultDate)
     m = Month(defaultDate)
    
     output = "<Select name='" & yearName & "'>"
     for i = (y - 10) to (y + 2)
          output = output & "<option value='" & i & "'"
          if i = y then output = output & " Selected"
          output = output & ">" & i & "</option>"
     next
     output = output & "</select>년 " & vbCrlf
    
     output = output & "<Select name='" & monthName & "'>"

     for i = 1 to 12
          output = output & "<option value='" & i & "'"
          if i = m then output = output & " Selected"
          output = output & ">" & i & "</option>"
     next
     output = output & "</select>월 "
    
     Response.write output
End Function

' ---------- 설정 년월일을 select 로 출력 -------------------------------
' defaultDate : 설정 년월(datetime)
' yearName : 년 select 이름
' monthName : 월 select 이름
' -------------------------------------------------------------------
Public Function printYMDSelect(defaultDate, yearName, monthName, dayName)
     Dim output, y, m, d, i
    
     y = Year(defaultDate)
     m = Month(defaultDate)
     d = Day(defaultDate)
    
     output = "<Select name='" & yearName & "' style='width:80px;'>"
     for i = (y - 10) to y
          output = output & "<option value='" & i & "'"
          if i = y then output = output & " Selected"
          output = output & ">" & i & "</option>"
     next
     output = output & "</select>년 " & vbCrlf
    
     output = output & "<Select name='" & monthName & "'  style='width:80px;'>"

     for i = 1 to 12
          If i < 10 Then
               output = output & "<option value='0" & i & "'"
               if i = m then output = output & " Selected"
               output = output & ">0" & i & "</option>"
          Else
               output = output & "<option value='" & i & "'"
               if i = m then output = output & " Selected"
               output = output & ">" & i & "</option>"
          End If
     next
     output = output & "</select>월 "

     output = output & "<Select name='" & dayName & "'  style='width:80px;'>"

     for i = 1 to 31
          If i < 10 Then
               output = output & "<option value='0" & i & "'"
               if i = d then output = output & " Selected"
               output = output & ">0" & i & "</option>"
          Else
               output = output & "<option value='" & i & "'"
               if i = d then output = output & " Selected"
               output = output & ">" & i & "</option>"
          End If
     next
     output = output & "</select>일 "
    
     Response.write output
End Function

' ---------- 생년월일을 select 로 출력 ----------------------------------
' defaultDate : 설정 년월(datetime)
' yearName : 년 select 이름
' monthName : 월 select 이름
' -------------------------------------------------------------------
Public Function printBirthdaySelect(yearName, monthName, dayName)
     Dim i

     Response.write "<Select name='" & yearName & "' style='width:100px;'>" & vbCrlf
     Response.write "<option value=''>선택</option>" & vbCrlf
     for i = (Year(Now) - 100) to Year(Now)
          Response.write "<option value='" & i & "'" & ">" & i & "</option>" & vbCrlf
     next
     Response.write "</select>년 " & vbCrlf
    
     Response.write "<Select name='" & monthName & "'  style='width:80px;'>" & vbCrlf
     Response.write "<option value=''>선택</option>" & vbCrlf
     for i = 1 to 12
          Response.write "<option value='" & Right("0" & i, 2) & "'" & ">" & i & "</option>" & vbCrlf
     next
     Response.write "</select>월 " & vbCrlf

     Response.write "<Select name='" & dayName & "'  style='width:80px;'>" & vbCrlf
     Response.write "<option value=''>선택</option>" & vbCrlf
     for i = 1 to 31
          Response.write "<option value='" & Right("0" & i, 2) & "'" & ">" & i & "</option>" & vbCrlf
     next
     Response.write "</select>일 " & vbCrlf
End Function

' -------------- 넘겨받은 폼 데이터를 모두 출력 -------------------------
' isEnd : true일 경우 ASP 수행 종료
' -------------------------------------------------------------------
Public Function printFormData(isEnd)
     dim item, i

     for each item in REQUEST.FORM
       for i=1 to REQUEST.FORM(item).Count
          response.write item & " : " & REQUEST.FORM(item)(i) & "<BR>" & vbcrlf
       next
     next
    
     if isEnd then response.end
end Function

' -------------- 문자열 길이 계산 ------------------------------
' text : 길이 계산할 문자열 (한글=2, 영숫자=1)
' ------------------------------------------------------------
function getKLength(text)
          dim realLen, lenText
          dim temp, i
         
          lenText = Len(text)
         
          realLen = 0
          for i = 1 to lenText
                    temp = Mid(text, i, 1)
                    if 4 < Len(Escape(temp)) then realLen = realLen + 1
                    realLen = realLen + 1
          next
         
          getKLength = realLen
end function

' -------------- 문자열 자르기 ---------------------------------
' text : 길이 계산할 문자열 (한글=2, 영숫자=1)
' length : 자를 문자열 길이    
' ------------------------------------------------------------
function getKLeft(text, length)
     dim realLen, realText, lenText
     dim temp, i
    
     lenText = Len(text)
    
     realLen = 0
     for i = 1 to lenText
          temp = Mid(text, i, 1)
          if 4 < Len(Escape(temp)) then realLen = realLen + 1

          realText = realText & temp
          realLen = realLen + 1

          if length <= realLen Then
               realText = realText & ".."
               Exit For
          End If
     next

     getKLeft = realText
end function    

' -------------- DB 내용으로 <select> 리스트 출력 --------------------------
' DB 쿼리하여 레코드셋 Open 한 상태라고 간주
' Select의 value=seqid
' DB : class.db.asp 의 DB 클래스
' name : Select의 name
' text : Select Option의 text로 설정될 DB 필드명
' value : Select Option의 value에 설정될 DB 필드명
' sel : 현재 선택되어 있는 항목
' nosel : 선택 항목이 없을때 디폴트 출력 메시지(전체,선택,등등), 사용안할때는=""
' ----------------------------------------------------------------------
Public Function printSelectCode(DB, name, text, value, sel, nosel)
     Dim post, addr, addr2, seqid
     Dim i, out

     with DB
          if Not .IsEOF then
               out = "<SELECT name='" & name & "'>" & vbCrLf
              
               if nosel <> "" then out = out & "<OPTION value=''>" & nosel & "</OPTION>" & vbCrLf

               Do Until .IsEOF
                    seqid = .getValue(value)
                    if Not IsNumeric(seqid) then seqid = RTrim(seqid)
                   
                    out = out & "<OPTION value='" & seqid & "'"
                    if seqid = sel then
                         out = out & " Selected"
                    end if
                    out = out & ">" & .getValue(text) & "</OPTION>" & vbCrLf

                    .MoveNext
               Loop
              
               out = out & "</SELECT>"
              
               Response.write out & vbCrLf
              
               printSelectCode = true
          else
               printSelectCode = false         
          end if
     End with
End Function


' -------------- DB 내용으로 <select> 리스트 출력 --------------------------
' DB 쿼리하여 레코드셋 Open 한 상태라고 간주
' Select의 value=seqid
' DB : class.db.asp 의 DB 클래스
' name : Select의 name
' text : Select Option의 text로 설정될 DB 필드명
' value : Select Option의 value에 설정될 DB 필드명
' sel : 현재 선택되어 있는 항목
' nosel : 선택 항목이 없을때 디폴트 출력 메시지(전체,선택,등등), 사용안할때는=""
' className : Select의 CSS 클래스명
' ----------------------------------------------------------------------
Public Function printSelectCodeClass(DB, name, text, value, sel, nosel, className)
     Dim post, addr, addr2, seqid
     Dim i, out

     with DB
          if Not .IsEOF then
               out = "<SELECT name='" & name & "' class='" & className & "'>" & vbCrLf
              
               if nosel <> "" then out = out & "<OPTION value=''>" & nosel & "</OPTION>" & vbCrLf

               Do Until .IsEOF

                    seqid = .getValue(value)
                    if Not IsNumeric(seqid) then seqid = RTrim(seqid)
                   
                    out = out & "<OPTION value='" & seqid & "'"
                    if seqid = sel then
                         out = out & " Selected"
                    end if
                    out = out & ">" & .getValue(text) & "</OPTION>" & vbCrLf

                    .MoveNext
               Loop
              
               out = out & "</SELECT>"
              
               Response.write out & vbCrLf
              
               printSelectCodeClass = true
          else
               printSelectCodeClass = false         
          end if
     End with
End Function


' -------------- DB 내용으로 <select> 리스트 출력 --------------------------
' DB 쿼리하여 레코드셋 Open 한 상태라고 간주
' Select의 value=seqid
' DB : class.db.asp 의 DB 클래스
' name : Select의 name
' text : Select Option의 text로 설정될 DB 필드명
' value : Select Option의 value에 설정될 DB 필드명
' sel : 현재 선택되어 있는 항목
' nosel : 선택 항목이 없을때 디폴트 출력 메시지(전체,선택,등등), 사용안할때는=""
' ----------------------------------------------------------------------
Public Function returnSelectCode(DB, name, text, value, sel, nosel)
     Dim post, addr, addr2, seqid
     Dim i, out

     with DB
          if Not .IsEOF then
               out = "<SELECT name='" & name & "'>" & vbCrLf
               if nosel <> "" then out = out & "<OPTION value=''>" & nosel & "</OPTION>" & vbCrLf

               Do Until .IsEOF
                    seqid = .getValue(value)
                    if Not IsNumeric(seqid) then seqid = RTrim(seqid)
                   
                    out = out & "<OPTION value='" & seqid & "'"
                    if seqid = sel then
                         out = out & " Selected"
                    end if
                    out = out & ">" & .getValue(text) & "</OPTION>" & vbCrLf

                    .MoveNext
               Loop
              
               out = out & "</SELECT>"
              
               returnSelectCode = out
          end if
     End with
End Function



' -------------- DB 내용으로 <select> 리스트 출력 --------------------------------------
' DB 쿼리하여 레코드셋 Open 한 상태라고 간주
' Select의 value=seqid
' DB : class.db.asp 의 DB 클래스
' name : Select의 name
' text : Select Option의 text로 설정될 DB 필드명
' value : Select Option의 value에 설정될 DB 필드명
' sel : 현재 선택되어 있는 항목
' nosel : 선택 항목이 없을때 디폴트 출력 메시지(전체,선택,등등), 사용안할때는=""
' onChange : JavaScript onChange 발생시 호출할 핸들러
' className : Select의 CSS 클래스명
' -----------------------------------------------------------------------------------
Public Function printSelectCodeOnChange(DB, name, text, value, sel, nosel, onChange, className)
     Dim post, addr, addr2, seqid
     Dim i, out

     with DB
          if Not .IsEOF then
               out = "<SELECT name='" & name & "' id='" & name & "' class='" & className & "' onChange='" & onChange & "'>" & vbCrLf
               if nosel <> "" then out = out & "<OPTION value=''>" & nosel & "</OPTION>" & vbCrLf
               If name = "colorlist" then out = out & "<OPTION value='c'>현재색상복사입력</OPTION>" & vbCrLf

               Do Until .IsEOF
                    seqid = .getValue(value)
                    if Not IsNumeric(seqid) then seqid = RTrim(seqid) Else seqid = CLng(seqid)
                   
                    out = out & "<OPTION value='" & seqid & "'"
                    if seqid = sel then
                         out = out & " Selected"
                    end if
                    out = out & ">" & .getValue(text) & "</OPTION>" & vbCrLf
                    .MoveNext
               Loop
              
               out = out & "</SELECT>"
                        
               printSelectCodeOnChange = true
          Else
               out = "<SELECT name='" & name & "' id='" & name & "' class='" & className & "' onChange='" & onChange & "' style='display:none'>" & vbCrLf
               out = out & "</SELECT>"

               printSelectCodeOnChange = false         
          end if

          Response.write out & vbCrLf
     End with
End Function

' -------------- DB 내용으로 <span> 리스트 출력 ---------------------------
' DB 쿼리하여 레코드셋 Open 한 상태라고 간주
' Select의 value=seqid
' DB : class.db.asp 의 DB 클래스
' ----------------------------------------------------------------------
Public Function printSpanSizecodeDetale(DB)
     Dim cateTemp,Out,i,displayValue,startSpan,endSpan,contentSpan
     Dim category,code
    
     With DB
          cateTemp = ""
          Out = ""
          i = 0
          displayValue = ""
          startSpan = ""
          endSpan = ""

          If Not .IsEOF Then
               Do Until .IsEOF
                    category = .GetValue("category")
                    code = .GetValue("code")
                    If i > 0 Then displayValue = "none"
                    If i > 0 Then endSpan = "</span>"

                    startSpan = "<span id='chk"& category &"' style='display:"& displayValue &"'>"
                    contentSpan = "<input type=checkbox name='"& category &"' value='"& code &"' />"& code

                    If category = cateTemp Then
                         Out = Out & contentSpan
                    Else
                         Out = Out & endSpan & startSpan & contentSpan
                    End If
               .MoveNext
               i=i+1
               cateTemp = category
               Loop
          End If

          Response.write Out &"</span>"& vbCrLf
     End With
End Function



' -------------- DB 내용으로 html 리스트 출력 ------------------------------
' 추가20140717
' DB 쿼리하여 레코드셋 Open 한 상태라고 간주
' DB : class.db.asp 의 DB 클래스
' name : Select의 name
' text : Select Option의 text로 설정될 DB 필드명
' value : Select Option의 value에 설정될 DB 필드명
' sel : 현재 선택되어 있는 항목
' selected : 선택된 항목인 경우 출력할 문자열
' 치환: {name}, {text}, {value}, {selected}, {class}
' ----------------------------------------------------------------------
Public Function printHTMLRow1(DB, name, text, value, sel, selected, lastClass, html)
     Dim post, addr, addr2, seqid,arrsel
     Dim i, out, outClass
    
     arrsel = Split(sel,",")
    
     with DB
          if Not .IsEOF then
              
               Do Until .IsEOF
                    out = html

                    out = Replace(out, "{name}", name)
                    out = Replace(out, "{text}", .getValue(text))
                    out = Replace(out, "{value}", .getValue(value))
                   
                    For i = 0 To UBound(arrsel) -1
                    If .getValue(value) = arrsel(i) Then out = Replace(out, "{selected}", selected)
                    Next
                   
                    .MoveNext
                   
                    If .IsEOF And lastClass <> "" Then outClass = lastClass Else outClass = ""

                    out = Replace(out, "{class}", outClass)

                    Response.write out
               Loop

          end if
     End with
End Function




' -------------- DB 내용으로 html 리스트 출력 --------------
' DB 쿼리하여 레코드셋 Open 한 상태라고 간주
' DB : class.db.asp 의 DB 클래스
' name : Select의 name
' text : Select Option의 text로 설정될 DB 필드명
' value : Select Option의 value에 설정될 DB 필드명
' sel : 현재 선택되어 있는 항목
' selected : 선택된 항목인 경우 출력할 문자열
' 치환: {name}, {text}, {value}, {selected}, {class}
Public Function printHTMLRow(DB, name, text, value, sel, selected, lastClass, html)
     Dim post, addr, addr2, seqid
     Dim i, out, outClass

     with DB
          if Not .IsEOF then
              
               Do Until .IsEOF
                    out = html

                    out = Replace(out, "{name}", name)
                    out = Replace(out, "{text}", .getValue(text))
                    out = Replace(out, "{value}", .getValue(value))

                    If .getValue(value) = sel Then out = Replace(out, "{selected}", selected)

                    .MoveNext
                   
                    If .IsEOF And lastClass <> "" Then outClass = lastClass Else outClass = ""

                    out = Replace(out, "{class}", outClass)

                    Response.write out
               Loop

          end if
     End with
End Function


' -------------- DB 내용으로 프린트용 리스트 출력 --------------
' DB 쿼리하여 레코드셋 Open 한 상태라고 간주
' DB : class.db.asp 의 DB 클래스
' name : Select의 name
' text : Select Option의 text로 설정될 DB 필드명
' value : Select Option의 value에 설정될 DB 필드명
' sel : 현재 선택되어 있는 항목
' normalText : 선택되지 않은 항목일 경우 출력할 문자열
' selectText : 선택된 항목인 경우 출력할 문자열
' lineMax : 한줄당 최대 출력 갯수
' lineBreak : 라인 종료 html
' 치환: {name}, {text}, {value}
Public Function printPrintableRow(DB, name, text, value, sel, normalText, selectText, lineMax, lineBreak)
     Dim i, out

     with DB
          if Not .IsEOF then
              
               i = 0
               Do Until .IsEOF
                    If .getValue(value) = sel Then
                         out = selectText
                    Else
                         out = normalText
                    End if

                    out = Replace(out, "{name}", name)
                    out = Replace(out, "{text}", .getValue(text))
                    out = Replace(out, "{value}", .getValue(value))
                   
                    Response.write out

                    i = i + 1

                    If (i Mod lineMax) = 0 Then Response.write lineBreak

                    .MoveNext
               Loop

          end if
     End with
End Function

' -------------- DB 내용으로 프린트용 리스트 출력(배열) --------------
' DB 쿼리하여 레코드셋 Open 한 상태라고 간주
' DB : class.db.asp 의 DB 클래스
' name : Select의 name
' text : Select Option의 text로 설정될 DB 필드명
' value : Select Option의 value에 설정될 DB 필드명
' sel : 현재 선택되어 있는 항목
' normalText : 선택되지 않은 항목일 경우 출력할 문자열
' selectText : 선택된 항목인 경우 출력할 문자열
' lineMax : 한줄당 최대 출력 갯수
' lineBreak : 라인 종료 html
' 치환: {name}, {text}, {value}
Public Function printPrintableRow1(DB, name, text, value, sel, normalText, selectText, lineMax, lineBreak)
     Dim i, Out , arrsel,ii,sel_select

     arrsel = Split(sel,",")

     with DB
          if Not .IsEOF then
              
              
               i = 0
               Do Until .IsEOF
                   
                    For ii = 0 To UBound(arrsel) -1
                         IF .getValue(value) = arrsel(ii) Then
                              sel_select = arrsel(ii)
                              Exit For
                         End If
                    NEXT    
                   
                    If .getValue(value) = sel_select Then
                         out = selectText
                    Else
                         out = normalText
                    End if
                   
                    out = Replace(out, "{name}", name)
                    out = Replace(out, "{text}", .getValue(text))
                    out = Replace(out, "{value}", .getValue(value))
                                       
                    Response.write out

                    i = i + 1
              
                    If (i Mod lineMax) = 0 Then Response.write lineBreak
                   
    
                    .MoveNext
              
               Loop
              


          end if
     End with
End Function

' ------------------ 대소문자 무시 치환 함수 ---------------------
' text : 원본 텍스트
' search : 찾을 단어
' rep : 바꿀 단어
' --------------------------------------------------------------
Public Function nReplaceAll(text, search, rep)
     Dim regEx

     Set regEx = New RegExp
    
     regEx.Pattern = search
     regEx.IgnoreCase = True
     regEx.Global = True
    
     nReplaceAll = regEx.Replace(text, rep)
End Function

' ----------------- 10진수값을 32진수로 코드로 변경 ------------------
' dec : 10진수
' 리턴 : 32진수 - arrGe 표에 근거한 숫자 표현
' 재귀호출 사용하므로 수정시 스택 오버플로우에 주의
' 함수 용도 : 시간에 근거한 코드 생성시(주문번호 등)
' --------------------------------------------------------------------
function DecToGe(dec)
     dim arrGeTable

     arrGeTable = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", _
                    "G", "H", "J", "K", "L", "M", "N", "P", "Q", "R", "S", "T", "W", "X", "Y", "Z")

     if 32 <= dec then
          DecToGe = DecToGe(dec \ 32) & arrGeTable(dec Mod 32)
     else
          DecToGe = arrGeTable(dec)
     end if
end Function

' ----------------- 시각을 초단위 절대값으로 ---------------
' datetime : 변환을 원하는 시각
' 리턴 : 1월 1일 0시 0분 0초부터 시작한 초단위 절대값
' ----------------------------------------------------------
function getYearSecond(datetime)
     dim d, h, n, s
    
     d = DatePart("y", datetime)
     h = (d * 24) + hour(datetime)
     n = (h * 24) + minute(datetime)
     s = (n * 24) + second(datetime)
    
     getYearSecond = s
end function

' ----------------- 주문번호생성 ---------------
' length : 주문번호 자리수(7 ~ 15)
' 리턴 : 생성된 주문번호 - 현재시간값 + 세션값
function makeOrderNo(length)
     if 6 < length then
          dim no
         
          no = "0000" & DecToGe( getYearSecond(now) )
    
          makeOrderNo = DecToGe(Right(Year(now), 1)) & Right(no, 5) & Right(session.sessionID, length - 6)
     else
          Err.Raise vbObjectError+1000, makeOrderNo, "makeOrderNo 함수의 인자 length는 6보다 커야 합니다."
     end if
end function

' ------------------ 텍스트 파일 읽기 ---------------------
' filepath : 파일 경로
' 리턴 : 텍스트 파일 전체 문자열 데이터
' ---------------------------------------------------------
Public Function readFile(filepath)
     Dim fso, path

     Set fso = Server.CreateObject("Scripting.FileSystemObject")
    
     path = Server.MapPath(filepath)
    
     readFile = fso.OpenTextFile(path).ReadAll
End Function

' -------------------- 출력 후 종료 -----------------------
' msg : 출력값
' 디버그용
' ---------------------------------------------------------
Public Function Trace(msg)
     Response.Write(msg)
    
     response.End()
End Function

' ------------------ 페이지 번호 출력 ---------------------
' page : 현재 페이지
' pageCount : 페이지 전체 갯수
' gotoCount : 한번에 출력할 페이지 번호  최대 갯수
' param : 추가 파라미터
' pagePrev : 이전 출력 표시용
' pageNext : 다음 출력 표시용
' ---------------------------------------------------------
Private Function printPageGoto(page, pageCount, gotoCount, param, pagePrev, pageNext)
     Dim blockpage, i, pageNo, loopLastPage
     Dim aList, out
    
     page = Cint(page)
     pageCount = Cint(pageCount)
     gotoCount = Cint(gotoCount)
    
     if gotoCount < pageCount  then
          loopLastPage = gotoCount
     else
          loopLastPage = pageCount
     end if

     blockpage=Int( (page-1) / gotoCount ) * gotoCount + 1

     if param = "" then
          aList = "?page="
     else
          aList = "?" & param & "&page="
     end if

     ' 시작 페이지
     pageNo = "<b>##</b>&nbsp;"

     ' 이전 gotoCount 개
     if blockPage = 1 Then
          out = out & nReplaceAll(pagePrev, "##", gotoCount) & " "
     Else
          out = out &"<a href='" & aList & blockPage - gotoCount & "'>" & nReplaceAll(pagePrev, "##", blockpage) & "</a> "
     End If

     ' 번호 부분
     For i = 1 To loopLastPage
          If blockpage=int(page) Then
               out = out & " <font size=2 color= gray>" & nReplaceAll(pageNo, "##", blockpage) & "</font> "
          Else
               out = out &" <a href='" & aList & blockpage & "'>" & nReplaceAll(pageNo, "##", blockpage) & "</a> "
          End If
         
          blockpage=blockpage+1
         
          if pageCount < blockpage Then Exit For
     Next

     ' 다음 gotoCount 개
     if i <= gotoCount  Then
          out = out & nReplaceAll(pageNext, "##", gotoCount) & " "
     Else
          out = out & "<a href='" & aList & blockpage & "'>" & nReplaceAll(pageNext, "##", blockpage) & "</a> "
     End If

     response.write out
End Function         



' ------------------ 이메일발송 (Window 2003) ---------------------
' tomail : 받는메일
' toname : 받는사람
' frommail : 보내는메일
' fromname : 보내는사람
' subject : 메일제목
' body : 메일내용
' Path : 첨부파일 경로
' ---------------------------------------------------------
Public Function CDOSendMail(tomail,toname,frommail,fromname,subject,body,Path)
     Dim CDOobjMail
        Set CDOobjMail = CreateObject("CDO.Message")

          With CDOobjMail
               .From= fromname & "<" & frommail & ">"
               .To = toname& "<" & tomail & ">" 
               .Subject = subject
               .HTMLBody = body
               IF Path<>"" THEN
               .AddAttachment(Path)
               END IF
               .Send           
             End With
        Set CDOobjMail = Nothing

        CDOSendMail=true
End Function


' ------------------ 이메일발송 (Window 2000) ---------------------
' tomail : 받는메일
' toname : 받는사람
' frommail : 보내는메일
' fromname : 보내는사람
' subject : 메일제목
' body : 메일내용
' ---------------------------------------------------------
Public Sub SendMail(tomail,toname,frommail,fromname,subject,body)
     Dim objMail
        Set objMail = Server.CreateObject("CDONTS.NewMail")
          With ObjMail
               .From= fromname & "<" & frommail & ">"
               .Subject = subject
               .To = toname & "<" & tomail & ">"
               .BodyFormat = 0  
               .MailFormat = 0  
               .body = body  
               .Send           
             End With
        Set objMail = Nothing
End Sub


' ------------------ 로그인 체크  ---------------------
' return_url : 로그인 후 돌아갈 페이지
' ---------------------------------------------------------
Public Sub Login_Chk(return_url)
     IF Request.Cookies(CK_LOGIN)(CK_USERID) = "" THEN
          Session("SS_RETURNURL") = return_url
          Alert = Alert("로그인 후 이용하세요", "/13_login/li_login.asp")    
          Response.End
     END IF         
End Sub



' ------------------ 숫자타입 변경  ---------------------
' 1자리 숫자앞에 0 추가
' ---------------------------------------------------------
Function setZero(val)
     If val < 10 Then
          setZero = "0" & val
     Else
          setZero = val
     End If
End Function


' ------------------ Request.Form 데이터 받기 ----------------
' dataName : Form 데이터 이름
' defaultValue : 초기값
' 리턴 : 받은 값이 있으면 해당값 리턴, 비어있으면 defaultValue 리턴
' ------------------------------------------------------------
Function RequestForm(dataName, defaultValue)
     Dim tmpData

     tmpData = Request.Form(dataName)

     if tmpData = "" then
          RequestForm = defaultValue
     else
          RequestForm = tmpData
     end if
End Function


' ------------------ Request.QueryString 데이터 받기 ----------------
' dataName : QueryString 데이터 이름
' defaultValue : 초기값
' 리턴 : 받은 값이 있으면 해당값 리턴, 비어있으면 defaultValue 리턴
' ------------------------------------------------------------
Function RequestQuery(dataName, defaultValue)
     Dim tmpData

     tmpData = Request.QueryString(dataName)

     if tmpData = "" then
          RequestQuery = defaultValue
     else
          RequestQuery = tmpData
     end if
End Function


' ------------------ Request.QueryString 데이터 받거나, 값이 없으면 쿠키값 설정 -----------------
' dataName : QueryString 데이터 이름 (cookie 이름)
' defaultValue : 초기값
' 리턴 : 받은 값이 있으면 해당값 리턴, 비어있으면 쿠키 저장값, 쿠키값도 없으면 defaultValue 리턴
' 현재 페이지의 파라미터 init=1 인 경우 쿠키 초기화
' -----------------------------------------------------------------------------------------------
Function RequestQueryCookie(dataName, defaultValue)
     Dim cookName : cookName = Request.ServerVariables("URL") & "?" & dataName
     Dim tmpData, isInit

     tmpData = Request.QueryString(dataName)

     If Request.QueryString("init") = "1" Then isInit = True Else isInit = False

     if tmpData = "" Then
          If Not isInit Then tmpData = Request.Cookies(cookName)

          if tmpData = "" Then
               tmpData = defaultValue
          End If
     end if

     Response.Cookies(cookName) = tmpData

     RequestQueryCookie = tmpData
End Function


' ------------------ 배열 출력 -------------------------
' arr : 출력할 배열
' begin_idx : 배열 시작값
' end_idx : 배열 종료값, 0보다 작은 경우 UBound 사용
' html : 출력할 HTML
'
' HTML 치환값
' {index} : 배열 인덱스
' {data} : 배열 값
' ----------------------------------------------------
Function printArray(arr, begin_idx, end_idx, html)
     Dim i, html_out

     If end_idx < 0 Then end_idx = UBound(arr)

     For i = begin_idx To end_idx
          html_out = Replace(html, "{index}", i)
          html_out = Replace(html_out, "{data}", arr(i))

          Response.write html_out
     Next
End Function

%>

