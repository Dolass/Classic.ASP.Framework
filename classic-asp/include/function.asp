<%
' ���� �Լ�

' -------------- ���� üũ�ؼ� �޽��� ��� -------------
' strErrorMsg : ����� �޽���
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

' ----------------- ���� ���� ��� -----------------
' filename : ����� ���ϸ�
' --------------------------------------------------
Public Function printFile(filename)
     Dim fso, path, data

     Set fso = Server.CreateObject("Scripting.FileSystemObject")
   
     path = Server.MapPath(".") & "\" & filename
     data = fso.OpenTextFile(path).ReadAll
   
     Response.write data

     Set fso = Nothing
End Function

' ----------------- ���ڿ����� ���ϸ� �������� -----------------
' text : ���� �ؽ�Ʈ
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

' ----------------- ���ڿ����� ���� Ȯ���� �������� ----------------
' text : ���� �ؽ�Ʈ
' --------------------------------------------------------------
Function getExt(fileName)
     Dim pos, ext

     pos = InstrRev(fileName, ".")

     GetExt = Mid(fileName, pos+1)
End Function

' --------- �ش� ���丮�� �ִ��� �˻��Ͽ� ������ ����  ----------
' dir : ������ ���丮
' ���� : ������ ���丮, ������ ��� ""
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

' --------------- �ش� ���� �ִ��� �˻�  -----------------
' path : ���� ���
' ���� : ������ ������ true
' -----------------------------------------------------
Public Function fileExists(path)
     Dim fso

     Set fso = Server.CreateObject("Scripting.FileSystemObject")
   
     fileExists = fso.FileExists(path)

     Set fso = Nothing
End Function

' --------------- �ش� ���� �ִ��� �˻�  -----------------
' url : �� �󿡼��� ���� ���
' ���� : ������ ������ true
' -----------------------------------------------------
Public Function urlExists(url)
     urlExists = fileExists( Server.MapPath(url) )
End Function

' --------------- ���� ����  ----------------------------
' path : ���� ���
' ���� : ������ ������ true
' -----------------------------------------------------
Public Function fileDelete(path)
     Dim fso

     Set fso = Server.CreateObject("Scripting.FileSystemObject")
   
     fileDelete = fso.FileExists(path)

     If fileDelete Then fso.DeleteFile(path)
    
     Set fso = Nothing
End Function

' --------------- ���� ����  ----------------------------
' path : ���� ���
' ���� : ������ ������ true
' -----------------------------------------------------
Public Function folderDelete(path)
     Dim fso

     Set fso = Server.CreateObject("Scripting.FileSystemObject")
   
     folderDelete = fso.FolderExists(path)

     If folderDelete Then fso.DeleteFolder(path)
    
     Set fso = Nothing
End Function

' --------------- ���� & ���� ����  -----------------------
' path : ���� ���
' ���� : ������ ������ true
' ------------------------------------------------------
Public Function fileFolderDelete(path)
     Call fileDelete(path)

     Dim fileName : fileName = getFilename(path)

     Call folderDelete( Replace(path, "\" & fileName, "") )
End Function
 
' ------------ ���ڿ��� Ȯ���ڸ� ���� �̹������� �˻� -----------
' text : ���� �ؽ�Ʈ
' -----------------------------------------------------------
Public Function isImage(text)
     Dim regEx, result

     Set regEx = New RegExp
    
     regEx.Pattern = ".(jpg|png|gif|jpeg)$"
     regEx.IgnoreCase = True
     regEx.Global = True
    
     isImage = regEx.Test(text)
End Function

' ---------- �迭���� �������� ã�� �ش� �迭 �ε��� ���� -------------
' array : �迭
' arrCount : �迭 ����
' ���� : �������� ������ �ִ� �迭�� �ε���, �������� 0�� ��� -1 ����
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

' ---------- �迭���� �ְ��� ã�� �ش� �迭 �ε��� ���� -------------
' array : �迭
' arrCount : �迭 ����
' ���� : �ְ��� ������ �ִ� �迭�� �ε���, �������� 0�� ��� -1 ����
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


' ---------- �迭�� Ư�� ���� �ִ��� üũ ---------------------------
' arr : �迭
' search : ã�� ��
' ���� : �ִ� ��� true, ������ false
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


' ---------- HTML ���ڿ��� Decode  ------------------------------
' encodedstring : Server.HTMLEncode �� ���ڵ��� ���ڿ�
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

' ---------- �ڹ� ��ũ��Ʈ alert ����ϰ� target ��ġ�� �̵� ---------------
' msg : ��� �޽���
' target : �̵��� ��ġ, ""=history.back
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

' ---------- ���� �ð��� select �� ��� --------------------------------
' defaultTime : ���� �ð�(datetime)
' hourName : �� select �̸�
' MinName : �� select �̸�
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
     output = output & "</select>�� " & vbCrlf
    
     output = output & "<Select name='" & MinName & "'>"
     ' 5�� �������� ���
     for i = 0 to 59 step 5
          output = output & "<option value='" & i & "'"
          if i = m then output = output & " Selected"
          output = output & ">" & i & "</option>"
     next
     output = output & "</select>�� "
    
     Response.write output
End Function

' ---------- ���� ����� select �� ��� -----------------------------
' defaultDate : ���� ���(datetime)
' yearName : �� select �̸�
' monthName : �� select �̸�
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
     output = output & "</select>�� " & vbCrlf
    
     output = output & "<Select name='" & monthName & "'>"

     for i = 1 to 12
          output = output & "<option value='" & i & "'"
          if i = m then output = output & " Selected"
          output = output & ">" & i & "</option>"
     next
     output = output & "</select>�� "
    
     Response.write output
End Function

' ---------- ���� ������� select �� ��� -------------------------------
' defaultDate : ���� ���(datetime)
' yearName : �� select �̸�
' monthName : �� select �̸�
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
     output = output & "</select>�� " & vbCrlf
    
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
     output = output & "</select>�� "

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
     output = output & "</select>�� "
    
     Response.write output
End Function

' ---------- ��������� select �� ��� ----------------------------------
' defaultDate : ���� ���(datetime)
' yearName : �� select �̸�
' monthName : �� select �̸�
' -------------------------------------------------------------------
Public Function printBirthdaySelect(yearName, monthName, dayName)
     Dim i

     Response.write "<Select name='" & yearName & "' style='width:100px;'>" & vbCrlf
     Response.write "<option value=''>����</option>" & vbCrlf
     for i = (Year(Now) - 100) to Year(Now)
          Response.write "<option value='" & i & "'" & ">" & i & "</option>" & vbCrlf
     next
     Response.write "</select>�� " & vbCrlf
    
     Response.write "<Select name='" & monthName & "'  style='width:80px;'>" & vbCrlf
     Response.write "<option value=''>����</option>" & vbCrlf
     for i = 1 to 12
          Response.write "<option value='" & Right("0" & i, 2) & "'" & ">" & i & "</option>" & vbCrlf
     next
     Response.write "</select>�� " & vbCrlf

     Response.write "<Select name='" & dayName & "'  style='width:80px;'>" & vbCrlf
     Response.write "<option value=''>����</option>" & vbCrlf
     for i = 1 to 31
          Response.write "<option value='" & Right("0" & i, 2) & "'" & ">" & i & "</option>" & vbCrlf
     next
     Response.write "</select>�� " & vbCrlf
End Function

' -------------- �Ѱܹ��� �� �����͸� ��� ��� -------------------------
' isEnd : true�� ��� ASP ���� ����
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

' -------------- ���ڿ� ���� ��� ------------------------------
' text : ���� ����� ���ڿ� (�ѱ�=2, ������=1)
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

' -------------- ���ڿ� �ڸ��� ---------------------------------
' text : ���� ����� ���ڿ� (�ѱ�=2, ������=1)
' length : �ڸ� ���ڿ� ����    
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

' -------------- DB �������� <select> ����Ʈ ��� --------------------------
' DB �����Ͽ� ���ڵ�� Open �� ���¶�� ����
' Select�� value=seqid
' DB : class.db.asp �� DB Ŭ����
' name : Select�� name
' text : Select Option�� text�� ������ DB �ʵ��
' value : Select Option�� value�� ������ DB �ʵ��
' sel : ���� ���õǾ� �ִ� �׸�
' nosel : ���� �׸��� ������ ����Ʈ ��� �޽���(��ü,����,���), �����Ҷ���=""
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


' -------------- DB �������� <select> ����Ʈ ��� --------------------------
' DB �����Ͽ� ���ڵ�� Open �� ���¶�� ����
' Select�� value=seqid
' DB : class.db.asp �� DB Ŭ����
' name : Select�� name
' text : Select Option�� text�� ������ DB �ʵ��
' value : Select Option�� value�� ������ DB �ʵ��
' sel : ���� ���õǾ� �ִ� �׸�
' nosel : ���� �׸��� ������ ����Ʈ ��� �޽���(��ü,����,���), �����Ҷ���=""
' className : Select�� CSS Ŭ������
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


' -------------- DB �������� <select> ����Ʈ ��� --------------------------
' DB �����Ͽ� ���ڵ�� Open �� ���¶�� ����
' Select�� value=seqid
' DB : class.db.asp �� DB Ŭ����
' name : Select�� name
' text : Select Option�� text�� ������ DB �ʵ��
' value : Select Option�� value�� ������ DB �ʵ��
' sel : ���� ���õǾ� �ִ� �׸�
' nosel : ���� �׸��� ������ ����Ʈ ��� �޽���(��ü,����,���), �����Ҷ���=""
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



' -------------- DB �������� <select> ����Ʈ ��� --------------------------------------
' DB �����Ͽ� ���ڵ�� Open �� ���¶�� ����
' Select�� value=seqid
' DB : class.db.asp �� DB Ŭ����
' name : Select�� name
' text : Select Option�� text�� ������ DB �ʵ��
' value : Select Option�� value�� ������ DB �ʵ��
' sel : ���� ���õǾ� �ִ� �׸�
' nosel : ���� �׸��� ������ ����Ʈ ��� �޽���(��ü,����,���), �����Ҷ���=""
' onChange : JavaScript onChange �߻��� ȣ���� �ڵ鷯
' className : Select�� CSS Ŭ������
' -----------------------------------------------------------------------------------
Public Function printSelectCodeOnChange(DB, name, text, value, sel, nosel, onChange, className)
     Dim post, addr, addr2, seqid
     Dim i, out

     with DB
          if Not .IsEOF then
               out = "<SELECT name='" & name & "' id='" & name & "' class='" & className & "' onChange='" & onChange & "'>" & vbCrLf
               if nosel <> "" then out = out & "<OPTION value=''>" & nosel & "</OPTION>" & vbCrLf
               If name = "colorlist" then out = out & "<OPTION value='c'>������󺹻��Է�</OPTION>" & vbCrLf

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

' -------------- DB �������� <span> ����Ʈ ��� ---------------------------
' DB �����Ͽ� ���ڵ�� Open �� ���¶�� ����
' Select�� value=seqid
' DB : class.db.asp �� DB Ŭ����
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



' -------------- DB �������� html ����Ʈ ��� ------------------------------
' �߰�20140717
' DB �����Ͽ� ���ڵ�� Open �� ���¶�� ����
' DB : class.db.asp �� DB Ŭ����
' name : Select�� name
' text : Select Option�� text�� ������ DB �ʵ��
' value : Select Option�� value�� ������ DB �ʵ��
' sel : ���� ���õǾ� �ִ� �׸�
' selected : ���õ� �׸��� ��� ����� ���ڿ�
' ġȯ: {name}, {text}, {value}, {selected}, {class}
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




' -------------- DB �������� html ����Ʈ ��� --------------
' DB �����Ͽ� ���ڵ�� Open �� ���¶�� ����
' DB : class.db.asp �� DB Ŭ����
' name : Select�� name
' text : Select Option�� text�� ������ DB �ʵ��
' value : Select Option�� value�� ������ DB �ʵ��
' sel : ���� ���õǾ� �ִ� �׸�
' selected : ���õ� �׸��� ��� ����� ���ڿ�
' ġȯ: {name}, {text}, {value}, {selected}, {class}
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


' -------------- DB �������� ����Ʈ�� ����Ʈ ��� --------------
' DB �����Ͽ� ���ڵ�� Open �� ���¶�� ����
' DB : class.db.asp �� DB Ŭ����
' name : Select�� name
' text : Select Option�� text�� ������ DB �ʵ��
' value : Select Option�� value�� ������ DB �ʵ��
' sel : ���� ���õǾ� �ִ� �׸�
' normalText : ���õ��� ���� �׸��� ��� ����� ���ڿ�
' selectText : ���õ� �׸��� ��� ����� ���ڿ�
' lineMax : ���ٴ� �ִ� ��� ����
' lineBreak : ���� ���� html
' ġȯ: {name}, {text}, {value}
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

' -------------- DB �������� ����Ʈ�� ����Ʈ ���(�迭) --------------
' DB �����Ͽ� ���ڵ�� Open �� ���¶�� ����
' DB : class.db.asp �� DB Ŭ����
' name : Select�� name
' text : Select Option�� text�� ������ DB �ʵ��
' value : Select Option�� value�� ������ DB �ʵ��
' sel : ���� ���õǾ� �ִ� �׸�
' normalText : ���õ��� ���� �׸��� ��� ����� ���ڿ�
' selectText : ���õ� �׸��� ��� ����� ���ڿ�
' lineMax : ���ٴ� �ִ� ��� ����
' lineBreak : ���� ���� html
' ġȯ: {name}, {text}, {value}
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

' ------------------ ��ҹ��� ���� ġȯ �Լ� ---------------------
' text : ���� �ؽ�Ʈ
' search : ã�� �ܾ�
' rep : �ٲ� �ܾ�
' --------------------------------------------------------------
Public Function nReplaceAll(text, search, rep)
     Dim regEx

     Set regEx = New RegExp
    
     regEx.Pattern = search
     regEx.IgnoreCase = True
     regEx.Global = True
    
     nReplaceAll = regEx.Replace(text, rep)
End Function

' ----------------- 10�������� 32������ �ڵ�� ���� ------------------
' dec : 10����
' ���� : 32���� - arrGe ǥ�� �ٰ��� ���� ǥ��
' ���ȣ�� ����ϹǷ� ������ ���� �����÷ο쿡 ����
' �Լ� �뵵 : �ð��� �ٰ��� �ڵ� ������(�ֹ���ȣ ��)
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

' ----------------- �ð��� �ʴ��� ���밪���� ---------------
' datetime : ��ȯ�� ���ϴ� �ð�
' ���� : 1�� 1�� 0�� 0�� 0�ʺ��� ������ �ʴ��� ���밪
' ----------------------------------------------------------
function getYearSecond(datetime)
     dim d, h, n, s
    
     d = DatePart("y", datetime)
     h = (d * 24) + hour(datetime)
     n = (h * 24) + minute(datetime)
     s = (n * 24) + second(datetime)
    
     getYearSecond = s
end function

' ----------------- �ֹ���ȣ���� ---------------
' length : �ֹ���ȣ �ڸ���(7 ~ 15)
' ���� : ������ �ֹ���ȣ - ����ð��� + ���ǰ�
function makeOrderNo(length)
     if 6 < length then
          dim no
         
          no = "0000" & DecToGe( getYearSecond(now) )
    
          makeOrderNo = DecToGe(Right(Year(now), 1)) & Right(no, 5) & Right(session.sessionID, length - 6)
     else
          Err.Raise vbObjectError+1000, makeOrderNo, "makeOrderNo �Լ��� ���� length�� 6���� Ŀ�� �մϴ�."
     end if
end function

' ------------------ �ؽ�Ʈ ���� �б� ---------------------
' filepath : ���� ���
' ���� : �ؽ�Ʈ ���� ��ü ���ڿ� ������
' ---------------------------------------------------------
Public Function readFile(filepath)
     Dim fso, path

     Set fso = Server.CreateObject("Scripting.FileSystemObject")
    
     path = Server.MapPath(filepath)
    
     readFile = fso.OpenTextFile(path).ReadAll
End Function

' -------------------- ��� �� ���� -----------------------
' msg : ��°�
' ����׿�
' ---------------------------------------------------------
Public Function Trace(msg)
     Response.Write(msg)
    
     response.End()
End Function

' ------------------ ������ ��ȣ ��� ---------------------
' page : ���� ������
' pageCount : ������ ��ü ����
' gotoCount : �ѹ��� ����� ������ ��ȣ  �ִ� ����
' param : �߰� �Ķ����
' pagePrev : ���� ��� ǥ�ÿ�
' pageNext : ���� ��� ǥ�ÿ�
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

     ' ���� ������
     pageNo = "<b>##</b>&nbsp;"

     ' ���� gotoCount ��
     if blockPage = 1 Then
          out = out & nReplaceAll(pagePrev, "##", gotoCount) & " "
     Else
          out = out &"<a href='" & aList & blockPage - gotoCount & "'>" & nReplaceAll(pagePrev, "##", blockpage) & "</a> "
     End If

     ' ��ȣ �κ�
     For i = 1 To loopLastPage
          If blockpage=int(page) Then
               out = out & " <font size=2 color= gray>" & nReplaceAll(pageNo, "##", blockpage) & "</font> "
          Else
               out = out &" <a href='" & aList & blockpage & "'>" & nReplaceAll(pageNo, "##", blockpage) & "</a> "
          End If
         
          blockpage=blockpage+1
         
          if pageCount < blockpage Then Exit For
     Next

     ' ���� gotoCount ��
     if i <= gotoCount  Then
          out = out & nReplaceAll(pageNext, "##", gotoCount) & " "
     Else
          out = out & "<a href='" & aList & blockpage & "'>" & nReplaceAll(pageNext, "##", blockpage) & "</a> "
     End If

     response.write out
End Function         



' ------------------ �̸��Ϲ߼� (Window 2003) ---------------------
' tomail : �޴¸���
' toname : �޴»��
' frommail : �����¸���
' fromname : �����»��
' subject : ��������
' body : ���ϳ���
' Path : ÷������ ���
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


' ------------------ �̸��Ϲ߼� (Window 2000) ---------------------
' tomail : �޴¸���
' toname : �޴»��
' frommail : �����¸���
' fromname : �����»��
' subject : ��������
' body : ���ϳ���
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


' ------------------ �α��� üũ  ---------------------
' return_url : �α��� �� ���ư� ������
' ---------------------------------------------------------
Public Sub Login_Chk(return_url)
     IF Request.Cookies(CK_LOGIN)(CK_USERID) = "" THEN
          Session("SS_RETURNURL") = return_url
          Alert = Alert("�α��� �� �̿��ϼ���", "/13_login/li_login.asp")    
          Response.End
     END IF         
End Sub



' ------------------ ����Ÿ�� ����  ---------------------
' 1�ڸ� ���ھտ� 0 �߰�
' ---------------------------------------------------------
Function setZero(val)
     If val < 10 Then
          setZero = "0" & val
     Else
          setZero = val
     End If
End Function


' ------------------ Request.Form ������ �ޱ� ----------------
' dataName : Form ������ �̸�
' defaultValue : �ʱⰪ
' ���� : ���� ���� ������ �ش簪 ����, ��������� defaultValue ����
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


' ------------------ Request.QueryString ������ �ޱ� ----------------
' dataName : QueryString ������ �̸�
' defaultValue : �ʱⰪ
' ���� : ���� ���� ������ �ش簪 ����, ��������� defaultValue ����
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


' ------------------ Request.QueryString ������ �ްų�, ���� ������ ��Ű�� ���� -----------------
' dataName : QueryString ������ �̸� (cookie �̸�)
' defaultValue : �ʱⰪ
' ���� : ���� ���� ������ �ش簪 ����, ��������� ��Ű ���尪, ��Ű���� ������ defaultValue ����
' ���� �������� �Ķ���� init=1 �� ��� ��Ű �ʱ�ȭ
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


' ------------------ �迭 ��� -------------------------
' arr : ����� �迭
' begin_idx : �迭 ���۰�
' end_idx : �迭 ���ᰪ, 0���� ���� ��� UBound ���
' html : ����� HTML
'
' HTML ġȯ��
' {index} : �迭 �ε���
' {data} : �迭 ��
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

