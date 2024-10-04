# Tạo một cửa sổ thông báo hiển thị "Hello, World!"
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
[System.Windows.Forms.MessageBox]::Show("Hello, World!", "Thông báo")
