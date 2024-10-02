Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Tạo đối tượng Bitmap với kích thước toàn bộ không gian ảo của các màn hình
$virtualScreenWidth = [System.Windows.Forms.SystemInformation]::VirtualScreen.Width
$virtualScreenHeight = [System.Windows.Forms.SystemInformation]::VirtualScreen.Height
$virtualScreenLeft = [System.Windows.Forms.SystemInformation]::VirtualScreen.Left
$virtualScreenTop = [System.Windows.Forms.SystemInformation]::VirtualScreen.Top

$screenshot = New-Object System.Drawing.Bitmap $virtualScreenWidth, $virtualScreenHeight
$graphics = [System.Drawing.Graphics]::FromImage($screenshot)

# Chụp toàn bộ không gian màn hình
$graphics.CopyFromScreen($virtualScreenLeft, $virtualScreenTop, 0, 0, $screenshot.Size)

# Lưu ảnh vào file, đảm bảo thư mục tồn tại
$folderPath = Join-Path $env:TEMP "ProgramPicture"
if (-not (Test-Path $folderPath)) {
    New-Item -Path $folderPath -ItemType Directory
}
$filePath = Join-Path $folderPath "screenshot.png"
$screenshot.Save($filePath, [System.Drawing.Imaging.ImageFormat]::Png)

# Giải phóng tài nguyên
$graphics.Dispose()
$screenshot.Dispose()

