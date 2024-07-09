# Define the URL and the output file path
New-Item -Path "C:\\Program Files (x86)\\MedTrainer" -ItemType Directory

# Log path and name
$logTime = get-date -format "yyy-MM-ddThh.mm.ss"
$logFile = "C:\\Program Files (x86)\\MedTrainer\\Logs\\"
Start-Transcript -IncludeInvocationHeader -Path $logFile$logTime".txt"

$url = "https://raw.githubusercontent.com/CasaNicos/FIleHosting/main/medtrainer_mt_compliance_logo%20(4).ico"
$outputFile = "C:\\Program Files (x86)\\MedTrainer\\file.ico"

# Download the file
Invoke-WebRequest -Uri $url -OutFile $outputFile

$shortcutName = "MedTrainer"
$shortcutURL = "https://lms.medtrainer.com/sso/"

# Function to find the user profile path
function Get-UserProfilePath {
    $users = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.Special -eq $false -and $_.LocalPath -match 'C:\\Users' }
    foreach ($user in $users) {
        $oneDrivePath = "$($user.LocalPath)\\OneDrive - House Works\\Desktop"
        if (Test-Path $oneDrivePath) {
            return $oneDrivePath
        }
    }
    return $null
}

$userProfilePath = Get-UserProfilePath
if ($userProfilePath -ne $null) {
    $shell = New-Object -ComObject WScript.Shell
    Write-Host "Shell Created $shell"
    $shortcut = $shell.CreateShortcut("$userProfilePath\\$shortcutName.lnk")
    Write-Host "Shortcut Path: $userProfilePath\\$shortcutName.lnk"
    $shortcut.TargetPath = $shortcutURL
    Write-Host "Shortcut Link: $shortcutURL"
    $shortcut.IconLocation = "C:\\Program Files (x86)\\MedTrainer\\file.ico"
    Write-Host "Icon Location: $shortcut.IconLocation"
    $shortcut.Save()
} else {
    Write-Host "User profile path not found"
}
