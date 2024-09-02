$path_arr = (New-Object -ComObject 'Shell.Application').Windows() | ForEach-Object { 
  $localPath = $_.Document.Folder.Self.Path 
  "C:\WINDOWS\explorer.exe `"$localPath`""
}

taskkill /F /IM explorer.exe

for ($i = 0; $i -lt $path_arr.Length; $i++) {
    Invoke-Expression $path_arr[$i]
    Start-Sleep 0.5
}

explorer.exe