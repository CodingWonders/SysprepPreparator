if (Test-Path -Path "$((Get-Location).Path)\bin\Debug\SysprepPreparator.exe" -PathType Leaf) {
	Write-Host "Packing release..."
	New-Item -Path "$((Get-Location).Path)\out" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
	Copy-Item -Path "$((Get-Location).Path)\bin\Debug\Microsoft.Dism.dll" -Destination "$((Get-Location).Path)\out\Microsoft.Dism.dll" -Verbose -Force
	Copy-Item -Path "$((Get-Location).Path)\bin\Debug\IniFileParser.dll" -Destination "$((Get-Location).Path)\out\IniFileParser.dll" -Verbose -Force
	New-Item -Path "$((Get-Location).Path)\out\Languages" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
	Copy-Item -Path "$((Get-Location).Path)\bin\Debug\Languages\*.ini" -Destination "$((Get-Location).Path)\out\Languages" -Verbose -Force
	Copy-Item -Path "$((Get-Location).Path)\bin\Debug\SysprepPreparator.exe*" -Destination "$((Get-Location).Path)\out" -Verbose -Force
	Compress-Archive -Path "$((Get-Location).Path)\out\*" -DestinationPath "$((Get-Location).Path)\SysprepPreparator.zip" -Force
} else {
	Write-Host "Output directory not found."
}