if (Test-Path -Path "$((Get-Location).Path)\bin\Debug\SysprepPreparator.exe" -PathType Leaf) {
	New-Item -Path ".\out" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
	Copy-Item -Path ".\bin\Debug\Microsoft.Dism.dll" -Destination ".\out\Microsoft.Dism.dll" -Verbose -Force
	Copy-Item -Path ".\bin\Debug\IniFileParser.dll" -Destination ".\out\IniFileParser.dll" -Verbose -Force
	New-Item -Path ".\out\Languages" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
	Copy-Item -Path ".\bin\Debug\Languages\*.ini" -Destination ".\out\Languages" -Verbose -Force
	Copy-Item -Path ".\bin\Debug\SysprepPreparator.exe*" -Destination ".\out" -Verbose -Force
	Compress-Archive -Path ".\out\*" -DestinationPath ".\SysprepPreparator.zip" -Force
}