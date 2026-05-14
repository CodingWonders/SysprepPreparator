param (
    [Parameter(Mandatory = $true, Position = 0)] [string] $appxFullNames
)

$appxFullNamesArray = $appxFullNames.Split(";")
Get-AppxPackage -AllUsers | Where-Object { $appxFullNamesArray.Contains($_.PackageFullName) } | Remove-AppxPackage -AllUsers