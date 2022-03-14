function iCopy-Files {
  param (
      [Parameter(Mandatory=$true)]
      $year
  )
  $targetPath = ".\backup\$year"
  $path = ".\filelist.xlsx"

  Import-Module ".\psexcel.1.0.2\PSExcel.psm1"
  $file = New-Object System.Collections.ArrayList
  foreach ($d in (Import-XLSX -Path $path -RowStart 1 -WarningAction silentlyContinue )) {
    if($d.path -ne $null){
      if(([DateTime]::ParseExact($str, "dd.MM.yyyy HH:mm:ss", $null)).Year -eq $year){
      Write-Host "Year: $year"
      $clean = $d.path.replace($d.share, "")
      $arr = $clean.split("\")
      $path = $clean.replace("\$($arr[-1])", "")
      $filename = $arr[-1]
      if(-Not (test-path "$targetPath\$path")) {
        md "$targetPath\$path" | Out-Null
      }
      if(-Not (Test-Path "$($targetPath)\$($path)\$filename")) {
        Try{
          $copied = Copy-Item $d.path -Destination "$($targetPath)\$($path)\$filename" -ErrorAction Stop
          "$([DateTime]::Now)" + "`t$copied`t >> $($targetPath)\$($path)\$filename" | Out-File ".\logs\copied-$year.log" -Append -Encoding utf8
          #Write-Host "(i) $($d.path)"
        }Catch{
          $_.Exception.Message | Out-File ".\logs\copied-$year-error.log" -Append -Encoding utf8
        }
      } else {
        #Write-Host "$($targetPath)\$($path)\$filename - exist"
      }
      }
    }
 }
}