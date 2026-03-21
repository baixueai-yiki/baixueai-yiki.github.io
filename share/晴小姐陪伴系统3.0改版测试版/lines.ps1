=0
Get-Content qing.ps1 | ForEach-Object { ++; if ( -ge 600 -and  -lt 710) { '{0}:{1}' -f ,  } }
