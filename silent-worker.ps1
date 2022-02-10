New-Item -Path "HKCU:\Software\Policies\Microsoft" -Name office
New-Item -Path "HKCU:\Software\Policies\Microsoft\office" -Name common
New-Item -Path "HKCU:\Software\Policies\Microsoft\office\common" -Name clienttelemetry
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\office\common\clienttelemetry" -Name "DisableTelemetry" -Value "1" -PropertyType "DWord"

New-ItemProperty -Path "HKCU:\SOFTWARE\Policies\Microsoft\office\common" -Name "sendcustomerdata" -Value "0" -PropertyType "DWord"

Add-Content -Path $env:windir\System32\drivers\etc\hosts -Value "`n0.0.0.0`thubblecontent.osi.office.net/contentsvc/api/telemetry" -Force