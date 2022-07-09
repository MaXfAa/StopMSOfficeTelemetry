New-Item -Path "HKCU:\Software\Policies\Microsoft" -Name "Office"
New-Item -Path "HKCU:\Software\Policies\Microsoft\Office" -Name "common"
New-Item -Path "HKCU:\Software\Policies\Microsoft\Office\common" -Name "clienttelemetry"
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Office\common\clienttelemetry" -Name "DisableTelemetry" -Value "1" -PropertyType "DWord"
New-Item -Path "HKCU:\Software\Policies\Microsoft\Office" -Name 16.0
New-Item -Path "HKCU:\Software\Policies\Microsoft\Office\16.0" -Name "common"
New-Item -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\common" -Name "clienttelemetry"
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\common\clienttelemetry" -Name "DisableTelemetry" -Value "1" -PropertyType "DWord"

New-Item -Path "HKCU:\Software\Policies\Microsoft\Office\16.0" -Name "OSM"
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\OSM" -Name "enablelogging" -Value "0" -PropertyType "DWord"
New-ItemProperty -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\common" -Name "sendcustomerdata" -Value "0" -PropertyType "DWord"
New-ItemProperty -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\common" -Name "sendcustomerdata" -Value "0" -PropertyType "DWord"
New-Item -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\OSM" -Name "preventedapplications"
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications" -Name "accesssolution" -Value "1" -PropertyType "DWord"
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications" -Name "olksolution" -Value "1" -PropertyType "DWord"
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications" -Name "onenotesolution" -Value "1" -PropertyType "DWord"
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications" -Name "pptsolution" -Value "1" -PropertyType "DWord"
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications" -Name "projectsolution" -Value "1" -PropertyType "DWord"
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications" -Name "publishersolution" -Value "1" -PropertyType "DWord"
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications" -Name "visiosolution" -Value "1" -PropertyType "DWord"
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications" -Name "wdsolution" -Value "1" -PropertyType "DWord"
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications" -Name "xlsolution" -Value "1" -PropertyType "DWord"

New-Item -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\OSM" -Name "preventedsolutiontypes"
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes" -Name "agave" -Value "1" -PropertyType "DWord"
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes" -Name "appaddins" -Value "1" -PropertyType "DWord"
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes" -Name "comaddins" -Value "1" -PropertyType "DWord"
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes" -Name "documentfiles" -Value "1" -PropertyType "DWord"
New-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes" -Name "templatefiles" -Value "1" -PropertyType "DWord"

Add-Content -Path $env:windir\System32\drivers\etc\hosts -Value "`n0.0.0.0`thubblecontent.osi.Office.net" -Force
Add-Content -Path $env:windir\System32\drivers\etc\hosts -Value "`n0.0.0.0`petrol.Office.microsoft.com" -Force
Add-Content -Path $env:windir\System32\drivers\etc\hosts -Value "`n0.0.0.0`upload.fp.measure.Office.com" -Force
Add-Content -Path $env:windir\System32\drivers\etc\hosts -Value "`n0.0.0.0`measure.Office.net" -Force