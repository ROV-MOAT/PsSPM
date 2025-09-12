Powershell SNMP Printer Monitoring and Reporting Script. <br>
<p align="center"><img src="https://github.com/ROV-MOAT/PsSPM/blob/main/PsSPM.png"/></p>

C# SNMP Library is used - https://github.com/lextudio/sharpsnmplib <br>
<br>
Checks printer status via SNMP and generates CSV, HTML report with toner levels, counters, and information. Sending a report by email.
Interface Mode - Console, FullGui(WPF). <br>

You can query any printer, you need to change/add the model and OID in the file "Lib\PsSPM_oid.psd1", and change/add the value in the table "$modelPatterns".

Example for console: <br>
powershell.exe -Command "& {D:\PsSPM_0.3.5b\PsSPM_0.3.5b.ps1 -InterfaceMode "Console" -ConsoleFile "D:\PsSPM_0.3.5b\IP\*.txt" -MailSend $true -MailUser "" -MailPass ""}" <br>
powershell.exe -File .\PsSPM_0.3.5b.ps1 -InterfaceMode Console -ConsoleFile D:\PsSPM_0.3.5b\IP\*.txt <br>
powershell.exe -File .\PsSPM_0.3.5b.ps1 -InterfaceMode FullGui <br>
powershell.exe -File .\PsSPM_0.3.5b.ps1 -InterfaceMode Console -ConsoleFile D:\PsSPM_0.3.5b\IP\*.txt -MailSend $false -MailUser User -MailPass Pass <br>

Sending a letter is convenient to use in Console mode, there are no mail settings in the GUI. <br>
Be careful not to store your account password in a script.
