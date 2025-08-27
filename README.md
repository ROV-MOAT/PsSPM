Powershell SNMP Printer Monitoring and Reporting Script.

![Window](https://github.com/ROV-MOAT/PsSPM/blob/main/PsSPM.png)

C# SNMP Library is used - https://github.com/lextudio/sharpsnmplib

Checks printer status via SNMP and generates CSV, HTML report with toner levels, counters, and information. Sending a report by email.
Interface Mode - Console, FullGui, LightGui.
Example - C:\0.3.5\PsSPM_0.3.5b.ps1 -InterfaceMode Console -ConsoleFile C:\0.3.5\IP\*.txt

You can query any printer, you need to change/add the model and OID in the file "Lib\PsSPM_oid.psd1", and change/add the value "$modelPatterns" in the function "Get-PrinterModelOIDSet".
