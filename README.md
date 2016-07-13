# CheckExport
A PowerShell script to check a directory for PDF files and if there are files, it will restart an export service.

The best way to use this is to create a batch script that calls the CheckReferralExport.ps1 file and then use Task Scheduler to trigger that batch script at whatever intervals desired.
```
@echo off
powershell C:\Users\maravedi\CheckReferralExport.ps1
```
