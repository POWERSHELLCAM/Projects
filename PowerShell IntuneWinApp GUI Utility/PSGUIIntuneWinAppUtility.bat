@ECHO OFF
powershell -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -Verb RunAs powershell -ArgumentList '-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File %~dpn0.ps1'"
