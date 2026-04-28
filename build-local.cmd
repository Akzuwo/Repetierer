@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0build-local.ps1" %*
