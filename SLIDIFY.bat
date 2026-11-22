@echo off
title SLIDIFY
echo Starting SLIDIFY... keep this window open while you use it. Close it to stop SLIDIFY.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0app\start.ps1"
if errorlevel 1 pause
