@echo off
cd /d "%~dp0"
title Insyn - lokal server
where py >nul 2>nul && (py server.py %*) || (python server.py %*)
pause
