@echo off
rem numbers-converter: saves Excel copies of Numbers files without opening the app's window.
rem Run "numbers-converter --help" to see how, or read "Command line" in the README.
"%~dp0resources\numbers-tool\numbers-tool.exe" %*
exit /b %errorlevel%
