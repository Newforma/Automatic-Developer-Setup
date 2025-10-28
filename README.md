# Automatic Developer Setup
A little Powershell tool to automate the [NPC developer setup](https://github.com/Newforma/enterprise-technical-documentation/tree/master/Newforma%20Dev%20Environment)
To get started, run this command in Powershell as admin:
```
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Newforma/Automatic-Developer-Setup/main/setup.ps1" -OutFile "$HOME\setup.ps1"; powershell -ExecutionPolicy Bypass -File "$HOME\setup.ps1"
```
