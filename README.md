# pexe
Quick Powershell Executor

![20241104_042645](https://github.com/user-attachments/assets/2cdb87b3-193c-4b76-8a74-0467802026fa)


# Environment
・Windows 10/11

・pwsh 11

・C#5

# Build
```
PS > C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /target:winexe /win32icon:pexe.ico /out:pexe.exe pexe.cs
```

# pexe.ini
`PosLeft` : Set the input box position on the taskbar. Default is 10 (assumes Windows 11).

`OutputColums`: Set the output text box columns. Default is 80.

`Font`: Set the font. Default is MS Gothic.

`WorkingDirectory`: Set the pwsh working directory. Default is the current directory.

`HistPath`: Set the history file path. Default is the current directory.

`ExpandButtonDark`: Set the expand button color type. 0: Light color (for Windows 11), 1: Dark color (for Windows 10).


# Run
Click `pexe.exe`

# Quit
Input the command `term` or close from `Close Window` on the right-click context menu.
