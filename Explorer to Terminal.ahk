#SingleInstance Force
#NoTrayIcon ; Remove this to show the tray icon

OpenTerminal(path, admin, linux)
{
    if (InStr(path, '::') == 1) { ; Some paths are invalid
        path := false
    }
    ; The following line of code requires a profile to be set up in Windows Terminal
    ; the profile name has to be 'Powershell (Elevated)'
    ; Windows 11 22H2 has broken this functionality
    ; cmd := 'wt.exe ' (admin ? 'new-tab -p "PowerShell (Elevated)" ' : (linux ? 'new-tab -p "Ubuntu" ' : '')) (path ? '-d "' path '"' : '')
    ; use this instead for now:
    cmd := (admin ? '*runas ' : '') 'wt.exe ' (linux ? 'new-tab -p "Ubuntu" ' : '') (path ? '-d "' path '"' : '') ; Replace 'Ubuntu' with your preferred profile
    OutputDebug(cmd)
    Run(cmd, , , &pid)
    if (WinWait('ahk_exe WindowsTerminal.exe', , 2)) {
        WinActivate('ahk_exe WindowsTerminal.exe')
    }
}

#t::  ; Win+T       | Opens terminal
#+t:: ; Win+Shift+T | Opens terminal as admin
#^t:: ; Win+Ctrl+T  | Opens terminal in Linux
{
    admin := RegExMatch(a_thisHotkey, '(\+)')
    linux := RegExMatch(a_thisHotkey, '(\^)')
    ; Open new tab if terminal is active, doesn't work, probably because of Win key
    ; Update (Win 11 22H2): Not needed anymore
    ;id := WinActive('ahk_exe WindowsTerminal.exe')
    ;if (id) {
    ;    if (!linux)
    ;        ControlSend('^+1', , 'ahk_id ' id)
    ;    else
    ;        ControlSend('^+4', , 'ahk_id ' id)
    ;    return
    ;}
    explorerHwnd := WinActive('ahk_class CabinetWClass')
    if (explorerHwnd) {
        for window in ComObject('Shell.Application').Windows {
            if (window.hwnd == explorerHwnd) {
                OpenTerminal(window.Document.Folder.Self.Path, admin, linux)
            }
        }
    } else {
        OpenTerminal(false, admin, linux)
    }
}