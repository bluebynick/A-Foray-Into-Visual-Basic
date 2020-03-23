Set objShell = WScript.CreateObject("WScript.Shell")
objShell.AppActivate ("taptip.exe") 'select the keyboard
objShell.SendKeys "^+%S" 'send the selected the keybinding (linking to "dual monitor tools")
'objShell.Moveable = false
objShell.SendKeys "%{ESC}" 'return focus to what it was before

'Goal: get on screen keyboard to be on tablet while I work on 'external monitor
'
'logic:
'
'1. Force it to open on primary
'
'Ans: https://superuser.com/questions/738081/how-to-make-applications-open-on-the-correct-monitor-when-using-multiple-monitor
'
'2. Prevent it from moving???
'
'end result:
'https://stackoverflow.com/questions/38815081/how-to-open-tabtip-keyboard-in-a-custom-location
'
'impossible due to app construction
'
'I JUST WASTED 6 HOURS