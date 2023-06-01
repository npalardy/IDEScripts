/// Copyright 2022 Norman Palardy 
/// npalardy@great-white-software.com
/// You may use this in any way you want but must provide attribution if this is included in other works

// KeyboardShortcut=CMD-OPT-SHIFT-T

Dim stripped As String
stripped = DoShellCommand("echo ""#Pragma warning \""NORM TODO\"" ""  | /usr/bin/tr -C -d ""[:print:]"" | /usr/bin/pbcopy -Prefer txt")

DoCommand("Paste")

