/// Copyright 2022 Norman Palardy 
/// npalardy@great-white-software.com
/// You may use this in any way you want but must provide attribution if this is included in other works

// KeyboardShortcut=CMD-OPT-SHIFT-D


Const IDESCript = True
Const debugThis = True
Const CommentMarker = "///"

DoInsertDoc()

Sub DoInsertDoc()

	Dim to_insert() As String

	to_insert.append "#if DebugBuild = false"
	to_insert.append "  #Pragma BackgroundTasks False"
	to_insert.append "  #Pragma BoundsChecking False"
	to_insert.append "  #Pragma BreakOnExceptions False"
	to_insert.append "  #Pragma NilObjectChecking False"
	to_insert.append "  #Pragma StackOverflowChecking False"
	to_insert.append "#EndIf"
	
	// Print(to_insert)
	selStart = 0
	SelLength = 0
	seltext = Join(to_insert, EndOfLine) + EndOfLine + EndOfLine

End Sub
