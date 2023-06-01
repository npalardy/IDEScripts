/// Copyright 2022 Norman Palardy 
/// npalardy@great-white-software.com
/// You may use this in any way you want but must provide attribution if this is included in other works

// KeyboardShortcut=CMD-OPT-SHIFT-D


Const IDESCript = True
Const debugThis = True
Const CommentMarker = "///"

DoInsertDoc()

Sub DoInsertDoc()

	Dim content As String
	Dim scope As String
	Dim subFunc As String
	Dim methodName As String
	Dim params As String
	Dim returntype As String

	doCommand("Copy")

	// Print( clipboard )

	if trim(clipboard) = "" Then
		return
	end if		
	
	Dim replaced_text As String = myreplacelineendings(clipboard, EndOfLine)
	Dim lines() As String = Split(replaced_text, EndOfLine)

	For Each line As String In lines
		If Trim(line) <> "" Then
			content = line
			Exit For
		End If
	Next

	// Print( content )

	If LanguageUtils.CrackMethodDeclaration(content, scope, subFunc, methodName, params, returntype) Then
		// Print("Its a method")

	ElseIf LanguageUtils.CrackEventDeclaration(content, scope, subFunc, methodName, params, returntype) Then
		// Print("Its an event")
	Else
		Return
	End If

	// Print (" params = " + params )

	Dim parms() As LanguageUtils.LocalVariable
	parms = LanguageUtils.CrackParams(params)

	Dim to_insert() As String

	to_insert.append commentMarker + " @Name : " + methodName 
	to_insert.append commentMarker 
	to_insert.append commentMarker + " @Purpose : <statement of purpose>"  
	to_insert.append commentMarker 

	For Each l As LanguageUtils.LocalVariable In parms

		Dim array_of As String = " "
		If l.isArray Then
			array_of = " array_of "
		End If

		to_insert.append commentMarker + " @Param : " + l.name + array_of + l.type + " " 
		to_insert.append commentMarker + "          <statement of purpose>"
		to_insert.append commentMarker 

	Next l

	if returnType <> "" Then
		to_insert.append commentMarker + " @Returns : " + returnType
	end if
	
	// Print(to_insert)
	selStart = 0
	SelLength = 0
	seltext = Join(to_insert, EndOfLine) + EndOfLine + EndOfLine

End Sub

Module LanguageUtils

	Protected Function CrackEventDeclaration(content As String, ByRef scope As String, ByRef subFunc As String, ByRef methodName As String, ByRef params As String, ByRef returntype As String) As Boolean

		// when we copy from the ide it coms in as 
		// Function KeyDown(Key As String) Handles KeyDown as Boolean
		// BUT when we read from a text project its like
		//   #tag Event
		//   Function KeyDown(Key As String) As Boolean
		//  
		//   End Function
		//   #tag EndEvent

		scope = ""
		subFunc = ""
		methodName = ""
		params = ""
		returntype = ""

		// strip content to make sure we basically got ONE line
		Dim lines() As String = Split(MyReplaceLineEndings(Trim(content), EndOfLine), EndOfLine)

		If lines.ubound < 0 Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure
		End If

		Dim tokens() As String = LanguageUtils.TokenizeLine(lines(0), False)

		scope = kScopePrivate

		// SUB FUNC ?
		If tokens(0) <> "SUB" And tokens(0) <> "FUNCTION" Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure
		End If

		subFunc = tokens(0)
		tokens.remove 0

		// empty string ?
		If tokens.ubound < 0 Then 
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure
		End If

		// out of tokens ?
		If tokens.Ubound < 0 Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure
		End If

		If IsIdentifier( tokens(0) ) = False Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure // the name is not an ident ?????
		End If

		methodName = tokens(0)
		tokens.remove 0

		// out of tokens ?
		If tokens.Ubound < 0 Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure // got maybe "scope name" but nothing else ?????????
		End If

		If tokens(0) <> "(" Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure 
		End If
		tokens.remove 0

		// out of tokens ?
		If tokens.Ubound < 0 Then    
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure // got maybe "scope name ( " but nothing else ?????????
		End If

		// we SHOULD be able to read any old token until we get to the closing matching paren
		// we're not really tryng to validate the code is going to compile
		// we just want to make sure that whatever is in ( ) is all picked up as "params"
		// and that the ( and ) are "balanced"
		Dim parenthesisCount As Integer = 1
		While tokens.ubound >= 0 And parenthesisCount > 0
  
		  If tokens(0) = "(" Then
			parenthesisCount = parenthesisCount + 1
			params = params + tokens(0)
		  ElseIf tokens(0) = ")" Then
			parenthesisCount = parenthesisCount - 1
			If parenthesisCount > 0 Then
			  params = params + tokens(0)
			End If
		  Else
	
			If tokens(0) = "." Then // a . we dont put spaces before
			ElseIf Right(params,1) = "." Then // and we dont put one after a . either
			Else
			  If params <> "" Then
				params = params + " "
			  End If
			End If
	
			params = params + tokens(0)
		  End If
  
		  tokens.remove 0
  
		Wend

		// IF we're a FUNCTION then we should have "as typename"
		// but as a SUB we exit here IF things are OK
		If parenthesisCount <> 0 Then 
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure  // failed parse for a SUB - unclosed (
		End If

		// out of tokens ?
		If tokens.Ubound < 0 Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseSuccess 
		End If

		// handles

		If tokens(0) <> "handles" Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure 
		End If
		tokens.remove 0
		// out of tokens ?
		If tokens.Ubound < 0 Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure // got maybe "scope name ( ) handles .... and no more ?
		End If

		// event
		If IsIdentifier(tokens(0)) = False Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure 
		End If
		tokens.remove 0

		If subFunc = "SUB" Then
		  If tokens.Count > 0 Then
			Return parseFailure  // there's something AFTER the end of the "handles event" 
		  End If
		  Return parseSuccess
		End If

		// we're a func !
		// should have an AS type
		If tokens(0) <> "as" Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure 
		End If
		tokens.remove 0
		// out of tokens ?
		If tokens.Ubound < 0 Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure // got maybe "scope name ( ) as .... and no more ?
		End If

		// the type name now .. accept dotted names ?????
		// this basically needs to be ident, or ident(), or ident(,,,)
		// or  ident . ident . ident, or  ident . ident . ident () 
		// or  ident . ident . ident(,,,,)
		Const kExpectIdent = 1
		Const kExpectDot = 2
		Const kExpectDotOrBracket = 3
		Const kExpectCloseBracket = 4
		Dim expect As Integer = kExpectIdent
		Dim returnsArrayType As Boolean

		While tokens.ubound >= 0 And tokens(0) <> "="
  
		  Select Case expect
	
		  Case kExpectIdent
			If IsIdentifier(tokens(0)) Or IsBaseType(tokens(0)) Or (tokens(0) = "Object") Then
			  returntype = returntype + tokens(0)
			  tokens.remove 0
			Else
			  #If debugThis
				Break
			  #EndIf
			  Return parseFailure // type name is not ident or ident.ident.ident
			End If
			expect = kExpectDotOrBracket
	
		  Case kExpectDotOrBracket
			If tokens(0) = "." Then
			  returntype = returntype + tokens(0)
			  tokens.remove 0
			  expect = kExpectIdent
			ElseIf tokens(0) = "(" Then
			  tokens.remove 0
			  returnsArrayType = True
			  expect = kExpectCloseBracket
			End If
	
		  Case kExpectCloseBracket
			If tokens(0) = ")" Then
	  
			  tokens.remove 0
			  Exit While
			Else
			  // eat tokens until we get to the close )
			  tokens.remove 0
			End If
	
		  End Select
  
		Wend

		If returnsArrayType Then
		  returntype = returntype + "()"
		End If

		// out of tokens ?
		// at this point we have a FULL Declaration so we can start saying "yup this is OK"
		If tokens.Ubound < 0 Then
		  Return parseSuccess // YAY !!!!!!!!! got "scope name( params ) handles eventname as type"
		End If

		// theres more ?????????
		#If debugThis
		  Break
		#EndIf

		Return parseFailure



	End Function

	Protected Function CrackMethodDeclaration(content As String, ByRef scope As String, ByRef subFunc As String, ByRef methodName As String, ByRef params As String, ByRef returntype As String) As Boolean
		// ident ... what can I say - any identified just not a reserved word
		//
		// paramlist we dont really care as we just match up the outmost ( and )
		// they just have to balance (if we crack parms then we would care)
		scope = ""
		subFunc = ""
		methodName = ""
		params = ""
		returntype = ""

		// strip content to make sure we basically got ONE line
		Dim trimmed_content As String = Trim(content)
		Dim cleaned_content As String = MyReplaceLineEndings(trimmed_content, EndOfLine)
		Dim lines() As String = Split(cleaned_content, EndOfLine)

		If lines.ubound < 0 Then
		  #If debugThis
			#If IDEScript
			  Print("lines < 0")
			#Else
			  Break
			#EndIf
		  #EndIf
		  Return parseFailure
		End If

		Dim tokens() As String = LanguageUtils.TokenizeLine(lines(0), False)

		// empty string ?
		If tokens.ubound < 0 Then 
		  #If debugThis
			#If IDEScript
			  Print("tokens.ubound < 0")
			#Else
			  Break
			#EndIf
		  #EndIf
		  Return parseFailure
		End If

		// default scope set up
		//   PRIVATE
		//   PUBLIC
		//   GLOBAL
		scope = kScopePublic
		If tokens(0) = kScopeGlobal _
		  Or tokens(0) = kScopeProtected _ 
		  Or tokens(0) = kScopePrivate _
		  Or tokens(0) = kScopePublic Then
		  scope = tokens(0) // got a valid scope 
		  tokens.remove 0
		End If

		// out of tokens ?
		If tokens.Ubound < 0 Then
		  #If debugThis
			#If IDEScript
			  Print("tokens.ubound < 0")
			#Else
			  Break
			#EndIf
		  #EndIf
		  Return parseFailure
		End If

		// SHARED ?
		If tokens(0) = "SHARED" Then
			#If debugThis
				Break
			#EndIf
			subFunc = tokens(0)
			tokens.remove 0
		End If

		// out of tokens ?
		If tokens.Ubound < 0 Then
		  #If debugThis
			#If IDEScript
			  Print("tokens.ubound < 0")
			#Else
			  Break
			#EndIf
		  #EndIf
		  Return parseFailure
		End If

		// SUB FUNC ?
		If tokens(0) <> "SUB" And tokens(0) <> "FUNCTION" Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure
		End If
		subFunc = tokens(0)
		tokens.remove 0

		// out of tokens ?
		If tokens.Ubound < 0 Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure
		End If

		If IsIdentifier( tokens(0) ) = False Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure // the name is not an ident ?????
		End If
		methodName = tokens(0)
		tokens.remove 0

		// out of tokens ?
		If tokens.Ubound < 0 Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure // got maybe "scope name" but nothing else ?????????
		End If

		If tokens(0) <> "(" Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure 
		End If
		tokens.remove 0

		// out of tokens ?
		If tokens.Ubound < 0 Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure // got maybe "scope name ( " but nothing else ?????????
		End If

		// we SHOULD be able to read any old token until we get to the closing matching paren
		// we're not really tryng to validate the code is going to compile
		// we just want to make sure that whatever is in ( ) is all picked up as "params"
		// and that the ( and ) are "balanced"
		Dim parenthesisCount As Integer = 1
		While tokens.ubound >= 0 And parenthesisCount > 0
  
		  If tokens(0) = "(" Then
			parenthesisCount = parenthesisCount + 1
			params = params + tokens(0)
		  ElseIf tokens(0) = ")" Then
			parenthesisCount = parenthesisCount - 1
			If parenthesisCount > 0 Then
			  params = params + tokens(0)
			End If
		  Else
	
			If tokens(0) = "." Then // a . we dont put spaces before
			ElseIf Right(params,1) = "." Then // and we dont put one after a . either
			Else
			  If params <> "" Then
				params = params + " "
			  End If
			End If
	
			params = params + tokens(0)
		  End If
		  tokens.remove 0
  
		Wend

		// IF we're a FUNCTION then we should have "as typename"
		// but as a SUB we exit here IF things are OK
		If subFunc = "SUB" Then
		  If tokens.ubound > -1 Then
			#If debugThis
			  Break
			#EndIf
			Return parseFailure  // failed parse for a SUB 
		  ElseIf parenthesisCount <> 0 Then 
			#If debugThis
			  Break
			#EndIf
			Return parseFailure  // failed parse for a SUB - unclosed (
		  Else
			Return parseSuccess // SUCCESSFUL FOR A SUB !!!!!!!!
		  End If
		End If

		// out of tokens ?
		If tokens.Ubound < 0 Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure // got maybe "scope name () " but nothing else ?????????
		End If

		// we're a func !
		// should have an AS type
		If tokens(0) <> "as" Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure 
		End If
		tokens.remove 0
		// out of tokens ?
		If tokens.Ubound < 0 Then
		  #If debugThis
			Break
		  #EndIf
		  Return parseFailure // got maybe "scope name ( ) as .... and no more ?
		End If

		// the type name now .. accept dotted names ?????
		// this basically needs to be ident, or ident(), or ident(,,,)
		// or  ident . ident . ident, or  ident . ident . ident () 
		// or  ident . ident . ident(,,,,)
		Const kExpectIdent = 1
		Const kExpectDot = 2
		Const kExpectDotOrBracket = 3
		Const kExpectCloseBracket = 4
		Dim expect As Integer = kExpectIdent
		Dim returnsArrayType As Boolean

		While tokens.ubound >= 0 And tokens(0) <> "="
  
		  Select Case expect
	
		  Case kExpectIdent
			If IsIdentifier(tokens(0)) Or IsBaseType(tokens(0)) Or (tokens(0) = "Object") Then
			  returntype = returntype + tokens(0)
			  tokens.remove 0
			Else
			  #If debugThis
				Break
			  #EndIf
			  Return parseFailure // type name is not ident or ident.ident.ident
			End If
			expect = kExpectDotOrBracket
	
		  Case kExpectDotOrBracket
			If tokens(0) = "." Then
			  returntype = returntype + tokens(0)
			  tokens.remove 0
			  expect = kExpectIdent
			ElseIf tokens(0) = "(" Then
			  tokens.remove 0
			  returnsArrayType = True
			  expect = kExpectCloseBracket
			End If
	
		  Case kExpectCloseBracket
			If tokens(0) = ")" Then
	  
			  tokens.remove 0
			  Exit While
			Else
			  // eat toekns until we get to the close )
			  tokens.remove 0
	  
			End If
	
		  End Select
  
		Wend
		If returnsArrayType Then
		  returntype = returntype + "()"
		End If

		// out of tokens ?
		// at this point we have a FULL Declaration so we can start saying "yup this is OK"
		If tokens.Ubound < 0 Then
		  Return parseSuccess // YAY !!!!!!!!! got "scope name( params ) as type"
		End If

		// theres more ?????????
		#If debugThis
		  Break
		#EndIf
		Return parseFailure

	End Function

	Function CrackParams(paramString As String) As LanguageUtils.LocalVariable()

		Dim result() As LanguageUtils.LocalVariable
		// we fake this into being a DIM line
		If paramString.Left(4) <> "DIM " Then
		  FindVarDeclarations("DIM " + paramString, 1, result)
		Else
		  FindVarDeclarations(paramString, 1, result)
		End If

		Return result
	End Function

	Function FindAnyInStr(startPos As Integer, source As String, findSet As String) As Integer
		// This is like InStr, except that instead of searching for just a single
		// character, we search for any character in a set.

		Dim i, maxi As Integer
		maxi = Len( source )
		Dim c As String
		For i = startPos To maxi
		  c = Mid( source, i, 1 )
		  If InStr( findSet, c ) > 0 Then
			// found one!
			Return i
		  End If
		Next

		Return 0

	End Function

	Sub FindVarDeclarations(sourceLine As String, lineNum As Integer, outVars() As LocalVariable)
		// Find variable declarations in the given line.
		// Append them to outVars.
		Redim outvars(-1)

		Dim tokens() As String = TokenizeLine(sourceLine, False)

		If tokens.ubound < 0 Then
		  Return
		End If

		Dim isConstDecl As Boolean

		// attributes ?
		// scope ?
		// func | sub ?
		If tokens(0) = "Public" Or tokens(0) = "Protected" Or tokens(0) = "Private" Or tokens(0) = "Function" Or tokens(0) = "Sub" Or tokens(0) = "Attributes" Then
  
		  // looks like a Sub/Function line; skip to the parameters...
		  While tokens.count > 0 And tokens(0) <> "Function" And tokens(0) <> "Sub" 
			tokens.remove 0
		  Wend
		  If tokens.count <= 0 Then
			Return
		  End If
		  While tokens.count > 0 And tokens(0) <> "(" 
			tokens.remove 0
		  Wend
		  tokens.remove 0 
		  If tokens.count <= 0 Then
			Return
		  End If
  
		ElseIf tokens(0) = "Const" Then
		  // We can just grab the next token as the local variable. 
		  isConstDecl = True
  
		  tokens.remove 0 
  
		ElseIf tokens(0) = "Dim" Then
  
		  tokens.remove 0 
  
		ElseIf tokens(0) = "Static" Then
  
		  tokens.remove 0 
  
		ElseIf tokens(0) = "Var" Then
  
		  tokens.remove 0 
  
		Else
  
		  Break
		  Return
  
		End If


		Const kIdentExpected = 1
		Const kAsExpected = 2
		Const kTypeExpected = 3
		Const kInValueExpr = 4
		Const kCommaExpected = 5
		Dim mode As Integer = kIdentExpected
		Dim openBrackets As Integer

		While tokens.count > 0
  
		  Dim token As String = tokens(0)
  
		  Select Case mode
	
		  Case kIdentExpected
	
			If token = "ByRef" Or token = "ByVal" Or token = "Assigns" Or token = "Extends" Or token = "Optional" Or token = "ParamArray" Then
			  // skip param modifier
			Else
			  outVars.Append New LocalVariable( token, "", lineNum )
	  
			  mode = kAsExpected
			End If
	
		  Case kAsExpected 
	
			If token = "(" Then
			  // skip everything until we get to a closing ")"
			  outVars(UBound(outVars)).isArray = True // (we'll prepend the actual type name below)
			  openBrackets = openBrackets + 1
	  
			ElseIf token = "As" Then
			  mode = kTypeExpected
	  
			ElseIf token = "New" Then
			  // skip "New"; it's just inserted before the type, which we should already expect
	  
			ElseIf token = ")" Then
			  openBrackets = openBrackets - 1
	  
			  If openBrackets = 0 Then
				mode = kAsExpected
			  End If
	  
			ElseIf token = "," Then
	  
			  If openBrackets = 0 Then
				mode = kIdentExpected
			  End If
	  
			ElseIf isConstDecl And token = "=" Then
	  
			  mode = kInValueExpr
	  
			End If
	
		  Case kTypeExpected
			If token = "New" Then
			  // skip the NEW whatever
			Else
			  // ok if the next token is a "." then we should accumylate all the "type as "name.name.name."
			  Dim actualType As String
	  
			  Do
		
				actualType = actualType + tokens(0)
				tokens.remove 0 
		
			  Loop Until tokens.count <= 0 Or (tokens(0) <> "."  And IsIdentifier(tokens(0)) = False)
	  
			  For i As Integer = outvars.Ubound DownTo 0
		
				If outVars(i).type = "" Then
				  outVars(i).type = actualType + If(outVars(i).IsArray,"()","")
				Else
				  Exit For
				End If
		
			  Next
	  
			  mode = kCommaExpected
	  
			  Continue // skip the pos + 1 below
			End If
	
		  Case kCommaExpected
	
			If token = "," Then
			  mode = kIdentExpected
			ElseIf token = "=" Then
			  mode = kInValueExpr
			End If
	
		  Case kInValueExpr
			Dim default As String 
	
			Do 
	  
			  default =  default + tokens(0)
			  tokens.remove 0
	  
			Loop Until tokens.count <= 0 Or tokens(0) = "," Or tokens(0) = ")"
	
			outVars(outVars.Ubound).default_value_str = default
	
			If tokens.count > 0 Then
			  If tokens(0) = "," Then
				tokens.remove 0
			  ElseIf tokens(0) = ")" Then
				tokens.remove 0
			  End If
			End If
			mode = kIdentExpected
	
			Continue // skip the pos + 1 below
	
		  End Select
  
		  tokens.remove 0
		Wend


		If isConstDecl Then
  
		  // check that the consts declared have types and if not then try the const expr to see what type they might be 
		  For i As Integer = 0 To outVars.Ubound
	
			If outVars(i).type = "" Then
	  
			  Select Case True
			  Case IsInteger(outVars(i).default_value_str)
				outVars(i).type = "Integer"
			  Case IsColor(outVars(i).default_value_str)
				outVars(i).type = "Color"
			  Case IsBoolean(outVars(i).default_value_str)
				outVars(i).type = "Boolean"
			  Case IsRealNumber(outVars(i).default_value_str)
				outVars(i).type = "Double"
			  Case outVars(i).default_value_str.Left(1)="""" And outVars(i).default_value_str.Right(1)="""" 
				outVars(i).type = "String"
			  Else
				Break
			  End Select
	  
			End If
	
		  Next
  
		End If

	End Sub

	Function IsBaseType(token As String) As Boolean
		// Return whether the given token is one of our base types.
		Return token = "Auto" Or _
		token = "Boolean" Or _
		token = "CFStringRef" Or token = "CGFloat" Or token = "Color" Or token = "CString" Or token = "Currency" Or _
		token = "Delegate" Or token = "Double" Or _
		token = "Int16" Or token = "Int32" Or token = "Int64" Or token = "Int8" Or token = "Integer" Or _
		token = "Object" Or token = "OSType" Or _
		token = "PString" Or token = "Ptr" Or _
		token = "Short" Or token = "Single" Or token = "String" Or _
		token = "Text" Or _
		token = "UInt16" Or token = "UInt32" Or token = "UInt64" Or token = "UInt8" Or token = "UInteger" Or _
		token = "Variant" Or _
		token = "WindowPtr" Or token = "WString" 

	End Function

	Function IsBoolean(token As String) As Boolean
		// Check to see whether the token is a boolean literal

		Return token = "true" Or token = "false"

	End Function

	Function IsColor(token As String) As Boolean
		// Check to see whether the token is a color literal

		// oddly all it needs its "&c" as the beginning and maybne nothing else :P
		// dont believe me just uncomment this
		' Dim c As Color 
		' c = &c

		Return token.Left(2) = "&c" 

	End Function

	Function IsDigit(c As String) As Boolean
		// Return whether the first character of c is a digit (0 - 9).

		Dim ascC As Integer = Asc(c)

		Return ascC >= &h30 And ascC <= &h39

	End Function

	Function IsIdentifier(token As String) As Boolean
		// Return whether the given identifier might be an identifier
		// rather than an operator or keyword.

		// NOTE: "me" and "self" return false here, because they're keywords,
		// even though they're also identifiers -- we may want to reconsider
		// this at some point.

		Dim c As String

		// first letter cannot be a number
		// or other punctuation
		c = Left( token, 1 )
		If InStr("0123456789~`!@#$%^&*()-+={}[]""':;<,>.?/|\", c) > 0 Then
			Return False
		End If

		// token must not be a keyword
		If KeywordDict.HasKey( token ) Then 
			Return False
		End If

		// otherwise, it's probably an identifier
		Return True

	End Function

	Protected Function IsInteger(token As String) As Boolean
		// Check to see whether the token is numeric first
		If Not IsNumeric( token ) Then 
			Return False
		End If

		// Now check to see if it's a real number
		If IsRealNumber( token ) Then 
			Return False
		End If

		// Finally, we know it's numeric, but not a real number,
		// so we're done
		Return True
	End Function

	Protected Function IsNumeric(token As String) As Boolean
		Dim chars() As String = Split(token,"")
		Dim have_dot As Boolean

		For Each char As String In chars
			If InStr(".0123456789", char) <= 0 Then
				Return False
			End If

			If char = "." Then
				If have_dot Then
					Return False
				End If

				have_dot = True
			End If

		Next

		Return True

	End Function

	Protected Function IsRealNumber(token As String) As Boolean
		// First, check to see if the token is numeric
		If Not IsNumeric( token ) Then 
			Return False
		End If

		// Now check to see whether the token has a "." or "e" in it
		Return InStr( token, "." ) > 0 Or InStr( token, "e" ) > 0
	End Function

	Protected Function IsWhitespace(c As String) As Boolean
		// Return whether the first character of c is a whitespace character.

		Return Asc(c) <= 32

	End Function

	Protected Function KeywordDict() As MyDictionary
		// Return a dictionary containing all the RB keywords as keys
		// (and nothing useful as values).

		If (mKeywordDict Is Nil) = False Then 
			Return mKeywordDict
		End If

		Dim keywords() As String
		keywords = Split("addressof alias And Array As Assigns Boolean Break" +_
		" ByRef ByVal Call Case Catch Class Color Const Declare Delegate" +_
		" Dim Do Double DownTo Each Else #Else Elseif #ElseIf End #EndIf" +_
		" Event Exception Exit Extends False Finally For Function GoTo" +_
		" Handles If #If Implements In Inherits Integer Interface Is" +_
		" IsA Lib Loop Me Mod Module Namespace New Next Nil Not Object" +_
		" Of Optional Or ParamArray #Pragma Private Protected Public" +_
		" Raise Redim Rem Return Return Select Self Shared Single Soft" +_
		" Static Step String Sub Super Then To True Try Until Wend While", " ")

		mKeywordDict = New MyDictionary
		Dim keyword As String
		For Each keyword In keywords
			mKeywordDict.Value( keyword ) = True
		Next

		Return mKeywordDict

	End Function

	Protected Function NextToken(source As String, startPos As Integer) As String
		// Get the next token in the given source code, starting at
		// the given char offset (as in Mid, first byte = 1).

		Dim i, j, maxi As Integer
		i = startPos
		maxi = Len( source )

		Dim c As String
		c = Mid( source, i, 1 )

		If IsWhitespace(c) Then
		  // run of whitespace
		  For j = i + 1 To maxi
			If Not IsWhitespace( Mid( source, j, 1 ) ) Then 
			  Exit
			End If
		  Next
		  Return Mid( source, i, j - i )
  
		ElseIf c = "'" Then
		  // comment from here to end of line
		  Return Mid( source, i )
  
		ElseIf c = "/" And i < maxi And Mid( source, i+1, 1 ) = "/" Then
		  // comment from here to end of line
		  Return Mid( source, i )
  
		ElseIf c = "R" And Mid( source, i, 3 ) = "Rem" _
		  And IsWhitespace( Mid( source, i+3, 1 ) ) Then
		  // comment from here to end of line
		  Return Mid( source, i )
  
		ElseIf c = """" Then
		  // string literal
		  j = i
		  While j > 0 And j < maxi
			j = InStr( j+1, source, """" )
			If j = 0 Then  // no more quotes found -- terminate at end of string
			  j = maxi
			  Exit
			End If
			If Mid( source, j+1, 1 ) = """" Then
			  j = j + 1   // doubled embedded quote -- skip it and continue
			Else
			  Exit        // found the close quote; exit the loop
			End If
		  Wend
		  Return Mid( source, i, j - i + 1 )
  
		ElseIf IsDigit( c ) Or _
		  (c = "." And i < maxi And IsDigit( Mid( source, i+1, 1 ) )) Then
		  // a number
		  j = i + 1
		  While j <= maxi
			c = Mid( source, j, 1 )
			If c <> "." And Not IsDigit( c ) Then 
			  Exit
			End If
			j = j + 1
		  Wend
		  Return Mid( source, i, j - i )
  
		ElseIf InStr( "<>+-*/\^=.,():"+EndOfLine, c ) > 0 Then
		  // an operator or paren 
		  // currently all our operators are one character 
		  // uh NO ! <= >= <> ????
		  Dim nextC As String = Mid(source, i+1, 1) 
		  If c = "<" And nextC = ">" Then
			Return c + nextC
		  ElseIf c = "<" And nextC = "=" Then
			Return c + nextC
		  ElseIf c = ">" And nextC = "=" Then
			Return c + nextC
		  End If
  
		  Return c
  
		Else
		  // anything else -- grab to next delimiter
		  j = FindAnyInStr( i+1, source, "<>""+-*/\^='.(), :"+EndOfLine )
		  If j < 1 Then 
			j = maxi+1
		  End If
		  Return Mid( source, i, j - i )
		End If

		Break // should never get here -- all cases above return something
	End Function

	Protected Function TokenizeLine(sourceLine As String, includeWhitespace As Boolean = True) As String()
		// Break the given line up into tokens.
		// If includeWhitespace = true, then include whitespace as well,
		// so the original line can be completely reconstructed.
		// Otherwise, skip over any whitespace tokens.

		Dim out() As String
		Dim pos As Integer = 1
		Dim maxpos As Integer = sourceLine.Len
		Dim token As String

		While pos <= maxpos
		  token = NextToken( sourceLine, pos )
		  If token = "" Then 
			Return out
		  End If
		  If includeWhitespace Or Not IsWhitespace(token) Then
			out.Append token
		  End If
		  pos = pos + token.Len
		Wend

		Return out

	End Function

	Protected mKeywordDict As MyDictionary

	Const kScopeGlobal = "Global"
	Const kScopePrivate = "Private"
	Const kScopeProtected = "Protected" 
	Const kScopePublic = "Public"
	Const parseFailure = False
	Const parseSuccess = True

	Protected Class LocalVariable
		Sub Constructor(varName As String, VarType As String, line As Integer)
			name = varName
			type = VarType
			firstLine = line

			// quick and simple but wrong in many cases
			If VarType.Right(2) = "()" Then
				isarray = True
			Else

				Dim tokens() As String = LanguageUtils.TokenizeLine(VarType)
				// no token for type ? it cant be an array !

				If tokens.ubound > 0 Then

					Dim typeStr As String
					Dim arrDimStr As String
					Dim bracketCount As Integer

					Dim state As Integer 

					Const kSeekOpen = 0
					Const kSeekClose = 1

					state = kSeekOpen

					// skip tokens until we get to (
					// then look for matched ( ) and we retain only the commas
					While tokens.ubound >= 0

						Select Case state

						Case kSeekOpen

							If tokens(0) = "(" Then
								bracketCount = bracketCount + 1
								arrDimStr = arrDimStr + tokens(0)
								tokens.remove 0
								state = kSeekClose
							Else
								typeStr = typeStr + tokens(0)
								tokens.remove 0
							End If

						Case kSeekClose

							If tokens(0) = "(" Then

								If bracketCount > 0 Then
									arrDimStr = arrDimStr + tokens(0)
								End If

								bracketCount = bracketCount + 1
								tokens.remove 0

							ElseIf tokens(0) = ")" Then

								If bracketCount > 0 Then
									arrDimStr = arrDimStr + tokens(0)
								End If

								bracketCount = bracketCount - 1
								tokens.remove 0

								If bracketCount = 0 Then
									state = kSeekOpen
								End If

							ElseIf tokens(0) = "," Then

								bounds = bounds + tokens(0)
								arrDimStr = arrDimStr + tokens(0)
								tokens.remove 0

							Else

								bounds = bounds + tokens(0)

								tokens.remove 0

							End If

						Else

						End Select

					Wend

					If arrDimStr <> "" Then
						type = typeStr + arrDimStr
						isarray = True
					Else
						type = typeStr
						bounds = ""
					End If

				End If

			End If

		End Sub

		default_value_str As String
		firstLine As Integer
		isarray As Boolean
		name As String
		type As String
		bounds as String
		
	End Class

	Protected Class MyDictionary
		Function HasKey(key As Variant) As Boolean
			Return keys.indexof(key) >= 0 
		End Function

		Sub Value(key As Variant, Assigns value As Variant)
			keys.append key
			values.append value
		End Sub

		keys() As Variant
		values() As Variant

	End Class

End Module

Module globalroutines

	Function MyReplaceLineEndings(in_string As String, replace_str As String) As String

		Dim r1 As String = ReplaceAll(in_string, &u0D + &u0A, EndOfLine )
		Dim r2 As String = ReplaceAll(r1, &u0D, EndOfLine )
		Dim r3 As String = ReplaceAll(r2, &u0A, EndOfLine )

		Return ReplaceAll(r3, EndOfLine, replace_str)

	End Function
End Module
