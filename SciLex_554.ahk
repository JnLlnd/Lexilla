#SingleInstance,Force
#Persistent

#Include %A_ScriptDir%\Lib\Scintilla.ahk ; from Adventure-3.0.4 aka AutoGui (https://www.autohotkey.com/boards/viewtopic.php?f=64&t=89901)
#Include %A_ScriptDir%\Lib\Scintilla_Util.ahk ; used by Scintilla class
#Include %A_ScriptDir%\Lib\Lexilla.ahk ; WORK IN PROGRESS, adapted from Scintilla.ahk

SetWorkingDir, %A_ScriptDir%

global g_blnUse554
global g_strScintillaDll
global g_strLexillaDll
global o_Sci ; class loading SciLexer.dll or Scintilla.dll according to g_blnUse554
global o_Lex ; class for Lexilla.dll
global o_Lexer ; lexer class using o_Sci or o_Lex according to g_blnUse554

g_blnUse554 := false

if (g_blnUse554)
{
	g_strScintillaDll := "lib\" . (A_PtrSize == 8 ? "Scintilla64.dll" : "Scintilla32.dll") ; v5.5.4
	g_strLexillaDll := "lib\" . (A_PtrSize == 8 ? "Lexilla64.dll" : "Lexilla32.dll") ; v5.5.4
}
else
{
	g_strScintillaDll := "lib\" . (A_PtrSize == 8 ? "SciLexer64.dll" : "SciLexer32.dll") ; v3.7.6
	g_strLexillaDll := ""
}

if (g_blnUse554)
{
	if (!LoadSciLexer(g_strScintillaDll))
	{
		MsgBox 0x10, %AppName% - Error
		, % "Failed to load library """ . g_strScintillaDll . """.`n`nThe program will exit."
		ExitApp
	}
	
	if (!LoadLexilla(g_strLexillaDll))
	{
		MsgBox 0x10, %AppName% - Error
		, % "Failed to load library """ . g_strScintillaDll . """.`n`nThe program will exit."
		ExitApp
	}
	o_Lexer := new Lexer()
}
else
{
	if (!LoadSciLexer(g_strScintillaDll)) ; from wrapper in Scintilla.ahk library
	{
		MsgBox 0x10, %AppName% - Error
		, % "Failed to load library """ . g_strScintillaDll . """.`n`nThe program will exit."
		ExitApp
	}
	Gui, New, +HwndhEditorHwnd +MinSize640x480, My Scintilla Editor
	o_Editor := new Editor("Scintilla", hEditorHwnd, true, false)
	o_Editor.NewScintilla() ; set o_Sci pointing to Scintilla class
	o_Lex := o_Sci ; o_Lex uses the same Scintilla wrapper v 3.7.6
	o_Lexer := new Lexer()
}


Gui, Add, Button, x10 w640 y500 gButtonClose, Close
Gui, Show, center

strBatchSample := "rem This is a sample`necho Echo... echo... echo...`nPause"
o_Editor.SetText(strBatchSample)
o_Editor.SetFont(true, 15)
o_Lexer.SetLexerType("BAT")

return

;-------------------------------------------------------------
ButtonClose:
GuiEscape:
GuiClose:

ExitApp
;-------------------------------------------------------------


;-------------------------------------------------------------
class Lexer
; based on Adventure IDE 3.0.4 developed by Alguimist (Gilberto Barbosa Babiretzki)
; Source: https://sourceforge.net/projects/autogui/
; Forum: https://www.autohotkey.com/boards/viewforum.php?f=64
;-------------------------------------------------------------
{
	aaSyntaxTypes := Object()
	aaSyntaxTypesByNames := Object()
	aaColors := Object()
	aaKeywords := Object()
	aaProperties := Object()
	aaDetectPatterns := Object()
	
	;---------------------------------------------------------
	__New()
	;---------------------------------------------------------
	{
		this.LoadSyntaxTypes()
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	SetLexerType(strLexerType)
	;---------------------------------------------------------
	{
		o_Lex.Type := strLexerType
		o_Lex.SetLexer(this.GetLexerByLexType(strLexerType))
		this.LoadTheme("Light")
		this.LoadLexerData(strLexerType, "Light")
		this.SetKeywords(strLexerType)
		this.ApplyTheme(strLexerType)
		o_Lex.COLOURISE(0, -1) ; 4003 SCI_COLOURISE, from start 0 to end -1
	}
	;---------------------------------------------------------
	
	;---------------------------------------------------------
	LoadSyntaxTypes()
	;---------------------------------------------------------
	{
		If (!LoadXMLEx(oXMLSyntaxTypes, A_WorkingDir . "\Lib\SyntaxTypes.xml"))
			Return

		oSyntaxTypes := oXMLSyntaxTypes.selectNodes("/syntaxtypes/types/type")
		For oType in oSyntaxTypes
		{
			Id := oType.getAttribute("id")
			Name := oType.getAttribute("name") ; Base filename
			this.aaSyntaxTypes[Id] := Object()
			this.aaSyntaxTypes[Id].Name := Name
			this.aaSyntaxTypes[Id].DN := oType.getAttribute("dn")
			this.aaSyntaxTypes[Id].Lexer := oType.getAttribute("lexer")
			this.aaSyntaxTypes[Id].Ext := oType.getAttribute("ext")

			this.aaColors[Id] := {}
			this.aaSyntaxTypesByNames[Name] := Id
		}
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	LoadTheme(strThemeName)
	; load values for themes "Light" or "Dark", not specific to language
	;---------------------------------------------------------
	{
		strThemeFile := A_WorkingDir . "\Lib\Themes.xml"
		LoadXMLEx(oXML, strThemeFile)
		If !IsObject(oXML)
		{
			Oops(0, "Error loading theme file`n~1~", strThemeFile)
			Return 0
		}
		
		Node := oXML.selectSingleNode("/themes/theme[@name='" . strThemeName . "']")

		this.aaColors["Default"] := {}
		this.aaColors["Default"].FC := this.GetThemeColor(Node, "default", "fc")
		this.aaColors["Default"].BC := this.GetThemeColor(Node, "default", "bc")
		
		this.aaColors["Caret"] := {}
		this.aaColors["Caret"].FC := this.GetThemeColor(Node, "caret", "fc")

		this.aaColors["Selection"] := {}
		this.aaColors["Selection"].FC := this.GetThemeColor(Node, "selection", "fc")
		this.aaColors["Selection"].BC := this.GetThemeColor(Node, "selection", "bc")
		this.aaColors["Selection"].Alpha := this.GetThemeValue(Node, "selection", "a")

		this.aaColors["NumbersMargin"] := {}
		this.aaColors["NumbersMargin"].FC := this.GetThemeColor(Node, "numbersmargin", "fc")
		this.aaColors["NumbersMargin"].BC := this.GetThemeColor(Node, "numbersmargin", "bc")

		this.aaColors["SymbolMargin"] := {}
		this.aaColors["SymbolMargin"].BC := this.GetThemeColor(Node, "symbolmargin", "bc")

		this.aaColors["Divider"] := {}
		this.aaColors["Divider"].BC := this.GetThemeColor(Node, "divider", "bc")
		this.aaColors["Divider"].Width := this.GetThemeValue(Node, "divider", "w")

		this.aaColors["FoldMargin"] := {}
		this.aaColors["FoldMargin"].DLC := this.GetThemeColor(Node, "foldmargin", "dlc") ; Drawing lines
		this.aaColors["FoldMargin"].BBC := this.GetThemeColor(Node, "foldmargin", "bbc") ; Button background
		this.aaColors["FoldMargin"].MBC := this.GetThemeColor(Node, "foldmargin", "mbc") ; Margin background

		this.aaColors["ActiveLine"] := {}
		this.aaColors["ActiveLine"].BC := this.GetThemeColor(Node, "activeline", "bc")

		this.aaColors["BraceMatch"] := {}
		this.aaColors["BraceMatch"].FC := this.GetThemeColor(Node, "bracematch", "fc")
		this.aaColors["BraceMatch"].Bold := this.GetThemeValue(Node, "bracematch", "b")
		this.aaColors["BraceMatch"].Italic := this.GetThemeValue(Node, "bracematch", "i")

		this.aaColors["MarkedText"] := {}
		this.aaColors["MarkedText"].Type := this.GetThemeValue(Node, "markers", "t")
		this.aaColors["MarkedText"].Color := this.GetThemeColor(Node, "markers", "c")
		this.aaColors["MarkedText"].Alpha := this.GetThemeValue(Node, "markers", "a")
		this.aaColors["MarkedText"].OutlineAlpha := this.GetThemeValue(Node, "markers", "oa")

		this.aaColors["IdenticalText"] := {}
		this.aaColors["IdenticalText"].Type := this.GetThemeValue(Node, "highlights", "t")
		this.aaColors["IdenticalText"].Color := this.GetThemeColor(Node, "highlights", "c")
		this.aaColors["IdenticalText"].Alpha := this.GetThemeValue(Node, "highlights", "a")
		this.aaColors["IdenticalText"].OutlineAlpha := this.GetThemeValue(Node, "highlights", "oa")

		this.aaColors["Calltip"] := {}
		this.aaColors["Calltip"].FC := this.GetThemeColor(Node, "calltip", "fc")
		this.aaColors["Calltip"].BC := this.GetThemeColor(Node, "calltip", "bc")

		this.aaColors["IndentGuide"] := {}
		this.aaColors["IndentGuide"].FC := this.GetThemeColor(Node, "indentguide", "fc")
		this.aaColors["IndentGuide"].BC := this.GetThemeColor(Node, "indentguide", "bc")

		this.aaColors["WhiteSpace"] := {}
		this.aaColors["WhiteSpace"].FC := this.GetThemeColor(Node, "whitespace", "fc")
		this.aaColors["WhiteSpace"].BC := this.GetThemeColor(Node, "whitespace", "bc")
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	GetThemeColor(oBaseNode, oNode, strAttrib)
	;---------------------------------------------------------
	{
		Local strValue := oBaseNode.selectSingleNode(oNode).getAttribute(strAttrib)
		Return strValue ? CvtClr(strValue) : strValue
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	GetThemeValue(oBaseNode, oNode, strAttrib)
	;---------------------------------------------------------
	{
		Return oBaseNode.selectSingleNode(oNode).getAttribute(strAttrib)
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	LoadLexerData(strLexerType, strThemeName)
	; Load specific language styles, keywords and properties
	;---------------------------------------------------------
	{
		strBaseName := this.GetNameByLexType(strLexerType)
		
		strThemeFile := A_WorkingDir . "\Lib\" . strBaseName . ".xml"
		LoadXMLEx(oXML, strThemeFile)
		If !IsObject(oXML)
		{
			Oops(0, "Error loading Lexer data`n~1~", strThemeFile)
			Return 0
		}

		; Styles
		oStyles := oXML.selectNodes("/scheme/theme[@name='" . strThemeName . "']/style")
		If (oStyles.length())
		{
			this.aaColors[strLexerType].Values := []

			For oStyle in oStyles
				this.LoadThemeStyles(strLexerType, oStyle)
		}

		; Keywords
		oKWGroups := oXML.selectNodes("/scheme/keywords/language[@id='" . strLexerType . "']/group") ; strLexerType is case sensitive (must be upper case, e.g. "AHK", "CSS")
		If (oKWGroups.length())
		{
			this.aaKeywords[strLexerType] := {}
			For oKWGroup in oKWGroups
				this.aaKeywords[strLexerType][oKWGroup.getAttribute("id")] := oKWGroup.getAttribute("keywords")
		}

		; Properties
		oProps := oXML.selectNodes("/scheme/properties/property")
		If (oProps.length())
		{
			this.aaProperties[strLexerType] := {}
			For oProp in oProps
				this.aaProperties[strLexerType][oProp.getAttribute("name")] := oProp.getAttribute("value")
		}
		
		return 1
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	LoadThemeStyles(strLexerType, oNode)
	;---------------------------------------------------------
	{
		Local strValue, strForegroundColor, strBackgroundColor

		strValue := oNode.getAttribute("v")
		If (strValue = "") ; do not use !v because we have a value 0
			Return

		this.aaColors[strLexerType][strValue] := {}
		this.aaColors[strLexerType].Values.Push(strValue)

		strForegroundColor := oNode.getAttribute("fc")
		If (strForegroundColor != "")
		{
			strForegroundColor := CvtClr(strForegroundColor)
			this.aaColors[strLexerType][strValue].FC := strForegroundColor
		}

		strBackgroundColor := oNode.getAttribute("bc")
		If (strBackgroundColor != "")
		{
			strBackgroundColor := CvtClr(strBackgroundColor)
			this.aaColors[strLexerType][strValue].BC := strBackgroundColor
		}

		this.aaColors[strLexerType][strValue].Bold := oNode.getAttribute("b")
		this.aaColors[strLexerType][strValue].Italic := oNode.getAttribute("i")
		this.aaColors[strLexerType][strValue].Under := oNode.getAttribute("u")
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	GetNameByLexType(strLexerType)
	; strLexerType is code (e.g. "AHK"), returns label name (e.g. "AutoHotkey")
	;---------------------------------------------------------
	{
		Return this.aaSyntaxTypes[strLexerType].Name ; Base filename
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	GetLexTypeByName(strTypeName)
	; strTypeName is name for label (e.g. "AutoHotkey"), returns lexer type code (e.g. "AHK")
	;---------------------------------------------------------
	{
		Return this.aaSyntaxTypesByNames[strTypeName]
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	GetLexerByLexType(strLexerType)
	; return the lexer number (e.g. 200) associated to this lexer type code (e.g. "AHK")
	;---------------------------------------------------------
	{
		Local intLexer := this.aaSyntaxTypes[strLexerType].Lexer
		Return intLexer != "" ? intLexer : 1
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	SetKeywords(Type)
	;---------------------------------------------------------
	{
		If (this.aaKeywords.HasKey(Type))
			For GrpType, Keywords in this.aaKeywords[Type]
				o_Lex.SetKeywords(GrpType, Keywords, 1)
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	ApplyTheme(Type := "")
	;---------------------------------------------------------
	{
		Local strValue, strForegroundColor, strBackgroundColor, strItalic, strBold, strUnder, intSelAlpha

		; Default color for text and background
		o_Lex.StyleSetFore(STYLE_DEFAULT, this.aaColors["Default"].FC)
		o_Lex.StyleSetBack(STYLE_DEFAULT, this.aaColors["Default"].BC)
		o_Lex.StyleClearAll() ; This message sets all styles to have the same attributes as STYLE_DEFAULT.

		; Caret
		o_Lex.SetCaretFore(this.aaColors["Caret"].FC)

		; Selection
		o_Lex.SetSelFore(1, this.aaColors["Selection"].FC)
		o_Lex.SetSelBack(1, this.aaColors["Selection"].BC)
		intSelAlpha := this.aaColors["Selection"].Alpha
		If (intSelAlpha != "")
			o_Lex.SetSelAlpha(intSelAlpha)

		; Margins
		; Line numbers
		o_Lex.StyleSetFore(STYLE_LINENUMBER, this.aaColors["NumbersMargin"].FC)
		o_Lex.StyleSetBack(STYLE_LINENUMBER, this.aaColors["NumbersMargin"].BC)
		; set or reset line number font
		o_Lex.STYLESETFONT(STYLE_LINENUMBER, "Courier New", 1) ; SCI_STYLESETFONT(int style, const char *fontName) ; STYLE_LINENUMBER:=33; ### must be followed by third parameter 1 (why?)
		o_Lex.STYLESETSIZE(STYLE_LINENUMBER, o_Edit.intLineNumberFontSizeCurrent) ; SCI_STYLESETSIZE(int style, int sizePoints) ; ### must be preceded by StyleSetFont command to work

		; Symbol margin and divider
		; o_Lex.SetMarginBackN(g_MarginSymbols, this.aaColors["SymbolMargin"].BC)
		; o_Lex.SetMarginBackN(g_MarginDivider, this.aaColors["Divider"].BC)
		; o_Lex.SetMarginWidthN(g_MarginDivider, this.aaColors["Divider"].Width)

		; Active line background color
		o_Lex.SetCaretLineBack(this.aaColors["ActiveLine"].BC)
		o_Lex.SetCaretLineVisible(o_Settings.EditorWindow.blnActiveLine.IniValue)
		o_Lex.SetCaretLineVisibleAlways(o_Settings.EditorWindow.blnActiveLine.IniValue)

		; Matching braces
		o_Lex.StyleSetBack(STYLE_BRACELIGHT, this.aaColors["ActiveLine"].BC)
		o_Lex.StyleSetFore(STYLE_BRACELIGHT, this.aaColors["BraceMatch"].FC)
		If (this.aaColors["BraceMatch"].Bold)
			o_Lex.StyleSetBold(STYLE_BRACELIGHT, True)
		If (this.aaColors["BraceMatch"].Italic)
			o_Lex.StyleSetItalic(STYLE_BRACELIGHT, True)

		; Calltips
		o_Lex.CalltipSetFore(this.aaColors["Calltip"].FC)
		o_Lex.CalltipSetBack(this.aaColors["Calltip"].BC)

		; Indentation guides
		o_Lex.StyleSetFore(37, this.aaColors["IndentGuide"].FC)
		o_Lex.StyleSetBack(37, this.aaColors["IndentGuide"].BC)

		; Language specifics
		Loop % (this.aaColors[Type].Values.Length())
		{
			strValue := this.aaColors[Type].Values[A_Index]

			strForegroundColor := this.aaColors[Type][strValue].FC
			If (strForegroundColor != "")
				o_Lex.StyleSetFore(strValue, strForegroundColor)

			strBackgroundColor := this.aaColors[Type][strValue].BC
			If (strBackgroundColor != "")
				o_Lex.StyleSetBack(strValue, strBackgroundColor)

			If (strItalic := this.aaColors[Type][strValue].Italic)
				o_Lex.StyleSetItalic(strValue, strItalic)

			If (strBold := this.aaColors[Type][strValue].Bold)
				o_Lex.StyleSetBold(strValue, strBold)

			If (strUnder := this.aaColors[Type][strValue].Under)
				o_Lex.StyleSetUnderline(strValue, strUnder)
		}
	}
	;---------------------------------------------------------

}
;-------------------------------------------------------------


;------------------------------------------------------------
SciOnNotify(wParam, lParam, msg, hWnd, obj)
; Scintilla messages when SC_MOD_INSERTTEXT or SC_MOD_DELETETEXT
;------------------------------------------------------------
{
	; Diag(A_ThisFunc, "obj.SCNCode", obj.SCNCode)
	; Diag(A_ThisFunc, "obj.modType", obj.modType)
	; Diag(A_ThisFunc, "obj.Ch", obj.Ch)
	; Diag(A_ThisFunc, "obj.modifiers", obj.modifiers)
	
    if (o_Settings.EditorWindow.blnAutoIndent.IniValue
		and obj.SCNCode == SCN_CHARADDED and obj.Ch = 10) ; SCN_CHARADDED:=2001, 10 for LF
	; intercept new line (LF), no need to send LF
	{
		; example from morosh  using PureBasic (https://www.purebasic.fr/english/viewtopic.php?p=625134&sid=0acb2962254d735d15786ce75510d69a#p625134)
		intPos := o_Sci.GETCURRENTPOS()
		intLine := o_Sci.LINEFROMPOSITION(intPos) - 1
		intLevel := o_Sci.GETLINEINDENTATION(intLine)
		if (intLevel > 0)
		{
			intLine++
			o_Sci.SETLINEINDENTATION(intLine, intLevel)
			intPos := o_Sci.GETLINEINDENTPOSITION(intLine)
			o_Sci.GOTOPOS(intPos)
		}
	}
	
    ; if (obj.SCNCode == SCN_MODIFIED) ; SCN_MODIFIED:=2008
		; and (obj.ModType != 20) ; SC_MOD_CHANGESTYLE := 0x4 | SC_PERFORMED_USER := 0x10 = 0x14 = 20
	; {
		; Diag(A_ThisFunc, "SCN_MODIFIED obj.ModType", obj.ModType)
        ; Gosub, EditorContentChanged
    ; }
}
;------------------------------------------------------------


;-------------------------------------------------------------
class Editor
;-------------------------------------------------------------
{
	;---------------------------------------------------------
	__New(strSource, intEditorHwnd, blnWordWrap, blnSeeInvisible)
	;---------------------------------------------------------
	{
		this.strSource := strSource
		this.intEditorHwnd := intEditorHwnd
		this.blnWordWrap := blnWordWrap
		this.blnSeeInvisible := blnSeeInvisible
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	NewScintilla()
	;---------------------------------------------------------
	{
		o_Sci := New Scintilla(this.intEditorHwnd, 10, 10, 640, 480)
		o_Sci.SETCODEPAGE(65001)
		
		if (o_Settings.EditorWindow.blnLineNumber.IniValue)
		{
			; margin 0 (line numbers) width is set in EditorContentChanged
			o_Sci.SETMARGINWIDTHN(1, 2) ; SCI_SETMARGINWIDTHN(int margin, int pixelWidth) ; Margin 1 (non-folding symbols), set width to 2 pixels for padding after line number (default is 16)
			o_Sci.SETMARGINLEFT("", 5) ; SCI_SETMARGINLEFT(<unused>, int pixelWidth)
			o_Sci.SETMARGINRIGHT("", 5) ; SCI_SETMARGINLEFT(<unused>, int pixelWidth)
		}
		else
			o_Sci.SETMARGINWIDTHN(1, 0) ; SCI_SETMARGINWIDTHN(int margin, int pixelWidth) ; Margin 1 (non-folding symbols), set width to 0 (default is 16)
		
		o_Sci.SETMODEVENTMASK(SCN_MODIFIED) ; SCI_SETMODEVENTMASK(int eventMask), eventMask can be SCN_MODIFIED or SC_MOD_INSERTTEXT|SC_MOD_DELETETEXT
		o_Sci.NOTIFY := "SciOnNotify"
		
		SC_POPUP_NEVER := 0 ; SC_POPUP_ALL := 1, SC_POPUP_TEXT := 2	(only if clicking on text area)
		o_Sci.USEPOPUP(SC_POPUP_NEVER)
		
		if (o_Settings.EditorWindow.blnMultipleSelection.IniValue)
		{
			; Multiple selection #### can cause application crash - to be stabilised
			o_Sci.SETMULTIPLESELECTION(true) ; SCI_SETMULTIPLESELECTION(bool multipleSelection)
			o_Sci.SETADDITIONALSELECTIONTYPING(true) ; SCI_SETADDITIONALSELECTIONTYPING(bool additionalSelectionTyping)
			o_Sci.SETMULTIPASTE(1) ; SCI_SETMULTIPASTE(int multiPaste); SC_MULTIPASTE_EACH=1
			o_Sci.SETVIRTUALSPACEOPTIONS(1) ; SCI_SETVIRTUALSPACEOPTIONS(int virtualSpaceOptions); SCVS_RECTANGULARSELECTION=1
		}
		
		this.SetWrapMode(this.blnWordWrap)
		this.SetSeeInvisible(this.blnSeeInvisible)
		this.intScintillaHwnd := o_Sci.hWnd
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	Clear()
	;---------------------------------------------------------
	{
		if (this.strSource = "Edit")
		{
		}
		else
			o_Sci.CLEAR()
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	Copy()
	;---------------------------------------------------------
	{
		o_Sci.COPY()
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	Cut()
	;---------------------------------------------------------
	{
		o_Sci.CUT()
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	GetCurrentPos()
	;---------------------------------------------------------
	{
		intCurrentPos := o_Sci.GETCURRENTPOS()
		
		return intCurrentPos
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	GetLineCount()
	;---------------------------------------------------------
	{
		return o_Sci.GETLINECOUNT()
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	GetLineStartPos(intLine)
	; my own method calling Edit_LineIndex() for Edit and PositionFromLine for Scintilla
	;---------------------------------------------------------
	{
		intLine := o_Sci.POSITIONFROMLINE(intLine)
		
		return intLine
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	GetSel(ByRef intStartSelPos := "", ByRef intEndSelPos := "")
	;---------------------------------------------------------
	{
		intStartSelPos := o_Sci.GETSELECTIONSTART()
		intEndSelPos := o_Sci.GETSELECTIONEND()
		
		return intStartSelPos
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	GetSelectionNEnd(intSelection)
	;---------------------------------------------------------
	{
		if (this.strSource = "Scintilla")
			return o_Sci.GETSELECTIONNEND(intSelection) ; SCI_GETSELECTIONNEND(int selection)
		; else Edit
		this.GetSel(intStartSelPos, intEndSelPos)
		return intEndSelPos
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	GetSelectionNStart(intSelection)
	;---------------------------------------------------------
	{
		if (this.strSource = "Scintilla")
			return o_Sci.GETSELECTIONNSTART(intSelection) ; SCI_GETSELECTIONNSTART(int selection)
		; else Edit
		this.GetSel(intStartSelPos, intEndSelPos)
		return intStartSelPos
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	GetSelectionsCount()
	;---------------------------------------------------------
	{
		if (this.strSource = "Scintilla")
			return o_Sci.GETSELECTIONS() ; SCI_GETSELECTIONS
		
		return 1 ; only one selection with Edit control
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	GetSelText()
	;---------------------------------------------------------
	{
		this.GetSel(intStartSelPos, intEndSelPos)
		intLen := intEndSelPos - intStartSelPos + 1 ; (+ 1 quiquired as in GetText()?)
		VarSetCapacity(strText, intLen, 0)
		o_Sci.GETSELTEXT(intLen, &strText)
		strText := StrGet(&strText, "UTF-8")
		
		return strText
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	GetText()
	;---------------------------------------------------------
	{
		intLen := o_Sci.GETLENGTH() + 1
		VarSetCapacity(strText, intLen, 0)
		o_Sci.GETTEXT(intLen, &strText)
		strText := StrGet(&strText, "UTF-8")

		return strText
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	GetTextLength()
	;---------------------------------------------------------
	{
		intLength := o_Sci.GETLENGTH()
		
		return intLength
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	GetTextRange(intStartPos := 0, intEndPos := -1)
	;---------------------------------------------------------
	{
		if (intEndPos = -1)
			intEndPos := this.GetTextLength()
		VarSetCapacity(strGetText, Abs(intStartPos - intEndPos) + 1, 0) ; Abs() in case values are reversed (see Editor.ahk from Adventure-3.0.4)
		VarSetCapacity(Sci_TextRange, 8 + A_PtrSize, 0) ; see https://scintilla.org/ScintillaDoc.html#Sci_TextRange
		NumPut(intStartPos, Sci_TextRange, 0, "UInt")
		NumPut(intEndPos, Sci_TextRange, 4, "UInt")
		NumPut(&strGetText, Sci_TextRange, 8, "Ptr")
		o_Sci.GETTEXTRANGE(0, &Sci_TextRange)
		strText := StrGet(&strGetText, "UTF-8")

		return strText
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	LineFromChar(intCharPos := -1)
	; return the line in the control that contains the specified character position
	; if intCharPos -1, the index of the current line is retrieved
	;---------------------------------------------------------
	{
		intLine := o_Sci.LINEFROMPOSITION(intCharPos)
		
		return intLine
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	MoveSelectedLines(strUpOrDown)
	;---------------------------------------------------------
	{
		if (strUpOrDown = "Up")
			o_Sci.MOVESELECTEDLINESUP()
		else ; Down
			o_Sci.MOVESELECTEDLINESDOWN()
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	Paste()
	;---------------------------------------------------------
	{
		o_Sci.PASTE()
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	PointXYFromPosition(intPosition := "")
	;---------------------------------------------------------
	{
		intX := o_Sci.POINTXFROMPOSITION("", intPosition)
		intY := o_Sci.POINTYFROMPOSITION("", intPosition)
		
		return {X: intX, Y:intY}
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	ReplaceSel(strText)
	;---------------------------------------------------------
	{
		o_Sci.REPLACESEL("", strText, 1) ; require 2nd parameter 1 not in documentation
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	ScrollCaret()
	;---------------------------------------------------------
	{
		o_Sci.SCROLLCARET()
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	SelectAll()
	;---------------------------------------------------------
	{
		o_Sci.SELECTALL()
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	SelectionIsRectangle()
	;---------------------------------------------------------
	{
		if (this.strSource = "Scintilla")
			return o_Sci.SELECTIONISRECTANGLE() ; SCI_SELECTIONISRECTANGLE
		
		return false
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	SetColors()
	;---------------------------------------------------------
	{
		aaTextColors := {"Light": o_Settings.FontsColors.strLightTextForeground.IniValue, "Dark": o_Settings.FontsColors.strDarkTextForeground.IniValue}
		aaBackgroundColors := {"Light": o_Settings.FontsColors.strLightTextBackground.IniValue, "Dark": o_Settings.FontsColors.strDarkTextBackground.IniValue}
		aaLineNumberForegroundColors := {"Light": o_Settings.FontsColors.strLightLineNumberForeground.IniValue, "Dark": o_Settings.FontsColors.strDarkLineNumberForeground.IniValue}
		aaLineNumberBackgroundColors := {"Light": o_Settings.FontsColors.strLightLineNumberBackground.IniValue, "Dark": o_Settings.FontsColors.strDarkLineNumberBackground.IniValue}
		
		strTheme := ((o_Settings.EditorWindow.strAppTheme.IniValue = "Windows" and !g_blnAppsUseLightTheme)
			or o_Settings.EditorWindow.strAppTheme.IniValue = "Dark" ? "Dark" : "Light")
			
		if (o_Settings.EditorWindow.blnSyntaxHighlighting.IniValue)
			o_Lexer.SetLexerType(o_Lexer.GetLexTypeByName(g_strCurrentLexerName))
		else
		{
			; Default color for text and background
			o_Sci.StyleSetFore(STYLE_DEFAULT, CvtClr("0x" . aaTextColors[strTheme]))
			o_Sci.StyleSetBack(STYLE_DEFAULT, CvtClr("0x" . aaBackgroundColors[strTheme]))
			o_Sci.StyleClearAll() ; This message sets all styles to have the same attributes as STYLE_DEFAULT.
			
			; Line numbers
			o_Sci.StyleSetFore(STYLE_LINENUMBER, CvtClr("0x" . aaLineNumberForegroundColors[strTheme]))
			o_Sci.StyleSetBack(STYLE_LINENUMBER, CvtClr("0x" . aaLineNumberBackgroundColors[strTheme]))
			
			; colors below are not in Options Gui, are only for Scintilla and only when no syntax highlighting

			; Caret
			o_Sci.SetCaretFore(CvtClr("0x" . (strTheme = "Light" ? o_Settings.FontsColors.strLightCaret.IniValue : o_Settings.FontsColors.strDarkCaret.IniValue)))

			; Selection
			o_Sci.SetSelFore(1, CvtClr("0x" . (strTheme = "Light" ? o_Settings.FontsColors.strLightSelectionForeground.IniValue : o_Settings.FontsColors.strDarkSelectionForeground.IniValue)))
			o_Sci.SetSelBack(1, CvtClr("0x" . (strTheme = "Light" ? o_Settings.FontsColors.strLightSelectionBackground.IniValue : o_Settings.FontsColors.strDarkSelectionBackground.IniValue)))
			o_Sci.SetSelAlpha(256)

			; Indentation guides
			o_Sci.StyleSetFore(STYLE_INDENTGUIDE, CvtClr("0x" . (strTheme = "Light" ? o_Settings.FontsColors.strLightIndentationForeground.IniValue
				: o_Settings.FontsColors.strDarkIndentationForeground.IniValue)))
			o_Sci.StyleSetBack(STYLE_INDENTGUIDE, CvtClr("0x" . (strTheme = "Light" ? o_Settings.FontsColors.strLightIndentationBackground.IniValue
				: o_Settings.FontsColors.strDarkIndentationBackground.IniValue)))
			
			; set or reset line number font
			o_Sci.STYLESETFONT(STYLE_LINENUMBER, "Courier New", 1) ; SCI_STYLESETFONT(int style, const char *fontName) ; STYLE_LINENUMBER:=33; ### must be followed by third parameter 1 (why?)
			o_Sci.STYLESETSIZE(STYLE_LINENUMBER, o_Edit.intLineNumberFontSizeCurrent) ; SCI_STYLESETSIZE(int style, int sizePoints) ; ### must be preceded by StyleSetFont command to work
		}
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	SetFocus()
	;---------------------------------------------------------
	{
		o_Sci.GRABFOCUS() ; use GrabFocus instead of o_Sci.SETFOCUS(true)
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	SetFont(blnFixedFont, intFontSize)
	;---------------------------------------------------------
	{
		this.blnFixedFontCurrent := blnFixedFont
		strFontName :=  (this.blnFixedFontCurrent ? "Courier New" : "Segoe UI")
		this.intFontSizeCurrent := intFontSize
				
		o_Sci.STYLESETFONT(STYLE_DEFAULT, "" . strFontName, 1) ; "" in case font name can be interpreted as number ; ### must be followed by third parameter 1 (why?)
		o_Sci.STYLESETSIZE(STYLE_DEFAULT, this.intFontSizeCurrent)
		o_Sci.STYLECLEARALL()
		; set or reset line number font
		this.intLineNumberFontSizeCurrent := Round(this.intFontSizeCurrent * 0.8)
		o_Sci.STYLESETFONT(STYLE_LINENUMBER, "Courier New", 1) ; SCI_STYLESETFONT(int style, const char *fontName) ; STYLE_LINENUMBER:=33; ### must be followed by third parameter 1 (why?)
		o_Sci.STYLESETSIZE(STYLE_LINENUMBER, this.intLineNumberFontSizeCurrent) ; SCI_STYLESETSIZE(int style, int sizePoints) ; ### must be preceded by StyleSetFont command to work
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	SetSel(intStartSelPos := 0, intEndSelPos := -1)
	; intStartSelPos and intEndSelPos are zero-based start (default 0) and end of the selection
	; if intStartSelPos and intEndSelPos are at the same position, there is no selection
	; if intStartSelPos is -1, remove any selection (set start and end to the same position, intEndSelPos is ignored)
	; if intEndSelPos is -1, end of the document (default)
	; only if this.strSource is "Scintilla", the caret is scrolled into view after this operation
	;---------------------------------------------------------
	{
		o_Sci.SETSEL(intStartSelPos, intEndSelPos)
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	SetTabs()
	;---------------------------------------------------------
	{
		o_Sci.SETUSETABS(!o_Settings.EditorWindow.blnTabSpace.IniValue)
		o_Sci.SETINDENTATIONGUIDES(o_Settings.EditorWindow.blnIndentShowGuides.IniValue ? 1 : 0) ; SC_IV_NONE* = 0, SC_IV_REAL* = 1, SC_IV_LOOKFORWARD* = 2, SC_IV_LOOKBOTH* = 3
		o_Sci.SETTABWIDTH(o_Settings.EditorWindow.intTabWidth.IniValue)
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	SetText(strText)
	;---------------------------------------------------------
	{
		o_Sci.SETTEXT("", strText, 1) ; require 2nd parameter 1 not in documentation
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	SetSeeInvisible(blnSeeInvisible)
	;---------------------------------------------------------
	{
		this.blnSeeInvisible := blnSeeInvisible
		
		; called for Scintilla only
		SCWS_INVISIBLE := 0
		SCWS_VISIBLEALWAYS := 1
		; SCWS_VISIBLEAFTERINDENT := 2, SCWS_VISIBLEONLYININDENT := 3
		o_Sci.SETVIEWWS(this.blnSeeInvisible ? SCWS_VISIBLEALWAYS : SCWS_INVISIBLE)
		o_Sci.SETVIEWEOL(this.blnSeeInvisible)
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	SetWrapMode(blnWrap)
	;---------------------------------------------------------
	{
		this.blnWordWrap := blnWrap
		; SC_WRAP_WORD (1) to enable wrapping on word or style boundaries, SC_WRAP_CHAR (2) to enable wrapping between any characters,
		; SC_WRAP_WHITESPACE (3) to enable wrapping on whitespace, and SC_WRAP_NONE (0) to disable line wrapping. 
		o_Sci.SETWRAPMODE(this.blnWordWrap ? 1 : 0)
	}

	;---------------------------------------------------------

	;---------------------------------------------------------
	TextIsSelected()
	;---------------------------------------------------------
	{
		intStart := o_Sci.GETSELECTIONSTART()
		intEnd := o_Sci.GETSELECTIONEND()
		blnSelected := intStart <> intEnd
		
		return blnSelected
	}
	;---------------------------------------------------------

	;---------------------------------------------------------
	WriteFile(strFilePath, strEncodingCode, strEolFormat)
	;---------------------------------------------------------
	{
		if not oFile := FileOpen(strFilePath, "w", (StrLen(strEncodingCode) ? strEncodingCode : A_FileEncoding))
			return -1

		strText := this.GetText()
		if (strEolFormat = "Unix") ; LF only
			strText := Edit_Convert2Unix(strText) ; use command already available in Edit library
		else if (strEolFormat = "Mac") ; CR only
			strText:=Edit_Convert2Mac(strText) ; use command already available in Edit library
		; else keep CRLF

		intResult := oFile.Write(strText)
		oFile.Close()

		return intResult
	}
	;---------------------------------------------------------
}
;-------------------------------------------------------------


;------------------------------------------------
Oops(varOwner, strMessage, objVariables*)
; varOwner can be a number or a string
;------------------------------------------------
{
	if (!varOwner)
		varOwner := 1
	Gui, %varOwner%:+OwnDialogs
	MsgBox, 48, % L(o_L["OopsTitle"], g_strAppNameText, g_strAppVersion), % L(strMessage, objVariables*)
}
;------------------------------------------------


;------------------------------------------------
L(strMessage, objVariables*)
;------------------------------------------------
{
	Loop
	{
		if InStr(strMessage, "~" . A_Index . "~")
			strMessage := StrReplace(strMessage, "~" . A_Index . "~", objVariables[A_Index])
 		else
			break
	}
	
	return strMessage
}
;------------------------------------------------


;-------------------------------------------------------------
LoadXMLEx(ByRef oXML, strFullpath)
;-------------------------------------------------------------
{
    oXML := ComObjCreate("MSXML2.DOMDocument.6.0")
    oXML.async := False

    If (!oXML.load(strFullpath))
	{
        Oops(0, "Failed to load XML file!"
        . "`n`nFilename: """ . strFullpath . """"
        . "`nError: " . Format("0x{:X}", oXML.parseError.errorCode & 0xFFFFFFFF)
        . "`nReason: " . oXML.parseError.reason)
        return 0
    }

    return 1
}
;-------------------------------------------------------------


