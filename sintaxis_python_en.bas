Attribute VB_Name = "sintaxis_python"

' Add in References "Microsoft VBScript Regular Expressions 5.5"
' On the web https://regex101.com/ regular expressions are explained
' Need to improve: -Function name "\def\s+(\w+)\("
' To improve, comment, suggest this project, write to rubend18@hotmail.com

' Declaration of a constant representing the total number of search patterns
Const total As Integer = 11

' Global declaration of the variable that will contain the selected range of text
Public texto As Range

Sub sintaxis_python()
    
    ' Turn off screen updating
    Application.ScreenUpdating = False
    
    ' Regular expressions
    Dim patrones(total) As String
    patrones(0) = "\b(__debug__)\b"
    patrones(1) = "\b(f|and|class|def|global|in|is|lambda|nonlocal|not|or|False|None|NotImplemented|True)\b"
    patrones(2) = "\b(as|assert|async|await|break|case|continue|del|elif|else|except|finally|for|from|if|import|match|pass|raise|return|try|while|with|yield)\b"
    patrones(3) = "\b(__import__|abs|all|any|ascii|bin|breakpoint|callable|chr|compile|delattr|dir|divmod|enumerate|eval|exec|filter|format|getattr|globals|hasattr|hash|help|hex|id|input|isinstance|issubclass|iter|len|locals|map|max|memoryview|min|next|oct|open|ord|pow|print|range|repr|reversed|round|setattr|sorted|sum|vars|zip)\b"
    patrones(4) = "\b(bool|bytearray|bytes|classmethod|complex|dict|float|frozenset|int|list|object|property|set|slice|staticmethod|str|tuple|type)\b"
    patrones(5) = "\b(Any|Callable|Coroutine|Dict|Ellipsis|List|Literal|Generic|Optional|Sequence|Set|Tuple|Type|Union|super)\b"
    patrones(6) = "(\+|\-|\*|\/|\%|\*\*|\/\/|>|<|==|>=|<=|!=|&|\||\^|~|>>|<<|=|\+=|-=|\*=|\/=|\%=|\*\*=|\/\/=|&=|\|=|\^=|>>=|<<=)"
    patrones(7) = "\b\d+\.?\w*"
    patrones(8) = "(['""]).*?\1"
    patrones(9) = "\def\s+(\w+)\(|^\s*@\w+"
    patrones(10) = "\#\s*(.*?)\s*\$\$"

    ' Color codes
    Dim colores(total) As Long
    colores(0) = RGB(0, 16, 128)   ' __debug__
    colores(1) = RGB(0, 0, 225)    ' Reserved words
    colores(2) = RGB(175, 0, 219)  ' Flow control reserved words
    colores(3) = RGB(121, 94, 38)  ' Built-in functions and methods
    colores(4) = RGB(37, 118, 147) ' Built-in data types
    colores(5) = RGB(0, 0, 0)      ' Special data types
    colores(6) = RGB(0, 0, 0)      ' Arithmetic, Comparison, and Logical Operators
    colores(7) = RGB(9, 129, 86)   ' Whole numbers or decimals
    colores(8) = RGB(163, 21, 21)  ' Text strings
    colores(9) = RGB(121, 94, 38)  ' Functions and decorators
    colores(10) = RGB(0, 128, 0)   ' Comments
    
    ' Text to apply syntactic highlighting
    Set texto = Selection.Range
    If texto = "" Then: MsgBox "Seleccione un texto", vbInformation, "Error": Exit Sub
    Debug.Print texto.Start & " " & texto.End
        
    ' Set flag line breaks
    ReemplazarTexto "^p", "$$^p"
    
    ' Remove tabs
    ReemplazarTexto "^t", "    "

    ' Apply initial format
    With texto
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle ' Spaces
        .ParagraphFormat.Alignment = wdAlignParagraphLeft ' Aligned
        .Font.Name = "consolas"
        .Font.Size = 8
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
        .Font.Italic = False
        .LanguageID = wdEnglishUS ' English
        .NoProofing = True ' No concealer
    End With

    ' Buscar patrones y aplicar formato
    Dim objRegExp As New RegExp
    objRegExp.MultiLine = True
    objRegExp.Global = True ' Busca todas las coincidencias

    Dim i As Integer
    For i = 0 To UBound(patrones) - 1
        objRegExp.Pattern = patrones(i)
        
        Dim objMatches As MatchCollection
        Set objMatches = objRegExp.Execute(texto)

        Dim objMatch As match
        For Each objMatch In objMatches
            Selection.Start = objMatch.FirstIndex + texto.Start
            Selection.End = objMatch.FirstIndex + objMatch.Length + texto.Start
            Selection.Font.Color = colores(i)
        Next objMatch
    Next i

    ' Remove flag for line breaks
    ReemplazarTexto "$$", ""
    
    ' Remove spaces before line breaks
    ReemplazarTexto "^w^p", "^p"
    
    ' Return to the beginning of the selected text
    Selection.Start = texto.Start
    Selection.End = texto.Start
    
    ' Activate screen refresh
    Application.ScreenUpdating = True
End Sub

' Function to Replace Text
Function ReemplazarTexto(ByVal buscar As String, ByVal reemplazar As String) As String
    texto.Find.Text = buscar
    texto.Find.Replacement.Text = reemplazar
    texto.Find.Execute Replace:=wdReplaceAll
End Function


