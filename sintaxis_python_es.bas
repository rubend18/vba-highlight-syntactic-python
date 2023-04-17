Attribute VB_Name = "sintaxis_python"

' Agregar en Referencias "Microsoft VBScript Regular Expressions 5.5"
' En la web https://regex101.com/ se explican las expresiones regulares
' Hay que mejorar: -Nombre de función "\def\s+(\w+)\("
' Para mejorar, comentar, sugerir este proyecto, escribir a rubend18@hotmail.com

' Declaración de una constante que representa el número total de patrones de búsqueda
Const total As Integer = 11

' Declaración global de la variable que contendrá el rango de texto seleccionado
Public texto As Range

Sub sintaxis_python()
    
    ' Desactivar la actualización de pantalla
    Application.ScreenUpdating = False
    
    ' Expresiones regulares
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

    ' Códigos de colores
    Dim colores(total) As Long
    colores(0) = RGB(0, 16, 128)   ' __debug__
    colores(1) = RGB(0, 0, 225)    ' Palabras reservadas
    colores(2) = RGB(175, 0, 219)  ' Palabras reservadas control de flujo
    colores(3) = RGB(121, 94, 38)  ' Funciones y métodos integrados
    colores(4) = RGB(37, 118, 147) ' Tipos de datos integrados
    colores(5) = RGB(0, 0, 0)      ' Tipos de datos especiales
    colores(6) = RGB(0, 0, 0)      ' Operadores aritméticos, de comparación y lógicos
    colores(7) = RGB(9, 129, 86)   ' Números enteros o decimales
    colores(8) = RGB(163, 21, 21)  ' Cadenas de texto
    colores(9) = RGB(121, 94, 38)  ' Funciones y decoradores
    colores(10) = RGB(0, 128, 0)   ' Comentarios
    
    ' Texto a aplicar resaltado sintáctico
    Set texto = Selection.Range
    If texto = "" Then: MsgBox "Seleccione un texto", vbInformation, "Error": Exit Sub
    Debug.Print texto.Start & " " & texto.End
        
    ' Poner bandera saltos de línea
    ReemplazarTexto "^p", "$$^p"
    
    ' Quitar tabulaciones
    ReemplazarTexto "^t", "    "

    ' Aplicar formato inicial
    With texto
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle ' Espaceado
        .ParagraphFormat.Alignment = wdAlignParagraphLeft ' Alineado
        .Font.Name = "consolas"
        .Font.Size = 8
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
        .Font.Italic = False
        .LanguageID = wdEnglishUS ' Inglés
        .NoProofing = True ' Sin corrector
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

    ' Quitar bandera para saltos de línea
    ReemplazarTexto "$$", ""
    
    ' Eliminar espacios antes de saltos de linea
    ReemplazarTexto "^w^p", "^p"
    
    ' Volver al inicio del texto seleccionado
    Selection.Start = texto.Start
    Selection.End = texto.Start
    
    ' Activar la actualización de pantalla
    Application.ScreenUpdating = True
End Sub

' Función para Reemplazar Texto
Function ReemplazarTexto(ByVal buscar As String, ByVal reemplazar As String) As String
    texto.Find.Text = buscar
    texto.Find.Replacement.Text = reemplazar
    texto.Find.Execute Replace:=wdReplaceAll
End Function


