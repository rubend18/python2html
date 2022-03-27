VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Python to HTML"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11055
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "ABRIR HTML"
      Height          =   495
      Left            =   9360
      TabIndex        =   8
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "GUARDAR HTML"
      Height          =   495
      Left            =   9360
      TabIndex        =   7
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "COPIAR HTML"
      Height          =   495
      Left            =   9360
      TabIndex        =   6
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4920
      Width           =   8895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NUEVO"
      Height          =   495
      Left            =   9360
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   8895
   End
   Begin VB.Label Label4 
      Caption         =   "Curso de Programación en Python - UNIR 2022"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   9240
      Width           =   6615
   End
   Begin VB.Label Label3 
      Caption         =   "Programado por Rubén Darío Jaramillo"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   9000
      Width           =   6615
   End
   Begin VB.Label Label2 
      Caption         =   "Código HTML"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4560
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese el código Python"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
    Text1.Text = Empty
    Text2.Text = Empty
    Text1.SetFocus
End Sub

Private Sub Command2_Click()
    Text2.SetFocus
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    Clipboard.Clear
    Clipboard.SetText Text2.SelText
    Text2.SetFocus
End Sub

Private Sub Command3_Click()
    Open App.Path & "/CODIGO.html" For Output As #1
    Print #1, Text2.Text
    Close #1
End Sub

Private Sub Command4_Click()
    Open App.Path & "/CODIGO.html" For Output As #1
    Print #1, Text2.Text
    Close #1
    ShellExecute Me.hwnd, "Open", App.Path & "/CODIGO.html", "", "", 1
End Sub

Function invertida(ByVal Cadena As String) As String
    Dim acumulador As String
    Dim caracter As String
    Dim i As Variant
    acumulador = ""
    badera = True
    For i = 1 To Len(Cadena)
        caracter = UCase(Mid(Cadena, i, 1))
        If caracter Like " " Then
            badera = True
            acumulador = acumulador & " "
        Else
            If badera = False Then
                acumulador = acumulador & UCase(caracter)
            Else
                acumulador = acumulador & LCase(caracter)
                badera = False
            End If
        End If
    Next i
    invertida = acumulador
End Function

Private Sub Form_Load()
    On Error Resume Next
    Text1.Text = GetSetting("Python2HTML", "seccion", "a", "")
    Text2.Text = GetSetting("Python2HTML", "seccion", "b", "")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SaveSetting "Python2HTML", "seccion", "a", Text1.Text
    SaveSetting "Python2HTML", "seccion", "b", Text2.Text
End Sub

Private Sub Text1_Change()
   
    cad1 = Text1.Text
    
    For i = 1 To Len(cad1)

        'ESPACIO
        If Mid(cad1, i, 1) = " " Then
            cad2 = cad2 & "&nbsp;"
            
        'CADENA
        ElseIf Mid(cad1, i, 1) = """" Then cad2 = cad2 & "<span style=""color:#ff0000;"">""</span>"

        'NUMEROS
        ElseIf Mid(cad1, i, 1) = "." Then cad2 = cad2 & "<span style=""color:#008800;"">.</span>"
        ElseIf Mid(cad1, i, 1) = "0" Then cad2 = cad2 & "<span style=""color:#008800;"">0</span>"
        ElseIf Mid(cad1, i, 1) = "1" Then cad2 = cad2 & "<span style=""color:#008800;"">1</span>"
        ElseIf Mid(cad1, i, 1) = "2" Then cad2 = cad2 & "<span style=""color:#008800;"">2</span>"
        ElseIf Mid(cad1, i, 1) = "3" Then cad2 = cad2 & "<span style=""color:#008800;"">3</span>"
        ElseIf Mid(cad1, i, 1) = "4" Then cad2 = cad2 & "<span style=""color:#008800;"">4</span>"
        ElseIf Mid(cad1, i, 1) = "5" Then cad2 = cad2 & "<span style=""color:#008800;"">5</span>"
        ElseIf Mid(cad1, i, 1) = "6" Then cad2 = cad2 & "<span style=""color:#008800;"">6</span>"
        ElseIf Mid(cad1, i, 1) = "7" Then cad2 = cad2 & "<span style=""color:#008800;"">7</span>"
        ElseIf Mid(cad1, i, 1) = "8" Then cad2 = cad2 & "<span style=""color:#008800;"">8</span>"
        ElseIf Mid(cad1, i, 1) = "9" Then cad2 = cad2 & "<span style=""color:#008800;"">9</span>"
        
        'OPERADORES
        'ElseIf Mid(cad1, i, 1) = "(" Then cad2 = cad2 & "<span style=""color:#000000;"">(</span>"
        'ElseIf Mid(cad1, i, 1) = ")" Then cad2 = cad2 & "<span style=""color:#000000;"">)</span>"
        'ElseIf Mid(cad1, i, 1) = "[" Then cad2 = cad2 & "<span style=""color:#000000;"">[</span>"
        'ElseIf Mid(cad1, i, 1) = "]" Then cad2 = cad2 & "<span style=""color:#000000;"">]</span>"
        'ElseIf Mid(cad1, i, 1) = "{" Then cad2 = cad2 & "<span style=""color:#000000;"">{</span>"
        'ElseIf Mid(cad1, i, 1) = "}" Then cad2 = cad2 & "<span style=""color:#000000;"">}</span>"
        'ElseIf Mid(cad1, i, 1) = ":" Then cad2 = cad2 & "<span style=""color:#000000;"">:</span>"
        'ElseIf Mid(cad1, i, 1) = "," Then cad2 = cad2 & "<span style=""color:#000000;"">,</span>"
        'ElseIf Mid(cad1, i, 1) = "_" Then cad2 = cad2 & "<span style=""color:#000000;"">_</span>"
        
        ElseIf Mid(cad1, i, 1) = "+" Then cad2 = cad2 & "<span style=""color:#aa22ff;""><strong>+</strong></span>"
        ElseIf Mid(cad1, i, 1) = "-" Then cad2 = cad2 & "<span style=""color:#aa22ff;""><strong>-</strong></span>"
        ElseIf Mid(cad1, i, 1) = "*" Then cad2 = cad2 & "<span style=""color:#aa22ff;""><strong>*</strong></span>"
        ElseIf Mid(cad1, i, 1) = "/" Then cad2 = cad2 & "<span style=""color:#aa22ff;""><strong>/</strong></span>"
        ElseIf Mid(cad1, i, 1) = "%" Then cad2 = cad2 & "<span style=""color:#aa22ff;""><strong>%</strong></span>"
        ElseIf Mid(cad1, i, 1) = "^" Then cad2 = cad2 & "<span style=""color:#aa22ff;""><strong>^</strong></span>"
        ElseIf Mid(cad1, i, 1) = "=" Then cad2 = cad2 & "<span style=""color:#aa22ff;""><strong>=</strong></span>"
        ElseIf Mid(cad1, i, 1) = ">" Then cad2 = cad2 & "<span style=""color:#aa22ff;""><strong>&gt;</strong></span>"
        ElseIf Mid(cad1, i, 1) = "<" Then cad2 = cad2 & "<span style=""color:#aa22ff;""><strong>&lt;</strong></span>"
        ElseIf Mid(cad1, i, 1) = "!" Then cad2 = cad2 & "<span style=""color:#aa22ff;""><strong>!</strong></span>"
        ElseIf Mid(cad1, i, 1) = "&" Then cad2 = cad2 & "<span style=""color:#aa22ff;""><strong>&</strong></span>"
        ElseIf Mid(cad1, i, 1) = "|" Then cad2 = cad2 & "<span style=""color:#aa22ff;""><strong>|</strong></span>"
        ElseIf Mid(cad1, i, 1) = "~" Then cad2 = cad2 & "<span style=""color:#aa22ff;""><strong>~</strong></span>"

        'PALABRAS RESERVADAS
        'https://eiposgrados.com/blog-python/palabras-reservadas-python/
        ElseIf Mid(cad1, i, 2) = "as" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>as</strong></span>": i = i + 1
        ElseIf Mid(cad1, i, 2) = "if" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>if</strong></span>": i = i + 1
        ElseIf Mid(cad1, i, 2) = "in" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>in</strong></span>": i = i + 1
        ElseIf Mid(cad1, i, 2) = "is" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>is</strong></span>": i = i + 1
        ElseIf Mid(cad1, i, 2) = "or" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>or</strong></span>": i = i + 1
        ElseIf Mid(cad1, i, 3) = "and" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>and</strong></span>": i = i + 2
        ElseIf Mid(cad1, i, 3) = "def" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>def</strong></span>": i = i + 2
        ElseIf Mid(cad1, i, 3) = "del" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>del</strong></span>": i = i + 2
        ElseIf Mid(cad1, i, 3) = "for" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>for</strong></span>": i = i + 2
        ElseIf Mid(cad1, i, 3) = "not" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>not</strong></span>": i = i + 2
        ElseIf Mid(cad1, i, 3) = "try" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>try</strong></span>": i = i + 2
        ElseIf Mid(cad1, i, 4) = "elif" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>elif</strong></span>": i = i + 3
        ElseIf Mid(cad1, i, 4) = "else" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>else</strong></span>": i = i + 3
        ElseIf Mid(cad1, i, 4) = "exec" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>exec</strong></span>": i = i + 3
        ElseIf Mid(cad1, i, 4) = "from" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>from</strong></span>": i = i + 3
        ElseIf Mid(cad1, i, 4) = "None" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>None</strong></span>": i = i + 3
        ElseIf Mid(cad1, i, 4) = "pass" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>pass</strong></span>": i = i + 3
        ElseIf Mid(cad1, i, 4) = "True" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>True</strong></span>": i = i + 3
        ElseIf Mid(cad1, i, 4) = "with" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>with</strong></span>": i = i + 3
        ElseIf Mid(cad1, i, 5) = "async" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>async</strong></span>": i = i + 4
        ElseIf Mid(cad1, i, 5) = "await" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>await</strong></span>": i = i + 4
        ElseIf Mid(cad1, i, 5) = "break" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>break</strong></span>": i = i + 4
        ElseIf Mid(cad1, i, 5) = "class" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>class</strong></span>": i = i + 4
        ElseIf Mid(cad1, i, 5) = "False" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>False</strong></span>": i = i + 4
        ElseIf Mid(cad1, i, 5) = "print" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>print</strong></span>": i = i + 4
        ElseIf Mid(cad1, i, 5) = "raise" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>raise</strong></span>": i = i + 4
        ElseIf Mid(cad1, i, 5) = "while" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>while</strong></span>": i = i + 4
        ElseIf Mid(cad1, i, 5) = "yield" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>yield</strong></span>": i = i + 4
        ElseIf Mid(cad1, i, 5) = "round" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>round</strong></span>": i = i + 4
        ElseIf Mid(cad1, i, 6) = "assert" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>assert</strong></span>": i = i + 5
        ElseIf Mid(cad1, i, 6) = "except" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>except</strong></span>": i = i + 5
        ElseIf Mid(cad1, i, 6) = "global" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>global</strong></span>": i = i + 5
        ElseIf Mid(cad1, i, 6) = "import" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>import</strong></span>": i = i + 5
        ElseIf Mid(cad1, i, 6) = "lambda" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>lambda</strong></span>": i = i + 5
        ElseIf Mid(cad1, i, 6) = "return" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>return</strong></span>": i = i + 5
        ElseIf Mid(cad1, i, 7) = "finally" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>finally</strong></span>": i = i + 6
        ElseIf Mid(cad1, i, 8) = "continue" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>continue</strong></span>": i = i + 7
        ElseIf Mid(cad1, i, 8) = "nonlocal" Then cad2 = cad2 & "<span style=""color:#008000;""><strong>nonlocal</strong></span>": i = i + 7
        
        'CASO CONTRARIO
        Else
            cad2 = cad2 & Mid(cad1, i, 1)
        End If
    Next i
    
    'CADENAS
    cad2 = cad2 & "<br>"
    cadInicio = "<pre><span style=""font-family:courier new,courier,monospace;"">"
    cadFin = "</span></pre>"
    coment1 = "<span style=""color:#007979;""><em>#&nbsp;Comentario</em></span><br>"
    coment2 = "<br><span style=""color:#ba2121;"">'''<br>Comentario<br>'''</span><br>"
    coment3 = coment1 & coment1 & coment1 & "<br>"
    
    'RESULTADO
    Text2.Text = cadInicio & coment3 & cad2 & coment2 & cadFin

End Sub

'nombrefuncion #0000ff azul
