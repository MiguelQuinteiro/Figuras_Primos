VERSION 5.00
Begin VB.Form frmFigurasPrimos 
   BackColor       =   &H00000000&
   Caption         =   "Figuras Primos"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtZoom 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdZoomMenos 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdZoomMas 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmFigurasPrimos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim zoom As Integer

Private Sub cmdZoomMas_Click()
  If zoom < 16384 Then
    zoom = zoom * 2
    txtZoom.Text = zoom
    Call Form_DblClick
  End If
End Sub

Private Sub cmdZoomMenos_Click()
  If zoom > 1 Then
    zoom = zoom / 2
    txtZoom.Text = zoom
    Call Form_DblClick
  End If
End Sub

' AL CARGAR EL FORMULARIO
Private Sub Form_Load()
  zoom = 0.51
End Sub

' AL HACER DOBLE CLICK
Private Sub Form_DblClick()
  Dim i As Long
  Dim p As Long
  Dim s As Boolean
  Dim x As Long
  Dim y As Long
  Dim ancho As Long
  Dim alto As Long

  x = 200
  y = 4000
  'x = frmFigurasPrimos.Width / 2
  'y = frmFigurasPrimos.Height / 2
  'zoom = 8
  alto = frmFigurasPrimos.Height * zoom
  ancho = frmFigurasPrimos.Width * zoom

  Cls

  ' Eje de coordenadas
  Line (x, y - alto)-(x, y + alto), vbRed
  Line (x - ancho, y)-(x + ancho, y), vbRed

  ' Pinta Primos

  For i = 1 To 20000
    If Primo(i) Then
      'If Primo(i + 2) Then
      Call DibujaCirculo(x, y, i * zoom, 2)
      'Call DibujaCirculo(x, y, i * zoom, (i Mod 3))
      'Line (x + i * zoom, y - alto)-(x + i * zoom, y + alto), vbWhite
      'Line (x - ancho, y + i * zoom)-(x + ancho, y + i * zoom), vbWhite

      'Line (x, y + i * zoom)-(x + i * zoom, y), vbGreen
      'x = x + i * zoom
      'y = y + i * zoom
      'End If
    End If

    '        'Caso especial
    '        If i = 4 Then
    '            Call DibujaCirculo(x, y, i * zoom, 10)
    '            Line (x + i * zoom, y - alto)-(x + i * zoom, y + alto), vbGreen
    '            Line (x - ancho, y + i * zoom)-(x + ancho, y + i * zoom), vbGreen
    '
    '            Line (x, y + i * zoom)-(x + i * zoom, y), vbRed
    '            Line (x + i * zoom, y + alto)-(x + ancho, y + i * zoom), vbRed
    '        End If
  Next i

  ' Pinta Pares


End Sub

' DIBUJA UN CIRCULO
Public Sub DibujaCirculo(ByVal pX As Long, ByVal pY As Long, ByVal pRadio As Long, ByVal pColor As Integer)
  Circle (pX, pY), pRadio, QBColor(pColor)
End Sub

' FUNCION PARA CALCULAR SI EL NUMERO ES PRIMO
Public Function Primo(ByVal pN As Long) As Boolean
  Dim i As Long
  Primo = True
  If pN = 1 Then
    Primo = False
  Else
    For i = 2 To Sqr(pN)
      If (pN / i) = Int(pN / i) Then
        Primo = False
      End If
    Next i
  End If
End Function

