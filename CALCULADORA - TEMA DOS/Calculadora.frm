VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton DIVISION 
      Caption         =   "/"
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton LIMPIAR 
      Caption         =   "C"
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton MULTIPLICAR 
      Caption         =   "*"
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton RESTA 
      Caption         =   "-"
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton SUMA 
      Caption         =   "+"
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton IGUAL 
      Caption         =   "="
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton PUNTO 
      Caption         =   "."
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton CERO 
      Caption         =   "0"
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton DOS 
      Caption         =   "2"
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton SEIS 
      Caption         =   "6"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton UNO 
      Caption         =   "1"
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton TRES 
      Caption         =   "3"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton NUEVE 
      Caption         =   "9"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton CUATRO 
      Caption         =   "4"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton CINCO 
      Caption         =   "5"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton OCHO 
      Caption         =   "8"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton SIETE 
      Caption         =   "7"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'En esta seccion de codigo se declaran las varibles a utilizar, en este caso son variables locales
Dim OPERACION As String
Dim A As Double
Dim B  As Double

'Se prepara el evento click el cual al suceder dicho evento se realizara la linea de codigo correspondiente la cual establece al ser precionado el boton en la caja de texto el numero asignado el simbolo & hace la funcion de concatenar los numeros insertados en la caja de texto,las lineas de codigo se implementan para los numeros del 1 al 9 incluyendo el boton de punto ya que su unica diferencia es que se asigana un numero diferente a cada boton
Private Sub CERO_Click()
Text1.Text = Text1.Text & "0"
End Sub

Private Sub CINCO_Click()
Text1.Text = Text1.Text & "5"
End Sub

Private Sub CUATRO_Click()
Text1.Text = Text1.Text & "4"
End Sub
'Dentro de estas secciones de codigo se lleva acabo la operacion matematica que se desa guardando los primeros valores en una variable llamada A, po siguinete se borra lo que esta en el cuadro de texto y se prepara para la operacion asignada, estas lineas de codigo nos sirven para los demas operaciones matematicas y solo se le cambia la operacion a realizar
Private Sub DIVISION_Click()
A = Text1.Text
Text1.Text = ""
OPERACION = "/"
End Sub

Private Sub DOS_Click()
Text1.Text = Text1.Text & "2"
End Sub
'Al oprimir el boton igual se efectuan las lineas de cogigo las cuales nos dicen que al tener las dos variables A y B procede a preguntar que boton es el que ha sido precionado para asi realizar la operasion asiganda
Private Sub IGUAL_Click()
B = Text1.Text
Text1.Text = ""
If OPERACION = "+" Then
Text1.Text = A + B
ElseIf OPERACION = "-" Then
Text1.Text = A - B
ElseIf OPERACION = "*" Then
Text1.Text = A * B
ElseIf OPERACION = "/" Then
Text1.Text = A / B
End If
End Sub
'Esta linea de codigo nos srive para limpiar la caja de texto
Private Sub LIMPIAR_Click()
Text1.Text = ""
End Sub

Private Sub MULTIPLICAR_Click()
A = Text1.Text
Text1.Text = ""
OPERACION = "*"
End Sub

Private Sub NUEVE_Click()
Text1.Text = Text1.Text & "9"
End Sub

Private Sub OCHO_Click()
Text1.Text = Text1.Text & "8"
End Sub

Private Sub PUNTO_Click()
Text1.Text = Text1.Text & "."
End Sub

Private Sub RESTA_Click()
A = Text1.Text
Text1.Text = ""
OPERACION = "-"
End Sub

Private Sub SEIS_Click()
Text1.Text = Text1.Text & "6"
End Sub

Private Sub SIETE_Click()
Text1.Text = Text1.Text & "7"
End Sub

Private Sub SUMA_Click()
A = Text1.Text
Text1.Text = ""
OPERACION = "+"
End Sub

Private Sub TRES_Click()
Text1.Text = Text1.Text & "3"
End Sub

Private Sub UNO_Click()
Text1.Text = Text1.Text & "1"
End Sub
