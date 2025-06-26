VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "Cargar"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox txtEdad 
      Height          =   285
      Left            =   2400
      TabIndex        =   9
      Top             =   3360
      Width           =   2295
   End
   Begin VB.ComboBox cboEstadoCivil 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2400
      List            =   "Form1.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtNombres 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtApellidos 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtCedula 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblEdad 
      Caption         =   "Edad"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblEstadoCivil 
      Caption         =   "Estado Civil"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblNombres 
      Caption         =   "Nombres"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblApellidos 
      Caption         =   "Apellidos"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblCedula 
      Caption         =   "Cédula"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCargar_Click()
Dim ci As String
Dim cCiudadano As New clsCiudadano

ci = InputBox("Ingrese una cedula")
cCiudadano.Load (ci)

If cCiudadano.HasError Then
    MsgBox "Ocurrio un error - " & cCiudadano.LastErrorMessage
    Exit Sub
End If

txtCedula.Text = cCiudadano.Cedula
txtApellidos.Text = cCiudadano.Apellidos
txtNombres.Text = cCiudadano.Nombres

cboEstadoCivil.ListIndex = -1
For i = 0 To cboEstadoCivil.ListCount - 1
    If cboEstadoCivil.List(i) = cCiudadano.EstadoCivil Then
        cboEstadoCivil.ListIndex = i
        Exit For
    End If
Next i


txtEdad.Text = CStr(cCiudadano.Edad)

MsgBox "Cargado"

End Sub

Private Sub cmdGuardar_Click()

Dim cCiudadano As New clsCiudadano

cCiudadano.Cedula = txtCedula.Text
cCiudadano.Apellidos = txtApellidos.Text
cCiudadano.Nombres = txtNombres.Text
cCiudadano.EstadoCivil = cboEstadoCivil.Text

If Not IsNumeric(txtEdad.Text) Then
MsgBox "El dato de la edad es incorrecto", vbCritical
Exit Sub
Else
cCiudadano.Edad = CLng(txtEdad.Text)
End If

cCiudadano.Save


If cCiudadano.HasError Then
MsgBox "Error al guardar el ciudadano " & cCiudadano.LastErrorMessage, vbCritical
Else
MsgBox "Ciudadano guardado correctamente", vbInformation

MsgBox cCiudadano.Pista.ToString

End If

Set cCiudadano = Nothing

End Sub
