VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_ErrorMessage As String
Private m_LastMessage As String
Private m_HasError As Boolean

Private Sub Class_Initialize()
    m_ErrorMessage = ""
    m_HasError = False
    m_LastMessage = ""
End Sub


' Método para obtener si hay error
Public Property Get HasError() As Boolean
    HasError = m_HasError
End Property

' Método para obtener el mensaje de error
Public Property Get ErrorMessage() As String
    ErrorMessage = m_ErrorMessage
End Property

' Método para establecer un error
Public Sub SetError(ByVal Message As String)
    m_HasError = True
    m_ErrorMessage = Message
End Sub

' Método para limpiar errores
Public Sub ClearError()
    m_HasError = False
    m_ErrorMessage = ""
End Sub

