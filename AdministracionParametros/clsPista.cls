VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Clase: ClsPista
' Descripción: Gestiona información sobre pistas o acciones con descripciones
'==============================================================================

Private m_Action As String
Private m_Desc As String
Private m_Desc2 As String

'------------------------------------------------------------------------------
' Constructor: Inicializa los valores por defecto
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Clear
End Sub

'==============================================================================
' PROPIEDADES DE SOLO LECTURA
'==============================================================================

'------------------------------------------------------------------------------
' HasPista: Indica si hay una acción definida
'------------------------------------------------------------------------------
Public Property Get HasPista() As Boolean
    HasPista = (Len(Trim(m_Action)) > 0)
End Property

'------------------------------------------------------------------------------
' IsComplete: Indica si todos los campos están completos
'------------------------------------------------------------------------------
Public Property Get IsComplete() As Boolean
    IsComplete = HasPista And _
                 (Len(Trim(m_Desc)) > 0) And _
                 (Len(Trim(m_Desc2)) > 0)
End Property

'==============================================================================
' PROPIEDADES DE LECTURA/ESCRITURA
'==============================================================================

'------------------------------------------------------------------------------
' Action: Acción principal
'------------------------------------------------------------------------------
Public Property Get Action() As String
    Action = m_Action
End Property

Public Property Let Action(ByVal valor As String)
    m_Action = Trim(valor)
End Property

'------------------------------------------------------------------------------
' Desc: Descripción principal
'------------------------------------------------------------------------------
Public Property Get Desc() As String
    Desc = m_Desc
End Property

Public Property Let Desc(ByVal valor As String)
    m_Desc = Trim(valor)
End Property

'------------------------------------------------------------------------------
' Desc2: Descripción secundaria
'------------------------------------------------------------------------------
Public Property Get Desc2() As String
    Desc2 = m_Desc2
End Property

Public Property Let Desc2(ByVal valor As String)
    m_Desc2 = Trim(valor)
End Property

'==============================================================================
' MÉTODOS PÚBLICOS
'==============================================================================

'------------------------------------------------------------------------------
' Clear: Limpia todos los valores
'------------------------------------------------------------------------------
Public Sub Clear()
    m_Action = ""
    m_Desc = ""
    m_Desc2 = ""
End Sub

'------------------------------------------------------------------------------
' SetPista: Establece todos los valores de una vez
'------------------------------------------------------------------------------
Public Sub SetPista(ByVal sAction As String, _
                    ByVal sDesc As String, _
                    Optional ByVal sDesc2 As String = "")
    Me.Action = sAction
    Me.Desc = sDesc
    Me.Desc2 = sDesc2
End Sub

'------------------------------------------------------------------------------
' ToString: Devuelve una representación en texto
'------------------------------------------------------------------------------
Public Function ToString() As String
    Dim result As String
    
    If HasPista Then
        result = "Acción: " & m_Action
        If Len(m_Desc) > 0 Then
            result = result & vbCrLf & "Descripción: " & m_Desc
        End If
        If Len(m_Desc2) > 0 Then
            result = result & vbCrLf & "Descripción 2: " & m_Desc2
        End If
    Else
        result = "[Sin pista definida]"
    End If
    
    ToString = result
End Function
