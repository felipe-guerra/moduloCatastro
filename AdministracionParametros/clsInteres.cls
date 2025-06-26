VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCiudadano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
' Clase: clsInteresMora
'   Gestiona CRUD sobre la tabla INTERES_MORA
'   (IntMo_Año, IntMo_Mes, IntMo_Valor)
'*****************************************************

Option Explicit


Private Const M_DEFAULT_MONTH As String = "00"

'------------------------
' Variables privadas
'------------------------
'Private m_CB     As clsConexionBuilder        ' Constructor de cadena de conexión
Private m_CnnDB  As ADODB.Connection       ' Conexión a la base de datos
Private m_EObj As clsError
Private m_Pista As clsPista


Private m_Cedula As String
Private m_Apellidos As String
Private m_Nombres As String
Private m_EstadoCivil As String
Private m_Edad As Long



Public Property Get Cedula() As String
Cedula = m_Cedula
End Property

Public Property Let Cedula(valor As String)

    If Not IsNumeric(valor) Then
    MsgBox "La cedula debe tener solo numeros "
    Exit Property
    End If
    
    m_Cedula = valor

End Property


Public Property Get Apellidos() As String
    Apellidos = m_Apellidos
End Property

Public Property Let Apellidos(valor As String)
    m_Apellidos = valor
End Property

Public Property Get Nombres() As String
    Nombres = m_Nombres
End Property

Public Property Let Nombres(valor As String)
    m_Nombres = valor
End Property

Public Property Get EstadoCivil() As String
    EstadoCivil = m_EstadoCivil
End Property

Public Property Let EstadoCivil(valor As String)
    m_EstadoCivil = valor
End Property

Public Property Get Edad() As String
    Edad = m_Edad
End Property

Public Property Let Edad(valor As String)
    m_Edad = valor
End Property


'------------------------
' Inicialización / limpieza
'------------------------
Private Sub Class_Initialize()
    ' Crear builder y establecer conexión
    'Set m_CB = New clsConexionBuilder
    Set m_CnnDB = New ADODB.Connection
    Set m_Pista = New clsPista
    'm_CnnDB.ConnectionString = m_CB.ConnectionString
    
    m_CnnDB.ConnectionString = "Provider=SQLOLEDB.1;Password=sql;Persist Security Info=True;User ID=sa;Initial Catalog=catastro;Data Source=."
    
    Set m_EObj = New clsError
End Sub

Private Sub Class_Terminate()
    ' Cerrar y liberar recursos
    CloseConnection
    Set m_CnnDB = Nothing
    'Set m_CB = Nothing
    Set m_EObj = Nothing
    Set m_Pista = Nothing
End Sub

Public Property Get LastErrorMessage() As String
    LastErrorMessage = m_EObj.ErrorMessage
End Property

Public Property Get HasError() As Boolean
    HasError = m_EObj.HasError
End Property

Public Property Get Pista() As clsPista
Set Pista = m_Pista
End Property


'------------------------
' Abrir / Cerrar conexión
'------------------------
Private Sub OpenConnection()
    If m_CnnDB.State <> adStateOpen Then
        m_CnnDB.Open
    End If
End Sub

Private Sub CloseConnection()
    If m_CnnDB.State = adStateOpen Then
        m_CnnDB.Close
    End If
End Sub

' La funcion extrae la información de la base de datos, tabla ciudadano, usando el campo Cedula
Public Sub Load(pCedula As String)

    Dim cmd As ADODB.Command
    Dim rs  As ADODB.Recordset
    'Dim val As Currency

    On Error GoTo ErrorHandler

    m_EObj.ClearError
    OpenConnection

    ' Preparar comando parametrizado
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = m_CnnDB
        .CommandType = adCmdText
        .CommandText = "SELECT * FROM CIUDADANO " & _
                            "WHERE Ciu_Cedula = ? "
        .Parameters.Append .CreateParameter("pcmdCedula", adVarChar, adParamInput, 10, pCedula)

    End With

    Set rs = cmd.Execute
    If Not rs.EOF Then
        Cedula = rs.Fields("Ciu_Cedula").Value
        Nombres = rs.Fields("Ciu_Nombres").Value
        Apellidos = rs.Fields("Ciu_Apellidos").Value
        EstadoCivil = rs.Fields("Ciu_EstadoCivil").Value
        Edad = rs.Fields("Ciu_Edad").Value
    Else
        m_EObj.SetError ("No se encontro información para la Cédula: " & pCedula)
    End If

    ' Cerrar objetos
    rs.Close
    CloseConnection
    Set rs = Nothing
    Set cmd = Nothing

    Exit Sub

ErrorHandler:
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    CloseConnection
    Set cmd = Nothing
    m_EObj.SetError ("Error al consultar el ciudadano: " & Err.Description)
    
End Sub




' Metodo para guardar
Public Sub Save()

    Dim cmd As ADODB.Command
       
    On Error GoTo ErrorHandler

    m_EObj.ClearError
    OpenConnection
    m_CnnDB.BeginTrans
    
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = m_CnnDB
        .CommandType = adCmdText
        .CommandText = "DELETE FROM CIUDADANO WHERE Ciu_Cedula = ? "
        .Parameters.Append .CreateParameter("pcmdCedula", adVarChar, adParamInput, 10, Me.Cedula)
    End With

    cmd.Execute
    Set cmd = Nothing

    
    ' Preparar comando parametrizado
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = m_CnnDB
        .CommandType = adCmdText
        .CommandText = "INSERT INTO CIUDADANO (Ciu_Cedula, Ciu_Apellidos, Ciu_Nombres, Ciu_EstadoCivil, Ciu_Edad ) " & _
                        "VALUES (?,?,?,?,?)"
        .Parameters.Append .CreateParameter("pcmdCedula", adVarChar, adParamInput, 10, Me.Cedula)
        .Parameters.Append .CreateParameter("pcmdApellidos", adVarChar, adParamInput, 35, Me.Apellidos)
        .Parameters.Append .CreateParameter("pcmdNombres", adVarChar, adParamInput, 35, Me.Nombres)
        .Parameters.Append .CreateParameter("pcmdEstadoCivil", adVarChar, adParamInput, 13, Me.EstadoCivil)
        .Parameters.Append .CreateParameter("pcmdEdad", adInteger, adParamInput, , Me.Edad)

    End With

    cmd.Execute
    m_CnnDB.CommitTrans
    
    CloseConnection
    
    Set cmd = Nothing
    
    
    m_Pista.Clear
    m_Pista.SetPista "AGREGAR CIUDADANO", "Agrega el Ciudadano de la Tabla de Ciudadano", "Cédula: " & Me.Cedula
    
    
    Exit Sub

ErrorHandler:
    If m_CnnDB.State = adStateOpen Then
    m_CnnDB.RollbackTrans
    End If
    
    CloseConnection
    Set cmd = Nothing
    m_EObj.SetError ("Error al guardar el ciudadano: " & Err.Description)
    
End Sub

Sub Delete(pCedula As String)
Dim cmd As ADODB.Command
    
    On Error GoTo ErrorHandler

    m_EObj.ClearError
    OpenConnection

    ' Preparar comando parametrizado
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = m_CnnDB
        .CommandType = adCmdText
        .CommandText = "DELETE FROM CIUDADANO " & _
                            "WHERE Ciu_Cedula = ? "
        .Parameters.Append .CreateParameter("pcmdCedula", adVarChar, adParamInput, 10, pCedula)

    End With

    cmd.Execute
    
    ' Cerrar objetos
    
    CloseConnection
    Set cmd = Nothing

    m_Pista.Clear
    m_Pista.SetPista "ELIMINAR CIUDADANO", "Elimina el Ciudadano de la Tabla de Ciudadano", "Cédula: " & pCedula

    Exit Sub

ErrorHandler:
        CloseConnection
    Set cmd = Nothing
    m_EObj.SetError ("Error al eliminar el ciudadano: " & Err.Description)

End Sub





'------------------------
' Verifica si existe un registro para un año y mes dados
'------------------------
Public Function HasRecord(ByVal year As String, Optional ByVal month As String = M_DEFAULT_MONTH) As Boolean
    Dim cmd As ADODB.Command
    Dim rs  As ADODB.Recordset

    On Error GoTo ErrorHandler

    m_EObj.ClearError
    OpenConnection
    
    ' Preparar comando parametrizado
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = m_CnnDB
        .CommandType = adCmdText
        .CommandText = "SELECT COUNT(*) AS Total FROM INTERES_MORA " & _
                            "WHERE IntMo_Año = ? AND IntMo_Mes = ?"
        .Parameters.Append .CreateParameter("pYear", adVarChar, adParamInput, 4, year)
        .Parameters.Append .CreateParameter("pMonth", adVarChar, adParamInput, 2, month)
    End With

    Set rs = cmd.Execute
    HasRecord = (rs.Fields("Total").Value > 0)

    ' Cerrar objetos
    rs.Close
    CloseConnection
    Set rs = Nothing
    Set cmd = Nothing
    Exit Function

ErrorHandler:
    ' Manejo de errores y limpieza
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    CloseConnection
    Set cmd = Nothing
    m_EObj.SetError ("Error al verificar existencia: " & Err.Description)
    HasRecord = False
End Function

'------------------------
' Inserta un nuevo registro
'------------------------
Public Function InsertRecord(ByVal year As String, ByVal month As String, ByVal iValue As Currency) As Boolean
    Dim cmd As ADODB.Command

    On Error GoTo ErrorHandler

    m_EObj.ClearError
    OpenConnection

    ' Preparar comando parametrizado
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = m_CnnDB
        .CommandType = adCmdText
        .CommandText = "INSERT INTO INTERES_MORA " & _
                            "(IntMo_Año, IntMo_Mes, IntMo_Valor) VALUES (?, ?, ?)"
        .Parameters.Append .CreateParameter("pYear", adVarChar, adParamInput, 4, year)
        .Parameters.Append .CreateParameter("pMonth", adVarChar, adParamInput, 2, month)
        .Parameters.Append .CreateParameter("pValue", adCurrency, adParamInput, , iValue)
        .Execute
    End With

    InsertRecord = True
    CloseConnection
    Set cmd = Nothing
    Exit Function

ErrorHandler:
    CloseConnection
    Set cmd = Nothing
    m_EObj.SetError ("Error al insertar registro: " & Err.Description)
    InsertRecord = False
End Function

'------------------------
' Actualiza un registro existente
'------------------------
Public Function UpdateRecord(ByVal year As String, ByVal month As String, ByVal iValue As Currency) As Boolean
    Dim cmd As ADODB.Command

    On Error GoTo ErrorHandler

    m_EObj.ClearError
    OpenConnection
    
    ' Preparar comando parametrizado
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = m_CnnDB
        .CommandType = adCmdText
        .CommandText = "UPDATE INTERES_MORA SET IntMo_Valor = ? " & _
                            "WHERE IntMo_Año = ? AND IntMo_Mes = ?"
        .Parameters.Append .CreateParameter("pValue", adCurrency, adParamInput, , iValue)
        .Parameters.Append .CreateParameter("pYear", adVarChar, adParamInput, 4, year)
        .Parameters.Append .CreateParameter("pMonth", adVarChar, adParamInput, 2, month)
        .Execute
    End With

    UpdateRecord = True
    CloseConnection
    Set cmd = Nothing
    Exit Function

ErrorHandler:
    CloseConnection
    Set cmd = Nothing
    m_EObj.SetError ("Error al actualizar registro: " & Err.Description)
    UpdateRecord = False
End Function

'------------------------
' Obtiene el interés para un año y mes dados
'------------------------
Public Function GetInteres(ByVal year As String, ByVal month As String) As Currency
    Dim cmd As ADODB.Command
    Dim rs  As ADODB.Recordset
    Dim val As Currency

    On Error GoTo ErrorHandler

    m_EObj.ClearError
    OpenConnection

    ' Preparar comando parametrizado
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = m_CnnDB
        .CommandType = adCmdText
        .CommandText = "SELECT IntMo_Valor FROM INTERES_MORA " & _
                            "WHERE IntMo_Año = ? AND IntMo_Mes = ?"
        .Parameters.Append .CreateParameter("pYear", adVarChar, adParamInput, 4, year)
        .Parameters.Append .CreateParameter("pMonth", adVarChar, adParamInput, 2, month)
    End With

    Set rs = cmd.Execute
    If Not rs.EOF And Not IsNull(rs.Fields("IntMo_Valor").Value) Then
        val = CCur(rs.Fields("IntMo_Valor").Value)
    Else
        m_EObj.SetError ("No se encontró el interés para " & month & "/" & year & ".")
        val = -1
    End If

    ' Cerrar objetos
    rs.Close
    CloseConnection
    Set rs = Nothing
    Set cmd = Nothing
    GetInteres = val
    Exit Function

ErrorHandler:
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    CloseConnection
    Set cmd = Nothing
    m_EObj.SetError ("Error al obtener interés: " & Err.Description)
    GetInteres = -1
End Function

'------------------------
' Obtiene el interés anual (mes = "00")
'------------------------
Public Function GetInteresAnual(ByVal year As String) As Currency
    Dim cmd As ADODB.Command
    Dim rs  As ADODB.Recordset
    Dim val As Currency

    On Error GoTo ErrorHandler


    ' Validación del parámetro año
'    If Len(year) <> 4 Then
'        MsgBox "El año debe tener 4 dígitos.", vbExclamation, TEXT_SIC
'        GetAnnualInterest = 0
'        Exit Function
'    End If

    m_EObj.ClearError
    OpenConnection
    
    ' Preparar comando parametrizado
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = m_CnnDB
        .CommandType = adCmdText
        .CommandText = "SELECT IntMo_Valor FROM INTERES_MORA " & _
                            "WHERE IntMo_Año = ? AND IntMo_Mes = '00'"
        .Parameters.Append .CreateParameter("pYear", adVarChar, adParamInput, 4, year)
    End With

    Set rs = cmd.Execute
    If Not rs.EOF And Not IsNull(rs.Fields("IntMo_Valor").Value) Then
        val = CCur(rs.Fields("IntMo_Valor").Value)
    Else
        m_EObj.SetError ("No se encontró el interés anual para el año " & year & ".")
        val = 0
    End If

    ' Cerrar objetos
    rs.Close
    CloseConnection
    Set rs = Nothing
    Set cmd = Nothing
    GetInteresAnual = val
    Exit Function

ErrorHandler:
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    CloseConnection
    Set cmd = Nothing
    m_EObj.SetError ("Error al obtener el interés anual: " & Err.Description)
    GetInteresAnual = -1
End Function


Public Property Get Min() As Long
    Dim rs  As New ADODB.Recordset
    Dim val As Long

    On Error GoTo ErrorHandler

    m_EObj.ClearError
    OpenConnection
    
    rs.Open "SELECT Min([IntMo_Año]) As A FROM INTERES_MORA ", m_CnnDB, adOpenForwardOnly
    
    If Not rs.EOF And Not IsNull(rs.Fields("A").Value) Then
        val = CLng(rs.Fields("A").Value)
    Else
        m_EObj.SetError ("Error al leer el menor año de interes")
        val = 0
    End If

    ' Cerrar objetos
    rs.Close
    CloseConnection
    Set rs = Nothing
    Min = val
    Exit Property

ErrorHandler:
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    CloseConnection
    m_EObj.SetError ("Error al obtener el minimo año: " & Err.Description)
    Min = 0

End Property

Public Property Get Max() As Long
    Dim rs  As New ADODB.Recordset
    Dim val As Long

    On Error GoTo ErrorHandler

    m_EObj.ClearError
    OpenConnection
    
    rs.Open "SELECT Max([IntMo_Año]) As A FROM INTERES_MORA ", m_CnnDB, adOpenForwardOnly
    
    If Not rs.EOF And Not IsNull(rs.Fields("A").Value) Then
        val = CLng(rs.Fields("A").Value)
    Else
        m_EObj.SetError ("Error al leer el mayor año de interes")
        val = 0
    End If

    ' Cerrar objetos
    rs.Close
    CloseConnection
    Set rs = Nothing
    Max = val
    Exit Property

ErrorHandler:
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    CloseConnection
    m_EObj.SetError ("Error al obtener el mayor año: " & Err.Description)
    Max = 0

End Property


Public Function GetTable() As Currency()
    Dim AInteres(1950 To 2200) As Currency
    Dim M_n, M_x As Long, i As Long
    M_n = Me.Min
    M_x = Me.Max
    
    For i = 1950 To 2200
        AInteres(i) = 0
        If M_n <> 0 And M_x <> 0 Then
            If i >= M_n And i <= M_x Then
                AInteres(i) = GetInteresAnual(CStr(i))
                If (AInteres(i) = 0) And HasError Then
                    AInteres(i) = 0
                End If
            End If
        End If
    Next i
    
    GetTable = AInteres
End Function

