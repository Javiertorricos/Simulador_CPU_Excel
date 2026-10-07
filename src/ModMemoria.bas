Attribute VB_Name = "ModMemoria"
Option Explicit

' ==============================================================================
'  UNIVERSIDAD CATOLICA BOLIVIANA "SAN PABLO"
'  Materia: Arquitectura de Computadoras (SIS-131)
'  Proyecto: Simulador de CPU von Neumann (8 bits) y Memoria Principal
'  Autor: Javier Torrico Sejas
'
'  MODULO: ModMemoria
'  Descripcion:
'    Gestiona la Memoria Principal (RAM) de 256 bytes (posiciones 00h a FFh).
'    Implementa la segmentacion logica entre Codigo (00h-7Fh) y Datos (80h-FFh).
'    Provee las operaciones primitivas requeridas: Read(address) y Write(address, value).
' ==============================================================================

' Constantes de segmentacion de memoria (256 bytes en total)
Public Const SEG_CODIGO_INICIO As Long = &H0      ' 0d
Public Const SEG_CODIGO_FIN As Long = &H7F        ' 127d
Public Const SEG_DATOS_INICIO As Long = &H80      ' 128d
Public Const SEG_DATOS_FIN As Long = &HFF         ' 255d

' Arreglo de almacenamiento fisico de la memoria RAM (256 celdas de 8 bits)
Public MemRAM(0 To 255) As Byte

' Variables para auditoria y visualizacion de la ultima operacion primitiva
Public UltimaOpMem As String       ' Ej: "Read(0x10) -> 0x05" o "Write(0x80, 0x2A)"
Public UltimaDirMem As Long        ' Direccion accedida (-1 si ninguna)
Public UltimoTipoMem As String     ' "R" (Lectura) o "W" (Escritura)

' ==============================================================================
'  OPERACIONES PRIMITIVAS DE MEMORIA (Requisito Especifico del Parcial)
' ==============================================================================

''' <summary>
''' Subrutina primitiva de lectura de memoria: Read(address).
''' Lee un byte de la direccion especificada de la RAM.
''' </summary>
Public Function Read(ByVal address As Long) As Byte
    Dim addrValida As Long
    addrValida = address And &HFF  ' Asegura rango de 8 bits (0..255)
    
    ' Lectura desde el arreglo fisico
    Read = MemRAM(addrValida)
    
    ' Registro de auditoria para la interfaz
    UltimaDirMem = addrValida
    UltimoTipoMem = "R"
    UltimaOpMem = "Read(0x" & Hex2(addrValida) & ") -> 0x" & Hex2(Read)
    
    ' Reflejar acceso en panel de primitivas de la hoja
    Call ModInterfaz.ActualizarPanelPrimitivas("READ", addrValida, CLng(Read))
    Call ModInterfaz.ResaltarCeldaMemoria(addrValida, "LECTURA")
End Function

''' <summary>
''' Subrutina primitiva de escritura de memoria: MemWrite(address, value).
''' NOTA: "Write" es palabra reservada en VBA (instruccion Write #).
''' Se implementa como MemWrite / WriteRAM para compatibilidad total.
''' </summary>
Public Sub MemWrite(ByVal address As Long, ByVal value As Byte)
    Dim addrValida As Long
    Dim valByte As Byte
    addrValida = address And &HFF
    valByte = value And &HFF
    
    ' Escritura en el arreglo fisico
    MemRAM(addrValida) = valByte
    
    ' Registro de auditoria para la interfaz
    UltimaDirMem = addrValida
    UltimoTipoMem = "W"
    UltimaOpMem = "Write(0x" & Hex2(addrValida) & ", 0x" & Hex2(valByte) & ")"
    
    ' Actualizar celda en la matriz 16x16 de la hoja
    Call ModMemoria.Memoria_ActualizarCeldaUI(addrValida)
    
    ' Reflejar acceso en panel de primitivas de la hoja
    Call ModInterfaz.ActualizarPanelPrimitivas("WRITE", addrValida, CLng(valByte))
    Call ModInterfaz.ResaltarCeldaMemoria(addrValida, "ESCRITURA")
End Sub

Public Sub WriteRAM(ByVal address As Long, ByVal value As Byte)
    Call MemWrite(address, value)
End Sub

' ==============================================================================
'  FUNCIONES AUXILIARES Y DE GESTION DE MEMORIA
' ==============================================================================

''' <summary>
''' Reinicia toda la memoria RAM a 00h y refresca la hoja Excel.
''' </summary>
Public Sub Memoria_Reset()
    Dim i As Long
    For i = 0 To 255
        MemRAM(i) = 0
    Next i
    UltimaDirMem = -1
    UltimoTipoMem = ""
    UltimaOpMem = "Memoria inicializada en 00h"
    Call Memoria_RefrescarTodaUI
    Call ModInterfaz.ActualizarPanelPrimitivas("LISTO", 0, 0)
End Sub

''' <summary>
''' Carga una lista de bytes a partir de una direccion inicial.
''' </summary>
Public Sub Memoria_CargarBytes(ByVal dirInicio As Long, ByRef datos() As Byte, ByVal cantidad As Long)
    Dim i As Long
    For i = 0 To cantidad - 1
        Dim dest As Long
        dest = (dirInicio + i) And &HFF
        MemRAM(dest) = datos(i)
    Next i
    Call Memoria_RefrescarTodaUI
End Sub

''' <summary>
''' Actualiza el valor de una celda especifica en la cuadricula matricial 16x16 de Excel.
''' La matriz se ubica en la hoja "Simulador" con filas 0..F y columnas 0..F.
''' </summary>
Public Sub Memoria_ActualizarCeldaUI(ByVal dir As Long)
    Dim ws As Worksheet
    Dim filaOffset As Long, colOffset As Long
    Dim f As Long, c As Long
    
    Set ws = ModInterfaz.HojaSim()
    If ws Is Nothing Then Exit Sub
    
    f = (dir \ 16)   ' Fila 0..15 (nibble alto)
    c = (dir Mod 16) ' Columna 0..15 (nibble bajo)
    
    ' Fila base = 6 (fila 00h esta en la fila 6 de Excel), Columna base = 18 (Col R)
    filaOffset = 6 + f
    colOffset = 18 + c
    
    ws.Cells(filaOffset, colOffset).NumberFormat = "@"
    ws.Cells(filaOffset, colOffset).Value = Hex2(MemRAM(dir))
End Sub

''' <summary>
''' Refresca todas las 256 celdas de la cuadricula matricial 16x16 de Excel.
''' </summary>
Public Sub Memoria_RefrescarTodaUI()
    Dim ws As Worksheet
    Dim f As Long, c As Long, dir As Long
    Dim matrizValores(1 To 16, 1 To 16) As Variant
    
    Set ws = ModInterfaz.HojaSim()
    If ws Is Nothing Then Exit Sub
    
    For f = 0 To 15
        For c = 0 To 15
            dir = f * 16 + c
            matrizValores(f + 1, c + 1) = Hex2(MemRAM(dir))
        Next c
    Next f
    
    ' Escritura en bloque para maxima velocidad y fluidez
    ws.Range("R6:AG21").NumberFormat = "@"
    ws.Range("R6:AG21").Value = matrizValores
End Sub

''' <summary>
''' Inspecciona el valor de una direccion de memoria y lo muestra en el panel inspector.
''' </summary>
Public Sub Memoria_Inspeccionar(ByVal dir As Long)
    Dim val As Long
    dir = dir And &HFF
    val = MemRAM(dir)
    Call ModInterfaz.MostrarInspectorMemoria(dir, val)
End Sub

' ==============================================================================
'  CONVERSORES FORMATO (Hexadecimal de 2 digitos y Binario de 8 bits)
' ==============================================================================

Public Function Hex2(ByVal v As Long) As String
    Dim h As String
    h = Hex$(v And &HFF)
    If Len(h) = 1 Then h = "0" & h
    Hex2 = h
End Function

Public Function Bin8(ByVal v As Long) As String
    Dim b As String
    Dim i As Long
    v = v And &HFF
    b = ""
    For i = 7 To 0 Step -1
        If (v And (2 ^ i)) <> 0 Then
            b = b & "1"
        Else
            b = b & "0"
        End If
    Next i
    Bin8 = b
End Function
