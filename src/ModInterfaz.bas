Attribute VB_Name = "ModInterfaz"
Option Explicit

' ==============================================================================
'  UNIVERSIDAD CATOLICA BOLIVIANA "SAN PABLO"
'  Materia: Arquitectura de Computadoras (SIS-131)
'  Proyecto: Simulador de CPU von Neumann (8 bits) y Memoria Principal
'  Autor: Javier Torrico Sejas
'
'  MODULO: ModInterfaz
'  Descripcion:
'    Controla la interfaz grafica del simulador en Excel:
'    - Iluminacion visual dinamica de las 4 Fases (Fetch, Decode, Execute, Store).
'    - Resaltado de buses, registros activos y celdas de memoria (PC, Lectura, Escritura).
'    - Panel de auditoria de operaciones primitivas Read/Write.
'    - Log cronologico de micro-operaciones en tiempo real.
'    - Inspector interactivo de memoria.
' ==============================================================================

Private mUltimaCeldaRAM As Range
Private mUltimaFilaEditor As Long

Public Function HojaSim() As Worksheet
    On Error Resume Next
    Set HojaSim = ThisWorkbook.Worksheets("Simulador")
    If HojaSim Is Nothing Then Set HojaSim = ActiveWorkbook.Worksheets("Simulador")
End Function

' ==============================================================================
'  MACROS VINCULADAS A BOTONES DE LA HOJA
' ==============================================================================

Public Sub Btn_CargarPrograma()
    Call ModEnsamblador.Ensamblar_DesdeHoja
End Sub

Public Sub Btn_PasoFase()
    Call ModCicloInstruccion.Paso_Fase
End Sub

Public Sub Btn_PasoInstruccion()
    Call ModCicloInstruccion.Paso_InstruccionCompleta
End Sub

Public Sub Btn_Run()
    Call ModCicloInstruccion.Modo_Continuo
End Sub

Public Sub Btn_Pausa()
    Call ModCicloInstruccion.Pausar_Ejecucion
End Sub

Public Sub Btn_Reset()
    Call ModCicloInstruccion.Pausar_Ejecucion
    Call ModCPU.CPU_Reset
    Call ModMemoria.Memoria_Reset
    Call ModInterfaz.RestaurarColoresEditor
    Call ModInterfaz.LogMensaje("SIMULADOR REINICIADO POR COMPLETO (Memoria y Registros a 00h).")
End Sub

Public Sub Btn_Ejemplo1(): Call ModEnsamblador.CargarProgramaDemo(1): End Sub
Public Sub Btn_Ejemplo2(): Call ModEnsamblador.CargarProgramaDemo(2): End Sub
Public Sub Btn_Ejemplo3(): Call ModEnsamblador.CargarProgramaDemo(3): End Sub
Public Sub Btn_Ejemplo4(): Call ModEnsamblador.CargarProgramaDemo(4): End Sub

' ==============================================================================
'  ILUMINACION DINAMICA DE FASES DEL CICLO DE RELOJ
' ==============================================================================

''' <summary>
''' Ilumina la tarjeta de la fase activa con colores vivos y atenua las demas.
''' </summary>
Public Sub ActualizarFaseUI(ByVal fase As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = HojaSim()
    If ws Is Nothing Then Exit Sub
    
    Dim colInactivo As Long, txtInactivo As Long, txtActivo As Long
    colInactivo = RGB(236, 240, 241) ' Gris pizarra suave
    txtInactivo = RGB(127, 140, 141)
    txtActivo = RGB(255, 255, 255)   ' Blanco
    
    ' Restablecer todas a inactivo
    ws.Range("H5:I6").Interior.Color = colInactivo: ws.Range("H5:I6").Font.Color = txtInactivo
    ws.Range("J5:K6").Interior.Color = colInactivo: ws.Range("J5:K6").Font.Color = txtInactivo
    ws.Range("L5:M6").Interior.Color = colInactivo: ws.Range("L5:M6").Font.Color = txtInactivo
    ws.Range("N5:O6").Interior.Color = colInactivo: ws.Range("N5:O6").Font.Color = txtInactivo
    
    ' Activar la fase correspondiente con su color tematico
    Select Case UCase$(Trim$(fase))
        Case "FETCH"
            ws.Range("H5:I6").Interior.Color = RGB(39, 174, 96)   ' Verde esmeralda
            ws.Range("H5:I6").Font.Color = txtActivo
        Case "DECODE"
            ws.Range("J5:K6").Interior.Color = RGB(41, 128, 185)  ' Azul cobalto
            ws.Range("J5:K6").Font.Color = txtActivo
        Case "EXECUTE"
            ws.Range("L5:M6").Interior.Color = RGB(230, 126, 34)  ' Naranja ambar
            ws.Range("L5:M6").Font.Color = txtActivo
        Case "STORE"
            ws.Range("N5:O6").Interior.Color = RGB(142, 68, 173)  ' Purpura
            ws.Range("N5:O6").Font.Color = txtActivo
    End Select
End Sub

''' <summary>
''' Actualiza los indicadores visuales de las banderas de estado (ZF, CF, SF).
''' </summary>
Public Sub ActualizarFlagsUI(ByVal zf As Long, ByVal cf As Long, ByVal sf As Long)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = HojaSim()
    If ws Is Nothing Then Exit Sub
    
    ' ZF: Celda L12
    If zf = 1 Then
        ws.Range("L12").Interior.Color = RGB(46, 204, 113) ' Verde activo
        ws.Range("L12").Font.Color = RGB(255, 255, 255)
    Else
        ws.Range("L12").Interior.Color = RGB(240, 243, 244)
        ws.Range("L12").Font.Color = RGB(149, 165, 166)
    End If
    
    ' CF: Celda M12
    If cf = 1 Then
        ws.Range("M12").Interior.Color = RGB(231, 76, 60)   ' Rojo alerta
        ws.Range("M12").Font.Color = RGB(255, 255, 255)
    Else
        ws.Range("M12").Interior.Color = RGB(240, 243, 244)
        ws.Range("M12").Font.Color = RGB(149, 165, 166)
    End If
    
    ' SF: Celda N12
    If sf = 1 Then
        ws.Range("N12").Interior.Color = RGB(52, 152, 219)  ' Azul activo
        ws.Range("N12").Font.Color = RGB(255, 255, 255)
    Else
        ws.Range("N12").Interior.Color = RGB(240, 243, 244)
        ws.Range("N12").Font.Color = RGB(149, 165, 166)
    End If
End Sub

Public Sub ActualizarPanelALU(ByVal opDesc As String, ByVal res As Long)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = HojaSim()
    If ws Is Nothing Then Exit Sub
    ws.Cells(14, 12).Value = opDesc
    ws.Cells(14, 16).Value = "0x" & ModMemoria.Hex2(res) & " (" & res & ")"
End Sub

Public Sub ActualizarPanelPrimitivas(ByVal tipo As String, ByVal dir As Long, ByVal val As Long)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = HojaSim()
    If ws Is Nothing Then Exit Sub
    ws.Cells(16, 8).Value = tipo
    ws.Cells(16, 10).Value = "0x" & ModMemoria.Hex2(dir) & " (" & dir & ")"
    ws.Cells(16, 12).Value = "0x" & ModMemoria.Hex2(val) & " (" & val & ")"
    ws.Cells(16, 14).Value = ModMemoria.UltimaOpMem
End Sub

Public Sub ResaltarBus(ByVal nombreBus As String, ByVal valor As Long)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = HojaSim()
    If ws Is Nothing Then Exit Sub
    If nombreBus = "DIRECCIONES" Then
        ws.Cells(5, 16).Value = "Bus Dir: 0x" & ModMemoria.Hex2(valor)
    Else
        ws.Cells(6, 16).Value = "Bus Dat: 0x" & ModMemoria.Hex2(valor)
    End If
End Sub

' ==============================================================================
'  RESALTADO DE CELDAS DE MEMORIA Y LINEAS DEL EDITOR
' ==============================================================================

Public Sub ResaltarCeldaMemoria(ByVal dir As Long, ByVal tipo As String)
    On Error Resume Next
    Dim ws As Worksheet
    Dim f As Long, c As Long
    Dim celda As Range
    
    Set ws = HojaSim()
    If ws Is Nothing Then Exit Sub
    
    f = (dir \ 16)
    c = (dir Mod 16)
    Set celda = ws.Cells(6 + f, 18 + c)
    
    ' Restaurar color de la celda previa segun su segmento
    If Not mUltimaCeldaRAM Is Nothing Then
        Dim dirAnt As Long
        dirAnt = (mUltimaCeldaRAM.Row - 6) * 16 + (mUltimaCeldaRAM.Column - 18)
        If dirAnt >= 0 And dirAnt <= 255 Then
            If dirAnt < 128 Then
                mUltimaCeldaRAM.Interior.Color = RGB(235, 245, 255) ' Segmento Codigo (Celeste sutil)
            Else
                mUltimaCeldaRAM.Interior.Color = RGB(235, 255, 235) ' Segmento Datos (Verde sutil)
            End If
        End If
    End If
    
    ' Aplicar nuevo color
    Select Case tipo
        Case "PC"
            celda.Interior.Color = RGB(133, 193, 233)  ' Azul Cian brillante
        Case "LECTURA"
            celda.Interior.Color = RGB(249, 231, 159)  ' Amarillo Ambar
        Case "ESCRITURA"
            celda.Interior.Color = RGB(245, 183, 177)  ' Coral / Rojo suave
    End Select
    
    Set mUltimaCeldaRAM = celda
    
    ' Actualizar Inspector de Memoria
    Call MostrarInspectorMemoria(dir, ModMemoria.MemRAM(dir))
End Sub

Public Sub ResaltarLineaEditor(ByVal dirPC As Long)
    On Error Resume Next
    Dim ws As Worksheet
    Dim fila As Long
    Dim strDir As String
    
    Set ws = HojaSim()
    If ws Is Nothing Then Exit Sub
    
    strDir = "0x" & ModMemoria.Hex2(dirPC)
    
    ' Restaurar fila anterior a blanco
    If mUltimaFilaEditor >= 7 And mUltimaFilaEditor <= 26 Then
        ws.Range(ws.Cells(mUltimaFilaEditor, 2), ws.Cells(mUltimaFilaEditor, 6)).Interior.Color = RGB(255, 255, 255)
    End If
    
    ' Buscar fila que contenga la direccion actual
    For fila = 7 To 26
        If Trim$(CStr(ws.Cells(fila, 2).Value)) = strDir Then
            ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 6)).Interior.Color = RGB(254, 249, 231) ' Amarillo claro
            mUltimaFilaEditor = fila
            Exit For
        End If
    Next fila
End Sub

Public Sub RestaurarColoresEditor()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = HojaSim()
    If ws Is Nothing Then Exit Sub
    ws.Range("B7:F26").Interior.Color = RGB(255, 255, 255)
    mUltimaFilaEditor = 0
End Sub

' ==============================================================================
'  LOG CRONOLOGICO DE MICRO-OPERACIONES
' ==============================================================================

Public Sub LogMensaje(ByVal msg As String)
    On Error Resume Next
    Dim ws As Worksheet, wsLog As Worksheet
    Dim f As Long
    
    Set ws = HojaSim()
    If ws Is Nothing Then Exit Sub
    
    ' Desplazar filas hacia abajo en el panel de log en pantalla (filas 19 a 27, columna H = 8)
    For f = 27 To 20 Step -1
        ws.Cells(f, 8).Value = ws.Cells(f - 1, 8).Value
    Next f
    
    ' Insertar nuevo mensaje en la primera linea
    ws.Cells(19, 8).Value = "[" & Format$(ModCPU.ContadorCiclos, "000") & "] " & msg
    
    ' Guardar tambien en hoja de Log completa si existe
    Set wsLog = ThisWorkbook.Worksheets("Log")
    If wsLog Is Nothing Then Set wsLog = ActiveWorkbook.Worksheets("Log")
    If Not wsLog Is Nothing Then
        Dim proxFila As Long
        proxFila = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
        wsLog.Cells(proxFila, 1).Value = ModCPU.ContadorCiclos
        wsLog.Cells(proxFila, 2).Value = ModCPU.FaseActual
        wsLog.Cells(proxFila, 3).Value = msg
        wsLog.Cells(proxFila, 4).Value = "0x" & ModMemoria.Hex2(ModCPU.PC)
        wsLog.Cells(proxFila, 5).Value = "0x" & ModMemoria.Hex2(ModCPU.AX)
        wsLog.Cells(proxFila, 6).Value = "0x" & ModMemoria.Hex2(ModCPU.BX)
        wsLog.Cells(proxFila, 7).Value = "ZF=" & ModCPU.ZF & " CF=" & ModCPU.CF & " SF=" & ModCPU.SF
    End If
End Sub

' ==============================================================================
'  INSPECTOR DE MEMORIA
' ==============================================================================

Public Sub MostrarInspectorMemoria(ByVal dir As Long, ByVal val As Long)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = HojaSim()
    If ws Is Nothing Then Exit Sub
    
    dir = dir And &HFF
    val = val And &HFF
    
    Dim segm As String
    If dir < 128 Then segm = "SEGMENTO DE CODIGO (CS)" Else segm = "SEGMENTO DE DATOS (DS)"
    
    Dim valConSigno As Long
    If val >= 128 Then valConSigno = val - 256 Else valConSigno = val
    
    ' Celdas del panel inspector en la hoja: R24 a AG25
    ws.Cells(24, 18).Value = "0x" & ModMemoria.Hex2(dir) & " (" & dir & "d)"
    ws.Cells(24, 21).Value = "0x" & ModMemoria.Hex2(val)
    ws.Cells(24, 24).Value = ModMemoria.Bin8(val) & "b"
    ws.Cells(24, 28).Value = val & " / " & valConSigno
    ws.Cells(24, 31).Value = segm
    ws.Cells(25, 18).Value = "Como Instruccion: " & ModCPU.DesensamblarOpcode(val, 0)
End Sub
