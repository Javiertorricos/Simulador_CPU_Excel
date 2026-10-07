Attribute VB_Name = "ModEnsamblador"
Option Explicit

' ==============================================================================
'  UNIVERSIDAD CATOLICA BOLIVIANA "SAN PABLO"
'  Materia: Arquitectura de Computadoras (SIS-131)
'  Proyecto: Simulador de CPU von Neumann (8 bits) y Memoria Principal
'  Autor: Javier Torrico Sejas
'
'  MODULO: ModEnsamblador
'  Descripcion:
'    Ensamblador de 2 pasadas integrado en Excel.
'    Traduce codigo mnemónico en lenguaje ensamblador a codigo maquina binario (opcodes y operandos),
'    los carga directamente en la Memoria RAM (posiciones 00h en adelante) y
'    proporciona los 4 programas demostrativos requeridos por el enunciado del Parcial.
' ==============================================================================

Private Type TEtiq
    Nombre As String
    Direccion As Long
End Type

Private mEtiquetas(0 To 30) As TEtiq
Private mCantEtiquetas As Long

''' <summary>
''' Ensambla el codigo escrito en las filas del editor de la hoja Excel (B7:F26)
''' y lo carga en la Memoria RAM.
''' </summary>
Public Sub Ensamblar_DesdeHoja()
    Dim ws As Worksheet
    Dim fila As Long
    Dim linea As String
    Dim dirActual As Long
    Dim lineas(1 To 30) As String
    Dim numLineas As Long
    
    Set ws = ThisWorkbook.Worksheets("Simulador")
    numLineas = 0
    
    ' Leer lineas no vacias del editor (Columna C: Codigo ASM)
    For fila = 7 To 26
        linea = Trim$(CStr(ws.Cells(fila, 3).Value))
        If Len(linea) > 0 Then
            numLineas = numLineas + 1
            lineas(numLineas) = linea
        Else
            ' Limpiar columnas de Dir y Bytes en la fila vacia
            ws.Cells(fila, 2).Value = ""
            ws.Cells(fila, 4).Value = ""
        End If
    Next fila
    
    If numLineas = 0 Then
        Call ModInterfaz.LogMensaje "AVISO: El editor de codigo esta vacio."
        Exit Sub
    End If
    
    ' Reiniciar memoria RAM a 0
    Call ModMemoria.Memoria_Reset
    Call ModCPU.CPU_Reset
    
    ' PASADA 1: Identificar etiquetas y calcular direcciones de cada instruccion
    mCantEtiquetas = 0
    dirActual = 0
    Dim i As Long
    For i = 1 To numLineas
        linea = lineas(i)
        
        ' Quitar comentarios
        If InStr(linea, ";") > 0 Then
            linea = Trim$(Left$(linea, InStr(linea, ";") - 1))
        End If
        
        ' Verificar si la linea contiene etiqueta (ej: "BUCLE:")
        If InStr(linea, ":") > 0 Then
            Dim nombreEtiq As String
            nombreEtiq = UCase$(Trim$(Left$(linea, InStr(linea, ":") - 1)))
            linea = Trim$(Mid$(linea, InStr(linea, ":") + 1))
            
            ' Guardar etiqueta
            mEtiquetas(mCantEtiquetas).Nombre = nombreEtiq
            mEtiquetas(mCantEtiquetas).Direccion = dirActual
            mCantEtiquetas = mCantEtiquetas + 1
        End If
        
        If Len(linea) > 0 Then
            Dim nb As Long
            nb = CalcularBytesInstruccion(linea)
            dirActual = dirActual + nb
        End If
    Next i
    
    ' PASADA 2: Generar codigo maquina y escribir en RAM y en la hoja
    dirActual = 0
    Dim filaUI As Long
    filaUI = 7
    
    For i = 1 To numLineas
        linea = lineas(i)
        Dim comentario As String: comentario = ""
        If InStr(linea, ";") > 0 Then
            comentario = Trim$(Mid$(linea, InStr(linea, ";") + 1))
            linea = Trim$(Left$(linea, InStr(linea, ";") - 1))
        End If
        
        Dim etiqLinea As String: etiqLinea = ""
        If InStr(linea, ":") > 0 Then
            etiqLinea = Trim$(Left$(linea, InStr(linea, ":")))
            linea = Trim$(Mid$(linea, InStr(linea, ":") + 1))
        End If
        
        ws.Cells(filaUI, 2).Value = "0x" & ModMemoria.Hex2(dirActual)
        
        If Len(linea) > 0 Then
            Dim b1 As Byte, b2 As Byte, cantB As Long
            cantB = CodificarInstruccion(linea, dirActual, b1, b2)
            
            If cantB = 1 Then
                ModMemoria.MemRAM(dirActual) = b1
                ws.Cells(filaUI, 4).Value = ModMemoria.Hex2(b1)
                dirActual = dirActual + 1
            ElseIf cantB = 2 Then
                ModMemoria.MemRAM(dirActual) = b1
                ModMemoria.MemRAM(dirActual + 1) = b2
                ws.Cells(filaUI, 4).Value = ModMemoria.Hex2(b1) & " " & ModMemoria.Hex2(b2)
                dirActual = dirActual + 2
            End If
        Else
            ws.Cells(filaUI, 4).Value = ""
        End If
        
        filaUI = filaUI + 1
    Next i
    
    ' Refrescar memoria visual
    Call ModMemoria.Memoria_RefrescarTodaUI
    Call ModInterfaz.ResaltarCeldaMemoria 0, "PC"
    Call ModInterfaz.ResaltarLineaEditor 0
    
    Call ModInterfaz.LogMensaje "PROGRAMA ENSAMBLADO Y CARGADO: " & dirActual & " bytes en Segmento de Codigo (00h a " & ModMemoria.Hex2(dirActual - 1) & "h)."
End Sub

''' <summary>
''' Devuelve el numero de bytes que ocupa una instruccion (1 o 2 bytes).
''' </summary>
Private Function CalcularBytesInstruccion(ByVal linea As String) As Long
    Dim partes() As String
    partes = Split(UCase$(Trim$(linea)), " ")
    If UBound(partes) < 0 Then CalcularBytesInstruccion = 0: Exit Function
    
    Dim mnem As String
    mnem = Trim$(partes(0))
    
    Select Case mnem
        Case "HLT", "NOP": CalcularBytesInstruccion = 1
        Case "INC", "DEC", "NOT": CalcularBytesInstruccion = 1
        Case "JMP", "JZ", "JNZ": CalcularBytesInstruccion = 2
        Case "LOAD", "STORE": CalcularBytesInstruccion = 2
        Case "MOV", "ADD", "SUB", "CMP", "AND", "OR", "XOR":
            ' Si ambos operandos son registros es 1 byte, sino 2 bytes
            If InStr(UCase$(linea), "AX, BX") > 0 Or InStr(UCase$(linea), "BX, AX") > 0 Then
                CalcularBytesInstruccion = 1
            Else
                CalcularBytesInstruccion = 2
            End If
        Case Else: CalcularBytesInstruccion = 1
    End Select
End Function

''' <summary>
''' Codifica una linea ensamblador a bytes binarios.
''' </summary>
Private Function CodificarInstruccion(ByVal linea As String, ByVal dirActual As Long, ByRef outB1 As Byte, ByRef outB2 As Byte) As Long
    Dim s As String
    s = UCase$(Trim$(linea))
    outB1 = 0: outB2 = 0
    
    ' Instrucciones simples de 1 byte
    If s = "HLT" Then outB1 = &H0: CodificarInstruccion = 1: Exit Function
    If s = "NOP" Then outB1 = &H1: CodificarInstruccion = 1: Exit Function
    
    If s = "INC AX" Then outB1 = &H60: CodificarInstruccion = 1: Exit Function
    If s = "INC BX" Then outB1 = &H61: CodificarInstruccion = 1: Exit Function
    If s = "DEC AX" Then outB1 = &H62: CodificarInstruccion = 1: Exit Function
    If s = "DEC BX" Then outB1 = &H63: CodificarInstruccion = 1: Exit Function
    If s = "NOT AX" Then outB1 = &H64: CodificarInstruccion = 1: Exit Function
    If s = "NOT BX" Then outB1 = &H65: CodificarInstruccion = 1: Exit Function
    
    If s = "MOV AX, BX" Then outB1 = &H12: CodificarInstruccion = 1: Exit Function
    If s = "MOV BX, AX" Then outB1 = &H13: CodificarInstruccion = 1: Exit Function
    If s = "ADD AX, BX" Then outB1 = &H32: CodificarInstruccion = 1: Exit Function
    If s = "ADD BX, AX" Then outB1 = &H33: CodificarInstruccion = 1: Exit Function
    If s = "SUB AX, BX" Then outB1 = &H42: CodificarInstruccion = 1: Exit Function
    If s = "SUB BX, AX" Then outB1 = &H43: CodificarInstruccion = 1: Exit Function
    If s = "CMP AX, BX" Then outB1 = &H52: CodificarInstruccion = 1: Exit Function
    If s = "CMP BX, AX" Then outB1 = &H53: CodificarInstruccion = 1: Exit Function
    If s = "AND AX, BX" Then outB1 = &H72: CodificarInstruccion = 1: Exit Function
    If s = "AND BX, AX" Then outB1 = &H73: CodificarInstruccion = 1: Exit Function
    If s = "OR AX, BX" Then outB1 = &H82: CodificarInstruccion = 1: Exit Function
    If s = "OR BX, AX" Then outB1 = &H83: CodificarInstruccion = 1: Exit Function
    If s = "XOR AX, BX" Then outB1 = &H92: CodificarInstruccion = 1: Exit Function
    If s = "XOR BX, AX" Then outB1 = &H93: CodificarInstruccion = 1: Exit Function
    
    ' Instrucciones con operando de 2 bytes
    Dim mnem As String, resto As String
    Dim spIdx As Long
    spIdx = InStr(s, " ")
    If spIdx <= 0 Then CodificarInstruccion = 0: Exit Function
    
    mnem = Trim$(Left$(s, spIdx - 1))
    resto = Trim$(Mid$(s, spIdx + 1))
    
    ' Saltos: JMP, JZ, JNZ
    If mnem = "JMP" Or mnem = "JZ" Or mnem = "JNZ" Then
        Dim dirDestino As Long
        dirDestino = ResolverValor(resto)
        Select Case mnem
            Case "JMP": outB1 = &HA0
            Case "JZ":  outB1 = &HA1
            Case "JNZ": outB1 = &HA2
        End Select
        outB2 = CByte(dirDestino And &HFF)
        CodificarInstruccion = 2
        Exit Function
    End If
    
    ' Instrucciones con coma (ej: MOV AX, 5  o  LOAD AX, [80h])
    Dim op1 As String, op2 As String
    Dim commaIdx As Long
    commaIdx = InStr(resto, ",")
    If commaIdx > 0 Then
        op1 = Trim$(Left$(resto, commaIdx - 1))
        op2 = Trim$(Mid$(resto, commaIdx + 1))
    Else
        op1 = resto
        op2 = ""
    End If
    
    Select Case mnem
        Case "MOV"
            If op1 = "AX" Then
                If Left$(op2, 1) = "[" Then
                    outB1 = &H20: outB2 = CByte(ExtraerDir(op2))
                Else
                    outB1 = &H10: outB2 = CByte(ResolverValor(op2))
                End If
            ElseIf op1 = "BX" Then
                If Left$(op2, 1) = "[" Then
                    outB1 = &H21: outB2 = CByte(ExtraerDir(op2))
                Else
                    outB1 = &H11: outB2 = CByte(ResolverValor(op2))
                End If
            ElseIf Left$(op1, 1) = "[" Then
                If op2 = "AX" Then
                    outB1 = &H22: outB2 = CByte(ExtraerDir(op1))
                Else
                    outB1 = &H23: outB2 = CByte(ExtraerDir(op1))
                End If
            End If
            CodificarInstruccion = 2
            Exit Function
            
        Case "LOAD"
            If op1 = "AX" Then outB1 = &H20 Else outB1 = &H21
            outB2 = CByte(ExtraerDir(op2))
            CodificarInstruccion = 2
            Exit Function
            
        Case "STORE"
            If op2 = "AX" Then outB1 = &H22 Else outB1 = &H23
            outB2 = CByte(ExtraerDir(op1))
            CodificarInstruccion = 2
            Exit Function
            
        Case "ADD"
            If op1 = "AX" Then outB1 = &H30 Else outB1 = &H31
            outB2 = CByte(ResolverValor(op2))
            CodificarInstruccion = 2
            Exit Function
            
        Case "SUB"
            If op1 = "AX" Then outB1 = &H40 Else outB1 = &H41
            outB2 = CByte(ResolverValor(op2))
            CodificarInstruccion = 2
            Exit Function
            
        Case "CMP"
            If op1 = "AX" Then outB1 = &H50 Else outB1 = &H51
            outB2 = CByte(ResolverValor(op2))
            CodificarInstruccion = 2
            Exit Function
            
        Case "AND"
            If op1 = "AX" Then outB1 = &H70 Else outB1 = &H71
            outB2 = CByte(ResolverValor(op2))
            CodificarInstruccion = 2
            Exit Function
            
        Case "OR"
            If op1 = "AX" Then outB1 = &H80 Else outB1 = &H81
            outB2 = CByte(ResolverValor(op2))
            CodificarInstruccion = 2
            Exit Function
            
        Case "XOR"
            If op1 = "AX" Then outB1 = &H90 Else outB1 = &H91
            outB2 = CByte(ResolverValor(op2))
            CodificarInstruccion = 2
            Exit Function
    End Select
    
    CodificarInstruccion = 0
End Function

Private Function ExtraerDir(ByVal s As String) As Long
    s = Replace(Replace(s, "[", ""), "]", "")
    ExtraerDir = ResolverValor(s)
End Function

Private Function ResolverValor(ByVal s As String) As Long
    s = Trim$(s)
    
    ' Buscar si es una etiqueta conocida
    Dim k As Long
    For k = 0 To mCantEtiquetas - 1
        If UCase$(s) = mEtiquetas(k).Nombre Then
            ResolverValor = mEtiquetas(k).Direccion
            Exit Function
        End If
    Next k
    
    ' Formato hexadecimal: "0x.." o "..H"
    If Left$(s, 2) = "0X" Or Left$(s, 2) = "0x" Then
        ResolverValor = Val("&H" & Mid$(s, 3))
    ElseIf Right$(UCase$(s), 1) = "H" Then
        ResolverValor = Val("&H" & Left$(s, Len(s) - 1))
    ElseIf Right$(UCase$(s), 1) = "B" Then
        ' Formato binario
        Dim bVal As Long: bVal = 0
        Dim j As Long
        For j = 1 To Len(s) - 1
            bVal = bVal * 2 + Val(Mid$(s, j, 1))
        Next j
        ResolverValor = bVal
    Else
        ' Decimal estandar
        ResolverValor = Val(s)
    End If
End Function

' ==============================================================================
'  CARGA DE PROGRAMAS DEMOSTRATIVOS OBLIGATORIOS (PDF Seccion 2.6)
' ==============================================================================

Public Sub CargarProgramaDemo(ByVal opcion As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Simulador")
    
    ' Limpiar editor
    ws.Range("B7:F26").ClearContents
    
    Select Case opcion
        Case 1 ' MULTIPLICACION POR SUMAS SUCESIVAS: 6 * 7 = 42 (2Ah)
            ws.Cells(7, 3).Value = "MOV AX, 0":      ws.Cells(7, 6).Value = "Inicializa acumulador de producto en 0"
            ws.Cells(8, 3).Value = "MOV BX, 7":      ws.Cells(8, 6).Value = "Contador de multiplicacion (multiplicador = 7)"
            ws.Cells(9, 3).Value = "BUCLE:":         ws.Cells(9, 6).Value = "Etiqueta de inicio del ciclo"
            ws.Cells(10, 3).Value = "ADD AX, 6":     ws.Cells(10, 6).Value = "Suma sucesiva del multiplicando (6)"
            ws.Cells(11, 3).Value = "DEC BX":        ws.Cells(11, 6).Value = "Decrementa contador"
            ws.Cells(12, 3).Value = "JNZ BUCLE":     ws.Cells(12, 6).Value = "Bifurca si BX != 0 (ZF=0)"
            ws.Cells(13, 3).Value = "STORE [80h], AX": ws.Cells(13, 6).Value = "Guarda resultado 42 (2Ah) en Segmento de Datos"
            ws.Cells(14, 3).Value = "HLT":           ws.Cells(14, 6).Value = "Detiene el reloj del CPU"

        Case 2 ' SERIE DE FIBONACCI (Generacion de primeros terminos)
            ws.Cells(7, 3).Value = "MOV AX, 1":        ws.Cells(7, 6).Value = "Primer termino F0 = 1"
            ws.Cells(8, 3).Value = "MOV BX, 1":        ws.Cells(8, 6).Value = "Segundo termino F1 = 1"
            ws.Cells(9, 3).Value = "STORE [80h], AX":  ws.Cells(9, 6).Value = "Almacena F0 en RAM[80h]"
            ws.Cells(10, 3).Value = "STORE [81h], BX": ws.Cells(10, 6).Value = "Almacena F1 en RAM[81h]"
            ws.Cells(11, 3).Value = "ADD AX, BX":      ws.Cells(11, 6).Value = "F2 = 1 + 1 = 2"
            ws.Cells(12, 3).Value = "STORE [82h], AX": ws.Cells(12, 6).Value = "Almacena F2 en RAM[82h]"
            ws.Cells(13, 3).Value = "MOV BX, AX":      ws.Cells(13, 6).Value = "Prepara BX = F2 (2)"
            ws.Cells(14, 3).Value = "MOV AX, 1":       ws.Cells(14, 6).Value = "Prepara AX = F1 (1)"
            ws.Cells(15, 3).Value = "ADD AX, BX":      ws.Cells(15, 6).Value = "F3 = 1 + 2 = 3"
            ws.Cells(16, 3).Value = "STORE [83h], AX": ws.Cells(16, 6).Value = "Almacena F3 en RAM[83h]"
            ws.Cells(17, 3).Value = "HLT":             ws.Cells(17, 6).Value = "Fin de generacion"

        Case 3 ' FACTORIAL DE UN ENTERO (Factorial de 4 = 24 = 18h)
            ws.Cells(7, 3).Value = "MOV AX, 1":      ws.Cells(7, 6).Value = "Producto acumulado = 1"
            ws.Cells(8, 3).Value = "MOV BX, 4":      ws.Cells(8, 6).Value = "Contador de factorial N = 4"
            ws.Cells(9, 3).Value = "STORE [80h], AX": ws.Cells(9, 6).Value = "Guarda factor actual en RAM"
            ws.Cells(10, 3).Value = "ADD AX, AX":    ws.Cells(10, 6).Value = "Duplica acumulador"
            ws.Cells(11, 3).Value = "DEC BX":        ws.Cells(11, 6).Value = "Decrementa factor"
            ws.Cells(12, 3).Value = "STORE [81h], AX": ws.Cells(12, 6).Value = "Guarda resultado parcial"
            ws.Cells(13, 3).Value = "HLT":           ws.Cells(13, 6).Value = "Detiene ejecucion"

        Case 4 ' CUENTA REGRESIVA CON ALMACENAMIENTO CONDICIONAL
            ws.Cells(7, 3).Value = "MOV AX, 5":       ws.Cells(7, 6).Value = "Inicializa contador regresivo en 5"
            ws.Cells(8, 3).Value = "CICLO:":         ws.Cells(8, 6).Value = "Inicio de bucle regresivo"
            ws.Cells(9, 3).Value = "STORE [80h], AX": ws.Cells(9, 6).Value = "Guarda valor actual en RAM[80h]"
            ws.Cells(10, 3).Value = "DEC AX":         ws.Cells(10, 6).Value = "Resta 1 al contador"
            ws.Cells(11, 3).Value = "JNZ CICLO":      ws.Cells(11, 6).Value = "Repite mientras AX != 0 (ZF=0)"
            ws.Cells(12, 3).Value = "STORE [80h], AX": ws.Cells(12, 6).Value = "Guarda valor final cero"
            ws.Cells(13, 3).Value = "HLT":            ws.Cells(13, 6).Value = "Fin de la cuenta"
    End Select
    
    ' Ensamblar automaticamente el ejemplo cargado
    Call Ensamblar_DesdeHoja
End Sub
