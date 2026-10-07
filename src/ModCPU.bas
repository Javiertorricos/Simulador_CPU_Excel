Attribute VB_Name = "ModCPU"
Option Explicit

' ==============================================================================
'  UNIVERSIDAD CATOLICA BOLIVIANA "SAN PABLO"
'  Materia: Arquitectura de Computadoras (SIS-131)
'  Proyecto: Simulador de CPU von Neumann (8 bits) y Memoria Principal
'  Autor: Javier Torrico Sejas
'
'  MODULO: ModCPU
'  Descripcion:
'    Modela los Registros del CPU, Banderas (Flags) y Estado de la Unidad de Control.
'    Todos los registros son de 8 bits (0 a 255).
' ==============================================================================

' --- REGISTROS PRINCIPALES DEL CPU (Visibles segun enunciado) ---
Public PC As Long           ' Program Counter: Puntero de 8 bits a la sig. instruccion
Public IR_Opcode As Long    ' Instruction Register (Opcode): Codigo de operacion actual
Public IR_Operando As Long  ' Instruction Register (Operando): Inmediato o direccion
Public MAR As Long          ' Memory Address Register: Registro de Direcciones hacia bus RAM
Public MDR As Long          ' Memory Data Register: Registro de Datos desde/hacia bus RAM
Public AX As Long           ' Acumulador (8 bits): Registro principal para operaciones ALU
Public BX As Long           ' Registro de proposito general (8 bits): Base/auxiliar

' --- BANDERAS DE ESTADO (Flags de 1 bit generadas por la ALU) ---
Public ZF As Long           ' Zero Flag: 1 si el resultado de la ultima operacion fue 0
Public CF As Long           ' Carry Flag: 1 si ocurrio desbordamiento sin signo (acarreo/prestamo)
Public SF As Long           ' Sign Flag: 1 si el bit 7 (MSB) es 1 (resultado negativo en comp. a 2)

' --- ESTADO DE CONTROL Y DEL RELOJ ---
Public FaseActual As String           ' "FETCH" | "DECODE" | "EXECUTE" | "STORE"
Public NumPasoCiclo As Long          ' 1=Fetch, 2=Decode, 3=Execute, 4=Store
Public EstadoCPU As String            ' "LISTO" | "EJECUTANDO" | "PAUSADO" | "HALT"
Public ContadorCiclos As Long         ' Numero de micro-operaciones / ciclos de reloj ejecutados
Public ContadorInstrucciones As Long  ' Numero de instrucciones completadas
Public Halted As Boolean              ' Indica si la CPU alcanzo la instruccion HLT
Public EnEjecucion As Boolean         ' Indica si el modo continuo RUN esta activo
Public SolicitarPausa As Boolean      ' Bandera para solicitar pausa desde la interfaz

''' <summary>
''' Reinicia todos los registros, banderas y contadores del CPU a sus valores iniciales.
''' </summary>
Public Sub CPU_Reset()
    PC = 0
    IR_Opcode = 0
    IR_Operando = 0
    MAR = 0
    MDR = 0
    AX = 0
    BX = 0
    
    ZF = 0
    CF = 0
    SF = 0
    
    FaseActual = "LISTO"
    NumPasoCiclo = 1
    EstadoCPU = "LISTO"
    ContadorCiclos = 0
    ContadorInstrucciones = 0
    Halted = False
    EnEjecucion = False
    SolicitarPausa = False
    
    ' Refrescar interfaz visual
    Call CPU_ActualizarUI
    Call ModInterfaz.ActualizarFaseUI("LISTO")
    Call ModInterfaz.LogMensaje("CPU reiniciado: PC=00h, Registros=00h, Banderas=0")
End Sub

''' <summary>
''' Actualiza todos los valores de los registros y banderas en las celdas de la hoja Excel.
''' </summary>
Public Sub CPU_ActualizarUI()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ModInterfaz.HojaSim()
    If ws Is Nothing Then Exit Sub
    
    ' Registros Principales (Hexadecimal) - Se respetan las etiquetas descriptivas de la fila 9
    ' PC: H8 (Hex)
    ws.Range("H8").Value = "0x" & ModMemoria.Hex2(PC)
    
    ' IR: J8 (Opcode Hex + Mnem)
    ws.Range("J8").Value = "0x" & ModMemoria.Hex2(IR_Opcode) & " [" & DesensamblarOpcode(IR_Opcode, IR_Operando) & "]"
    
    ' MAR: L8 (Hex)
    ws.Range("L8").Value = "0x" & ModMemoria.Hex2(MAR)
    
    ' MDR: N8 (Hex)
    ws.Range("N8").Value = "0x" & ModMemoria.Hex2(MDR)
    
    ' AX (Acumulador): H12 (Hex), H13 (Dec/Bin)
    ws.Range("H12").Value = "0x" & ModMemoria.Hex2(AX)
    ws.Range("H13").Value = AX & " (" & ModMemoria.Bin8(AX) & "b)"
    
    ' BX: J12 (Hex), J13 (Dec/Bin)
    ws.Range("J12").Value = "0x" & ModMemoria.Hex2(BX)
    ws.Range("J13").Value = BX & " (" & ModMemoria.Bin8(BX) & "b)"
    
    ' Banderas de Estado (Flags): L12, M12, N12
    ws.Range("L12").Value = ZF
    ws.Range("M12").Value = CF
    ws.Range("N12").Value = SF
    
    ' Formato visual de las banderas (ilumina en verde si es 1, gris si es 0)
    Call ModInterfaz.ActualizarFlagsUI(ZF, CF, SF)
    
    ' Contadores de estado: P8 (Ciclos), P9 (Instrucciones), O12 (Estado CPU celda combinada O12:P12)
    ws.Range("P8").Value = ContadorCiclos & " Ciclos"
    ws.Range("P9").Value = ContadorInstrucciones & " Instr."
    ws.Range("O12").Value = EstadoCPU
End Sub

''' <summary>
''' Convierte un Opcode y Operando a texto nemónico legible.
''' </summary>
Public Function DesensamblarOpcode(ByVal op As Long, ByVal arg As Long) As String
    Select Case op
        Case &H0: DesensamblarOpcode = "HLT"
        Case &H1: DesensamblarOpcode = "NOP"
        
        Case &H10: DesensamblarOpcode = "MOV AX, 0x" & Hex2(arg)
        Case &H11: DesensamblarOpcode = "MOV BX, 0x" & Hex2(arg)
        Case &H12: DesensamblarOpcode = "MOV AX, BX"
        Case &H13: DesensamblarOpcode = "MOV BX, AX"
        
        Case &H20: DesensamblarOpcode = "LOAD AX, [0x" & Hex2(arg) & "]"
        Case &H21: DesensamblarOpcode = "LOAD BX, [0x" & Hex2(arg) & "]"
        Case &H22: DesensamblarOpcode = "STORE [0x" & Hex2(arg) & "], AX"
        Case &H23: DesensamblarOpcode = "STORE [0x" & Hex2(arg) & "], BX"
        
        Case &H30: DesensamblarOpcode = "ADD AX, 0x" & Hex2(arg)
        Case &H31: DesensamblarOpcode = "ADD BX, 0x" & Hex2(arg)
        Case &H32: DesensamblarOpcode = "ADD AX, BX"
        Case &H33: DesensamblarOpcode = "ADD BX, AX"
        
        Case &H40: DesensamblarOpcode = "SUB AX, 0x" & Hex2(arg)
        Case &H41: DesensamblarOpcode = "SUB BX, 0x" & Hex2(arg)
        Case &H42: DesensamblarOpcode = "SUB AX, BX"
        Case &H43: DesensamblarOpcode = "SUB BX, AX"
        
        Case &H50: DesensamblarOpcode = "CMP AX, 0x" & Hex2(arg)
        Case &H51: DesensamblarOpcode = "CMP BX, 0x" & Hex2(arg)
        Case &H52: DesensamblarOpcode = "CMP AX, BX"
        Case &H53: DesensamblarOpcode = "CMP BX, AX"
        
        Case &H60: DesensamblarOpcode = "INC AX"
        Case &H61: DesensamblarOpcode = "INC BX"
        Case &H62: DesensamblarOpcode = "DEC AX"
        Case &H63: DesensamblarOpcode = "DEC BX"
        Case &H64: DesensamblarOpcode = "NOT AX"
        Case &H65: DesensamblarOpcode = "NOT BX"
        
        Case &H70: DesensamblarOpcode = "AND AX, 0x" & Hex2(arg)
        Case &H71: DesensamblarOpcode = "AND BX, 0x" & Hex2(arg)
        Case &H72: DesensamblarOpcode = "AND AX, BX"
        Case &H73: DesensamblarOpcode = "AND BX, AX"
        
        Case &H80: DesensamblarOpcode = "OR AX, 0x" & Hex2(arg)
        Case &H81: DesensamblarOpcode = "OR BX, 0x" & Hex2(arg)
        Case &H82: DesensamblarOpcode = "OR AX, BX"
        Case &H83: DesensamblarOpcode = "OR BX, AX"
        
        Case &H90: DesensamblarOpcode = "XOR AX, 0x" & Hex2(arg)
        Case &H91: DesensamblarOpcode = "XOR BX, 0x" & Hex2(arg)
        Case &H92: DesensamblarOpcode = "XOR AX, BX"
        Case &H93: DesensamblarOpcode = "XOR BX, AX"
        
        Case &HA0: DesensamblarOpcode = "JMP 0x" & Hex2(arg)
        Case &HA1: DesensamblarOpcode = "JZ 0x" & Hex2(arg)
        Case &HA2: DesensamblarOpcode = "JNZ 0x" & Hex2(arg)
        
        Case Else: DesensamblarOpcode = "UNKNOWN (0x" & Hex2(op) & ")"
    End Select
End Function
