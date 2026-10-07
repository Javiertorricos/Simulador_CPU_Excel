Attribute VB_Name = "ModCicloInstruccion"
Option Explicit

' ==============================================================================
'  UNIVERSIDAD CATOLICA BOLIVIANA "SAN PABLO"
'  Materia: Arquitectura de Computadoras (SIS-131)
'  Proyecto: Simulador de CPU von Neumann (8 bits) y Memoria Principal
'  Autor: Javier Torrico Sejas
'
'  MODULO: ModCicloInstruccion
'  Descripcion:
'    Controla el Ciclo de Instruccion Completo descompuesto en sus 4 Fases:
'      1. FETCH   (Busqueda de la instruccion en memoria RAM hacia IR)
'      2. DECODE  (Decodificacion del Opcode y preparacion de operandos)
'      3. EXECUTE (Ejecucion en la ALU o evaluacion de bifurcaciones)
'      4. STORE   (Almacenamiento/Write-back en registro destino o RAM)
'    Permite ejecucion Paso a Paso (Step) y Modo Continuo (Run) con retardo.
' ==============================================================================

Private mResultadoALU As Long
Private mCondicionSalto As Boolean

' ==============================================================================
'  PUNTOS DE ENTRADA PARA CONTROLES DE LA INTERFAZ
' ==============================================================================

''' <summary>
''' Avanza una unica fase del ciclo de instruccion (Fetch -> Decode -> Execute -> Store).
''' Modo Paso a Paso esencial para la defensa oral.
''' </summary>
Public Sub Paso_Fase()
    If ModCPU.Halted Then
        Call ModInterfaz.LogMensaje "CPU DETENIDO (HLT): Presione REINICIAR."
        Exit Sub
    End If
    
    ModCPU.ContadorCiclos = ModCPU.ContadorCiclos + 1
    
    Select Case ModCPU.NumPasoCiclo
        Case 1: Call Fase_Fetch
        Case 2: Call Fase_Decode
        Case 3: Call Fase_Execute
        Case 4: Call Fase_Store
    End Select
    
    Call ModCPU.CPU_ActualizarUI
End Sub

''' <summary>
''' Ejecuta una instruccion completa (las 4 fases completas hasta terminar la instruccion).
''' </summary>
Public Sub Paso_InstruccionCompleta()
    If ModCPU.Halted Then Exit Sub
    
    Dim faseInicial As Long
    faseInicial = ModCPU.NumPasoCiclo
    
    ' Ejecuta las fases restantes hasta volver a Fetch (fase 1)
    Do
        Call Paso_Fase
        If ModCPU.Halted Then Exit Do
    Loop While ModCPU.NumPasoCiclo <> 1
End Sub

''' <summary>
''' Modo Continuo (Run / Play): Ejecucion secuencial automatica con retardo.
''' </summary>
Public Sub Modo_Continuo()
    If ModCPU.Halted Then
        Call ModInterfaz.LogMensaje "CPU DETENIDO (HLT): Presione REINICIAR antes de ejecutar."
        Exit Sub
    End If
    
    ModCPU.EnEjecucion = True
    ModCPU.SolicitarPausa = False
    ModCPU.EstadoCPU = "EJECUTANDO"
    Call ModCPU.CPU_ActualizarUI
    
    Dim retardoSeg As Double
    retardoSeg = ModCicloInstruccion.ObtenerRetardo()
    
    Do While (Not ModCPU.Halted) And (Not ModCPU.SolicitarPausa)
        Call Paso_Fase
        
        ' Retardo configurable para permitir observar la animacion
        Call ModCicloInstruccion.DormirMs retardoSeg
        DoEvents
    Loop
    
    ModCPU.EnEjecucion = False
    If ModCPU.Halted Then
        ModCPU.EstadoCPU = "HALT"
        Call ModInterfaz.LogMensaje "PROGRAMA FINALIZADO CON EXITO: Instruccion HLT alcanzada."
    Else
        ModCPU.EstadoCPU = "PAUSADO"
        Call ModInterfaz.LogMensaje "Ejecucion pausada por el usuario."
    End If
    
    Call ModCPU.CPU_ActualizarUI
End Sub

''' <summary>
''' Solicita pausar la ejecucion continua al terminar la fase actual.
''' </summary>
Public Sub Pausar_Ejecucion()
    ModCPU.SolicitarPausa = True
    ModCPU.EstadoCPU = "PAUSADO"
    Call ModCPU.CPU_ActualizarUI
End Sub

' ==============================================================================
'  LAS 4 FASES DEL CICLO DE INSTRUCCION
' ==============================================================================

''' <summary>
''' FASE 1: FETCH (Busqueda)
'''  - La direccion en PC se carga en MAR.
'''  - Se lee la RAM hacia MDR usando la primitiva Read(MAR).
'''  - El dato leido pasa al registro de instruccion IR_Opcode.
'''  - Se incrementa el contador de programa PC.
''' </summary>
Public Sub Fase_Fetch()
    ModCPU.FaseActual = "FETCH"
    ModCPU.EstadoCPU = "FETCH"
    Call ModInterfaz.ActualizarFaseUI "FETCH"
    
    ' 1. Transferencia PC -> MAR
    ModCPU.MAR = ModCPU.PC
    Call ModInterfaz.ResaltarBus "DIRECCIONES", ModCPU.PC
    
    ' 2. Lectura primitiva desde memoria RAM: MDR <- Read(MAR)
    ModCPU.MDR = ModMemoria.Read(ModCPU.MAR)
    Call ModInterfaz.ResaltarBus "DATOS", ModCPU.MDR
    
    ' 3. Cargar Opcode en el Registro de Instruccion
    ModCPU.IR_Opcode = ModCPU.MDR
    ModCPU.IR_Operando = 0
    
    ' 4. Incrementar Program Counter: PC <- PC + 1
    Dim pcAnterior As Long
    pcAnterior = ModCPU.PC
    ModCPU.PC = (ModCPU.PC + 1) And &HFF
    
    ' Resaltar celda del nuevo PC en la matriz
    Call ModInterfaz.ResaltarCeldaMemoria ModCPU.PC, "PC"
    
    Call ModInterfaz.LogMensaje "[FETCH] PC=" & Hex2(pcAnterior) & " -> MAR=" & Hex2(ModCPU.MAR) & _
                                " | MDR=Read(" & Hex2(ModCPU.MAR) & ")=" & Hex2(ModCPU.MDR) & _
                                " -> IR_Opcode=" & Hex2(ModCPU.IR_Opcode) & " | PC++ a " & Hex2(ModCPU.PC)
                                
    ModCPU.NumPasoCiclo = 2
End Sub

''' <summary>
''' FASE 2: DECODE (Decodificacion)
'''  - La Unidad de Control interpreta el Opcode en IR.
'''  - Si la instruccion tiene 2 bytes (operando o direccion), realiza la busqueda
'''    del segundo byte en memoria: MAR <- PC, MDR <- Read(MAR), IR_Operando <- MDR, PC++.
''' </summary>
Public Sub Fase_Decode()
    ModCPU.FaseActual = "DECODE"
    ModCPU.EstadoCPU = "DECODE"
    Call ModInterfaz.ActualizarFaseUI "DECODE"
    
    Dim necesitaSegundoByte As Boolean
    necesitaSegundoByte = RequiereSegundoByte(ModCPU.IR_Opcode)
    
    If necesitaSegundoByte Then
        ' Lectura del operando inmediato o direccion de salto / memoria
        ModCPU.MAR = ModCPU.PC
        ModCPU.MDR = ModMemoria.Read(ModCPU.MAR)
        ModCPU.IR_Operando = ModCPU.MDR
        
        Dim pcAnt As Long
        pcAnt = ModCPU.PC
        ModCPU.PC = (ModCPU.PC + 1) And &HFF
        
        Call ModInterfaz.ResaltarCeldaMemoria ModCPU.PC, "PC"
        
        Call ModInterfaz.LogMensaje "[DECODE] Opcode 0x" & Hex2(ModCPU.IR_Opcode) & " (" & _
                                    ModCPU.DesensamblarOpcode(ModCPU.IR_Opcode, ModCPU.IR_Operando) & _
                                    ") | Operando 2do byte: MAR=0x" & Hex2(ModCPU.MAR) & _
                                    ", MDR=0x" & Hex2(ModCPU.MDR) & " | PC++ a 0x" & Hex2(ModCPU.PC)
    Else
        ModCPU.IR_Operando = 0
        Call ModInterfaz.LogMensaje "[DECODE] Opcode 0x" & Hex2(ModCPU.IR_Opcode) & " (" & _
                                    ModCPU.DesensamblarOpcode(ModCPU.IR_Opcode, 0) & _
                                    ") | Instruccion de 1 byte (sin operando en RAM)"
    End If
    
    ModCPU.NumPasoCiclo = 3
End Sub

''' <summary>
''' FASE 3: EXECUTE (Ejecucion)
'''  - La ALU procesa la operacion aritmetico-logica (ADD, SUB, INC, DEC, AND, OR, XOR, NOT, CMP).
'''  - O se evaluan las condiciones de salto (JMP, JZ, JNZ).
'''  - O se preparan las transferencias de memoria (LOAD / STORE).
''' </summary>
Public Sub Fase_Execute()
    ModCPU.FaseActual = "EXECUTE"
    ModCPU.EstadoCPU = "EXECUTE"
    Call ModInterfaz.ActualizarFaseUI "EXECUTE"
    
    mCondicionSalto = False
    mResultadoALU = 0
    
    Select Case ModCPU.IR_Opcode
        ' --- CONTROL ---
        Case &H0 ' HLT
            ModCPU.Halted = True
            Call ModInterfaz.LogMensaje "[EXECUTE] HLT detectado: Deteniendo reloj del CPU."
            
        Case &H1 ' NOP
            Call ModInterfaz.LogMensaje "[EXECUTE] NOP: No operacion."

        ' --- TRANSFERENCIA DE DATOS (MOV) ---
        Case &H10 ' MOV AX, imm
            mResultadoALU = ModCPU.IR_Operando
            Call ModInterfaz.LogMensaje "[EXECUTE] MOV: Preparando carga de inmediato 0x" & Hex2(mResultadoALU) & " en AX"
            
        Case &H11 ' MOV BX, imm
            mResultadoALU = ModCPU.IR_Operando
            Call ModInterfaz.LogMensaje "[EXECUTE] MOV: Preparando carga de inmediato 0x" & Hex2(mResultadoALU) & " en BX"
            
        Case &H12 ' MOV AX, BX
            mResultadoALU = ModCPU.BX
            Call ModInterfaz.LogMensaje "[EXECUTE] MOV: Copiando valor de BX (0x" & Hex2(mResultadoALU) & ") hacia AX"
            
        Case &H13 ' MOV BX, AX
            mResultadoALU = ModCPU.AX
            Call ModInterfaz.LogMensaje "[EXECUTE] MOV: Copiando valor de AX (0x" & Hex2(mResultadoALU) & ") hacia BX"

        ' --- ACCESO A MEMORIA (LOAD / STORE) ---
        Case &H20, &H21 ' LOAD AX/BX, [dir]
            ModCPU.MAR = ModCPU.IR_Operando
            ModCPU.MDR = ModMemoria.Read(ModCPU.MAR)
            mResultadoALU = ModCPU.MDR
            Call ModInterfaz.LogMensaje "[EXECUTE] LOAD: Leyendo RAM[0x" & Hex2(ModCPU.MAR) & "] = 0x" & Hex2(ModCPU.MDR)
            
        Case &H22 ' STORE [dir], AX
            ModCPU.MAR = ModCPU.IR_Operando
            ModCPU.MDR = ModCPU.AX
            Call ModInterfaz.LogMensaje "[EXECUTE] STORE: Preparando escritura de AX (0x" & Hex2(ModCPU.AX) & ") en RAM[0x" & Hex2(ModCPU.MAR) & "]"
            
        Case &H23 ' STORE [dir], BX
            ModCPU.MAR = ModCPU.IR_Operando
            ModCPU.MDR = ModCPU.BX
            Call ModInterfaz.LogMensaje "[EXECUTE] STORE: Preparando escritura de BX (0x" & Hex2(ModCPU.BX) & ") en RAM[0x" & Hex2(ModCPU.MAR) & "]"

        ' --- ARITMETICA Y LOGICA (ALU) ---
        Case &H30: mResultadoALU = ModALU.ALU_Operar("ADD", ModCPU.AX, ModCPU.IR_Operando)
        Case &H31: mResultadoALU = ModALU.ALU_Operar("ADD", ModCPU.BX, ModCPU.IR_Operando)
        Case &H32: mResultadoALU = ModALU.ALU_Operar("ADD", ModCPU.AX, ModCPU.BX)
        Case &H33: mResultadoALU = ModALU.ALU_Operar("ADD", ModCPU.BX, ModCPU.AX)
        
        Case &H40: mResultadoALU = ModALU.ALU_Operar("SUB", ModCPU.AX, ModCPU.IR_Operando)
        Case &H41: mResultadoALU = ModALU.ALU_Operar("SUB", ModCPU.BX, ModCPU.IR_Operando)
        Case &H42: mResultadoALU = ModALU.ALU_Operar("SUB", ModCPU.AX, ModCPU.BX)
        Case &H43: mResultadoALU = ModALU.ALU_Operar("SUB", ModCPU.BX, ModCPU.AX)
        
        Case &H50: mResultadoALU = ModALU.ALU_Operar("CMP", ModCPU.AX, ModCPU.IR_Operando)
        Case &H51: mResultadoALU = ModALU.ALU_Operar("CMP", ModCPU.BX, ModCPU.IR_Operando)
        Case &H52: mResultadoALU = ModALU.ALU_Operar("CMP", ModCPU.AX, ModCPU.BX)
        Case &H53: mResultadoALU = ModALU.ALU_Operar("CMP", ModCPU.BX, ModCPU.AX)
        
        Case &H60: mResultadoALU = ModALU.ALU_Operar("INC", ModCPU.AX, 0)
        Case &H61: mResultadoALU = ModALU.ALU_Operar("INC", ModCPU.BX, 0)
        Case &H62: mResultadoALU = ModALU.ALU_Operar("DEC", ModCPU.AX, 0)
        Case &H63: mResultadoALU = ModALU.ALU_Operar("DEC", ModCPU.BX, 0)
        Case &H64: mResultadoALU = ModALU.ALU_Operar("NOT", ModCPU.AX, 0)
        Case &H65: mResultadoALU = ModALU.ALU_Operar("NOT", ModCPU.BX, 0)
        
        Case &H70: mResultadoALU = ModALU.ALU_Operar("AND", ModCPU.AX, ModCPU.IR_Operando)
        Case &H71: mResultadoALU = ModALU.ALU_Operar("AND", ModCPU.BX, ModCPU.IR_Operando)
        Case &H72: mResultadoALU = ModALU.ALU_Operar("AND", ModCPU.AX, ModCPU.BX)
        Case &H73: mResultadoALU = ModALU.ALU_Operar("AND", ModCPU.BX, ModCPU.AX)
        
        Case &H80: mResultadoALU = ModALU.ALU_Operar("OR", ModCPU.AX, ModCPU.IR_Operando)
        Case &H81: mResultadoALU = ModALU.ALU_Operar("OR", ModCPU.BX, ModCPU.IR_Operando)
        Case &H82: mResultadoALU = ModALU.ALU_Operar("OR", ModCPU.AX, ModCPU.BX)
        Case &H83: mResultadoALU = ModALU.ALU_Operar("OR", ModCPU.BX, ModCPU.AX)
        
        Case &H90: mResultadoALU = ModALU.ALU_Operar("XOR", ModCPU.AX, ModCPU.IR_Operando)
        Case &H91: mResultadoALU = ModALU.ALU_Operar("XOR", ModCPU.BX, ModCPU.IR_Operando)
        Case &H92: mResultadoALU = ModALU.ALU_Operar("XOR", ModCPU.AX, ModCPU.BX)
        Case &H93: mResultadoALU = ModALU.ALU_Operar("XOR", ModCPU.BX, ModCPU.AX)

        ' --- CONTROL DE FLUJO (SALTOS) ---
        Case &HA0 ' JMP dir
            mCondicionSalto = True
            Call ModInterfaz.LogMensaje "[EXECUTE] JMP: Salto incondicional hacia 0x" & Hex2(ModCPU.IR_Operando)
            
        Case &HA1 ' JZ dir (Salto si Zero=1)
            If ModCPU.ZF = 1 Then
                mCondicionSalto = True
                Call ModInterfaz.LogMensaje "[EXECUTE] JZ: Condicion cumplida (ZF=1). Bifurcando hacia 0x" & Hex2(ModCPU.IR_Operando)
            Else
                mCondicionSalto = False
                Call ModInterfaz.LogMensaje "[EXECUTE] JZ: Condicion no cumplida (ZF=0). No salta."
            End If
            
        Case &HA2 ' JNZ dir (Salto si Zero=0)
            If ModCPU.ZF = 0 Then
                mCondicionSalto = True
                Call ModInterfaz.LogMensaje "[EXECUTE] JNZ: Condicion cumplida (ZF=0). Bifurcando hacia 0x" & Hex2(ModCPU.IR_Operando)
            Else
                mCondicionSalto = False
                Call ModInterfaz.LogMensaje "[EXECUTE] JNZ: Condicion no cumplida (ZF=1). No salta."
            End If
            
        Case Else
            Call ModInterfaz.LogMensaje "[EXECUTE] Opcode no implementado: 0x" & Hex2(ModCPU.IR_Opcode)
    End Select
    
    ModCPU.NumPasoCiclo = 4
End Sub

''' <summary>
''' FASE 4: STORE / WRITE-BACK (Almacenamiento)
'''  - El resultado se guarda en el registro destino (AX, BX).
'''  - O se escribe en la memoria RAM via la primitiva Write(MAR, MDR).
'''  - O se actualiza el PC en caso de salto efectivo.
'''  - Concluye la instruccion y prepara el siguiente ciclo.
''' </summary>
Public Sub Fase_Store()
    ModCPU.FaseActual = "STORE"
    ModCPU.EstadoCPU = "STORE"
    Call ModInterfaz.ActualizarFaseUI "STORE"
    
    Select Case ModCPU.IR_Opcode
        ' Escritura en Registro AX
        Case &H10, &H12, &H20, &H30, &H32, &H40, &H42, &H60, &H62, &H64, &H70, &H72, &H80, &H82, &H90, &H92
            ModCPU.AX = mResultadoALU
            Call ModInterfaz.LogMensaje "[STORE] Registro AX <- 0x" & Hex2(ModCPU.AX)
            
        ' Escritura en Registro BX
        Case &H11, &H13, &H21, &H31, &H33, &H41, &H43, &H61, &H63, &H65, &H71, &H73, &H81, &H83, &H91, &H93
            ModCPU.BX = mResultadoALU
            Call ModInterfaz.LogMensaje "[STORE] Registro BX <- 0x" & Hex2(ModCPU.BX)
            
        ' Escritura en Memoria RAM via Primitiva Write()
        Case &H22, &H23
            Call ModMemoria.Write(ModCPU.MAR, CByte(ModCPU.MDR))
            Call ModInterfaz.LogMensaje "[STORE] Write(0x" & Hex2(ModCPU.MAR) & ", 0x" & Hex2(ModCPU.MDR) & ") completado en RAM."
            
        ' Actualizacion de PC por Salto
        Case &HA0, &HA1, &HA2
            If mCondicionSalto Then
                ModCPU.PC = ModCPU.IR_Operando
                Call ModInterfaz.ResaltarCeldaMemoria ModCPU.PC, "PC"
                Call ModInterfaz.LogMensaje "[STORE] Salto efectuado: PC actualizado a 0x" & Hex2(ModCPU.PC)
            Else
                Call ModInterfaz.LogMensaje "[STORE] No hubo salto: PC continua en 0x" & Hex2(ModCPU.PC)
            End If
            
        ' Instrucciones que no modifican registros ni RAM (CMP, HLT, NOP)
        Case &H50, &H51, &H52, &H53
            Call ModInterfaz.LogMensaje "[STORE] CMP: Solo banderas actualizadas. Registros intactos."
            
        Case &H0
            Call ModInterfaz.LogMensaje "[STORE] Fin de programa (HLT)."
    End Select
    
    ModCPU.ContadorInstrucciones = ModCPU.ContadorInstrucciones + 1
    
    ' Resalta en el editor la linea correspondiente a la siguiente instruccion
    Call ModInterfaz.ResaltarLineaEditor ModCPU.PC
    
    ' Vuelve a la Fase 1 (Fetch) para la siguiente instruccion
    ModCPU.NumPasoCiclo = 1
End Sub

' ==============================================================================
'  UTILIDADES INTERNAS DEL CICLO
' ==============================================================================

Public Function RequiereSegundoByte(ByVal op As Long) As Boolean
    Select Case op
        ' Instrucciones con operando inmediato, direccion de memoria o salto:
        Case &H10, &H11, _
             &H20, &H21, &H22, &H23, _
             &H30, &H31, _
             &H40, &H41, _
             &H50, &H51, _
             &H70, &H71, _
             &H80, &H81, _
             &H90, &H91, _
             &HA0, &HA1, &HA2
            RequiereSegundoByte = True
        Case Else
            RequiereSegundoByte = False
    End Select
End Function

Public Function ObtenerRetardo() As Double
    Dim ws As Worksheet
    Dim valVel As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Simulador")
    valVel = CLng(ws.Range("E30").Value)  ' Celda con valor de velocidad (1 a 5)
    If valVel <= 0 Then valVel = 3
    
    ' Velocidad 1: 0.8s, Velocidad 3: 0.3s, Velocidad 5: 0.05s
    Select Case valVel
        Case 1: ObtenerRetardo = 0.8
        Case 2: ObtenerRetardo = 0.5
        Case 3: ObtenerRetardo = 0.3
        Case 4: ObtenerRetardo = 0.15
        Case 5: ObtenerRetardo = 0.05
        Case Else: ObtenerRetardo = 0.3
    End Select
End Function

Public Sub DormirMs(ByVal segundos As Double)
    Dim tInicio As Double
    tInicio = Timer
    Do While Timer < tInicio + segundos
        DoEvents
    Loop
End Sub
