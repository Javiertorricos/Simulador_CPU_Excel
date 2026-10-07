Attribute VB_Name = "ModALU"
Option Explicit

' ==============================================================================
'  UNIVERSIDAD CATOLICA BOLIVIANA "SAN PABLO"
'  Materia: Arquitectura de Computadoras (SIS-131)
'  Proyecto: Simulador de CPU von Neumann (8 bits) y Memoria Principal
'  Autor: Javier Torrico Sejas
'
'  MODULO: ModALU
'  Descripcion:
'    Implementa la Unidad Aritmetico-Logica (ALU) de 8 bits.
'    Procesa operaciones aritmeticas (ADD, SUB, INC, DEC) y logicas (AND, OR, XOR, NOT, CMP).
'    Calcula rigurosamente las banderas de estado: ZF (Zero), CF (Carry) y SF (Sign).
' ==============================================================================

Public UltimaOpALU As String       ' Descripcion de la operacion actual para la UI
Public UltimoResALU As Long        ' Valor resultante de 8 bits

''' <summary>
''' Ejecuta una operacion en la ALU sobre operandos de 8 bits (0..255).
''' Actualiza las banderas ZF, CF y SF segun la semantica de arquitectura de computadoras.
''' Devuelve el resultado truncado a 8 bits (0..255).
''' </summary>
Public Function ALU_Operar(ByVal op As String, ByVal valA As Long, ByVal valB As Long) As Long
    Dim resRaw As Long
    Dim res8 As Long
    
    valA = valA And &HFF
    valB = valB And &HFF
    
    Select Case UCase$(Trim$(op))
        ' --- OPERACIONES ARITMETICAS ---
        Case "ADD"
            resRaw = valA + valB
            res8 = resRaw And &HFF
            
            ' Carry Flag: Acarreo si la suma excede 255
            If resRaw > 255 Then ModCPU.CF = 1 Else ModCPU.CF = 0
            
            ' Zero Flag y Sign Flag
            If res8 = 0 Then ModCPU.ZF = 1 Else ModCPU.ZF = 0
            If (res8 And &H80) <> 0 Then ModCPU.SF = 1 Else ModCPU.SF = 0
            
            UltimaOpALU = "ADD: 0x" & Hex2(valA) & " + 0x" & Hex2(valB) & " = 0x" & Hex2(res8)

        Case "SUB", "CMP"
            resRaw = valA - valB
            res8 = resRaw And &HFF
            
            ' Carry Flag (Prestamo en resta sin signo): Se activa si valA < valB
            If valA < valB Then ModCPU.CF = 1 Else ModCPU.CF = 0
            
            ' Zero Flag y Sign Flag
            If res8 = 0 Then ModCPU.ZF = 1 Else ModCPU.ZF = 0
            If (res8 And &H80) <> 0 Then ModCPU.SF = 1 Else ModCPU.SF = 0
            
            If UCase$(Trim$(op)) = "CMP" Then
                UltimaOpALU = "CMP: Compara 0x" & Hex2(valA) & " con 0x" & Hex2(valB) & " (Actualiza ZF, CF, SF)"
            Else
                UltimaOpALU = "SUB: 0x" & Hex2(valA) & " - 0x" & Hex2(valB) & " = 0x" & Hex2(res8)
            End If

        Case "INC"
            resRaw = valA + 1
            res8 = resRaw And &HFF
            
            ' En x86 INC no afecta CF (mantiene el acarreo previo)
            If res8 = 0 Then ModCPU.ZF = 1 Else ModCPU.ZF = 0
            If (res8 And &H80) <> 0 Then ModCPU.SF = 1 Else ModCPU.SF = 0
            
            UltimaOpALU = "INC: 0x" & Hex2(valA) & " + 1 = 0x" & Hex2(res8)

        Case "DEC"
            resRaw = valA - 1
            res8 = resRaw And &HFF
            
            ' En x86 DEC no afecta CF
            If res8 = 0 Then ModCPU.ZF = 1 Else ModCPU.ZF = 0
            If (res8 And &H80) <> 0 Then ModCPU.SF = 1 Else ModCPU.SF = 0
            
            UltimaOpALU = "DEC: 0x" & Hex2(valA) & " - 1 = 0x" & Hex2(res8)

        ' --- OPERACIONES LOGICAS ---
        Case "AND"
            res8 = (valA And valB) And &HFF
            ModCPU.CF = 0   ' En x86 operaciones logicas limpian CF
            If res8 = 0 Then ModCPU.ZF = 1 Else ModCPU.ZF = 0
            If (res8 And &H80) <> 0 Then ModCPU.SF = 1 Else ModCPU.SF = 0
            
            UltimaOpALU = "AND: 0x" & Hex2(valA) & " AND 0x" & Hex2(valB) & " = 0x" & Hex2(res8)

        Case "OR"
            res8 = (valA Or valB) And &HFF
            ModCPU.CF = 0
            If res8 = 0 Then ModCPU.ZF = 1 Else ModCPU.ZF = 0
            If (res8 And &H80) <> 0 Then ModCPU.SF = 1 Else ModCPU.SF = 0
            
            UltimaOpALU = "OR: 0x" & Hex2(valA) & " OR 0x" & Hex2(valB) & " = 0x" & Hex2(res8)

        Case "XOR"
            res8 = (valA Xor valB) And &HFF
            ModCPU.CF = 0
            If res8 = 0 Then ModCPU.ZF = 1 Else ModCPU.ZF = 0
            If (res8 And &H80) <> 0 Then ModCPU.SF = 1 Else ModCPU.SF = 0
            
            UltimaOpALU = "XOR: 0x" & Hex2(valA) & " XOR 0x" & Hex2(valB) & " = 0x" & Hex2(res8)

        Case "NOT"
            res8 = (Not valA) And &HFF
            ' NOT no altera banderas en x86
            UltimaOpALU = "NOT: NOT 0x" & Hex2(valA) & " = 0x" & Hex2(res8)

        Case Else
            res8 = valA
            UltimaOpALU = "ALU Pasante: 0x" & Hex2(res8)
    End Select
    
    UltimoResALU = res8
    ALU_Operar = res8
    
    ' Mostrar calculo en el panel ALU de la hoja
    Call ModInterfaz.ActualizarPanelALU(UltimaOpALU, res8)
End Function
