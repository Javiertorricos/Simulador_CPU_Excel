# ==============================================================================
#  UNIVERSIDAD CATOLICA BOLIVIANA "SAN PABLO"
#  Materia: Arquitectura de Computadoras (SIS-131)
#  Proyecto: Simulador de CPU von Neumann (8 bits) y Memoria Principal
#  Autor: Javier Torrico Sejas
#
#  Script de generacion automatizada del libro Excel con macros (.xlsm)
# ==============================================================================

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BaseDir = Split-Path -Parent $ScriptDir
$SrcDir = Join-Path $BaseDir 'src'
$OutputFile = Join-Path $BaseDir 'Simulador_CPU_8bits_Javier_Torrico.xlsm'

Write-Host "Generando Simulador de CPU de 8 bits para Javier Torrico Sejas..."
Write-Host "Directorio base: $BaseDir"
Write-Host "Archivo destino: $OutputFile"

# Asegurar permiso de acceso a modelo de objetos de VBA (AccessVBOM)
$secPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
if (-not (Test-Path $secPath)) { New-Item -Path $secPath -Force | Out-Null }
Set-ItemProperty -Path $secPath -Name "AccessVBOM" -Value 1 -Type DWord

# Crear instancia de Excel
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1

try {
    # Crear libro nuevo
    $wb = $xl.Workbooks.Add()
    
    # Configurar primera hoja: Simulador
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Simulador"
    
    # Crear segunda hoja: Log
    $wsLog = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws)
    $wsLog.Name = "Log"
    
    # Configurar hoja Log
    $wsLog.Range("A1:G1").Value = @("Ciclo", "Fase", "Micro-operacion", "PC", "AX", "BX", "Flags")
    $wsLog.Range("A1:G1").Font.Bold = $true
    $wsLog.Range("A1:G1").Interior.Color = 0x222222
    $wsLog.Range("A1:G1").Font.Color = 0xFFFFFF
    $wsLog.Columns.Item(1).ColumnWidth = 8
    $wsLog.Columns.Item(2).ColumnWidth = 12
    $wsLog.Columns.Item(3).ColumnWidth = 55
    $wsLog.Columns.Item(4).ColumnWidth = 8
    $wsLog.Columns.Item(5).ColumnWidth = 8
    $wsLog.Columns.Item(6).ColumnWidth = 8
    $wsLog.Columns.Item(7).ColumnWidth = 18

    # -------------------------------------------------------------------------
    #  DISEÑO DE LA HOJA PRINCIPAL: Simulador
    # -------------------------------------------------------------------------
    $ws.Activate()
    
    # Ancho de columnas
    $ws.Columns.Item("A").ColumnWidth = 2
    $ws.Columns.Item("B").ColumnWidth = 7
    $ws.Columns.Item("C").ColumnWidth = 18
    $ws.Columns.Item("D").ColumnWidth = 9
    $ws.Columns.Item("E").ColumnWidth = 12
    $ws.Columns.Item("F").ColumnWidth = 22
    $ws.Columns.Item("G").ColumnWidth = 2
    
    $ws.Columns.Item("H").ColumnWidth = 11
    $ws.Columns.Item("I").ColumnWidth = 11
    $ws.Columns.Item("J").ColumnWidth = 11
    $ws.Columns.Item("K").ColumnWidth = 11
    $ws.Columns.Item("L").ColumnWidth = 11
    $ws.Columns.Item("M").ColumnWidth = 11
    $ws.Columns.Item("N").ColumnWidth = 11
    $ws.Columns.Item("O").ColumnWidth = 11
    $ws.Columns.Item("P").ColumnWidth = 14
    $ws.Columns.Item("Q").ColumnWidth = 2
    
    # Columnas de Memoria RAM (R a AG = 16 columnas)
    for ($col = 18; $col -le 33; $col++) {
        $ws.Columns.Item($col).ColumnWidth = 4.2
    }
    
    # 1. ENCABEZADO PRINCIPAL (Fila 2 a 3)
    $rngHeader = $ws.Range("B2:AG3")
    $rngHeader.Merge()
    $rngHeader.Value = "UNIVERSIDAD CATOLICA BOLIVIANA «SAN PABLO»  |  SIMULADOR DE CPU VON NEUMANN (8 BITS) Y MEMORIA PRINCIPAL`nEstudiante: Javier Torrico Sejas  |  Docente: Ing. Paulo Cesar Loayza Carrasco  |  Materia: Arquitectura de Computadoras (SIS-131)"
    $rngHeader.Font.Name = "Segoe UI"
    $rngHeader.Font.Size = 11
    $rngHeader.Font.Bold = $true
    $rngHeader.Font.Color = 0xFFFFFF
    $rngHeader.Interior.Color = 0x332211 # Azul oscuro/marino
    $rngHeader.HorizontalAlignment = 3  # Centrado
    $rngHeader.VerticalAlignment = 2    # Centrado vertical
    
    # 2. PANEL IZQUIERDO: EDITOR DE ENSAMBLADOR (B5:F26)
    $ws.Range("B5:F5").Merge()
    $ws.Range("B5:F5").Value = "EDITOR DE CODIGO ENSAMBLADOR (x86 8-bits)"
    $ws.Range("B5:F5").Font.Bold = $true
    $ws.Range("B5:F5").Interior.Color = 0x554433
    $ws.Range("B5:F5").Font.Color = 0xFFFFFF
    $ws.Range("B5:F5").HorizontalAlignment = 3
    
    $ws.Range("B6:F6").Value = @("Dir", "Instruccion ASM", "Bytes", "Operacion", "Comentario")
    $ws.Range("B6:F6").Font.Bold = $true
    $ws.Range("B6:F6").Interior.Color = 0xE0E0E0
    $ws.Range("B6:F6").HorizontalAlignment = 3
    
    # Bordes de tabla de ensamblador
    $ws.Range("B6:F26").Borders.LineStyle = 1
    $ws.Range("B6:F26").Font.Name = "Consolas"
    $ws.Range("B6:F26").Font.Size = 9.5
    
    # 3. CICLO DE INSTRUCCION (4 FASES) - Fila 5 a 6 en H:O
    $ws.Range("H4:P4").Merge()
    $ws.Range("H4:P4").Value = "CICLO DE INSTRUCCION DE 4 FASES (CLOCK / RELOJ)"
    $ws.Range("H4:P4").Font.Bold = $true
    $ws.Range("H4:P4").Interior.Color = 0x223344
    $ws.Range("H4:P4").Font.Color = 0xFFFFFF
    $ws.Range("H4:P4").HorizontalAlignment = 3
    
    $fases = @(
        @("H5:I6", "1. FETCH`n(Busqueda RAM -> IR)", 0xECF0F1),
        @("J5:K6", "2. DECODE`n(Decodifica Opcode)", 0xECF0F1),
        @("L5:M6", "3. EXECUTE`n(ALU / Saltos)", 0xECF0F1),
        @("N5:O6", "4. STORE`n(Escribe Destino/RAM)", 0xECF0F1)
    )
    foreach ($f in $fases) {
        $rngF = $ws.Range($f[0])
        $rngF.Merge()
        $rngF.Value = $f[1]
        $rngF.Font.Bold = $true
        $rngF.Font.Size = 9
        $rngF.HorizontalAlignment = 3
        $rngF.VerticalAlignment = 2
        $rngF.Interior.Color = $f[2]
        $rngF.Borders.LineStyle = 1
    }
    
    # Indicador de Buses en P5 y P6
    $ws.Range("P5").Value = "Bus Dir: 0x00"
    $ws.Range("P6").Value = "Bus Dat: 0x00"
    $ws.Range("P5:P6").Font.Size = 8.5
    $ws.Range("P5:P6").Font.Bold = $true
    $ws.Range("P5:P6").Interior.Color = 0xFAF0E6
    $ws.Range("P5:P6").Borders.LineStyle = 1

    # 4. BANCO DE REGISTROS DEL CPU (Fila 7 a 13)
    $ws.Range("H7:P7").Merge()
    $ws.Range("H7:P7").Value = "BANCO DE REGISTROS DE LA CPU (8 BITS) Y BANDERAS (FLAGS)"
    $ws.Range("H7:P7").Font.Bold = $true
    $ws.Range("H7:P7").Interior.Color = 0x334455
    $ws.Range("H7:P7").Font.Color = 0xFFFFFF
    $ws.Range("H7:P7").HorizontalAlignment = 3
    
    # Fila 8-9: PC, IR, MAR, MDR, Ciclos/Instrucciones
    $ws.Range("H8:I8").Merge(); $ws.Range("H8:I8").Value = "0x00"; $ws.Range("H8:I8").Font.Bold = $true; $ws.Range("H8:I8").HorizontalAlignment = 3
    $ws.Range("H9:I9").Merge(); $ws.Range("H9:I9").Value = "PC (Program Counter)"; $ws.Range("H9:I9").Font.Size = 8; $ws.Range("H9:I9").HorizontalAlignment = 3
    
    $ws.Range("J8:K8").Merge(); $ws.Range("J8:K8").Value = "0x00 [HLT]"; $ws.Range("J8:K8").Font.Bold = $true; $ws.Range("J8:K8").HorizontalAlignment = 3
    $ws.Range("J9:K9").Merge(); $ws.Range("J9:K9").Value = "IR (Instruction Reg)"; $ws.Range("J9:K9").Font.Size = 8; $ws.Range("J9:K9").HorizontalAlignment = 3
    
    $ws.Range("L8:M8").Merge(); $ws.Range("L8:M8").Value = "0x00"; $ws.Range("L8:M8").Font.Bold = $true; $ws.Range("L8:M8").HorizontalAlignment = 3
    $ws.Range("L9:M9").Merge(); $ws.Range("L9:M9").Value = "MAR (Mem Address Reg)"; $ws.Range("L9:M9").Font.Size = 8; $ws.Range("L9:M9").HorizontalAlignment = 3
    
    $ws.Range("N8:O8").Merge(); $ws.Range("N8:O8").Value = "0x00"; $ws.Range("N8:O8").Font.Bold = $true; $ws.Range("N8:O8").HorizontalAlignment = 3
    $ws.Range("N9:O9").Merge(); $ws.Range("N9:O9").Value = "MDR (Mem Data Reg)"; $ws.Range("N9:O9").Font.Size = 8; $ws.Range("N9:O9").HorizontalAlignment = 3
    
    $ws.Range("P8").Value = "0 Ciclos"
    $ws.Range("P9").Value = "0 Instr."
    $ws.Range("P8:P9").Font.Size = 8.5
    $ws.Range("P8:P9").HorizontalAlignment = 3

    # Fila 11-13: AX, BX, FLAGS (ZF, CF, SF) y ESTADO
    $ws.Range("H11:I11").Merge(); $ws.Range("H11:I11").Value = "AX (Acumulador)"; $ws.Range("H11:I11").Font.Size = 8; $ws.Range("H11:I11").HorizontalAlignment = 3
    $ws.Range("H12:I12").Merge(); $ws.Range("H12:I12").Value = "0x00"; $ws.Range("H12:I12").Font.Bold = $true; $ws.Range("H12:I12").HorizontalAlignment = 3
    $ws.Range("H13:I13").Merge(); $ws.Range("H13:I13").Value = "0 (00000000b)"; $ws.Range("H13:I13").Font.Size = 8; $ws.Range("H13:I13").HorizontalAlignment = 3
    
    $ws.Range("J11:K11").Merge(); $ws.Range("J11:K11").Value = "BX (General)"; $ws.Range("J11:K11").Font.Size = 8; $ws.Range("J11:K11").HorizontalAlignment = 3
    $ws.Range("J12:K12").Merge(); $ws.Range("J12:K12").Value = "0x00"; $ws.Range("J12:K12").Font.Bold = $true; $ws.Range("J12:K12").HorizontalAlignment = 3
    $ws.Range("J13:K13").Merge(); $ws.Range("J13:K13").Value = "0 (00000000b)"; $ws.Range("J13:K13").Font.Size = 8; $ws.Range("J13:K13").HorizontalAlignment = 3
    
    $ws.Range("L11").Value = "ZF (Zero)"; $ws.Range("L11").Font.Size = 8; $ws.Range("L11").HorizontalAlignment = 3
    $ws.Range("L12").Value = "0"; $ws.Range("L12").Font.Bold = $true; $ws.Range("L12").HorizontalAlignment = 3
    
    $ws.Range("M11").Value = "CF (Carry)"; $ws.Range("M11").Font.Size = 8; $ws.Range("M11").HorizontalAlignment = 3
    $ws.Range("M12").Value = "0"; $ws.Range("M12").Font.Bold = $true; $ws.Range("M12").HorizontalAlignment = 3
    
    $ws.Range("N11").Value = "SF (Sign)"; $ws.Range("N11").Font.Size = 8; $ws.Range("N11").HorizontalAlignment = 3
    $ws.Range("N12").Value = "0"; $ws.Range("N12").Font.Bold = $true; $ws.Range("N12").HorizontalAlignment = 3
    
    $ws.Range("O11:P11").Merge(); $ws.Range("O11:P11").Value = "ESTADO CPU"; $ws.Range("O11:P11").Font.Size = 8; $ws.Range("O11:P11").HorizontalAlignment = 3
    $ws.Range("O12:P12").Merge(); $ws.Range("O12:P12").Value = "LISTO"; $ws.Range("O12:P12").Font.Bold = $true; $ws.Range("O12:P12").HorizontalAlignment = 3
    
    $ws.Range("H8:P13").Borders.LineStyle = 1
    $ws.Range("H8:P13").Interior.Color = 0xF9F9F9

    # 5. PANEL ALU (Fila 14)
    $ws.Range("H14:K14").Merge(); $ws.Range("H14:K14").Value = "ALU (Unidad Aritmetico-Logica):"; $ws.Range("H14:K14").Font.Bold = $true; $ws.Range("H14:K14").Font.Size = 8.5
    $ws.Range("L14:O14").Merge(); $ws.Range("L14:O14").Value = "En espera de operacion"; $ws.Range("L14:O14").Font.Size = 8.5
    $ws.Range("P14").Value = "Res: 0x00"; $ws.Range("P14").Font.Bold = $true; $ws.Range("P14").Font.Size = 8.5
    $ws.Range("H14:P14").Borders.LineStyle = 1
    $ws.Range("H14:P14").Interior.Color = 0xEFF3F8

    # 6. PANEL DE PRIMITIVAS DE MEMORIA (Fila 15 a 16)
    $ws.Range("H15:P15").Merge()
    $ws.Range("H15:P15").Value = "OPERACIONES PRIMITIVAS DE MEMORIA: Read(address) / Write(address, value)"
    $ws.Range("H15:P15").Font.Bold = $true
    $ws.Range("H15:P15").Font.Size = 9
    $ws.Range("H15:P15").Interior.Color = 0x445566
    $ws.Range("H15:P15").Font.Color = 0xFFFFFF
    $ws.Range("H15:P15").HorizontalAlignment = 3
    
    $ws.Range("H16:I16").Merge(); $ws.Range("H16:I16").Value = "LISTO"; $ws.Range("H16:I16").Font.Bold = $true; $ws.Range("H16:I16").HorizontalAlignment = 3
    $ws.Range("J16:K16").Merge(); $ws.Range("J16:K16").Value = "MAR: 0x00 (0)"; $ws.Range("J16:K16").HorizontalAlignment = 3; $ws.Range("J16:K16").Font.Size = 8.5
    $ws.Range("L16:M16").Merge(); $ws.Range("L16:M16").Value = "MDR: 0x00 (0)"; $ws.Range("L16:M16").HorizontalAlignment = 3; $ws.Range("L16:M16").Font.Size = 8.5
    $ws.Range("N16:P16").Merge(); $ws.Range("N16:P16").Value = "Sin accesos recientes"; $ws.Range("N16:P16").HorizontalAlignment = 3; $ws.Range("N16:P16").Font.Size = 8.5
    $ws.Range("H16:P16").Borders.LineStyle = 1
    $ws.Range("H16:P16").Interior.Color = 0xFCFBF7

    # 7. LOG DE MICRO-OPERACIONES (Fila 18 a 27)
    $ws.Range("H18:P18").Merge()
    $ws.Range("H18:P18").Value = "LOG CRONOLOGICO DE MICRO-OPERACIONES (TRAZA DEL RELOJ)"
    $ws.Range("H18:P18").Font.Bold = $true
    $ws.Range("H18:P18").Font.Size = 9
    $ws.Range("H18:P18").Interior.Color = 0x222222
    $ws.Range("H18:P18").Font.Color = 0xFFFFFF
    $ws.Range("H18:P18").HorizontalAlignment = 3
    
    for ($f = 19; $f -le 27; $f++) {
        $ws.Range("H$f:P$f").Merge()
        $ws.Range("H$f:P$f").Font.Name = "Consolas"
        $ws.Range("H$f:P$f").Font.Size = 8.5
        $ws.Range("H$f:P$f").Interior.Color = 0xF5F5F5
        $ws.Range("H$f:P$f").Borders.LineStyle = 1
    }
    $ws.Range("H19:P19").Value = "[000] Sistema inicializado y listo para ejecutar."

    # 8. MATRIZ DE MEMORIA RAM 16x16 (R4 a AG21)
    $ws.Range("R4:AG4").Merge()
    $ws.Range("R4:AG4").Value = "MEMORIA PRINCIPAL RAM (256 BYTES: 00h - FFh)"
    $ws.Range("R4:AG4").Font.Bold = $true
    $ws.Range("R4:AG4").Font.Size = 10
    $ws.Range("R4:AG4").Interior.Color = 0x2C3E50
    $ws.Range("R4:AG4").Font.Color = 0xFFFFFF
    $ws.Range("R4:AG4").HorizontalAlignment = 3
    
    # Encabezados de columnas de la matriz: +0 a +F
    $colsHex = @("+0", "+1", "+2", "+3", "+4", "+5", "+6", "+7", "+8", "+9", "+A", "+B", "+C", "+D", "+E", "+F")
    for ($c = 0; $c -lt 16; $c++) {
        $cell = $ws.Cells.Item(5, 18 + $c)
        $cell.Value = $colsHex[$c]
        $cell.Font.Bold = $true
        $cell.Font.Size = 9
        $cell.HorizontalAlignment = 3
        $cell.Interior.Color = 0xD5DBDB
    }
    
    # Filas de la matriz: 00h a F0h
    # Filas 6 a 13 (direcciones 00h-7Fh): Segmento de Codigo (Azul sutil: 0xFFF5EB = celeste suave)
    # Filas 14 a 21 (direcciones 80h-FFh): Segmento de Datos (Verde sutil: 0xEBFFEB = verde suave)
    for ($f = 0; $f -lt 16; $f++) {
        $rowExcel = 6 + $f
        $colorFila = if ($f -lt 8) { 0xFFF5EB } else { 0xEBFFEB }
        
        for ($c = 0; $c -lt 16; $c++) {
            $cell = $ws.Cells.Item($rowExcel, 18 + $c)
            $cell.Value = "00"
            $cell.Font.Name = "Consolas"
            $cell.Font.Size = 9
            $cell.HorizontalAlignment = 3
            $cell.Interior.Color = $colorFila
        }
    }
    $ws.Range("R5:AG21").Borders.LineStyle = 1

    # Leyenda de Segmentacion de Memoria (Fila 22)
    $ws.Range("R22:V22").Merge(); $ws.Range("R22:V22").Value = "00h-7Fh: Segmento Codigo (CS)"; $ws.Range("R22:V22").Font.Size = 8; $ws.Range("R22:V22").Interior.Color = 0xFFF5EB; $ws.Range("R22:V22").Borders.LineStyle = 1
    $ws.Range("W22:AA22").Merge(); $ws.Range("W22:AA22").Value = "80h-FFh: Segmento Datos (DS)"; $ws.Range("W22:AA22").Font.Size = 8; $ws.Range("W22:AA22").Interior.Color = 0xEBFFEB; $ws.Range("W22:AA22").Borders.LineStyle = 1
    $ws.Range("AB22:AG22").Merge(); $ws.Range("AB22:AG22").Value = "Celeste: PC | Ambar: Read | Rojo: Write"; $ws.Range("AB22:AG22").Font.Size = 8; $ws.Range("AB22:AG22").HorizontalAlignment = 3

    # 9. PANEL INSPECTOR DE MEMORIA (Fila 23 a 25)
    $ws.Range("R23:AG23").Merge()
    $ws.Range("R23:AG23").Value = "INSPECTOR DE MEMORIA (VALORES EN HEX, BIN, DEC Y NEMONICO)"
    $ws.Range("R23:AG23").Font.Bold = $true
    $ws.Range("R23:AG23").Font.Size = 8.5
    $ws.Range("R23:AG23").Interior.Color = 0x34495E
    $ws.Range("R23:AG23").Font.Color = 0xFFFFFF
    $ws.Range("R23:AG23").HorizontalAlignment = 3
    
    $ws.Range("R24:T24").Merge(); $ws.Range("R24:T24").Value = "0x00 (0d)"; $ws.Range("R24:T24").Font.Bold = $true; $ws.Range("R24:T24").HorizontalAlignment = 3
    $ws.Range("U24:W24").Merge(); $ws.Range("U24:W24").Value = "0x00"; $ws.Range("U24:W24").Font.Bold = $true; $ws.Range("U24:W24").HorizontalAlignment = 3
    $ws.Range("X24:AA24").Merge(); $ws.Range("X24:AA24").Value = "00000000b"; $ws.Range("X24:AA24").Font.Name = "Consolas"; $ws.Range("X24:AA24").HorizontalAlignment = 3
    $ws.Range("AB24:AD24").Merge(); $ws.Range("AB24:AD24").Value = "0 / 0"; $ws.Range("AB24:AD24").HorizontalAlignment = 3
    $ws.Range("AE24:AG24").Merge(); $ws.Range("AE24:AG24").Value = "SEG. CODIGO"; $ws.Range("AE24:AG24").HorizontalAlignment = 3; $ws.Range("AE24:AG24").Font.Size = 8
    
    $ws.Range("R25:AG25").Merge(); $ws.Range("R25:AG25").Value = "Como Instruccion: HLT (Detiene el reloj)"; $ws.Range("R25:AG25").Font.Size = 8.5; $ws.Range("R25:AG25").HorizontalAlignment = 3
    $ws.Range("R24:AG25").Borders.LineStyle = 1
    $ws.Range("R24:AG25").Interior.Color = 0xFAF8F5

    # 10. PANEL DE CONTROL Y BOTONES DE ACCION (Fila 28 a 34)
    # Velocidad en E30
    $ws.Range("B30:D30").Merge(); $ws.Range("B30:D30").Value = "Velocidad de Reloj (1=Lento, 5=Rapido):"; $ws.Range("B30:D30").Font.Bold = $true; $ws.Range("B30:D30").Font.Size = 8.5
    $ws.Range("E30").Value = 3
    $ws.Range("E30").Font.Bold = $true
    $ws.Range("E30").HorizontalAlignment = 3
    $ws.Range("E30").Interior.Color = 0xD5F5E3
    $ws.Range("E30").Borders.LineStyle = 1

    # Crear botones de Excel (Form Shapes / Buttons)
    function Add-Boton($ws, [string]$nombre, [string]$texto, [string]$macro, [double]$left, [double]$top, [double]$w, [double]$h, [int]$colorFill) {
        $btn = $ws.Shapes.AddShape(5, $left, $top, $w, $h) # msoShapeRoundedRectangle = 5
        $btn.Name = $nombre
        $btn.TextFrame2.TextRange.Text = $texto
        $btn.TextFrame2.TextRange.Font.Name = "Segoe UI"
        $btn.TextFrame2.TextRange.Font.Size = 9.5
        $btn.TextFrame2.TextRange.Font.Bold = -1
        $btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
        $btn.TextFrame2.TextRange.ParagraphFormat.Alignment = 2 # Centrado
        $btn.TextFrame2.VerticalAnchor = 3                      # Centrado
        $btn.Fill.ForeColor.RGB = $colorFill
        $btn.Line.Visible = 0
        $btn.OnAction = $macro
        return $btn
    }

    # Coordenadas calculadas usando celdas
    $cColB = $ws.Columns.Item(2).Left
    $cColD = $ws.Columns.Item(4).Left
    $cColH = $ws.Columns.Item(8).Left
    $cColJ = $ws.Columns.Item(10).Left
    $cColL = $ws.Columns.Item(12).Left
    $cColN = $ws.Columns.Item(14).Left
    $cColP = $ws.Columns.Item(16).Left
    
    $rFila28 = $ws.Rows.Item(28).Top
    $rFila32 = $ws.Rows.Item(32).Top
    $rFila34 = $ws.Rows.Item(34).Top

    # Botones Principales de Ejecucion
    Add-Boton $ws "btn_load"  "CARGAR PROGRAMA"  "Btn_CargarPrograma"   $cColB ($rFila28) 140 28 0x7A431E # Azul oscuro
    Add-Boton $ws "btn_step"  "PASO A PASO"      "Btn_PasoFase"         ($cColB + 150) ($rFila28) 110 28 0x3E8E41 # Verde
    Add-Boton $ws "btn_instr" "INSTRUCCION >>"   "Btn_PasoInstruccion"  ($cColB + 270) ($rFila28) 120 28 0x5B6B2E # Verde azulado
    Add-Boton $ws "btn_run"   "CONTINUO (RUN)"   "Btn_Run"              ($cColB + 400) ($rFila28) 120 28 0x228B22 # Verde bosque
    Add-Boton $ws "btn_pause" "PAUSAR"           "Btn_Pausa"            ($cColB + 530) ($rFila28) 90  28 0x207CCA # Naranja
    Add-Boton $ws "btn_reset" "REINICIAR"        "Btn_Reset"            ($cColB + 630) ($rFila28) 100 28 0x3B3BCE # Rojo tomate

    # Botones de Carga Rapida de Ejemplos
    $rFilaEjemplos = $ws.Rows.Item(32).Top
    $ws.Range("B31:G31").Merge()
    $ws.Range("B31:G31").Value = "PROGRAMAS DEMOSTRATIVOS OBLIGATORIOS (CLICK PARA CARGAR):"
    $ws.Range("B31:G31").Font.Bold = $true
    $ws.Range("B31:G31").Font.Size = 9
    
    Add-Boton $ws "btn_ex1" "1. Multiplicacion (6*7)" "Btn_Ejemplo1" $cColB ($rFilaEjemplos) 160 25 0x4A4A4A
    Add-Boton $ws "btn_ex2" "2. Serie Fibonacci"      "Btn_Ejemplo2" ($cColB + 170) ($rFilaEjemplos) 150 25 0x4A4A4A
    Add-Boton $ws "btn_ex3" "3. Factorial (4!)"       "Btn_Ejemplo3" ($cColB + 330) ($rFilaEjemplos) 140 25 0x4A4A4A
    Add-Boton $ws "btn_ex4" "4. Cuenta Regresiva"     "Btn_Ejemplo4" ($cColB + 480) ($rFilaEjemplos) 150 25 0x4A4A4A

    # -------------------------------------------------------------------------
    #  IMPORTACION DE CODIGO VBA
    # -------------------------------------------------------------------------
    Write-Host "Importando modulos VBA desde $SrcDir..."
    $vbaProj = $wb.VBProject
    $modulos = @(
        'ModMemoria.bas',
        'ModCPU.bas',
        'ModALU.bas',
        'ModCicloInstruccion.bas',
        'ModEnsamblador.bas',
        'ModInterfaz.bas'
    )
    foreach ($m in $modulos) {
        $mPath = Join-Path $SrcDir $m
        $nombre = [System.IO.Path]::GetFileNameWithoutExtension($m)
        Write-Host " - Importando $nombre..."
        
        $texto = Get-Content -Path $mPath -Raw -Encoding UTF8
        # Quitar encabezado Attribute VB_Name
        $texto = ($texto -split "`r?`n" | Where-Object { $_ -notmatch '^Attribute\s+VB_' }) -join "`r`n"
        
        $comp = $null
        foreach ($c in $vbaProj.VBComponents) {
            if ($c.Name -eq $nombre) { $comp = $c; break }
        }
        if ($null -eq $comp) {
            $comp = $vbaProj.VBComponents.Add(1) # vbext_ct_StdModule = 1
            $comp.Name = $nombre
        }
        $cm = $comp.CodeModule
        if ($cm.CountOfLines -gt 0) {
            $cm.DeleteLines(1, $cm.CountOfLines)
        }
        $cm.AddFromString($texto)
    }

    # Pre-cargar el Ejemplo 1 (Multiplicacion 6 * 7 = 42) en el editor y en la RAM
    Write-Host "Inicializando ejemplo 1 (Multiplicacion por sumas sucesivas)..."
    $ws.Cells.Item(7, 2).Value = "0x00"; $ws.Cells.Item(7, 3).Value = "MOV AX, 0";       $ws.Cells.Item(7, 4).Value = "10 00"; $ws.Cells.Item(7, 5).Value = "MOV";   $ws.Cells.Item(7, 6).Value = "Inicializa acumulador de producto en 0"
    $ws.Cells.Item(8, 2).Value = "0x02"; $ws.Cells.Item(8, 3).Value = "MOV BX, 7";       $ws.Cells.Item(8, 4).Value = "11 07"; $ws.Cells.Item(8, 5).Value = "MOV";   $ws.Cells.Item(8, 6).Value = "Contador (multiplicador = 7)"
    $ws.Cells.Item(9, 2).Value = "0x04"; $ws.Cells.Item(9, 3).Value = "BUCLE:";          $ws.Cells.Item(9, 4).Value = "";      $ws.Cells.Item(9, 5).Value = "ETIQ";  $ws.Cells.Item(9, 6).Value = "Inicio de bucle de suma"
    $ws.Cells.Item(10, 2).Value = "0x04"; $ws.Cells.Item(10, 3).Value = "ADD AX, 6";     $ws.Cells.Item(10, 4).Value = "30 06"; $ws.Cells.Item(10, 5).Value = "ADD";  $ws.Cells.Item(10, 6).Value = "Suma sucesiva del multiplicando (6)"
    $ws.Cells.Item(11, 2).Value = "0x06"; $ws.Cells.Item(11, 3).Value = "DEC BX";        $ws.Cells.Item(11, 4).Value = "63";    $ws.Cells.Item(11, 5).Value = "DEC";  $ws.Cells.Item(11, 6).Value = "Decrementa contador"
    $ws.Cells.Item(12, 2).Value = "0x07"; $ws.Cells.Item(12, 3).Value = "JNZ BUCLE";     $ws.Cells.Item(12, 4).Value = "A2 04"; $ws.Cells.Item(12, 5).Value = "JNZ";  $ws.Cells.Item(12, 6).Value = "Bifurca si BX != 0 (ZF=0)"
    $ws.Cells.Item(13, 2).Value = "0x09"; $ws.Cells.Item(13, 3).Value = "STORE [80h], AX"; $ws.Cells.Item(13, 4).Value = "22 80"; $ws.Cells.Item(13, 5).Value = "STORE"; $ws.Cells.Item(13, 6).Value = "Guarda producto en RAM[80h] (Datos)"
    $ws.Cells.Item(14, 2).Value = "0x0B"; $ws.Cells.Item(14, 3).Value = "HLT";           $ws.Cells.Item(14, 4).Value = "00";    $ws.Cells.Item(14, 5).Value = "HLT";  $ws.Cells.Item(14, 6).Value = "Detiene el reloj del CPU"

    # Bytes en la memoria RAM (00h a 0Bh)
    $bytesProg = @("10", "00", "11", "07", "30", "06", "63", "A2", "04", "22", "80", "00")
    for ($i = 0; $i -lt $bytesProg.Count; $i++) {
        $ws.Cells.Item(6, 18 + $i).Value = $bytesProg[$i]
    }
    
    # Resaltar celda 00h (PC inicial)
    $ws.Cells.Item(6, 18).Interior.Color = 0xFFA500 # Cian/Azul PC inicial

    # Guardar libro como .xlsm (52 = xlOpenXMLWorkbookMacroEnabled)
    Write-Host "Guardando libro en: $OutputFile"
    if (Test-Path $OutputFile) {
        try { Remove-Item $OutputFile -Force } catch {}
    }
    $wb.SaveAs($OutputFile, 52)
    
    Write-Host "PROCESO COMPLETADO EXITOSAMENTE!"
    Write-Host "Archivo generado: $OutputFile"
}
finally {
    if ($null -ne $wb) {
        try { $wb.Close($false) } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } catch {}
    }
    if ($null -ne $xl) {
        try { $xl.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch {}
    }
}
