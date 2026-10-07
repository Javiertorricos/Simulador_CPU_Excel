# GitHub Projects Kanban Board: Issues, Descriptions & Attachments (English)

**Project:** 8-Bit von Neumann CPU & Main Memory Simulator (Excel VBA)  
**Author:** Javier Torrico Sejas  
**Repository:** [https://github.com/Javiertorricos/Simulador_CPU_Excel](https://github.com/Javiertorricos/Simulador_CPU_Excel)  
**Academic Course:** Computer Architecture (SIS-131) — Exam 1  
**Institution:** Universidad Católica Boliviana "San Pablo"  
**Professor:** Eng. Paulo César Loayza Carrasco  

---

## Kanban Board Overview & Workflow

The GitHub Projects board is organized into five sequential columns following agile methodology:
1. **Backlog**: Planned tasks, requirements breakdown, and conceptual designs.
2. **To Do**: Prioritized tasks ready for implementation in the current iteration.
3. **In Progress**: Tasks actively being coded, drafted, or tested.
4. **In Review / Testing**: Completed features undergoing unit testing, rubric compliance checks, or code review.
5. **Done**: Fully verified tasks with acceptance criteria fulfilled, code committed, and deliverables attached.

Below is the complete, comprehensive specification of all project issues. You can copy and paste each issue directly into **GitHub Issues** and link them to your **GitHub Project**.

---

### Issue #1: `[DESIGN] Technical Architecture & von Neumann Specification`
- **Status:** Done
- **Labels:** `architecture`, `documentation`, `design`
- **Milestone:** Phase 1: Architectural Design
- **Description:**
  Formalize the theoretical model and hardware block specifications for the simulated 8-bit computer architecture based on the classical von Neumann model and x86-compatible conventions.
- **Acceptance Criteria:**
  - [x] Define functional block diagram: Control Unit (Sequencer, IR, Decoder, PC), Datapath (ALU, AX, BX, Flags), Memory Interface (MAR, MDR), and Main Memory.
  - [x] Specify register sizes (8 bits, range `00h` to `FFh` / 0 to 255).
  - [x] Specify ALU condition flags: Zero Flag (ZF), Carry Flag (CF), and Sign Flag (SF).
  - [x] Define system buses: 8-bit Address Bus, 8-bit Data Bus, and Control Lines (READ, WRITE, CLOCK).
- **Deliverables & Attachments to Include:**
  - Attach Mermaid architecture diagram block code.
  - Attach screenshot/diagram of von Neumann blocks and bus interconnection.
  - Link to architecture section in [README.md](file:///c:/Users/Javier/OneDrive/Documentos/Arquitectura%20Computadora/Simulador_primer_Examen/von-neumann-simulador-cpu-8bits-excel-vba-main/von-neumann-simulador-cpu-8bits-excel-vba-main/Proyecto_CPU/README.md).
- **Semantic Commit:** `docs(arch): define von Neumann 8-bit architecture and system blocks`

---

### Issue #2: `[DESIGN] Memory Map Specification & Logical Segmentation`
- **Status:** Done
- **Labels:** `memory`, `architecture`, `design`
- **Milestone:** Phase 1: Architectural Design
- **Description:**
  Design the 256-byte main memory organization (00h to FFh) with clear logical segmentation separating program instructions from runtime variable storage.
- **Acceptance Criteria:**
  - [x] Divide memory space: Code Segment (CS: `00h` - `7Fh`, 128 bytes) and Data Segment (DS: `80h` - `FFh`, 128 bytes).
  - [x] Define memory cell data width: 1 byte (8 bits) per addressable location.
  - [x] Define color-coding conventions for GUI matrix: Soft Blue for Code Segment, Soft Green for Data Segment, Amber for Memory Read, Coral for Memory Write, Cyan for PC pointer.
  - [x] Define primitive function signatures: `Read(address)` and `Write(address, value)`.
- **Deliverables & Attachments to Include:**
  - Memory map table diagram (00h-7Fh vs 80h-FFh).
  - Specification of read/write memory timing and bus latching.
- **Semantic Commit:** `docs(memory): specify 256-byte memory mapping and code/data segmentation`

---

### Issue #3: `[DESIGN] 8-Bit Instruction Set Architecture (ISA) & Opcode Matrix`
- **Status:** Done
- **Labels:** `isa`, `assembly`, `design`
- **Milestone:** Phase 1: Architectural Design
- **Description:**
  Establish the formal Instruction Set Architecture (ISA) for the 8-bit CPU. Map each mnemonic to a systematic 1-byte opcode, byte lengths (1 or 2 bytes), operand addressing modes, and affected flags.
- **Acceptance Criteria:**
  - [x] Data transfer instructions: `MOV reg, imm` (2B), `MOV reg, reg` (1B), `LOAD reg, [dir]` (2B), `STORE [dir], reg` (2B).
  - [x] Arithmetic & logic instructions: `ADD` (1B/2B), `SUB` (1B/2B), `INC` (1B), `DEC` (1B), `CMP` (1B/2B), `AND` (1B/2B), `OR` (1B/2B), `XOR` (1B/2B), `NOT` (1B).
  - [x] Control flow instructions: `JMP dir` (2B), `JZ dir` (2B), `JNZ dir` (2B), `HLT` (1B, opcode `00h`), `NOP` (1B).
  - [x] Systematic opcode encoding scheme: high nibble indicates operation family; low nibble indicates operand mode.
- **Deliverables & Attachments to Include:**
  - Full markdown table with all opcodes, mnemonics, byte sizes, and RTL behavior.
- **Semantic Commit:** `docs(isa): specify formal 8-bit opcode table and instruction encoding`

---

### Issue #4: `[FEAT] Main Memory Subsystem & Primitives (ModMemoria.bas)`
- **Status:** Done
- **Labels:** `feature`, `vba`, `memory`
- **Milestone:** Phase 2: Core Processing Engine
- **Description:**
  Implement the memory subsystem in VBA with an internal array of 256 bytes, dedicated primitive subroutines `Read(address)` and `Write(address, value)`, and interactive synchronization with the Excel worksheet.
- **Acceptance Criteria:**
  - [x] Create module `ModMemoria.bas` with global array `MemRAM(0 To 255) As Byte`.
  - [x] Implement `Read(ByVal address As Long) As Byte` which updates audit logs and highlights memory read cells.
  - [x] Implement `Write(ByVal address As Long, ByVal value As Byte)` which updates `MemRAM`, updates worksheet cell, and highlights memory write cells.
  - [x] Implement helper functions: `Memoria_Reset()`, `Memoria_RefrescarTodaUI()`, `Hex2(v)`, `Bin8(v)`.
- **Deliverables & Attachments to Include:**
  - Source code file: [ModMemoria.bas](file:///c:/Users/Javier/OneDrive/Documentos/Arquitectura%20Computadora/Simulador_primer_Examen/von-neumann-simulador-cpu-8bits-excel-vba-main/von-neumann-simulador-cpu-8bits-excel-vba-main/Proyecto_CPU/src/ModMemoria.bas).
  - Unit test confirmation demonstrating `Write(0x80, 42)` and `Read(0x80) == 42`.
- **Semantic Commit:** `feat(memory): implement 256-byte RAM and Read/Write primitive subroutines`

---

### Issue #5: `[FEAT] CPU Registers & Control State Management (ModCPU.bas)`
- **Status:** Done
- **Labels:** `feature`, `vba`, `cpu`
- **Milestone:** Phase 2: Core Processing Engine
- **Description:**
  Model all physical CPU registers (`PC`, `IR_Opcode`, `IR_Operando`, `MAR`, `MDR`, `AX`, `BX`), status flags (`ZF`, `CF`, `SF`), and clock state in clean, modular VBA code.
- **Acceptance Criteria:**
  - [x] Declare 8-bit registers: `PC`, `MAR`, `MDR`, `AX`, `BX`, `IR_Opcode`, `IR_Operando`.
  - [x] Declare 1-bit flags: `ZF` (Zero), `CF` (Carry), `SF` (Sign).
  - [x] Declare machine execution state: `FaseActual`, `NumPasoCiclo`, `ContadorCiclos`, `ContadorInstrucciones`, `Halted`, `EnEjecucion`.
  - [x] Implement `CPU_Reset()` to restore all registers to `00h` and reset flags.
  - [x] Implement `CPU_ActualizarUI()` to refresh register cells on the Excel sheet in Hex, Dec, and Binary formats.
- **Deliverables & Attachments to Include:**
  - Source code file: [ModCPU.bas](file:///c:/Users/Javier/OneDrive/Documentos/Arquitectura%20Computadora/Simulador_primer_Examen/von-neumann-simulador-cpu-8bits-excel-vba-main/von-neumann-simulador-cpu-8bits-excel-vba-main/Proyecto_CPU/src/ModCPU.bas).
  - Screenshot of CPU register bank dashboard on Excel.
- **Semantic Commit:** `feat(cpu): implement CPU register bank, status flags, and state management`

---

### Issue #6: `[FEAT] Arithmetic Logic Unit (ALU) & Flag Calculations (ModALU.bas)`
- **Status:** Done
- **Labels:** `feature`, `vba`, `alu`
- **Milestone:** Phase 2: Core Processing Engine
- **Description:**
  Construct the 8-bit Arithmetic Logic Unit (ALU) supporting all required arithmetic and bitwise operations with mathematically exact flag updates (ZF, CF, SF).
- **Acceptance Criteria:**
  - [x] Implement arithmetic operations: `ADD`, `SUB`, `INC`, `DEC`, and `CMP`.
  - [x] Implement logical operations: `AND`, `OR`, `XOR`, and `NOT`.
  - [x] ZF calculation: `ZF = 1` if 8-bit result is 0, else 0.
  - [x] CF calculation: `CF = 1` if unsigned addition exceeds 255 or unsigned subtraction requires borrow (`valA < valB`).
  - [x] SF calculation: `SF = 1` if MSB (bit 7) is 1 (negative in two's complement).
  - [x] Verify `CMP` updates flags identically to `SUB` without modifying destination register.
- **Deliverables & Attachments to Include:**
  - Source code file: [ModALU.bas](file:///c:/Users/Javier/OneDrive/Documentos/Arquitectura%20Computadora/Simulador_primer_Examen/von-neumann-simulador-cpu-8bits-excel-vba-main/von-neumann-simulador-cpu-8bits-excel-vba-main/Proyecto_CPU/src/ModALU.bas).
  - Test table with corner cases (e.g., `255 + 1`, `42 - 42`, `5 - 7`, bitwise operations).
- **Semantic Commit:** `feat(alu): implement 8-bit ALU operations and precise ZF, CF, SF calculations`

---

### Issue #7: `[FEAT] Complete 4-Phase Instruction Cycle Engine (ModCicloInstruccion.bas)`
- **Status:** Done
- **Labels:** `feature`, `vba`, `control-unit`, `cycle`
- **Milestone:** Phase 2: Core Processing Engine
- **Description:**
  Decompose CPU execution into the four authentic hardware clock phases: Fetch, Decode, Execute, and Store (Write-back). Implement single-phase step, instruction step, and continuous run modes.
- **Acceptance Criteria:**
  - [x] `Fase_Fetch()`: `MAR <- PC`, `MDR <- Read(MAR)`, `IR_Opcode <- MDR`, `PC <- PC + 1`.
  - [x] `Fase_Decode()`: Decode opcode in IR; if instruction takes 2 bytes, fetch operand: `MAR <- PC`, `MDR <- Read(MAR)`, `IR_Operando <- MDR`, `PC <- PC + 1`.
  - [x] `Fase_Execute()`: Call ALU for arithmetic/logic, evaluate conditional branch conditions (`JMP`, `JZ`, `JNZ`), or prepare memory address for `LOAD`/`STORE`.
  - [x] `Fase_Store()`: Write result into destination register (`AX`, `BX`) or write to RAM via `Write(MAR, MDR)` for `STORE`; update PC on branch taken.
  - [x] Implement `Paso_Fase()` (single micro-step button), `Paso_InstruccionCompleta()`, `Modo_Continuo()` with adjustable delay, and `Pausar_Ejecucion()`.
- **Deliverables & Attachments to Include:**
  - Source code file: [ModCicloInstruccion.bas](file:///c:/Users/Javier/OneDrive/Documentos/Arquitectura%20Computadora/Simulador_primer_Examen/von-neumann-simulador-cpu-8bits-excel-vba-main/von-neumann-simulador-cpu-8bits-excel-vba-main/Proyecto_CPU/src/ModCicloInstruccion.bas).
  - State machine transition diagram for Fetch -> Decode -> Execute -> Store.
- **Semantic Commit:** `feat(cycle): implement 4-phase instruction cycle and execution modes`

---

### Issue #8: `[FEAT] Two-Pass Assembler & Machine Code Loader (ModEnsamblador.bas)`
- **Status:** Done
- **Labels:** `feature`, `vba`, `assembler`
- **Milestone:** Phase 3: Assembler & Test Programs
- **Description:**
  Create an in-sheet two-pass assembler that parses mnemonic text, resolves labels (e.g. `BUCLE:`), generates machine bytes, writes them to the sheet table, and loads them into RAM starting at `00h`.
- **Acceptance Criteria:**
  - [x] Pass 1: Parse labels, strip comments (`;`), and compute instruction memory offsets.
  - [x] Pass 2: Translate instructions into machine bytes (1-byte or 2-byte opcodes) and resolve jump target addresses.
  - [x] Parse numbers in hexadecimal (`0x..`, `..h`), binary (`..b`), and decimal formats.
  - [x] Support memory operand brackets (e.g. `[80h]`, `[0x80]`).
  - [x] Load machine code directly into `MemRAM` and refresh the 16x16 grid on the worksheet.
- **Deliverables & Attachments to Include:**
  - Source code file: [ModEnsamblador.bas](file:///c:/Users/Javier/OneDrive/Documentos/Arquitectura%20Computadora/Simulador_primer_Examen/von-neumann-simulador-cpu-8bits-excel-vba-main/von-neumann-simulador-cpu-8bits-excel-vba-main/Proyecto_CPU/src/ModEnsamblador.bas).
  - Assembler test table demonstrating encoding for MOV, ADD, SUB, STORE, LOAD, JNZ.
- **Semantic Commit:** `feat(asm): implement 2-pass in-sheet assembler and label resolution`

---

### Issue #9: `[FEAT] Mandatory Demonstration Programs Suite`
- **Status:** Done
- **Labels:** `programs`, `assembly`, `demo`
- **Milestone:** Phase 3: Assembler & Test Programs
- **Description:**
  Embed the 4 mandatory demonstration programs with loops, branching, and memory writes as required by Section 2.6 of the exam guidelines.
- **Acceptance Criteria:**
  - [x] Program 1: Multiplication by successive addition (`6 * 7 = 42`), saving `42` (`2Ah`) into Data Segment `RAM[80h]`.
  - [x] Program 2: Fibonacci sequence generation, saving consecutive terms into `RAM[80h]`, `RAM[81h]`, `RAM[82h]`, `RAM[83h]`.
  - [x] Program 3: Factorial computation (`4! = 24`), saving intermediate and final product to RAM.
  - [x] Program 4: Countdown loop (`5` down to `0`) with conditional data storage in RAM.
  - [x] Provide 1-click loading buttons for each program in the Excel user interface.
- **Deliverables & Attachments to Include:**
  - Assembly source files and comments for all 4 programs.
  - Trace of registers and memory for Program 1 (`6 * 7 = 42`).
- **Semantic Commit:** `feat(demo): embed 4 mandatory test programs with 1-click loader buttons`

---

### Issue #10: `[FEAT] Excel Graphical User Interface & Interactive Matrix (ModInterfaz.bas)`
- **Status:** Done
- **Labels:** `ui`, `excel`, `vba`, `ergonomics`
- **Milestone:** Phase 4: GUI & Visual Real-Time Animation
- **Description:**
  Design an ergonomic, single-screen dashboard in Excel featuring dynamic phase card illumination, memory cell highlights, register badges, real-time micro-operation logging, and memory inspector.
- **Acceptance Criteria:**
  - [x] 16x16 Main Memory matrix displaying 256 cells with hexadecimal values.
  - [x] Color-coded logical segmentation (Code Segment 00h-7Fh vs Data Segment 80h-FFh).
  - [x] Dynamic cell highlighting for PC pointer (Cyan), Memory Read (Amber), and Memory Write (Coral).
  - [x] 4 Phase cards (Fetch, Decode, Execute, Store) dynamically illuminated with Emerald Green, Cobalt Blue, Orange, and Purple during execution.
  - [x] Chronological micro-operations log panel updating each clock step (`[001] FETCH: ...`).
  - [x] Memory inspector card showing Hex, Decimal, Binary, and Mnemonic interpretation of selected cell.
  - [x] Macro buttons: `CARGAR PROGRAMA`, `PASO A PASO`, `INSTRUCCION >>`, `CONTINUO (RUN)`, `PAUSAR`, `REINICIAR`.
- **Deliverables & Attachments to Include:**
  - Source code file: [ModInterfaz.bas](file:///c:/Users/Javier/OneDrive/Documentos/Arquitectura%20Computadora/Simulador_primer_Examen/von-neumann-simulador-cpu-8bits-excel-vba-main/von-neumann-simulador-cpu-8bits-excel-vba-main/Proyecto_CPU/src/ModInterfaz.bas).
  - Screenshots of the GUI running in Fetch, Decode, Execute, and Store states.
- **Semantic Commit:** `feat(gui): build Excel simulator dashboard with real-time dynamic highlights`

---

### Issue #11: `[INFRA] Automated Workbook Builder Script (Generar-Simulador-Javier.ps1)`
- **Status:** Done
- **Labels:** `automation`, `powershell`, `build`
- **Milestone:** Phase 5: Build Automation & Testing
- **Description:**
  Create an automated PowerShell script that builds the complete macro-enabled workbook `Simulador_CPU_8bits_Javier_Torrico.xlsm` from scratch using Excel COM, formatting all cells, drawing buttons, and importing VBA modules.
- **Acceptance Criteria:**
  - [x] Automated script: [Generar-Simulador-Javier.ps1](file:///c:/Users/Javier/OneDrive/Documentos/Arquitectura%20Computadora/Simulador_primer_Examen/von-neumann-simulador-cpu-8bits-excel-vba-main/von-neumann-simulador-cpu-8bits-excel-vba-main/Proyecto_CPU/scripts/Generar-Simulador-Javier.ps1).
  - [x] Formats sheet headers, register boxes, memory matrix, and buttons with exact RGB styling.
  - [x] Imports all 6 VBA modules into the VBProject without syntax or compilation errors.
  - [x] Pre-populates Example 1 (Multiplication) into editor and RAM.
  - [x] Saves workbook cleanly as `.xlsm` (format 52).
- **Deliverables & Attachments to Include:**
  - Script file: `scripts/Generar-Simulador-Javier.ps1`.
  - Console execution log demonstrating exit code 0.
- **Semantic Commit:** `infra(build): create PowerShell automated Excel workbook builder script`

---

### Issue #12: `[TEST] Unit Verification Suite & Cycle Trace Validation`
- **Status:** Done
- **Labels:** `testing`, `validation`, `qa`
- **Milestone:** Phase 5: Build Automation & Testing
- **Description:**
  Execute comprehensive verification testing across the ALU, memory read/write primitives, 2-pass assembler, and end-to-end execution of Program 1 to ensure 100% correctness according to rubric.
- **Acceptance Criteria:**
  - [x] ALU tests: addition with carry (`200 + 60 -> CF=1`), subtraction (`42 - 42 -> ZF=1`, `5 - 7 -> CF=1, SF=1`), logic ops (`AND`, `OR`, `XOR`, `NOT`).
  - [x] Memory tests: `Write(0x80, 42)` and `Read(0x80)` returns 42.
  - [x] Instruction cycle tests: 4 phases execute sequentially for `MOV AX, 0`.
  - [x] Program 1 test: runs to `HLT` with `AX = 42` and `RAM[80h] = 42` (`2Ah`).
- **Deliverables & Attachments to Include:**
  - Test execution logs.
  - Mathematical trace table for each step of the multiplication loop.
- **Semantic Commit:** `test(core): verify ALU flags, memory primitives, and cycle execution`

---

### Issue #13: `[DOCS] Comprehensive Technical Documentation (README.md)`
- **Status:** Done
- **Labels:** `documentation`, `readme`, `mermaid`
- **Milestone:** Phase 6: Documentation & Deliverables
- **Description:**
  Author a professional, high-standard `README.md` in Markdown containing architecture diagrams in Mermaid, formal ISA opcode table, step-by-step user guide, program mathematical trace, and Kanban links under author **Javier Torrico Sejas**.
- **Acceptance Criteria:**
  - [x] Academic header: Universidad Católica Boliviana "San Pablo", SIS-131, Author Javier Torrico Sejas, Professor Ing. Paulo César Loayza Carrasco.
  - [x] Mermaid architecture diagram showing CPU, ALU, registers, buses, and RAM.
  - [x] Formal ISA table (opcodes, sizes, flags, description).
  - [x] Step-by-step user guide for opening `.xlsm`, loading code, stepping, running, and inspecting.
  - [x] Complete cycle-by-cycle mathematical trace for the demo program.
  - [x] Links to GitHub repository and GitHub Projects Kanban board.
- **Deliverables & Attachments to Include:**
  - [README.md](file:///c:/Users/Javier/OneDrive/Documentos/Arquitectura%20Computadora/Simulador_primer_Examen/von-neumann-simulador-cpu-8bits-excel-vba-main/von-neumann-simulador-cpu-8bits-excel-vba-main/Proyecto_CPU/README.md) file.
- **Semantic Commit:** `docs(readme): author comprehensive technical documentation with Mermaid and ISA tables`

---

### Issue #14: `[DOCS] 15-Minute Oral Defense Script & Conceptual Q&A Guide`
- **Status:** Done
- **Labels:** `documentation`, `defense`, `presentation`
- **Milestone:** Phase 6: Documentation & Deliverables
- **Description:**
  Prepare a strict 15-minute oral defense preparation guide adhering to the evaluation rubric (Minutes 0-2: Architecture, 2-4: GitHub & Kanban, 4-10: Live Simulator Demo, 10-15: Code Q&A and Live Memory Modification).
- **Acceptance Criteria:**
  - [x] Minute-by-minute speaking script for the student Javier Torrico Sejas.
  - [x] Step-by-step walkthrough of what buttons to click and what to highlight during the live demo.
  - [x] Anticipated questions from the professor regarding VBA code lines with precise technical answers.
  - [x] Practical live modification recipe: how to change an instruction in live RAM or in the editor on the spot.
- **Deliverables & Attachments to Include:**
  - [GUIA_DEFENSA_ORAL.md](file:///c:/Users/Javier/OneDrive/Documentos/Arquitectura%20Computadora/Simulador_primer_Examen/von-neumann-simulador-cpu-8bits-excel-vba-main/von-neumann-simulador-cpu-8bits-excel-vba-main/Proyecto_CPU/GUIA_DEFENSA_ORAL.md).
- **Semantic Commit:** `docs(defense): prepare 15-minute oral defense script and live Q&A guide`

---

## Summary of Files to Attach / Include in the GitHub Repository

| Folder / File | Type | Purpose & Content |
|---|---|---|
| [`Simulador_CPU_8bits_Javier_Torrico.xlsm`](file:///c:/Users/Javier/OneDrive/Documentos/Arquitectura%20Computadora/Simulador_primer_Examen/von-neumann-simulador-cpu-8bits-excel-vba-main/von-neumann-simulador-cpu-8bits-excel-vba-main/Proyecto_CPU/Simulador_CPU_8bits_Javier_Torrico.xlsm) | Excel Workbook | Macro-enabled simulator ready to run with GUI, memory matrix, controls, and pre-loaded programs. |
| [`README.md`](file:///c:/Users/Javier/OneDrive/Documentos/Arquitectura%20Computadora/Simulador_primer_Examen/von-neumann-simulador-cpu-8bits-excel-vba-main/von-neumann-simulador-cpu-8bits-excel-vba-main/Proyecto_CPU/README.md) | Technical Docs | Main project documentation with Mermaid diagram, formal ISA table, user manual, and execution trace. |
| [`GITHUB_ISSUES_KANBAN_EN.md`](file:///c:/Users/Javier/OneDrive/Documentos/Arquitectura%20Computadora/Simulador_primer_Examen/von-neumann-simulador-cpu-8bits-excel-vba-main/von-neumann-simulador-cpu-8bits-excel-vba-main/Proyecto_CPU/GITHUB_ISSUES_KANBAN_EN.md) | Agile Docs | Complete list of all 14 project issues in English with descriptions, acceptance criteria, and attachments. |
| [`GUIA_DEFENSA_ORAL.md`](file:///c:/Users/Javier/OneDrive/Documentos/Arquitectura%20Computadora/Simulador_primer_Examen/von-neumann-simulador-cpu-8bits-excel-vba-main/von-neumann-simulador-cpu-8bits-excel-vba-main/Proyecto_CPU/GUIA_DEFENSA_ORAL.md) | Defense Script | 15-minute defense guide with minute-by-minute script, code explanation, and live modification recipe. |
| `src/ModMemoria.bas` | VBA Module | 256-byte RAM, primitive subroutines `Read()` and `Write()`, logical segmentation. |
| `src/ModCPU.bas` | VBA Module | Physical CPU registers (`PC`, `IR`, `MAR`, `MDR`, `AX`, `BX`), flags (`ZF`, `CF`, `SF`), CPU reset. |
| `src/ModALU.bas` | VBA Module | Arithmetic & logic unit (`ADD`, `SUB`, `INC`, `DEC`, `AND`, `OR`, `XOR`, `NOT`, `CMP`) and flag logic. |
| `src/ModCicloInstruccion.bas` | VBA Module | 4-phase instruction cycle (`Fetch`, `Decode`, `Execute`, `Store`), Step mode, and Run mode. |
| `src/ModEnsamblador.bas` | VBA Module | In-sheet 2-pass assembler, label resolver, and 4 mandatory demo programs. |
| `src/ModInterfaz.bas` | VBA Module | GUI formatting, dynamic phase illumination, memory cell highlights, and log panel. |
| `scripts/Generar-Simulador-Javier.ps1` | PowerShell | Automated build script to generate the entire Excel simulator from scratch. |
