# Guía Estratégica para la Defensa Oral (15 Minutos)

**Materia:** Arquitectura de Computadoras (SIS-131) — Primer Parcial  
**Docente:** Ing. Paulo César Loayza Carrasco  
**Estudiante:** Javier Torrico Sejas  
**Proyecto:** Simulador de CPU von Neumann (8 bits) y Memoria Principal en Excel + VBA  

---

## 1. Cronograma Estricto de la Defensa (15 Minutos)

| Bloque de Tiempo | Fase | Objetivo Clave | Qué Mostrar en Pantalla |
|:---:|:---|:---|:---|
| **00:00 – 02:00** (2 min) | **Arquitectura y Hardware** | Explicar el modelo von Neumann, registros de 8 bits, mapa de memoria 16×16 y la ISA. | Diagrama Mermaid en `README.md` y vista general del simulador en Excel. |
| **02:00 – 04:00** (2 min) | **GitHub y Kanban** | Demostrar metodología ágil, tablero Kanban en GitHub Projects y commits semánticos. | Pestaña del navegador con GitHub Projects (5 columnas) y log de commits en Git. |
| **04:00 – 10:00** (6 min) | **Demostración en Vivo** | Ejecutar el programa paso a paso (Fetch -> Decode -> Execute -> Store), bucles y flags. | Excel en vivo pulsando `PASO A PASO` y luego `CONTINUO (RUN)`. |
| **10:00 – 15:00** (5 min) | **Preguntas y Código en Vivo** | Responder preguntas conceptuales, explicar líneas de VBA y modificar una instrucción en vivo. | Editor de VBA (Alt + F11) y modificación de celda en la matriz de RAM en vivo. |

---

## 2. Libreto Paso a Paso para la Exposición

### Minutos 0:00 a 2:00 — Explicación de la Arquitectura
> *"Buenos días, Ingeniero Loayza. Mi nombre es Javier Torrico Sejas y presento el Simulador de CPU de 8 bits bajo arquitectura von Neumann con memoria principal en Excel y VBA.*  
> *El sistema implementa los bloques fundamentales de hardware:*
> 1. *La **Unidad de Control (UC)**, compuesta por el Contador de Programa (`PC`), el Registro de Instrucción (`IR`), el Decodificador y el Secuenciador de Fases del Reloj.*
> 2. *La **Unidad Aritmético-Lógica (ALU)** de 8 bits, que procesa operaciones como `ADD`, `SUB`, `INC`, `DEC`, `CMP` y operaciones lógicas, actualizando estrictamente las banderas de estado: `ZF` (Zero), `CF` (Carry) y `SF` (Sign).*
> 3. *El **Banco de Registros**, con el Acumulador `AX`, el registro general `BX`, y los registros de acople al bus de memoria: `MAR` (Memory Address Register) y `MDR` (Memory Data Register).*
> 4. *La **Memoria Principal (RAM)** de 256 bytes (`00h` a `FFh`) visualizada en una matriz 16×16, con segmentación lógica clara: Segmento de Código (`00h` a `7Fh`) en azul suave y Segmento de Datos (`80h` a `FFh`) en verde suave, interactuando exclusivamente mediante las primitivas dedicadas `Read(address)` y `Write(address, value)`."*

---

### Minutos 2:00 a 4:00 — Evidencia de Gestión en GitHub Projects
> *(Cambia la pestaña del navegador a tu repositorio y tablero Kanban)*  
> *"La gestión del proyecto se llevó a cabo bajo metodología ágil mediante un **Tablero Kanban en GitHub Projects** organizado en 5 columnas: Backlog, To Do, In Progress, In Review/Testing y Done.*  
> *Se desglosaron 14 tareas técnicas con criterios de aceptación explícitos: matriz de memoria, banco de registros, ALU con flags, el secuenciador del ciclo de 4 fases, el ensamblador de dos pasadas y la batería de pruebas.*  
> *El control de versiones siguió un flujo de commits semánticos (`feat:`, `fix:`, `docs:`, `test:`, `infra:`), garantizando un historial atómico y trazable a lo largo de todo el desarrollo."*

---

### Minutos 4:00 a 10:00 — Demostración en Vivo del Simulador
> *(Pasa a Microsoft Excel con `Simulador_CPU_8bits_Javier_Torrico.xlsm`)*  
> 
> **Paso 1: Mostrar el programa cargado**  
> *"Para esta demostración he cargado el Programa Demostrativo Obligatorio: Multiplicación por sumas sucesivas (`6 × 7 = 42`). El programa inicializa `AX = 0`, `BX = 7`, y mediante un bucle suma `6` y decrementa `BX` hasta que `ZF = 1`, guardando el resultado en `RAM[80h]`.*
>
> **Paso 2: Ejecución Paso a Paso (Una Fase a la vez)**  
> *(Pulsa el botón **PASO A PASO** 4 veces para mostrar una instrucción completa)*:
> - **1er Clic (FETCH):** *"Al pulsar Paso a Paso, entramos en la fase **FETCH** (iluminada en verde). El `PC` (`00h`) se transfiere a `MAR`, se invoca la primitiva `Read(00h)` cargando `10h` en `MDR`. El dato pasa a `IR_Opcode` y el `PC` se incrementa a `01h`."*
> - **2do Clic (DECODE):** *"En la fase **DECODE** (iluminada en azul), la Unidad de Control reconoce que el opcode `10h` corresponde a `MOV AX, imm` (2 bytes). Lee el segundo byte de memoria: `MAR = 01h`, `MDR = 00h` (el valor inmediato), y `PC` incrementa a `02h`."*
> - **3er Clic (EXECUTE):** *"En la fase **EXECUTE** (iluminada en naranja), la ALU y el bus preparan el valor inmediato para su transferencia."*
> - **4to Clic (STORE):** *"En la fase **STORE** (iluminada en púrpura), el acumulador `AX` recibe el valor `00h`. La instrucción ha concluido y el panel de log registra cada micro-operación."*
>
> **Paso 3: Ejecución Continua y Salto Condicional**  
> *(Pulsa **CONTINUO (RUN)**)*  
> *"Ahora activamos el modo continuo con control de retardo de reloj. Observemos cómo en cada iteración la celda `BX` decrementa (`6, 5, 4, 3, 2, 1`), la bandera `ZF` permanece en 0, y la instrucción `JNZ BUCLE` efectúa el salto dinámico modificando el `PC` nuevamente a `04h`.*  
> *En la séptima iteración, `BX` llega a 0: la ALU activa la bandera `ZF = 1` (indicador verde encendido). Al llegar a `JNZ`, la condición de salto ya no se cumple y el flujo continúa secuencialmente hacia `STORE [80h], AX`.*  
> *Se activa la primitiva `Write(80h, 2Ah)` en el panel y la celda `80h` de la matriz en el Segmento de Datos se ilumina en rojo coral mostrando el resultado `2A` (42 en decimal).*  
> *Finalmente, se ejecuta `HLT` (opcode `00h`) y el reloj se detiene de forma precisa."*

---

### Minutos 10:00 a 15:00 — Preguntas del Docente y Modificación en Vivo

A continuación tienes las preguntas técnicas exactas que el Ing. Loayza puede realizar y las respuestas que debes dar con total seguridad:

#### Pregunta 1: *"¿Dónde se implementa la primitiva Read y Write en el código VBA?"*
- **Respuesta:**  
  *"Se encuentran en el módulo [`ModMemoria.bas`](file:///c:/Users/Javier/OneDrive/Documentos/Arquitectura%20Computadora/Simulador_primer_Examen/von-neumann-simulador-cpu-8bits-excel-vba-main/von-neumann-simulador-cpu-8bits-excel-vba-main/Proyecto_CPU/src/ModMemoria.bas):*  
  - *La función `Read(ByVal address As Long) As Byte` recibe una dirección de 8 bits, lee del arreglo `MemRAM(addrValida)`, registra la auditoría y llama a `ResaltarCeldaMemoria addrValida, "LECTURA"`.*  
  - *La subrutina `Write(ByVal address As Long, ByVal value As Byte)` escribe en `MemRAM(addrValida) = valByte`, actualiza inmediatamente la celda en la matriz 16×16 en pantalla mediante `Memoria_ActualizarCeldaUI(addrValida)` y llama a `ResaltarCeldaMemoria addrValida, "ESCRITURA"`."*

#### Pregunta 2: *"¿Por qué el PC se incrementa en su propio sumador y no usa la ALU?"*
- **Respuesta:**  
  *"Porque en esta arquitectura el acumulador `AX` es un registro visible para el programador. Si usáramos la ALU para hacer `PC + 1`, sobreescribiríamos o destruiríamos el valor almacenado en `AX` o alteraríamos las banderas de estado (`ZF`, `CF`, `SF`). Por ello, la Unidad de Control dispone de un circuito incrementador dedicado: `PC = (PC + 1) And &HFF`."*

#### Pregunta 3: *"¿Cómo calcula la ALU las banderas ZF, CF y SF en una resta o CMP?"*
- **Respuesta:**  
  *"En [`ModALU.bas`](file:///c:/Users/Javier/OneDrive/Documentos/Arquitectura%20Computadora/Simulador_primer_Examen/von-neumann-simulador-cpu-8bits-excel-vba-main/von-neumann-simulador-cpu-8bits-excel-vba-main/Proyecto_CPU/src/ModALU.bas):*  
  - *Se calcula `resRaw = valA - valB` y `res8 = resRaw And &HFF`.*  
  - *`ZF` (Zero Flag) vale `1` si `res8 = 0`.*  
  - *`CF` (Carry / Préstamo en resta sin signo) vale `1` si `valA < valB` (ocurre un borrow).*  
  - *`SF` (Sign Flag) evalúa el bit más significativo (MSB, bit 7): `If (res8 And &H80) <> 0 Then SF = 1 Else SF = 0`.*  
  - *En `CMP`, el cálculo es idéntico al de `SUB`, pero el resultado no se escribe en ningún registro destino, preservando los datos del programador."*

#### Pregunta 4: *"Modifica una instrucción en vivo en la memoria o en el editor para cambiar el comportamiento del programa."*
- **Cómo hacerlo en vivo en 5 segundos:**  
  1. Si te pide cambiar el multiplicador: en el editor (celda C8), cambia `MOV BX, 7` por `MOV BX, 5`, pulsa **CARGAR PROGRAMA** y pulsa **RUN ▶**. Verás que en `RAM[80h]` ahora se guarda `30` (`1Eh` = `6 × 5 = 30`).  
  2. Si te pide cambiarlo directamente en memoria RAM sin reensamblar:  
     - Ve a la matriz de memoria en la celda **U6** (dirección `03h`, que es el operando inmediato de `MOV BX`).  
     - Escribe `03`.  
     - Pulsa **REINICIAR** y luego **RUN ▶**.  
     - El programa multiplicará `6 × 3 = 18` (`12h`) y lo guardará en `RAM[80h]`.  
  *Esto demuestra ante el docente dominio conceptual absoluto tanto del hardware como de la memoria física.*
