# 📊 Sistema de Gestión de Workforce — Excel VBA

Sistema modular de gestión operacional para **Contact Centers**, desarrollado en Excel con macros VBA. Permite administrar turnos, agentes, cobertura, colaciones y asistencia de forma integrada.

---

## 📁 Archivos del Proyecto

| Archivo | Descripción |
|---|---|
| `Agentes_por_dia_v1.xlsm` | Resumen diario de agentes activos por línea de negocio |
| `Turnos_v1.xlsm` | Gestión y asignación de turnos por agente |
| `Matriz_Cobertura_v1.xlsm` | Matriz de cobertura horaria por intervalos |
| `Malla_Shift_v1.xlsm` | Malla de turnos semanal/mensual |
| `Colaciones_v1.xlsm` | Control de horarios de colación por agente y canal |
| `Asistencia_Cierre_v1.xlsm` | Cierre de asistencia diaria con estados y minutos en cola |

---

## 🔧 Funcionalidades Principales

### 🧑‍💼 Agentes por Día
- Tabla pivot con máximo de agentes activos por día
- Segmentado por líneas: **AT**, **CROSS**, **P**
- Vista de Grand Total diario

### 🕐 Colaciones
- Control de actividades de descanso (comida, break)
- Canales: **FO** (Front Office) y **RRSS** (Redes Sociales)
- Registro de inicio, fin y duración en minutos
- Ajuste de zona horaria a **Santiago (UTC-3/UTC-4)**

### ✅ Asistencia y Cierre
- Estados de asistencia: `Presente`, `Día Libre`, `Festivo`, `Desvinculación`, entre otros
- Minutos en cola por agente y por día
- Procesamiento automatizado con macros VBA

---

## 🚀 Cómo Usar

1. **Descargar** los archivos `.xlsm` desde este repositorio
2. **Abrir** en Microsoft Excel (se recomienda Excel 2016 o superior)
3. Al abrir, **habilitar macros** cuando Excel lo solicite
4. Seguir el orden lógico de carga: Turnos → Malla → Agentes → Cobertura → Colaciones → Asistencia

> ⚠️ **Nota:** Los archivos contienen macros VBA. Asegúrate de habilitarlas para que el sistema funcione correctamente. Descarga solo desde esta fuente confiable.

---

## 🛠️ Requisitos

- Microsoft Excel 2016 o superior
- Macros VBA habilitadas
- Sistema operativo: Windows (recomendado) o macOS con Excel

---

## 📌 Casos de Uso

- Planificación de turnos en operaciones de Contact Center
- Seguimiento de asistencia y estados de agentes
- Control de cobertura por franja horaria
- Gestión de colaciones y tiempos de descanso

---

## 👤 Autor

Desarrollado para uso operacional en Contact Center.  
Compartido con fines educativos y de comunidad en LinkedIn.

---

## 📄 Licencia

Este proyecto se comparte de forma libre para uso personal y educativo.  
Para uso comercial, contactar al autor.
