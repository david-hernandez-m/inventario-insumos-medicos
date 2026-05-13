# Inventario de Insumos Médicos v4.3

Aplicación de escritorio desarrollada en **Python + Tkinter + SQLite** para gestionar inventario de medicamentos e insumos médicos, registrar movimientos de stock, generar reportes y trabajar con carga masiva desde Excel.

Este proyecto fue pensado como una solución simple, local y práctica para centros de salud, bodegas clínicas, postas, ambulancias, CESFAM o cualquier entorno donde se necesite controlar insumos de forma rápida y ordenada.

---

## Tabla de contenidos

1. [Descripción general](#descripción-general)
2. [Funciones principales](#funciones-principales)
3. [Tecnologías utilizadas](#tecnologías-utilizadas)
4. [Estructura del proyecto](#estructura-del-proyecto)
5. [Requisitos](#requisitos)
6. [Instalación](#instalación)
7. [Cómo ejecutar la aplicación](#cómo-ejecutar-la-aplicación)
8. [Inicio de sesión](#inicio-de-sesión)
9. [Uso de la aplicación](#uso-de-la-aplicación)
10. [Carga masiva desde Excel](#carga-masiva-desde-excel)
11. [Exportaciones](#exportaciones)
12. [Cambio de contraseña](#cambio-de-contraseña)
13. [Respaldo de base de datos](#respaldo-de-base-de-datos)
14. [Base de datos](#base-de-datos)
15. [Problemas comunes y soluciones](#problemas-comunes-y-soluciones)
16. [Ideas de mejora futura](#ideas-de-mejora-futura)
17. [Autor y propósito](#autor-y-propósito)

---

## Descripción general

**Inventario de Insumos Médicos v4.3** es una aplicación de escritorio que permite:

- registrar medicamentos e insumos médicos,
- controlar stock actual y stock mínimo,
- registrar entradas y salidas,
- revisar movimientos históricos,
- visualizar alertas por bajo stock y vencimientos,
- importar datos desde Excel,
- exportar datos a CSV, Excel y PDF,
- generar respaldos de la base de datos,
- administrar el acceso mediante usuario y contraseña.

La aplicación funciona de forma **local**, sin necesidad de internet ni servidor externo, usando una base de datos **SQLite** que se crea automáticamente en la carpeta del proyecto.

---

## Funciones principales

### 1. Gestión de inventario
Permite registrar insumos con los siguientes datos:

- código
- nombre
- categoría
- stock actual
- stock mínimo
- unidad
- fecha de vencimiento
- ubicación
- proveedor
- lote
- observaciones

También permite editar o eliminar registros existentes.

### 2. Registro de movimientos
Cada vez que se necesita agregar o retirar stock, se puede registrar un movimiento indicando:

- insumo
- tipo de movimiento (`Entrada` o `Salida`)
- cantidad
- motivo
- usuario que realizó la acción

La aplicación actualiza el stock automáticamente y deja trazabilidad.

### 3. Alertas
La app muestra:

- insumos con stock igual o menor al mínimo,
- insumos próximos a vencer,
- insumos ya vencidos.

### 4. Importación desde Excel
Puedes cargar una lista de medicamentos o insumos de forma masiva desde un archivo `.xlsx`.

Si el código del medicamento no existe, se crea un registro nuevo.  
Si el código ya existe, se actualiza el registro y se ajusta el stock.

### 5. Exportación de información
Puedes exportar inventario y movimientos en:

- CSV
- Excel
- PDF

### 6. Seguridad básica
Incluye:

- inicio de sesión,
- cambio de contraseña,
- cierre de sesión.

### 7. Respaldo
Permite generar una copia de seguridad de la base de datos en la carpeta `respaldos/`.

---

## Tecnologías utilizadas

- **Python 3**
- **Tkinter** para la interfaz gráfica
- **SQLite** para la base de datos local
- **openpyxl** para importación y exportación en Excel
- **reportlab** para exportación en PDF

---

## Estructura del proyecto

Ejemplo de estructura esperada:

```bash
inventario_salud/
│
├── app_inventario_salud_v43.py
├── README.md
├── plantilla_medicamentos_v43.xlsx
├── inventario_salud_v43.db
└── respaldos/