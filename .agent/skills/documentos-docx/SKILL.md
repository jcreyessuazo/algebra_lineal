---
name: documentos-docx
description: >-
  Usa esta skill siempre que el usuario pida crear, redactar o generar cualquier
  documento, guía de ejercicios, guía de materia, resumen de estudio, apuntes,
  evaluación o informe. Garantiza que todo documento generado se entregue en
  formato Microsoft Word (.docx) con formato profesional.
---

# Generación de Documentos y Guías en Formato DOCX

Todo documento, guía de ejercicios o guía de materia solicitado por el usuario **debe** entregarse como un archivo `.docx` real en el directorio de trabajo (nunca únicamente como texto plano o Markdown, a menos que el usuario pida explícitamente otro formato).

## Flujo de Trabajo

1. **Diseñar el contenido según el tipo de documento:**
   - **Guía de Materia / Estudio:**
     - Título principal y metadatos (Asignatura, Tema, Fecha).
     - Objetivos de aprendizaje.
     - Desarrollo teórico estructurado con subtítulos (`Heading 1`, `Heading 2`, `Heading 3`), definiciones destacadas, tablas comparativas y ejemplos paso a paso.
     - Resumen o puntos clave al final.
   - **Guía de Ejercicios:**
     - Encabezado con título, asignatura, instrucciones generales y tiempo estimado o puntaje (si aplica).
     - Ejercicios organizados por nivel de dificultad (Básico, Intermedio, Avanzado) o por ítem/sección.
     - Enunciados claros y numerados.
     - Sección final de **Solucionario / Respuestas** (desarrollada paso a paso o con respuestas finales, según corresponda).
   - **Documento General / Informe:**
     - Portada o encabezado formal, introducción, secciones de desarrollo, tablas de datos y conclusión.

2. **Generar el archivo `.docx`:**
   - **Opción A (Conversión rápida desde Markdown):** Escribir o pasar un archivo Markdown al script auxiliar [md_to_docx.py](./scripts/md_to_docx.py):
     ```powershell
     python ".agent/skills/documentos-docx/scripts/md_to_docx.py" "entrada.md" "salida.docx"
     ```
   - **Opción B (Script `python-docx` a medida):** Para documentos con diseño avanzado (tablas personalizadas, celdas con color, encabezados/pies de página complejos), ejecutar un script en Python usando `python-docx` aplicando las reglas de estilo de abajo.

3. **Verificación:**
   - Comprobar que el archivo `.docx` existe en la ruta esperada y que su tamaño es mayor a 0 bytes.
   - Si se creó algún archivo temporal `.md` o `.py` para generar el `.docx`, eliminarlo o guardarlo en la carpeta `scratch/` para mantener limpio el espacio de trabajo del usuario.
   - Entregar al usuario el enlace directo al archivo `.docx` creado.

## Estándares de Estilo en `python-docx`

- **Tipografía:** `Calibri` o `Arial`, tamaño `11 pt` para cuerpo de texto, interlineado `1.15`, espaciado posterior de `6 pt`.
- **Títulos:**
  - Título del documento: `18 pt`, Negrita, color azul oscuro (`#1F4E79`), centrado.
  - `Heading 1`: `14 pt`, Negrita, color `#1F4E79`.
  - `Heading 2`: `12 pt`, Negrita, color `#2E75B6`.
- **Tablas:** Estilo `Table Grid`, fila de encabezado con fondo `#1F4E79` y texto blanco en negrita.
- **Bloques de fórmulas, código o notas importantes:** Fuente `Consolas` (`10 pt`) o párrafo con sangría y borde/sombreado suave (`#F2F2F2`).
