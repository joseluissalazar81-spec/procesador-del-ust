# Procesador de Planificaciones DEL — UST 2026-1

Herramienta para revisar y corregir planificaciones didácticas antes de enviarlas a revisión DEL.

## Qué hace

- Corrige automáticamente procedimientos, instrumentos, tipos de evaluación, lenguaje e incoherencias formales
- Verifica los 41 criterios de la escala de apreciación UST
- Cruza la planificación contra el programa oficial (PDF)
- Verifica hitos del lineamiento A+Se si corresponde
- Genera el archivo corregido con cambios en azul, filas Formativa en lila y Sumativa en amarillo

## Archivos

| Archivo | Descripción |
|---------|-------------|
| `app_del.py` | Aplicación Streamlit (interfaz web) |
| `revisar_planificaciones.py` | Script revisor principal |
| `Procesar_DEL.ipynb` | Notebook Google Colab (alternativa sin instalación) |
| `requirements.txt` | Dependencias Python |

## Uso local

```bash
pip install -r requirements.txt
streamlit run app_del.py
```

## Despliegue en Streamlit Cloud

1. Fork o sube este repositorio a GitHub
2. Ir a [share.streamlit.io](https://share.streamlit.io)
3. Conectar el repositorio → seleccionar `app_del.py`
4. Compartir la URL con el equipo DEL

## Flujo de trabajo

```
Docente entrega PDF programa + xlsx planificación
              ↓
     Subir a la herramienta
              ↓
   Correcciones automáticas aplicadas
   Verificación 41 criterios UST
   Verificación vs programa oficial
   (opcional) Verificación hitos A+Se
              ↓
   Descargar xlsx corregido + log
              ↓
     Enviar a revisión DEL
```

## Opciones

| Opción | Cuándo activar |
|--------|---------------|
| `ES_AS = True` (Colab) / checkbox A+S (app) | Solo si la asignatura tiene lineamiento Aprendizaje + Servicio |
| Decreto de actualización | Próximamente — pendiente de ejemplo real |
