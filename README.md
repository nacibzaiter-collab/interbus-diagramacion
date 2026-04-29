# Diagramación de Servicios — INTERBUS SRL

## Archivos del proyecto
```
interbus/
├── app.py              ← Backend Flask (toda la lógica)
├── requirements.txt    ← Dependencias Python
├── render.yaml         ← Config para Render.com
└── templates/
    └── index.html      ← Frontend (Android-friendly)
```

## Cómo deployar en Render.com (gratis)

1. Crear cuenta en https://render.com
2. Subir estos archivos a un repositorio GitHub (nuevo repo, público o privado)
3. En Render → "New Web Service" → conectar el repo
4. Render detecta el `render.yaml` automáticamente
5. Click en "Deploy" — en 2-3 minutos tenés la URL

## Uso de la app

### Primera vez
1. Abrís la URL en el celular
2. Subís el archivo **PLANILLAS MAESTRO.xlsm** (se guarda en el servidor)
3. Subís la **Planilla de Citación** semanal
4. Seleccionás el tipo de semana
5. Generás → descargás el Excel

### Semanas siguientes
1. Solo subís la nueva **Planilla de Citación**
2. El Maestro ya está guardado (solo lo actualizás cuando cambia)

## Notas técnicas
- El Maestro se guarda en `/tmp/interbus_uploads/maestro_actual.xlsm`
- En Render free tier `/tmp` se resetea con cada deploy — si actualizan el código,
  hay que volver a subir el Maestro una vez
- El contador de planillas se guarda en localStorage del navegador
- Sesiones de Flask guardan las planillas generadas para el download
