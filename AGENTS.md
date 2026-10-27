# Reglas del repositorio

## Prioridad máxima en redacción

Antes de redactar o editar cualquier texto del proyecto, lee y aplica la skill `unslop` ubicada en `.agents/skills/unslop/SKILL.md`.

Toda edición sobre el documento o archivos del repositorio debe cumplir estas reglas de escritura:

- Elimina vocabulario típico de IA, frases de relleno y tono promocional.
- Usa voz activa, verbos concretos y datos medibles.
- No uses rayas, guiones largos ni paréntesis como conectores en la prosa. Separa las ideas con punto o coma.
- No uses dos puntos en mitad de una oración. Úsalos solo antes de una lista o tabla.
- Escribe los encabezados con mayúscula solo en la primera palabra y en los nombres propios.
- Evita abusar de las negritas dentro de los párrafos.

## Fuente de verdad y flujo de entrega

El archivo editable y única fuente de verdad del trabajo práctico es `Proyecto_Inversion_Injertos_Patagonia.md`.

- Todos los cambios de contenido, estructura, tablas y bibliografía se hacen únicamente sobre `Proyecto_Inversion_Injertos_Patagonia.md`.
- El archivo `docx-og/Proyecto_Inversion_Injertos_Patagonia_Entregable_v2.5.docx` se conserva sin cambios como respaldo del formato original.
- Las dependencias de Python están declaradas en `requirements.txt` para funcionar en Windows y Linux.

### Ejecución en Linux y Windows

El script `scripts/export_entregable.py` detecta el sistema operativo, crea el entorno virtual `.venv` si no existe, instala `requirements.txt` automáticamente y genera los archivos `entregas/Proyecto_Inversion_Injertos_Patagonia_Entregable.docx` y `entregas/Proyecto_Inversion_Injertos_Patagonia_Entregable.pdf`.

En Linux:

```bash
python3 scripts/export_entregable.py
```

En Windows:

```powershell
python scripts\export_entregable.py
```

Si prefieres inicializar o usar el entorno virtual de forma manual:

- Linux: `python3 -m venv .venv && .venv/bin/pip install -r requirements.txt && .venv/bin/python scripts/export_entregable.py`
- Windows: `python -m venv .venv && .venv\Scripts\pip install -r requirements.txt && .venv\Scripts\python scripts\export_entregable.py`
