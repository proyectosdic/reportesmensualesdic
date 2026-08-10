
# DIC | Sistema de informes mensuales — V1

Prototipo funcional en Streamlit para la Dirección de Integración Comunitaria (DIC) del ITESO.

## Unidades incluidas
- CUE
- COINCIDE
- CUDJ
- CUI
- CEJUVEN
- CPC
- CEFSI

## Funcionalidad V1
- Captura por mes y año.
- 5 actividades iniciales + botón para agregar más.
- Máximo 250 palabras por descripción.
- Ranking Top 1 / Top 2 / Top 3 sin duplicados.
- Campo de participantes/alcance.
- Carga de fotografías preparada.
- Guardar borrador / enviar reporte.
- Validación de correo @iteso.mx.
- Panel DIC de enviados, borradores y pendientes.
- Edición del texto para consolidado sin borrar el texto original.
- Exportación a Word y PDF.
- Buscador histórico por palabra/tema.
- Supabase como base de datos; si no se configura, funciona en modo demo durante la sesión.

## Cómo probarlo localmente
1. Instala Python 3.11 o superior.
2. Abre una terminal dentro de esta carpeta.
3. Ejecuta:

   pip install -r requirements.txt

4. Luego:

   streamlit run app.py

## Conectar Supabase
1. Crea un proyecto en Supabase.
2. Ejecuta `schema.sql` en el SQL Editor.
3. Crea `.streamlit/secrets.toml`.
4. Copia los valores desde `.streamlit/secrets.toml.example`.
5. Sustituye SUPABASE_URL y SUPABASE_KEY.

## Próxima fase recomendada
- Subir fotografías a Supabase Storage.
- Correo real de confirmación.
- Login institucional/OIDC.
- Row Level Security por centro.
- Recordatorios automáticos a centros pendientes.
- Mejorar el diseño del Word/PDF para que replique el formato institucional.
