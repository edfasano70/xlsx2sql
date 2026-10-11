# xlsx2sql

Aplicación de escritorio con PyQt6 que convierte archivos Excel (`.xlsx`) en
sentencias SQL `INSERT`.

## Características

- Lee todas las hojas de un `.xlsx` y genera un `INSERT` por fila.
- Respeta el tipo real de cada celda (números sin comillas, textos con comillas,
  `NULL` para celdas vacías, booleanos como `1`/`0`, fechas en formato ISO).
- Escapa comillas simples y descarta filas completamente vacías.
- Tema claro/oscuro y persistencia de la geometría de la ventana en
  `~/.config/xlsx2sql/config.xml`.

## Uso

```bash
python main.py
```

1. **Abrir Excel...** (`Ctrl+O`) y selecciona un archivo `.xlsx`.
2. La aplicación muestra el SQL generado.
3. **Guardar SQL...** (`Ctrl+S`) para escribirlo en un archivo `.sql`.

## Desarrollo

```bash
python3 -m pip install --user --break-system-packages PyQt6 pytest ruff openpyxl
pytest
ruff check .
ruff format .
```

## Estructura

```
xlsx2sql/
├── main.py            # toda la app
├── resources/         # logo e iconos
├── src/               # vacío, para módulos futuros
├── tests/             # tests con pytest
├── pyproject.toml
└── README.md
```
