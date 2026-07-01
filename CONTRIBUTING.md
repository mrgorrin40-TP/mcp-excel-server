# Contribuir a MCP Excel Server

Gracias por tu interés en contribuir a MCP Excel Server. Este documento proporciona guías y requisitos para contribuir al proyecto.

## Desarrollo Local

### Prerrequisitos

- Python 3.10+
- Poetry
- Git

### Configuración

```bash
# Clonar el repositorio
git clone https://github.com/usuario/mcp-excel-server.git
cd mcp-excel-server

# Instalar dependencias
poetry install

# Instalar dependencias de desarrollo
poetry install --with dev

# Instalar xlwings (opcional, para soporte VBA)
poetry install -E vba
```

### Ejecutar Tests

```bash
# Ejecutar todos los tests
poetry run pytest

# Ejecutar tests con cobertura
poetry run pytest --cov=mcp_excel

# Ejecutar un test específico
poetry run pytest tests/unit/test_cache.py
```

### Verificación de Código

```bash
# Linting
poetry run ruff check src/

# Type checking
poetry run mypy src/

# Auto-fix linting issues
poetry run ruff check --fix src/
```

## Estructura del Proyecto

```
mcp-excel-server/
├── src/mcp_excel/
│   ├── server.py              # Servidor MCP principal
│   ├── config.py              # Configuración
│   ├── backends/              # Backends de Excel
│   │   ├── base.py            # Backend abstracto
│   │   ├── openpyxl_backend.py
│   │   ├── xlwings_backend.py
│   │   └── factory.py
│   ├── tools/                 # Tools MCP organizados por dominio
│   │   ├── read.py
│   │   ├── write.py
│   │   ├── formulas.py
│   │   ├── inspect.py
│   │   ├── charts.py
│   │   ├── tables.py
│   │   └── vba.py
│   └── utils/                 # Utilidades compartidas
│       ├── backend.py         # Helper get_backend()
│       ├── cache.py           # Cache singleton
│       ├── common.py          # Funciones comunes
│       ├── validation.py      # Validación de paths
│       ├── audit.py           # Audit logging
│       ├── rate_limiter.py    # Rate limiting
│       ├── vba_templates.py   # Templates VBA
│       └── vba_validator.py   # Validación VBA
├── tests/
│   ├── unit/                  # Tests unitarios
│   ├── integration/           # Tests de integración
│   └── contract/              # Tests de contrato
└── docs/
    └── TOOLS.md               # Documentación de tools
```

## Convenciones de Código

### Estilo

- Seguir PEP 8 para Python
- Usar Ruff para linting
- Type hints en todas las funciones públicas
- Docstrings en formato Google para funciones públicas

### Commits

Usar [Conventional Commits](https://www.conventionalcommits.org/es/v1.0.0/) en español:

```
tipo(alcance): descripción

[opcional corpo]

[opcional pie de notas]
```

Tipos:
- `feat`: Nueva funcionalidad
- `fix`: Corrección de bug
- `refactor`: Refactorización sin cambio de funcionalidad
- `test`: Agregar o modificar tests
- `docs`: Documentación
- `chore`: Tareas de mantenimiento

### Ejemplo

```bash
git commit -m "feat(tools): agregar soporte para gráficos 3D"
git commit -m "fix(cache): corregir memory leak en cache de workbooks"
git commit -m "refactor(vba): extraer templates a módulo separado"
```

## Agregar Nuevo Tool

1. Crear función en el módulo correspondiente (`tools/`)
2. Decorar con `@tools.tool()`
3. Agregar tipo de retorno como `dict[str, Any]`
4. Manejar errores con try/except
5. Usar `get_backend()` para obtener el backend
6. Agregar tests en `tests/integration/`
7. Documentar en `docs/TOOLS.md`

### Ejemplo de Tool

```python
@tools.tool(
    name="mi_nuevo_tool",
    description="Descripción del tool",
    tags={"excel", "read"},
)
async def mi_nuevo_tool(
    file_path: Annotated[str, Field(description="Ruta al archivo")],
    param: Annotated[str, Field(description="Parámetro")],
) -> dict[str, Any]:
    """Docstring del tool."""
    try:
        backend = get_backend(file_path)
        # Lógica del tool
        return {"success": True, "result": result}
    except Exception as e:
        logger.error("Error: %s", e)
        return {"success": False, "error": str(e)}
```

## Agregar Nuevo Backend

1. Crear clase que herede de `ExcelBackend` en `backends/`
2. Implementar todos los métodos abstractos
3. Agregar conditionales en `factory.py`
4. Agregar tests unitarios
5. Actualizar documentación

## Reportar Bugs

Usar el issue tracker de GitHub. Incluir:
- Versión de Python
- Versión del paquete
- Pasos para reproducir
- Comportamiento esperado vs actual
- Logs relevantes

## Licencia

Al contribuir, aceptas que tus contribuciones se licencien bajo la MIT License.
