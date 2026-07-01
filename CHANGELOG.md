# Changelog

Todas las alteraciones notables en MCP Excel Server se documentan en este archivo.

El formato está basado en [Keep a Changelog](https://keepachangelog.com/es/1.0.0/),
y este proyecto adhiere a [Semantic Versioning](https://semver.org/lang/es/).

## [Unreleased]

## [0.4.0] - 2026-07-01

### Added
- Validación de paths anti-traversal con `allowed_directories` config
- Timeout real para ejecución de macros VBA con `ThreadPoolExecutor`
- Audit logging para operaciones de tools y eventos de seguridad
- Rate limiting configurable por cliente para requests y macros
- Detección de código VBA peligroso (Shell, FileSystemObject, SendKeys, etc.)
- Tests de seguridad (26 tests para validation, audit, rate limiter, VBA safety)

### Changed
- Extraído `get_backend()` a `utils/backend.py` para eliminar código duplicado
- Extraído `VBA_TEMPLATES` a `utils/vba_templates.py`
- Extraído `validate_vba_code()` a `utils/vba_validator.py`
- Movido `col_letter_to_index()` a `utils/common.py`
- Backend methods `get_max_row`, `get_max_column` añadidos a `base.py`

### Fixed
- Circular import en `get_backend()` moviendo a `utils/backend.py`
- Mock paths en tests de VBA

## [0.3.0] - 2026-07-01

### Added
- Soporte para macros VBA con xlwings backend
- 12 tools de VBA: `list_vba_modules`, `get_vba_code`, `set_vba_code`, `add_vba_module`, `delete_vba_module`, `rename_vba_module`, `run_macro`, `list_macros`, `get_vba_templates`, `validate_vba_code`, `import_vba_module`, `export_vba_module`
- 12 templates de VBA para tareas comunes
- Configuración VBA: `vba_enabled`, `vba_macro_timeout`, `vba_trust_access`, `vba_show_excel`, `vba_audit_log`
- Soporte para formatos `.xlsm`, `.xlsb`, `.xlam`
- Backend xlwings con ejecución de macros y acceso a proyectos VBA

## [0.2.0] - 2026-07-01

### Added
- 4 tools de gráficos: `create_chart`, `modify_chart`, `list_charts`, `delete_chart`
- 4 tools de tablas estructuradas: `create_table`, `list_tables`, `get_table_data`, `add_table_row`
- 2 MCP Resources: `workbook://summary`, `workbook://sheets`
- 5 MCP Prompts: `analyze_data`, `create_report`, `compare_sheets`, `clean_data`, `document_workbook`

### Changed
- Refactorización completa de la arquitectura
- Eliminación de herramientas duplicadas
- Unificación de `WorkbookCache` en singleton `shared_cache`

## [0.1.0] - 2026-07-01

### Added
- 14 tools iniciales para lectura, escritura, fórmulas e inspección
- Backend openpyxl con cache por hashes de archivos
- Paginación de celdas con límite configurable
- Respuestas con headers consistentes
