# RELAC TX - Energy System Optimization Model

Sistema de modelado de optimización energética basado en OSeMOSYS para América Latina y el Caribe.

## Descripción

Este proyecto implementa un pipeline automatizado para la ejecución de modelos de optimización energética utilizando OSeMOSYS. El sistema soporta múltiples solvers (GLPK, CBC, CPLEX, Gurobi) y está diseñado para garantizar reproducibilidad completa de los resultados.

## Características Principales

- **Pipeline Automatizado**: Gestión completa del flujo de trabajo con DVC
- **Múltiples Solvers**: Soporte para GLPK, CBC, CPLEX y Gurobi
- **Reproducibilidad Garantizada**: Seeds configurables para resultados determinísticos
- **Medición de Rendimiento**: Timer integrado para monitorear tiempos de ejecución
- **Gestión Automática de Entorno**: Creación y actualización automática del entorno Conda

## Requisitos del Sistema

- Windows 10 o superior
- Git para Windows
- Miniconda o Anaconda
- Al menos un solver: GLPK, CBC, CPLEX o Gurobi

## Inicio Rápido

```bash
# Clonar el repositorio
git clone https://github.com/clg-admin/relac_tx.git
cd relac_tx

# Ejecutar el modelo (desde Anaconda Prompt)
python run.py
```

El script `run.py` gestiona automáticamente:
- Creación del entorno Conda
- Instalación de dependencias
- Ejecución del pipeline completo
- Generación de archivos de salida

## Documentación

Para instrucciones detalladas de instalación y configuración, consulta la guía completa:
- **Guía de Instalación y Ejecución**: `RELAC_TX_Guia_instalacion_ejecucion.md`

## Estructura de Archivos de Salida

Los resultados se generan en `t1_confection/` con los siguientes archivos:
- `RELAC_TX_Inputs.csv` / `RELAC_TX_Inputs_YYYY-MM-DD.csv`
- `RELAC_TX_Outputs.csv` / `RELAC_TX_Outputs_YYYY-MM-DD.csv`
- `RELAC_TX_Combined_Inputs_Outputs.csv` / `RELAC_TX_Combined_Inputs_Outputs_YYYY-MM-DD.csv`

Los archivos con fecha mantienen un histórico completo de ejecuciones.

## Configuración

El archivo principal de configuración es `t1_confection/MOMF_T1_AB.yaml`, donde puedes ajustar:
- Solver a utilizar (`solver: 'cplex'`)
- Número de threads para solvers comerciales
- Seeds para reproducibilidad
- Anualización de capital (`annualize_capital`)

## Editor de Tecnologías Secundarias

El proyecto incluye un sistema para facilitar la edición de tecnologías secundarias (Secondary Techs) en los archivos de parametrización.

### Uso del Editor

1. **Generar plantilla de edición**:
   ```bash
   python t1_confection/D1_generate_editor_template.py
   ```
   Esto crea el archivo `Secondary_Techs_Editor.xlsx` con listas desplegables para facilitar la edición.

2. **Editar valores**:
   - Abrir `Secondary_Techs_Editor.xlsx`
   - Seleccionar: Escenario (BAU, NDC, NDC+ELC, NDC_NoRPO, o ALL)
   - Seleccionar: País, Tecnología y Parámetro
   - Ingresar los valores para los años deseados

3. **Aplicar cambios**:
   ```bash
   python t1_confection/D2_update_secondary_techs.py
   ```
   Este script:
   - Crea respaldos automáticos de los archivos originales
   - Aplica los cambios a los escenarios correspondientes
   - Actualiza automáticamente el campo `Projection.Mode` a "User defined"
   - Genera un log detallado de todas las operaciones

### Características del Editor

- **Listas desplegables**: Facilitan la selección de escenarios, países, tecnologías y parámetros
- **Validación automática**: Verifica que los datos sean consistentes antes de aplicar cambios
- **Respaldos automáticos**: Crea copias de seguridad con timestamp antes de modificar archivos
- **Aplicación a múltiples escenarios**: Usa "ALL" para aplicar cambios a todos los escenarios a la vez
- **Logs detallados**: Registro completo de cambios aplicados y errores encontrados

## Licencia

Este proyecto está licenciado bajo la Licencia MIT - consulta el archivo [LICENSE](LICENSE) para más detalles.

Copyright (c) 2025 Climate Lead Group

Este proyecto está desarrollado por Climate Lead Group para análisis de sistemas energéticos en América Latina y el Caribe.
