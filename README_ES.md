# BODs XLS Export

![Interfaz del Contador de BODs](image.png)

Sistema automatizado para contar y exportar BODs (Bulk Order Deeds) de Ultima Online a hojas de cálculo Excel.

## 📋 Características

- **Interfaz gráfica moderna** con tema oscuro
- **Conteo automático** de BODs en bolsas del juego
- **Exportación a Excel** con formato colorido
- **Múltiples colecciones** (Verite, Agapite, Gold)
- **Selección de directorio** de exportación personalizada
- **Estado en tiempo real** del proceso
- **Botón para abrir** carpeta de exportación
- **Validación de datos** y manejo de errores

## 🚀 Cómo Usar

### Prerrequisitos

- **UO Stealth** instalado y configurado
- **Python 3.8+** (ya incluido en UO Stealth)
- **Dependencias Python** instaladas

### Instalación

1. **Clona o descarga** este repositorio
2. **Instala las dependencias**:
   ```bash
   pip install -r requirements.txt
   ```
3. **Configura los IDs de las bolsas** en el archivo `bs_config.py` (si es necesario)

### Ejecución

1. **Abre UO Stealth**
2. **Conéctate al servidor** Ultima Online
3. **Carga el script** `countBodsGenXLS.py` en UO Stealth
4. **Haz clic en "Play"** para ejecutar
5. **Configura el directorio** de exportación (si es necesario)
6. **Selecciona la colección** de BODs deseada
7. **Haz clic en "Start"** para iniciar el conteo

### Configuración

- **Directorio de exportación**: Por defecto, los archivos se guardan en `~/Documents/BOD_Reports/`
- **IDs de las bolsas**: Configura en el archivo `bs_config.py` según tus bolsas en el juego
- **Colecciones disponibles**: Verite, Agapite, Gold

## 📁 Estructura del Proyecto

```
bodsxlsexport/
├── countBodsGenXLS.py    # Script principal con interfaz gráfica
├── xlsGenerator.py       # Generador de hojas de cálculo Excel
├── bs_config.py          # Configuraciones y listas de BODs
├── modules/              # Módulos utilitarios
│   ├── common_utils.py   # Funciones utilitarias generales
│   └── connection.py     # Utilidades de conexión
├── requirements.txt      # Dependencias Python
├── README.md            # Documentación en inglés
├── README_PT.md         # Documentación en portugués
├── README_ES.md         # Documentación en español
├── .gitignore           # Archivos ignorados por Git
└── image.png            # Captura de pantalla de la interfaz
```

## 🔧 Solución de Problemas

### El script no inicia
- Verifica si UO Stealth está funcionando
- Confirma si Python está configurado correctamente
- Verifica si todas las dependencias están instaladas

### No se encuentran BODs
- Asegúrate de que las bolsas estén en el suelo cerca del personaje
- Verifica si los IDs de las bolsas en `bs_config.py` son correctos
- Confirma si el personaje está lo suficientemente cerca de las bolsas

### Error al guardar archivo
- Verifica si el directorio de exportación existe
- Confirma si hay permisos de escritura en el directorio
- Intenta seleccionar un directorio diferente

## 📊 Formato de la Hoja de Cálculo

La hoja de cálculo Excel generada contiene:
- **Columnas por material y cantidad** (ej: "VERITE 20e", "AGAPITE 15e")
- **Filas por tipo de item** (LBOD, COIF, LEGS, TUNIC, etc.)
- **Colores diferenciados** por material:
  - 🔵 Azul: Valorite
  - 🟢 Verde: Verite
  - 🟣 Morado: Agapite
  - 🟡 Amarillo: Gold
  - 🔴 Rojo: Valores en cero

## 🎮 Compatibilidad

- **Servidor**: Astraroth (Ultima Online)
- **Cliente**: UO Stealth
- **Sistemas**: Windows, macOS, Linux

## 📝 Registro de Cambios

### v1.1.0
- ✨ Interfaz gráfica moderna con tema oscuro
- ✨ Selección de directorio de exportación
- ✨ Estado en tiempo real
- ✨ Botón para abrir carpeta de exportación
- ✨ Validación de datos mejorada
- ✨ Manejo de errores mejorado
- ✨ Interfaz más compacta y responsiva

### v1.0.0
- 🎉 Versión inicial
- ✨ Conteo automático de BODs
- ✨ Exportación a Excel
- ✨ Múltiples colecciones de BODs

## 👨‍💻 Autor

**Hecho para Astraroth** - Servidor Ultima Online

## 📄 Licencia

Este proyecto es de uso libre para la comunidad de Ultima Online.

---

*Para soporte o dudas, contacta a través de los canales oficiales del servidor Astraroth.*
