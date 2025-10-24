"""
escritorios-virtuales: Implementación con ctypes puro para gestión de escritorios virtuales de Windows
=====================================================================================================

Una librería Python ligera para gestionar escritorios virtuales de Windows 10/11
sin dependencias externas (sin comtypes, sin pywin32).

Uso Básico:
-----------
    >>> from escritorios_virtuales import EscritorioVirtual, VistaAplicacion
    
    >>> # Listar todos los escritorios
    >>> escritorios = EscritorioVirtual.obtener_todos()
    >>> print(f"Tienes {len(escritorios)} escritorios")
    
    >>> # Obtener escritorio actual
    >>> actual = EscritorioVirtual.actual()
    >>> print(f"Escritorio actual: {actual.numero}")
    
    >>> # Crear un nuevo escritorio
    >>> nuevo_escritorio = EscritorioVirtual.crear()
    >>> nuevo_escritorio.ir()  # Cambiar a él
    
    >>> # Mover una ventana
    >>> ventana = VistaAplicacion.actual()  # Obtener ventana enfocada
    >>> ventana.mover_a_escritorio(escritorios[0])  # Mover al primer escritorio
    
    >>> # Anclar una ventana (mostrar en todos los escritorios)
    >>> ventana.anclar()

Para más ejemplos, consulta el directorio ejemplos/.
"""

__version__ = "1.0.0"
__author__ = "XeBoLaX"
__license__ = "MIT"

# Importar clases principales para acceso fácil
from .nucleo import (
    EscritorioVirtual,
    VistaAplicacion,
    GestorEscritorios,
    ExcepcionEVD,
    ErrorInicializacionCOM,
    VersionWindowsNoSoportada,
    OperacionNoSoportada,
)

# Exportar símbolos públicos
__all__ = [
    'EscritorioVirtual',
    'VistaAplicacion',
    'GestorEscritorios',
    'ExcepcionEVD',
    'ErrorInicializacionCOM',
    'VersionWindowsNoSoportada',
    'OperacionNoSoportada',
]
