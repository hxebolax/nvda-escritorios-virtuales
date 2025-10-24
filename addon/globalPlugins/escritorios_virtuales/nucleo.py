"""
Capa de lógica de negocio para gestión de escritorios virtuales.
Proporciona clases de alto nivel para trabajar con escritorios y ventanas.
"""

import ctypes
from ctypes import pointer, c_void_p, POINTER
from ctypes.wintypes import HWND, BOOL, UINT
from typing import List, Optional
import logging

from .com import (
    GestorCOM, VersionWindows, GUID, S_OK,
    IVirtualDesktop, IApplicationView, IObjectArray,
    obtener_elementos_array_objetos, user32
)

logger = logging.getLogger(__name__)


class ExcepcionEVD(Exception):
    """Excepción base para errores de Escritorios Virtuales"""
    pass


class ErrorInicializacionCOM(ExcepcionEVD):
    """Error al inicializar COM"""
    pass


class VersionWindowsNoSoportada(ExcepcionEVD):
    """Versión de Windows no soportada"""
    pass


class OperacionNoSoportada(ExcepcionEVD):
    """Operación no soportada en esta versión"""
    pass


class EscritorioVirtual:
    """Representa un escritorio virtual de Windows"""
    
    def __init__(self, puntero_escritorio: c_void_p, gestor: 'GestorEscritorios'):
        self._escritorio = ctypes.cast(puntero_escritorio, POINTER(IVirtualDesktop))
        self._gestor = gestor
        self._id = None
        self._numero = None
    
    @property
    def id(self) -> str:
        """Obtener GUID del escritorio"""
        if self._id is None:
            guid = GUID()
            vtbl = self._escritorio.contents.lpVtbl.contents
            resultado = vtbl.GetID(self._escritorio, pointer(guid))
            if resultado == S_OK:
                self._id = str(guid)
            else:
                raise ExcepcionEVD(f"Error al obtener ID del escritorio: {resultado:#x}")
        return self._id
    
    @property
    def numero(self) -> int:
        """Obtener número del escritorio (1-based)"""
        if self._numero is None:
            escritorios = self._gestor.obtener_escritorios()
            mi_id = self.id
            for i, escritorio in enumerate(escritorios, 1):
                if escritorio.id == mi_id:
                    self._numero = i
                    break
            if self._numero is None:
                raise ExcepcionEVD("Escritorio no encontrado en la lista")
        return self._numero
    
    @property
    def nombre(self) -> str:
        """Obtener nombre del escritorio (si está soportado)"""
        if not self._gestor.version.soporta_renombrar():
            return f"Escritorio {self.numero}"
        
        # En versiones que soportan nombres, el nombre está en la interfaz
        # Por simplicidad, retornamos el número por ahora
        return f"Escritorio {self.numero}"
    
    def renombrar(self, nuevo_nombre: str):
        """Renombrar escritorio"""
        if not self._gestor.version.soporta_renombrar():
            raise OperacionNoSoportada("Renombrar no soportado en esta versión de Windows")
        
        # La implementación completa requeriría IVirtualDesktopManagerInternal2
        # Por ahora lanzamos excepción
        raise OperacionNoSoportada("Funcionalidad de renombrar no implementada aún")
    
    def ir(self):
        """Cambiar a este escritorio"""
        # Llamar AllowSetForegroundWindow para mejor comportamiento de foco
        user32.AllowSetForegroundWindow(-1)  # ASFW_ANY
        
        vtbl = self._gestor.gestor_interno.contents.lpVtbl.contents
        resultado = vtbl.SwitchDesktop(self._gestor.gestor_interno, self._escritorio)
        
        if resultado != S_OK:
            raise ExcepcionEVD(f"Error al cambiar de escritorio: {resultado:#x}")
    
    def eliminar(self, respaldo: Optional['EscritorioVirtual'] = None):
        """Eliminar este escritorio"""
        if respaldo is None:
            # Usar el primer escritorio como respaldo
            escritorios = self._gestor.obtener_escritorios()
            if len(escritorios) <= 1:
                raise ExcepcionEVD("No se puede eliminar el último escritorio")
            respaldo = escritorios[0] if escritorios[0].id != self.id else escritorios[1]
        
        vtbl = self._gestor.gestor_interno.contents.lpVtbl.contents
        resultado = vtbl.RemoveDesktop(
            self._gestor.gestor_interno,
            self._escritorio,
            respaldo._escritorio
        )
        
        if resultado != S_OK:
            raise ExcepcionEVD(f"Error al eliminar escritorio: {resultado:#x}")
    
    def obtener_ventanas(self) -> List['VistaAplicacion']:
        """Obtener ventanas en este escritorio"""
        todas_ventanas = self._gestor.obtener_ventanas()
        ventanas_escritorio = []
        
        for ventana in todas_ventanas:
            try:
                if ventana.id_escritorio == self.id or ventana.esta_anclada():
                    ventanas_escritorio.append(ventana)
            except:
                pass
        
        return ventanas_escritorio
    
    @classmethod
    def actual(cls) -> 'EscritorioVirtual':
        """Obtener el escritorio actual"""
        gestor = GestorEscritorios()
        return gestor.obtener_escritorio_actual()
    
    @classmethod
    def crear(cls) -> 'EscritorioVirtual':
        """Crear un nuevo escritorio virtual"""
        gestor = GestorEscritorios()
        return gestor.crear_escritorio()
    
    @classmethod
    def obtener_todos(cls) -> List['EscritorioVirtual']:
        """Obtener lista de todos los escritorios virtuales"""
        gestor = GestorEscritorios()
        return gestor.obtener_escritorios()
    
    def __str__(self):
        return f"EscritorioVirtual(numero={self.numero}, id={self.id})"
    
    def __repr__(self):
        return self.__str__()



class VistaAplicacion:
    """Representa una ventana (ApplicationView)"""
    
    def __init__(self, puntero_vista: c_void_p, gestor: 'GestorEscritorios'):
        self._vista = ctypes.cast(puntero_vista, POINTER(IApplicationView))
        self._gestor = gestor
        self._hwnd = None
        self._titulo = None
        self._id_aplicacion = None
        self._id_escritorio = None
    
    @property
    def hwnd(self) -> int:
        """Handle de ventana"""
        if self._hwnd is None:
            hwnd = HWND()
            vtbl = self._vista.contents.lpVtbl.contents
            resultado = vtbl.GetThumbnailWindow(self._vista, pointer(hwnd))
            if resultado == S_OK:
                self._hwnd = hwnd.value
            else:
                self._hwnd = 0
        return self._hwnd
    
    @property
    def titulo(self) -> str:
        """Título de ventana"""
        if self._titulo is None:
            hwnd = self.hwnd
            if hwnd:
                longitud = user32.GetWindowTextLengthW(hwnd)
                if longitud > 0:
                    buffer = ctypes.create_unicode_buffer(longitud + 1)
                    user32.GetWindowTextW(hwnd, buffer, longitud + 1)
                    self._titulo = buffer.value
                else:
                    self._titulo = ""
            else:
                self._titulo = ""
        return self._titulo
    
    @property
    def id_aplicacion(self) -> str:
        """ID de aplicación"""
        if self._id_aplicacion is None:
            try:
                puntero_id_app = ctypes.c_wchar_p()
                vtbl = self._vista.contents.lpVtbl.contents
                resultado = vtbl.GetAppUserModelId(self._vista, pointer(puntero_id_app))
                if resultado == S_OK and puntero_id_app:
                    self._id_aplicacion = puntero_id_app.value
                else:
                    self._id_aplicacion = ""
            except:
                self._id_aplicacion = ""
        return self._id_aplicacion
    
    @property
    def id_escritorio(self) -> str:
        """ID del escritorio donde está la ventana"""
        if self._id_escritorio is None:
            guid = GUID()
            vtbl = self._vista.contents.lpVtbl.contents
            resultado = vtbl.GetVirtualDesktopId(self._vista, pointer(guid))
            if resultado == S_OK:
                self._id_escritorio = str(guid)
            else:
                self._id_escritorio = ""
        return self._id_escritorio
    
    @property
    def escritorio(self) -> EscritorioVirtual:
        """Obtener el escritorio virtual donde está la ventana"""
        id_escritorio = self.id_escritorio
        if not id_escritorio:
            raise ExcepcionEVD("La ventana no tiene ID de escritorio")
        
        # Buscar el escritorio con este ID
        escritorios = self._gestor.obtener_escritorios()
        for escritorio in escritorios:
            if escritorio.id == id_escritorio:
                return escritorio
        
        raise ExcepcionEVD(f"Escritorio con ID {id_escritorio} no encontrado")
    
    def mover_a_escritorio(self, escritorio: EscritorioVirtual):
        """Mover ventana a otro escritorio"""
        vtbl = self._gestor.gestor_interno.contents.lpVtbl.contents
        resultado = vtbl.MoveViewToDesktop(
            self._gestor.gestor_interno,
            self._vista,
            escritorio._escritorio
        )
        
        if resultado != S_OK:
            raise ExcepcionEVD(f"Error al mover ventana: {resultado:#x}")
        
        # Invalidar cache
        self._id_escritorio = None
    
    def anclar(self):
        """Anclar ventana (mostrar en todos los escritorios)"""
        vtbl = self._gestor.aplicaciones_ancladas.contents.lpVtbl.contents
        resultado = vtbl.PinView(self._gestor.aplicaciones_ancladas, self._vista)
        
        if resultado != S_OK:
            raise ExcepcionEVD(f"Error al anclar ventana: {resultado:#x}")
    
    def desanclar(self):
        """Desanclar ventana"""
        vtbl = self._gestor.aplicaciones_ancladas.contents.lpVtbl.contents
        resultado = vtbl.UnpinView(self._gestor.aplicaciones_ancladas, self._vista)
        
        if resultado != S_OK:
            raise ExcepcionEVD(f"Error al desanclar ventana: {resultado:#x}")
    
    def esta_anclada(self) -> bool:
        """Verificar si la ventana está anclada"""
        esta_anclada = BOOL()
        vtbl = self._gestor.aplicaciones_ancladas.contents.lpVtbl.contents
        resultado = vtbl.IsViewPinned(
            self._gestor.aplicaciones_ancladas,
            self._vista,
            pointer(esta_anclada)
        )
        
        if resultado == S_OK:
            return bool(esta_anclada.value)
        return False
    
    def enfocar(self):
        """Enfocar ventana"""
        # Obtener el escritorio de la ventana
        escritorio_ventana = self.escritorio
        escritorio_actual = self._gestor.obtener_escritorio_actual()
        
        # Si la ventana está en otro escritorio, cambiar primero
        if escritorio_ventana.id != escritorio_actual.id:
            escritorio_ventana.ir()
            # Esperar un momento para que el cambio se complete
            import time
            time.sleep(0.3)
        
        # Usar SwitchTo en lugar de SetFocus para evitar problemas de permisos
        vtbl = self._vista.contents.lpVtbl.contents
        resultado = vtbl.SwitchTo(self._vista)
        
        if resultado != S_OK:
            raise ExcepcionEVD(f"Error al enfocar ventana: {resultado:#x}")
    
    def se_muestra_en_alternador(self) -> bool:
        """Verificar si se muestra en alt-tab"""
        mostrado = UINT()
        vtbl = self._vista.contents.lpVtbl.contents
        resultado = vtbl.GetShowInSwitchers(self._vista, pointer(mostrado))
        
        if resultado == S_OK:
            return bool(mostrado.value)
        return False
    
    @classmethod
    def actual(cls) -> Optional['VistaAplicacion']:
        """Obtener la ventana actualmente enfocada"""
        gestor = GestorEscritorios()
        return gestor.obtener_ventana_actual()
    
    def __str__(self):
        return f"VistaAplicacion(hwnd={self.hwnd}, titulo='{self.titulo}')"
    
    def __repr__(self):
        return self.__str__()


class GestorEscritorios:
    """Gestor principal para escritorios virtuales"""
    
    def __init__(self):
        try:
            self.gestor_com = GestorCOM()
            self.version = self.gestor_com.version
            
            # Inicializar interfaces COM
            self.gestor_interno = self.gestor_com.obtener_gestor_escritorios_interno()
            self.coleccion_vistas = self.gestor_com.obtener_coleccion_vistas_aplicacion()
            self.aplicaciones_ancladas = self.gestor_com.obtener_aplicaciones_ancladas()
            
            logger.info("GestorEscritorios inicializado exitosamente")
            
        except Exception as e:
            logger.error(f"Error al inicializar GestorEscritorios: {e}")
            raise ErrorInicializacionCOM(f"Error al inicializar interfaces COM: {e}")
    
    def obtener_escritorios(self) -> List[EscritorioVirtual]:
        """Obtener lista de todos los escritorios virtuales"""
        escritorios = []
        
        try:
            # Obtener array de escritorios
            puntero_array = POINTER(IObjectArray)()
            vtbl = self.gestor_interno.contents.lpVtbl.contents
            resultado = vtbl.GetDesktops(self.gestor_interno, pointer(puntero_array))
            
            if resultado != S_OK:
                raise ExcepcionEVD(f"Error al obtener escritorios: {resultado:#x}")
            
            # Obtener elementos del array
            elementos = obtener_elementos_array_objetos(puntero_array, self.version.guid_escritorio)
            
            # Crear objetos EscritorioVirtual
            for elemento in elementos:
                escritorio = EscritorioVirtual(elemento, self)
                escritorios.append(escritorio)
            
            # Liberar array
            self.gestor_com.liberar_interfaz(puntero_array)
            
        except Exception as e:
            logger.error(f"Error obteniendo escritorios: {e}")
            raise
        
        return escritorios
    
    def obtener_escritorio_actual(self) -> EscritorioVirtual:
        """Obtener escritorio actual"""
        puntero_escritorio = POINTER(IVirtualDesktop)()
        vtbl = self.gestor_interno.contents.lpVtbl.contents
        resultado = vtbl.GetCurrentDesktop(self.gestor_interno, pointer(puntero_escritorio))
        
        if resultado != S_OK:
            raise ExcepcionEVD(f"Error al obtener escritorio actual: {resultado:#x}")
        
        return EscritorioVirtual(ctypes.cast(puntero_escritorio, c_void_p), self)
    
    def crear_escritorio(self) -> EscritorioVirtual:
        """Crear nuevo escritorio virtual"""
        puntero_escritorio = POINTER(IVirtualDesktop)()
        vtbl = self.gestor_interno.contents.lpVtbl.contents
        resultado = vtbl.CreateDesktopW(self.gestor_interno, pointer(puntero_escritorio))
        
        if resultado != S_OK:
            raise ExcepcionEVD(f"Error al crear escritorio: {resultado:#x}")
        
        return EscritorioVirtual(ctypes.cast(puntero_escritorio, c_void_p), self)
    
    def obtener_ventanas(self, escritorio: Optional[EscritorioVirtual] = None) -> List[VistaAplicacion]:
        """Obtener ventanas (todas o de un escritorio específico)"""
        ventanas = []
        
        try:
            # Obtener array de vistas ordenadas por Z
            puntero_array = POINTER(IObjectArray)()
            vtbl = self.coleccion_vistas.contents.lpVtbl.contents
            resultado = vtbl.GetViewsByZOrder(self.coleccion_vistas, pointer(puntero_array))
            
            if resultado != S_OK:
                raise ExcepcionEVD(f"Error al obtener ventanas: {resultado:#x}")
            
            # Obtener elementos del array
            from .com import IID_IApplicationView
            elementos = obtener_elementos_array_objetos(puntero_array, IID_IApplicationView)
            
            # Crear objetos VistaAplicacion
            for elemento in elementos:
                ventana = VistaAplicacion(elemento, self)
                
                # Filtrar ventanas que no se muestran en alternador
                if not ventana.se_muestra_en_alternador():
                    continue
                
                # Si se especificó un escritorio, filtrar por él
                if escritorio is not None:
                    try:
                        if ventana.id_escritorio != escritorio.id and not ventana.esta_anclada():
                            continue
                    except:
                        continue
                
                ventanas.append(ventana)
            
            # Liberar array
            self.gestor_com.liberar_interfaz(puntero_array)
            
        except Exception as e:
            logger.error(f"Error obteniendo ventanas: {e}")
            raise
        
        return ventanas
    
    def obtener_ventana_actual(self) -> Optional[VistaAplicacion]:
        """Obtener ventana actualmente enfocada"""
        try:
            puntero_vista = POINTER(IApplicationView)()
            vtbl = self.coleccion_vistas.contents.lpVtbl.contents
            resultado = vtbl.GetViewInFocus(self.coleccion_vistas, pointer(puntero_vista))
            
            if resultado == S_OK and puntero_vista:
                return VistaAplicacion(ctypes.cast(puntero_vista, c_void_p), self)
        except Exception as e:
            logger.error(f"Error obteniendo ventana actual: {e}")
        
        return None
    
    def obtener_cantidad_escritorios(self) -> int:
        """Obtener número de escritorios"""
        cantidad = UINT()
        vtbl = self.gestor_interno.contents.lpVtbl.contents
        resultado = vtbl.GetCount(self.gestor_interno, pointer(cantidad))
        
        if resultado == S_OK:
            return cantidad.value
        return 0
    
    def __del__(self):
        """Cleanup al destruir"""
        try:
            if hasattr(self, 'gestor_interno'):
                self.gestor_com.liberar_interfaz(self.gestor_interno)
            if hasattr(self, 'coleccion_vistas'):
                self.gestor_com.liberar_interfaz(self.coleccion_vistas)
            if hasattr(self, 'aplicaciones_ancladas'):
                self.gestor_com.liberar_interfaz(self.aplicaciones_ancladas)
        except:
            pass
