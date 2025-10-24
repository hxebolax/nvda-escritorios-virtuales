"""
Capa COM con ctypes puro para acceso a escritorios virtuales de Windows.
Implementación sin dependencias externas (sin comtypes ni pywin32).
"""

import sys
import ctypes
from ctypes import (
    Structure, POINTER, pointer, c_void_p, c_ulong, c_ushort, c_ubyte,
    c_uint, c_int, c_ulonglong, c_wchar_p, c_bool, HRESULT, windll
)
from ctypes.wintypes import HWND, BOOL, UINT, DWORD, LPVOID, LPCWSTR, INT, RECT, SIZE, ULONG
from typing import Optional, List, Any
import logging

logger = logging.getLogger(__name__)

# Constantes COM
COINIT_APARTMENTTHREADED = 0x2
COINIT_MULTITHREADED = 0x0
CLSCTX_LOCAL_SERVER = 0x4
S_OK = 0
E_NOINTERFACE = 0x80004002

# Cargar librerías
ole32 = windll.ole32
user32 = windll.user32
combase = windll.combase


class GUID(Structure):
    """Estructura GUID de Windows"""
    _fields_ = [
        ("Data1", c_ulong),
        ("Data2", c_ushort),
        ("Data3", c_ushort),
        ("Data4", c_ubyte * 8)
    ]
    
    def __init__(self, cadena_guid: Optional[str] = None):
        super().__init__()
        if cadena_guid:
            self._parsear_guid(cadena_guid)
    
    def _parsear_guid(self, cadena_guid: str):
        """Parsear cadena GUID formato {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}"""
        cadena_guid = cadena_guid.strip('{}')
        partes = cadena_guid.split('-')
        
        self.Data1 = int(partes[0], 16)
        self.Data2 = int(partes[1], 16)
        self.Data3 = int(partes[2], 16)
        
        data4_str = partes[3] + partes[4]
        for i in range(8):
            self.Data4[i] = int(data4_str[i*2:i*2+2], 16)
    
    def __str__(self):
        data4_str = ''.join(f'{b:02X}' for b in self.Data4)
        return f"{{{self.Data1:08X}-{self.Data2:04X}-{self.Data3:04X}-{data4_str[:4]}-{data4_str[4:]}}}"


# GUIDs de interfaces y clases COM
CLSID_ImmersiveShell = GUID("{C2F03A33-21F5-47FA-B4BB-156362A2F239}")
CLSID_VirtualDesktopManagerInternal = GUID("{C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B}")
CLSID_VirtualDesktopPinnedApps = GUID("{B5A399E7-1C87-46B8-88E9-FC5747B171BD}")

# IServiceProvider
IID_IServiceProvider = GUID("{6D5140C1-7436-11CE-8034-00AA006009FA}")

# IObjectArray
IID_IObjectArray = GUID("{92CA9DCD-5622-4BBA-A805-5E9F541BD8C9}")

# IApplicationView
IID_IApplicationView = GUID("{372E1D3B-38D3-42E4-A15B-8AB2B178F513}")

# IApplicationViewCollection
IID_IApplicationViewCollection = GUID("{1841C6D7-4F9D-42C0-AF41-8747538F10E5}")

# IVirtualDesktopPinnedApps
IID_IVirtualDesktopPinnedApps = GUID("{4CE81583-1E4C-4632-A621-07A53543148F}")

# GUIDs de IVirtualDesktop por versión
GUID_IVirtualDesktop_9000 = GUID("{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}")
GUID_IVirtualDesktop_20231 = GUID("{62FDF88B-11CA-4AFB-8BD8-2296DFAE49E2}")
GUID_IVirtualDesktop_21313 = GUID("{536D3495-B208-4CC9-AE26-DE8111275BF8}")
GUID_IVirtualDesktop_22621 = GUID("{3F07F4BE-B107-441A-AF0F-39D82529072C}")
GUID_IVirtualDesktop_22631 = GUID("{3F07F4BE-B107-441A-AF0F-39D82529072C}")
GUID_IVirtualDesktop_26100 = GUID("{3F07F4BE-B107-441A-AF0F-39D82529072C}")

# GUIDs de IVirtualDesktopManagerInternal por versión
GUID_IVirtualDesktopManagerInternal_9000 = GUID("{F31574D6-B682-4CDC-BD56-1827860ABEC6}")
GUID_IVirtualDesktopManagerInternal_20231 = GUID("{094AFE11-44F2-4BA0-976F-29A97E263EE0}")
GUID_IVirtualDesktopManagerInternal_21313 = GUID("{B2F925B9-5A0F-4D2E-9F4D-2B1507593C10}")
GUID_IVirtualDesktopManagerInternal_22621 = GUID("{A3175F2D-239C-4BD2-8AA0-EEBA8B0B138E}")
GUID_IVirtualDesktopManagerInternal_22631 = GUID("{4970BA3D-FD4E-4647-BEA3-D89076EF4B9C}")
GUID_IVirtualDesktopManagerInternal_26100 = GUID("{53F5CA0B-158F-4124-900C-057158060B27}")


class VersionWindows:
    """Detecta la versión de Windows y selecciona GUIDs apropiados"""
    
    def __init__(self):
        info_version = sys.getwindowsversion()
        self.mayor = info_version.major
        self.menor = info_version.minor
        self.compilacion = info_version.build
        
        logger.info(f"Versión de Windows detectada: {self.mayor}.{self.menor}.{self.compilacion}")
        
        # Determinar versión y GUIDs
        if self.compilacion >= 26100:
            self.nombre_version = "26100+"
            self.guid_escritorio = GUID_IVirtualDesktop_26100
            self.guid_gestor = GUID_IVirtualDesktopManagerInternal_26100
        elif self.compilacion >= 22631:
            self.nombre_version = "22631+"
            self.guid_escritorio = GUID_IVirtualDesktop_22631
            self.guid_gestor = GUID_IVirtualDesktopManagerInternal_22631
        elif self.compilacion >= 22621:
            self.nombre_version = "22621+"
            self.guid_escritorio = GUID_IVirtualDesktop_22621
            self.guid_gestor = GUID_IVirtualDesktopManagerInternal_22621
        elif self.compilacion >= 21313:
            self.nombre_version = "21313+"
            self.guid_escritorio = GUID_IVirtualDesktop_21313
            self.guid_gestor = GUID_IVirtualDesktopManagerInternal_21313
        elif self.compilacion >= 20231:
            self.nombre_version = "20231+"
            self.guid_escritorio = GUID_IVirtualDesktop_20231
            self.guid_gestor = GUID_IVirtualDesktopManagerInternal_20231
        else:
            self.nombre_version = "9000+"
            self.guid_escritorio = GUID_IVirtualDesktop_9000
            self.guid_gestor = GUID_IVirtualDesktopManagerInternal_9000
        
        logger.info(f"Usando perfil de versión: {self.nombre_version}")
    
    def soporta_renombrar(self) -> bool:
        """Renombrar escritorios soportado desde build 19041"""
        return self.compilacion >= 19041
    
    def soporta_fondo_pantalla(self) -> bool:
        """Fondo de pantalla por escritorio soportado desde build 21313"""
        return self.compilacion >= 21313


class HSTRING_HEADER(Structure):
    """Header para HSTRING"""
    _fields_ = [
        ("flags", c_uint),
        ("length", c_uint),
        ("padding1", c_uint),
        ("padding2", c_uint),
        ("data", c_void_p)
    ]


class HSTRING:
    """Wrapper para HSTRING de Windows Runtime"""
    
    def __init__(self, valor: str = ""):
        self.valor = valor
        self.handle = c_void_p()
        if valor:
            self.crear()
    
    def crear(self):
        """Crear HSTRING desde string"""
        resultado = combase.WindowsCreateString(
            c_wchar_p(self.valor),
            len(self.valor),
            pointer(self.handle)
        )
        if resultado != S_OK:
            raise Exception(f"Error al crear HSTRING: {resultado}")
        return self.handle
    
    def obtener_valor(self) -> str:
        """Obtener string desde HSTRING"""
        if not self.handle:
            return ""
        
        longitud = c_uint()
        buffer = combase.WindowsGetStringRawBuffer(self.handle, pointer(longitud))
        if buffer:
            return ctypes.wstring_at(buffer, longitud.value)
        return ""
    
    def eliminar(self):
        """Liberar HSTRING"""
        if self.handle:
            combase.WindowsDeleteString(self.handle)
            self.handle = c_void_p()
    
    def __del__(self):
        self.eliminar()
    
    @staticmethod
    def desde_handle(handle: c_void_p) -> str:
        """Convertir handle HSTRING a string"""
        if not handle:
            return ""
        longitud = c_uint()
        buffer = combase.WindowsGetStringRawBuffer(handle, pointer(longitud))
        if buffer:
            return ctypes.wstring_at(buffer, longitud.value)
        return ""


# Definición de vtable para IUnknown
class IUnknownVtbl(Structure):
    pass

class IUnknown(Structure):
    """Interfaz base COM IUnknown"""
    pass

IUnknown._fields_ = [("lpVtbl", POINTER(IUnknownVtbl))]

IUnknownVtbl._fields_ = [
    ("QueryInterface", ctypes.WINFUNCTYPE(HRESULT, POINTER(IUnknown), POINTER(GUID), POINTER(c_void_p))),
    ("AddRef", ctypes.WINFUNCTYPE(c_ulong, POINTER(IUnknown))),
    ("Release", ctypes.WINFUNCTYPE(c_ulong, POINTER(IUnknown)))
]


# IServiceProvider
class IServiceProviderVtbl(Structure):
    pass

class IServiceProvider(Structure):
    pass

IServiceProvider._fields_ = [("lpVtbl", POINTER(IServiceProviderVtbl))]

IServiceProviderVtbl._fields_ = [
    ("QueryInterface", ctypes.WINFUNCTYPE(HRESULT, POINTER(IServiceProvider), POINTER(GUID), POINTER(c_void_p))),
    ("AddRef", ctypes.WINFUNCTYPE(c_ulong, POINTER(IServiceProvider))),
    ("Release", ctypes.WINFUNCTYPE(c_ulong, POINTER(IServiceProvider))),
    ("QueryService", ctypes.WINFUNCTYPE(HRESULT, POINTER(IServiceProvider), POINTER(GUID), POINTER(GUID), POINTER(c_void_p)))
]


# IObjectArray
class IObjectArrayVtbl(Structure):
    pass

class IObjectArray(Structure):
    pass

IObjectArray._fields_ = [("lpVtbl", POINTER(IObjectArrayVtbl))]

IObjectArrayVtbl._fields_ = [
    ("QueryInterface", ctypes.WINFUNCTYPE(HRESULT, POINTER(IObjectArray), POINTER(GUID), POINTER(c_void_p))),
    ("AddRef", ctypes.WINFUNCTYPE(c_ulong, POINTER(IObjectArray))),
    ("Release", ctypes.WINFUNCTYPE(c_ulong, POINTER(IObjectArray))),
    ("GetCount", ctypes.WINFUNCTYPE(HRESULT, POINTER(IObjectArray), POINTER(UINT))),
    ("GetAt", ctypes.WINFUNCTYPE(HRESULT, POINTER(IObjectArray), UINT, POINTER(GUID), POINTER(c_void_p)))
]


# IApplicationView
class IApplicationViewVtbl(Structure):
    pass

class IApplicationView(Structure):
    pass

IApplicationView._fields_ = [("lpVtbl", POINTER(IApplicationViewVtbl))]

# IApplicationView tiene muchos métodos, definimos todos hasta GetShowInSwitchers
IApplicationViewVtbl._fields_ = [
    ("QueryInterface", ctypes.WINFUNCTYPE(HRESULT, POINTER(IApplicationView), POINTER(GUID), POINTER(c_void_p))),
    ("AddRef", ctypes.WINFUNCTYPE(c_ulong, POINTER(IApplicationView))),
    ("Release", ctypes.WINFUNCTYPE(c_ulong, POINTER(IApplicationView))),
    # IInspectable methods
    ("GetIids", c_void_p),
    ("GetRuntimeClassName", c_void_p),
    ("GetTrustLevel", c_void_p),
    # IApplicationView methods
    ("SetFocus", ctypes.WINFUNCTYPE(HRESULT, POINTER(IApplicationView))),
    ("SwitchTo", ctypes.WINFUNCTYPE(HRESULT, POINTER(IApplicationView))),
    ("TryInvokeBack", c_void_p),
    ("GetThumbnailWindow", ctypes.WINFUNCTYPE(HRESULT, POINTER(IApplicationView), POINTER(HWND))),
    ("GetMonitor", c_void_p),
    ("GetVisibility", ctypes.WINFUNCTYPE(HRESULT, POINTER(IApplicationView), POINTER(UINT))),
    ("SetCloak", c_void_p),
    ("GetPosition", c_void_p),
    ("SetPosition", c_void_p),
    ("InsertAfterWindow", c_void_p),
    ("GetExtendedFramePosition", c_void_p),
    ("GetAppUserModelId", ctypes.WINFUNCTYPE(HRESULT, POINTER(IApplicationView), POINTER(c_wchar_p))),
    ("SetAppUserModelId", c_void_p),
    ("IsEqualByAppUserModelId", c_void_p),
    ("GetViewState", c_void_p),
    ("SetViewState", c_void_p),
    ("GetNeediness", c_void_p),
    ("GetLastActivationTimestamp", c_void_p),
    ("SetLastActivationTimestamp", c_void_p),
    ("GetVirtualDesktopId", ctypes.WINFUNCTYPE(HRESULT, POINTER(IApplicationView), POINTER(GUID))),
    ("SetVirtualDesktopId", c_void_p),
    ("GetShowInSwitchers", ctypes.WINFUNCTYPE(HRESULT, POINTER(IApplicationView), POINTER(UINT))),
    ("SetShowInSwitchers", c_void_p),
    ("GetScaleFactor", c_void_p),
    ("CanReceiveInput", c_void_p),
    ("GetCompatibilityPolicyType", c_void_p),
    ("SetCompatibilityPolicyType", c_void_p),
    ("GetSizeConstraints", c_void_p),
    ("GetSizeConstraintsForDpi", c_void_p),
    ("SetSizeConstraintsForDpi", c_void_p),
    ("OnMinSizePreferencesUpdated", c_void_p),
    ("ApplyOperation", c_void_p),
    ("IsTray", c_void_p),
    ("IsInHighZOrderBand", c_void_p),
    ("IsSplashScreenPresented", c_void_p),
    ("Flash", c_void_p),
    ("GetRootSwitchableOwner", c_void_p),
    ("EnumerateOwnershipTree", c_void_p),
    ("GetEnterpriseId", c_void_p),
    ("IsMirrored", c_void_p),
]


# IVirtualDesktop
class IVirtualDesktopVtbl(Structure):
    pass

class IVirtualDesktop(Structure):
    pass

IVirtualDesktop._fields_ = [("lpVtbl", POINTER(IVirtualDesktopVtbl))]

IVirtualDesktopVtbl._fields_ = [
    ("QueryInterface", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktop), POINTER(GUID), POINTER(c_void_p))),
    ("AddRef", ctypes.WINFUNCTYPE(c_ulong, POINTER(IVirtualDesktop))),
    ("Release", ctypes.WINFUNCTYPE(c_ulong, POINTER(IVirtualDesktop))),
    ("IsViewVisible", c_void_p),
    ("GetID", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktop), POINTER(GUID))),
]


# IVirtualDesktopManagerInternal
class IVirtualDesktopManagerInternalVtbl(Structure):
    pass

class IVirtualDesktopManagerInternal(Structure):
    pass

IVirtualDesktopManagerInternal._fields_ = [("lpVtbl", POINTER(IVirtualDesktopManagerInternalVtbl))]

IVirtualDesktopManagerInternalVtbl._fields_ = [
    ("QueryInterface", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktopManagerInternal), POINTER(GUID), POINTER(c_void_p))),
    ("AddRef", ctypes.WINFUNCTYPE(c_ulong, POINTER(IVirtualDesktopManagerInternal))),
    ("Release", ctypes.WINFUNCTYPE(c_ulong, POINTER(IVirtualDesktopManagerInternal))),
    ("GetCount", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktopManagerInternal), POINTER(UINT))),
    ("MoveViewToDesktop", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktopManagerInternal), POINTER(IApplicationView), POINTER(IVirtualDesktop))),
    ("CanViewMoveDesktops", c_void_p),
    ("GetCurrentDesktop", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktopManagerInternal), POINTER(POINTER(IVirtualDesktop)))),
    ("GetDesktops", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktopManagerInternal), POINTER(POINTER(IObjectArray)))),
    ("GetAdjacentDesktop", c_void_p),
    ("SwitchDesktop", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktopManagerInternal), POINTER(IVirtualDesktop))),
    ("SwitchDesktopAndMoveForegroundView", c_void_p),
    ("CreateDesktopW", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktopManagerInternal), POINTER(POINTER(IVirtualDesktop)))),
    ("MoveDesktop", c_void_p),
    ("RemoveDesktop", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktopManagerInternal), POINTER(IVirtualDesktop), POINTER(IVirtualDesktop))),
    ("FindDesktop", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktopManagerInternal), POINTER(GUID), POINTER(POINTER(IVirtualDesktop)))),
    ("GetDesktopSwitchIncludeExcludeViews", c_void_p),
    ("SetName", c_void_p),
    ("SetWallpaper", c_void_p),
    ("SetWallpaperForAllDesktops", c_void_p),
    ("CopyDesktopState", c_void_p),
    ("CreateRemoteDesktop", c_void_p),
    ("pDesktop", c_void_p),
    ("SwitchRemoteDesktop", c_void_p),
    ("SwitchDesktopWithAnimation", c_void_p),
    ("GetLastActiveDesktop", c_void_p),
    ("WaitForAnimationToComplete", c_void_p),
]


# IApplicationViewCollection
class IApplicationViewCollectionVtbl(Structure):
    pass

class IApplicationViewCollection(Structure):
    pass

IApplicationViewCollection._fields_ = [("lpVtbl", POINTER(IApplicationViewCollectionVtbl))]

IApplicationViewCollectionVtbl._fields_ = [
    ("QueryInterface", ctypes.WINFUNCTYPE(HRESULT, POINTER(IApplicationViewCollection), POINTER(GUID), POINTER(c_void_p))),
    ("AddRef", ctypes.WINFUNCTYPE(c_ulong, POINTER(IApplicationViewCollection))),
    ("Release", ctypes.WINFUNCTYPE(c_ulong, POINTER(IApplicationViewCollection))),
    ("GetViews", c_void_p),
    ("GetViewsByZOrder", ctypes.WINFUNCTYPE(HRESULT, POINTER(IApplicationViewCollection), POINTER(POINTER(IObjectArray)))),
    ("GetViewsByAppUserModelId", c_void_p),
    ("GetViewForHwnd", ctypes.WINFUNCTYPE(HRESULT, POINTER(IApplicationViewCollection), HWND, POINTER(POINTER(IApplicationView)))),
    ("GetViewForApplication", c_void_p),
    ("GetViewForAppUserModelId", c_void_p),
    ("GetViewInFocus", ctypes.WINFUNCTYPE(HRESULT, POINTER(IApplicationViewCollection), POINTER(POINTER(IApplicationView)))),
]


# IVirtualDesktopPinnedApps
class IVirtualDesktopPinnedAppsVtbl(Structure):
    pass

class IVirtualDesktopPinnedApps(Structure):
    pass

IVirtualDesktopPinnedApps._fields_ = [("lpVtbl", POINTER(IVirtualDesktopPinnedAppsVtbl))]

IVirtualDesktopPinnedAppsVtbl._fields_ = [
    ("QueryInterface", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktopPinnedApps), POINTER(GUID), POINTER(c_void_p))),
    ("AddRef", ctypes.WINFUNCTYPE(c_ulong, POINTER(IVirtualDesktopPinnedApps))),
    ("Release", ctypes.WINFUNCTYPE(c_ulong, POINTER(IVirtualDesktopPinnedApps))),
    ("IsAppIdPinned", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktopPinnedApps), LPCWSTR, POINTER(BOOL))),
    ("PinAppID", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktopPinnedApps), LPCWSTR)),
    ("UnpinAppID", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktopPinnedApps), LPCWSTR)),
    ("IsViewPinned", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktopPinnedApps), POINTER(IApplicationView), POINTER(BOOL))),
    ("PinView", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktopPinnedApps), POINTER(IApplicationView))),
    ("UnpinView", ctypes.WINFUNCTYPE(HRESULT, POINTER(IVirtualDesktopPinnedApps), POINTER(IApplicationView))),
]


class GestorCOM:
    """Gestor de inicialización y creación de objetos COM"""
    
    def __init__(self):
        self.inicializado = False
        self.version = VersionWindows()
        self._inicializar_com()
    
    def _inicializar_com(self):
        """Inicializar COM"""
        try:
            resultado = ole32.CoInitializeEx(None, COINIT_APARTMENTTHREADED)
            if resultado == S_OK or resultado == 1:  # S_FALSE = ya inicializado
                self.inicializado = True
                logger.info("COM inicializado exitosamente")
            else:
                logger.warning(f"Inicialización COM retornó: {resultado}")
                self.inicializado = True  # Continuar de todos modos
        except Exception as e:
            logger.error(f"Error al inicializar COM: {e}")
            raise
    
    def crear_instancia(self, clsid: GUID, iid: GUID):
        """Crear instancia COM usando CoCreateInstance"""
        ppv = c_void_p()
        resultado = ole32.CoCreateInstance(
            pointer(clsid),
            None,
            CLSCTX_LOCAL_SERVER,
            pointer(iid),
            pointer(ppv)
        )
        
        if resultado != S_OK:
            raise Exception(f"CoCreateInstance falló con código: {resultado:#x}")
        
        return ppv
    
    def consultar_servicio(self, proveedor_servicio: POINTER(IServiceProvider), 
                          guid_servicio: GUID, guid_interfaz: GUID):
        """Obtener servicio usando IServiceProvider::QueryService"""
        ppv = c_void_p()
        
        vtbl = proveedor_servicio.contents.lpVtbl.contents
        resultado = vtbl.QueryService(
            proveedor_servicio,
            pointer(guid_servicio),
            pointer(guid_interfaz),
            pointer(ppv)
        )
        
        if resultado != S_OK:
            raise Exception(f"QueryService falló con código: {resultado:#x}")
        
        return ppv
    
    def obtener_immersive_shell(self) -> POINTER(IServiceProvider):
        """Obtener IServiceProvider del ImmersiveShell"""
        ppv = self.crear_instancia(CLSID_ImmersiveShell, IID_IServiceProvider)
        return ctypes.cast(ppv, POINTER(IServiceProvider))
    
    def obtener_gestor_escritorios_interno(self) -> POINTER(IVirtualDesktopManagerInternal):
        """Obtener IVirtualDesktopManagerInternal"""
        proveedor_servicio = self.obtener_immersive_shell()
        
        ppv = self.consultar_servicio(
            proveedor_servicio,
            CLSID_VirtualDesktopManagerInternal,
            self.version.guid_gestor
        )
        
        # Liberar proveedor de servicio
        vtbl = proveedor_servicio.contents.lpVtbl.contents
        vtbl.Release(proveedor_servicio)
        
        return ctypes.cast(ppv, POINTER(IVirtualDesktopManagerInternal))
    
    def obtener_coleccion_vistas_aplicacion(self) -> POINTER(IApplicationViewCollection):
        """Obtener IApplicationViewCollection"""
        proveedor_servicio = self.obtener_immersive_shell()
        
        ppv = self.consultar_servicio(
            proveedor_servicio,
            IID_IApplicationViewCollection,
            IID_IApplicationViewCollection
        )
        
        # Liberar proveedor de servicio
        vtbl = proveedor_servicio.contents.lpVtbl.contents
        vtbl.Release(proveedor_servicio)
        
        return ctypes.cast(ppv, POINTER(IApplicationViewCollection))
    
    def obtener_aplicaciones_ancladas(self) -> POINTER(IVirtualDesktopPinnedApps):
        """Obtener IVirtualDesktopPinnedApps"""
        proveedor_servicio = self.obtener_immersive_shell()
        
        ppv = self.consultar_servicio(
            proveedor_servicio,
            CLSID_VirtualDesktopPinnedApps,
            IID_IVirtualDesktopPinnedApps
        )
        
        # Liberar proveedor de servicio
        vtbl = proveedor_servicio.contents.lpVtbl.contents
        vtbl.Release(proveedor_servicio)
        
        return ctypes.cast(ppv, POINTER(IVirtualDesktopPinnedApps))
    
    def liberar_interfaz(self, puntero_interfaz):
        """Liberar referencia a interfaz COM"""
        if puntero_interfaz:
            try:
                vtbl = puntero_interfaz.contents.lpVtbl.contents
                vtbl.Release(puntero_interfaz)
            except:
                pass
    
    def __del__(self):
        """Cleanup COM al destruir"""
        if self.inicializado:
            try:
                ole32.CoUninitialize()
            except:
                pass


def obtener_elementos_array_objetos(array: POINTER(IObjectArray), iid_elemento: GUID) -> List[c_void_p]:
    """Obtener elementos de un IObjectArray"""
    elementos = []
    
    vtbl = array.contents.lpVtbl.contents
    
    # Obtener count
    count = UINT()
    resultado = vtbl.GetCount(array, pointer(count))
    if resultado != S_OK:
        return elementos
    
    # Obtener cada elemento
    for i in range(count.value):
        ppv = c_void_p()
        resultado = vtbl.GetAt(array, i, pointer(iid_elemento), pointer(ppv))
        if resultado == S_OK and ppv:
            elementos.append(ppv)
    
    return elementos
