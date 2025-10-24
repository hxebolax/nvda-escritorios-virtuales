# -*- coding: utf-8 -*-
"""
Complemento de Escritorios Virtuales para NVDA
Permite gestionar escritorios virtuales de Windows 10/11 con comandos de teclado

Autor: XeBoLaX
Versión: 1.0.0
"""

import globalPluginHandler
import ui
import scriptHandler
import wx
import addonHandler

# Inicializar traducciones
addonHandler.initTranslation()

# Importar la librería de escritorios virtuales
try:
	from .escritorios_virtuales import EscritorioVirtual, VistaAplicacion, GestorEscritorios, ExcepcionEVD
	LIBRERIA_DISPONIBLE = True
except ImportError:
	LIBRERIA_DISPONIBLE = False


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	"""Plugin global para gestión de escritorios virtuales"""
	
	def __init__(self):
		super(GlobalPlugin, self).__init__()
		
		if not LIBRERIA_DISPONIBLE:
			wx.CallAfter(
				ui.message,
				_("Error: La librería escritorios-virtuales no está instalada. "
				  "Por favor, instálala con: pip install escritorios-virtuales")
			)
			return
		
		# Inicializar gestor
		try:
			self.gestor = GestorEscritorios()
		except Exception as e:
			wx.CallAfter(
				ui.message,
				_("Error al inicializar el gestor de escritorios: {error}").format(error=str(e))
			)
			self.gestor = None
	
	def terminate(self):
		"""Limpieza al cerrar NVDA"""
		super(GlobalPlugin, self).terminate()
	
	@scriptHandler.script(
		description=_("Anuncia el escritorio virtual actual"),
		category=_("Escritorios Virtuales"),
		gesture="kb:NVDA+control+d"
	)
	def script_anunciarEscritorioActual(self, gesture):
		"""Anuncia el escritorio actual"""
		if not LIBRERIA_DISPONIBLE or not self.gestor:
			ui.message(_("Gestor de escritorios no disponible"))
			return
		
		try:
			actual = self.gestor.obtener_escritorio_actual()
			escritorios = self.gestor.obtener_escritorios()
			total = len(escritorios)
			
			mensaje = _("Escritorio {numero} de {total}").format(
				numero=actual.numero,
				total=total
			)
			ui.message(mensaje)
		except Exception as e:
			ui.message(_("Error: {error}").format(error=str(e)))
	
	@scriptHandler.script(
		description=_("Lista todos los escritorios virtuales"),
		category=_("Escritorios Virtuales"),
		gesture="kb:NVDA+control+shift+d"
	)
	def script_listarEscritorios(self, gesture):
		"""Lista todos los escritorios"""
		if not LIBRERIA_DISPONIBLE or not self.gestor:
			ui.message(_("Gestor de escritorios no disponible"))
			return
		
		try:
			escritorios = self.gestor.obtener_escritorios()
			actual = self.gestor.obtener_escritorio_actual()
			
			mensaje = _("Total de escritorios: {total}. ").format(total=len(escritorios))
			
			for escritorio in escritorios:
				es_actual = _(" (actual)") if escritorio.id == actual.id else ""
				ventanas = escritorio.obtener_ventanas()
				mensaje += _("Escritorio {numero}: {ventanas} ventanas{actual}. ").format(
					numero=escritorio.numero,
					ventanas=len(ventanas),
					actual=es_actual
				)
			
			ui.message(mensaje)
		except Exception as e:
			ui.message(_("Error: {error}").format(error=str(e)))
	
	@scriptHandler.script(
		description=_("Crea un nuevo escritorio virtual"),
		category=_("Escritorios Virtuales"),
		gesture="kb:NVDA+control+n"
	)
	def script_crearEscritorio(self, gesture):
		"""Crea un nuevo escritorio"""
		if not LIBRERIA_DISPONIBLE or not self.gestor:
			ui.message(_("Gestor de escritorios no disponible"))
			return
		
		try:
			nuevo = self.gestor.crear_escritorio()
			ui.message(_("Escritorio {numero} creado").format(numero=nuevo.numero))
		except Exception as e:
			ui.message(_("Error al crear escritorio: {error}").format(error=str(e)))
	
	@scriptHandler.script(
		description=_("Cambia al escritorio virtual anterior"),
		category=_("Escritorios Virtuales"),
		gesture="kb:NVDA+control+leftArrow"
	)
	def script_escritorioAnterior(self, gesture):
		"""Cambia al escritorio anterior"""
		if not LIBRERIA_DISPONIBLE or not self.gestor:
			ui.message(_("Gestor de escritorios no disponible"))
			return
		
		try:
			escritorios = self.gestor.obtener_escritorios()
			actual = self.gestor.obtener_escritorio_actual()
			
			# Encontrar índice actual
			indice_actual = None
			for i, escritorio in enumerate(escritorios):
				if escritorio.id == actual.id:
					indice_actual = i
					break
			
			if indice_actual is None:
				ui.message(_("Error: No se pudo determinar el escritorio actual"))
				return
			
			# Calcular índice anterior (circular)
			indice_anterior = (indice_actual - 1) % len(escritorios)
			escritorio_anterior = escritorios[indice_anterior]
			
			escritorio_anterior.ir()
			ui.message(_("Escritorio {numero}").format(numero=escritorio_anterior.numero))
		except Exception as e:
			ui.message(_("Error: {error}").format(error=str(e)))
	
	@scriptHandler.script(
		description=_("Cambia al escritorio virtual siguiente"),
		category=_("Escritorios Virtuales"),
		gesture="kb:NVDA+control+rightArrow"
	)
	def script_escritorioSiguiente(self, gesture):
		"""Cambia al escritorio siguiente"""
		if not LIBRERIA_DISPONIBLE or not self.gestor:
			ui.message(_("Gestor de escritorios no disponible"))
			return
		
		try:
			escritorios = self.gestor.obtener_escritorios()
			actual = self.gestor.obtener_escritorio_actual()
			
			# Encontrar índice actual
			indice_actual = None
			for i, escritorio in enumerate(escritorios):
				if escritorio.id == actual.id:
					indice_actual = i
					break
			
			if indice_actual is None:
				ui.message(_("Error: No se pudo determinar el escritorio actual"))
				return
			
			# Calcular índice siguiente (circular)
			indice_siguiente = (indice_actual + 1) % len(escritorios)
			escritorio_siguiente = escritorios[indice_siguiente]
			
			escritorio_siguiente.ir()
			ui.message(_("Escritorio {numero}").format(numero=escritorio_siguiente.numero))
		except Exception as e:
			ui.message(_("Error: {error}").format(error=str(e)))
	
	@scriptHandler.script(
		description=_("Mueve la ventana actual al escritorio anterior"),
		category=_("Escritorios Virtuales"),
		gesture="kb:NVDA+control+shift+leftArrow"
	)
	def script_moverVentanaAnterior(self, gesture):
		"""Mueve la ventana actual al escritorio anterior"""
		if not LIBRERIA_DISPONIBLE or not self.gestor:
			ui.message(_("Gestor de escritorios no disponible"))
			return
		
		try:
			ventana = self.gestor.obtener_ventana_actual()
			if not ventana:
				ui.message(_("No hay ventana enfocada"))
				return
			
			escritorios = self.gestor.obtener_escritorios()
			actual = self.gestor.obtener_escritorio_actual()
			
			# Encontrar índice actual
			indice_actual = None
			for i, escritorio in enumerate(escritorios):
				if escritorio.id == actual.id:
					indice_actual = i
					break
			
			if indice_actual is None:
				ui.message(_("Error: No se pudo determinar el escritorio actual"))
				return
			
			# Calcular índice anterior
			indice_anterior = (indice_actual - 1) % len(escritorios)
			escritorio_destino = escritorios[indice_anterior]
			
			ventana.mover_a_escritorio(escritorio_destino)
			ui.message(_("Ventana movida al escritorio {numero}").format(numero=escritorio_destino.numero))
		except Exception as e:
			ui.message(_("Error: {error}").format(error=str(e)))
	
	@scriptHandler.script(
		description=_("Mueve la ventana actual al escritorio siguiente"),
		category=_("Escritorios Virtuales"),
		gesture="kb:NVDA+control+shift+rightArrow"
	)
	def script_moverVentanaSiguiente(self, gesture):
		"""Mueve la ventana actual al escritorio siguiente"""
		if not LIBRERIA_DISPONIBLE or not self.gestor:
			ui.message(_("Gestor de escritorios no disponible"))
			return
		
		try:
			ventana = self.gestor.obtener_ventana_actual()
			if not ventana:
				ui.message(_("No hay ventana enfocada"))
				return
			
			escritorios = self.gestor.obtener_escritorios()
			actual = self.gestor.obtener_escritorio_actual()
			
			# Encontrar índice actual
			indice_actual = None
			for i, escritorio in enumerate(escritorios):
				if escritorio.id == actual.id:
					indice_actual = i
					break
			
			if indice_actual is None:
				ui.message(_("Error: No se pudo determinar el escritorio actual"))
				return
			
			# Calcular índice siguiente
			indice_siguiente = (indice_actual + 1) % len(escritorios)
			escritorio_destino = escritorios[indice_siguiente]
			
			ventana.mover_a_escritorio(escritorio_destino)
			ui.message(_("Ventana movida al escritorio {numero}").format(numero=escritorio_destino.numero))
		except Exception as e:
			ui.message(_("Error: {error}").format(error=str(e)))
	
	@scriptHandler.script(
		description=_("Ancla o desancla la ventana actual (mostrar en todos los escritorios)"),
		category=_("Escritorios Virtuales"),
		gesture="kb:NVDA+control+p"
	)
	def script_anclarVentana(self, gesture):
		"""Ancla o desancla la ventana actual"""
		if not LIBRERIA_DISPONIBLE or not self.gestor:
			ui.message(_("Gestor de escritorios no disponible"))
			return
		
		try:
			ventana = self.gestor.obtener_ventana_actual()
			if not ventana:
				ui.message(_("No hay ventana enfocada"))
				return
			
			if ventana.esta_anclada():
				ventana.desanclar()
				ui.message(_("Ventana desanclada"))
			else:
				ventana.anclar()
				ui.message(_("Ventana anclada en todos los escritorios"))
		except Exception as e:
			ui.message(_("Error: {error}").format(error=str(e)))
	
	@scriptHandler.script(
		description=_("Anuncia información de la ventana actual"),
		category=_("Escritorios Virtuales"),
		gesture="kb:NVDA+control+w"
	)
	def script_infoVentana(self, gesture):
		"""Anuncia información de la ventana actual"""
		if not LIBRERIA_DISPONIBLE or not self.gestor:
			ui.message(_("Gestor de escritorios no disponible"))
			return
		
		try:
			ventana = self.gestor.obtener_ventana_actual()
			if not ventana:
				ui.message(_("No hay ventana enfocada"))
				return
			
			escritorio = ventana.escritorio
			anclada = _("anclada") if ventana.esta_anclada() else _("no anclada")
			
			mensaje = _("Ventana: {titulo}. Escritorio {numero}. {anclada}").format(
				titulo=ventana.titulo,
				numero=escritorio.numero,
				anclada=anclada
			)
			ui.message(mensaje)
		except Exception as e:
			ui.message(_("Error: {error}").format(error=str(e)))
	
	@scriptHandler.script(
		description=_("Ir a un escritorio específico (abre diálogo)"),
		category=_("Escritorios Virtuales"),
		gesture="kb:NVDA+control+g"
	)
	def script_irAEscritorio(self, gesture):
		"""Abre diálogo para ir a un escritorio específico"""
		if not LIBRERIA_DISPONIBLE or not self.gestor:
			ui.message(_("Gestor de escritorios no disponible"))
			return
		
		try:
			escritorios = self.gestor.obtener_escritorios()
			
			# Crear diálogo con mejor accesibilidad
			def mostrarDialogo():
				dlg = wx.TextEntryDialog(
					None,
					_("Introduce el número de escritorio (1-{max}):").format(max=len(escritorios)),
					_("Ir a Escritorio")
				)
				
				# Hacer que NVDA lea el diálogo
				dlg.SetFocus()
				
				if dlg.ShowModal() == wx.ID_OK:
					try:
						numero = int(dlg.GetValue())
						if 1 <= numero <= len(escritorios):
							escritorios[numero - 1].ir()
							wx.CallAfter(ui.message, _("Escritorio {numero}").format(numero=numero))
						else:
							wx.CallAfter(ui.message, _("Número de escritorio inválido"))
					except ValueError:
						wx.CallAfter(ui.message, _("Debes introducir un número"))
				
				dlg.Destroy()
			
			wx.CallAfter(mostrarDialogo)
		except Exception as e:
			ui.message(_("Error: {error}").format(error=str(e)))
	
	@scriptHandler.script(
		description=_("Elimina el escritorio actual"),
		category=_("Escritorios Virtuales"),
		gesture="kb:NVDA+control+delete"
	)
	def script_eliminarEscritorioActual(self, gesture):
		"""Elimina el escritorio actual"""
		if not LIBRERIA_DISPONIBLE or not self.gestor:
			ui.message(_("Gestor de escritorios no disponible"))
			return
		
		try:
			escritorios = self.gestor.obtener_escritorios()
			
			if len(escritorios) <= 1:
				ui.message(_("No se puede eliminar el último escritorio"))
				return
			
			actual = self.gestor.obtener_escritorio_actual()
			numero_actual = actual.numero
			
			# Confirmar eliminación
			def confirmarEliminacion():
				dlg = wx.MessageDialog(
					None,
					_("¿Estás seguro de que quieres eliminar el escritorio {numero}?").format(numero=numero_actual),
					_("Confirmar Eliminación"),
					wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION
				)
				
				if dlg.ShowModal() == wx.ID_YES:
					try:
						# Encontrar escritorio de respaldo (el primero que no sea el actual)
						respaldo = None
						for escritorio in escritorios:
							if escritorio.id != actual.id:
								respaldo = escritorio
								break
						
						if respaldo:
							actual.eliminar(respaldo=respaldo)
							wx.CallAfter(ui.message, _("Escritorio {numero} eliminado").format(numero=numero_actual))
						else:
							wx.CallAfter(ui.message, _("Error: No se encontró escritorio de respaldo"))
					except Exception as e:
						wx.CallAfter(ui.message, _("Error al eliminar: {error}").format(error=str(e)))
				
				dlg.Destroy()
			
			wx.CallAfter(confirmarEliminacion)
		except Exception as e:
			ui.message(_("Error: {error}").format(error=str(e)))
	
	@scriptHandler.script(
		description=_("Elimina un escritorio específico (abre diálogo)"),
		category=_("Escritorios Virtuales"),
		gesture="kb:NVDA+control+shift+delete"
	)
	def script_eliminarEscritorioEspecifico(self, gesture):
		"""Abre diálogo para eliminar un escritorio específico"""
		if not LIBRERIA_DISPONIBLE or not self.gestor:
			ui.message(_("Gestor de escritorios no disponible"))
			return
		
		try:
			escritorios = self.gestor.obtener_escritorios()
			
			if len(escritorios) <= 1:
				ui.message(_("No se puede eliminar el último escritorio"))
				return
			
			def mostrarDialogo():
				# Diálogo para seleccionar escritorio
				dlg = wx.TextEntryDialog(
					None,
					_("Introduce el número del escritorio a eliminar (1-{max}):").format(max=len(escritorios)),
					_("Eliminar Escritorio")
				)
				
				if dlg.ShowModal() == wx.ID_OK:
					try:
						numero = int(dlg.GetValue())
						if 1 <= numero <= len(escritorios):
							escritorio_eliminar = escritorios[numero - 1]
							
							# Confirmar eliminación
							dlg2 = wx.MessageDialog(
								None,
								_("¿Estás seguro de que quieres eliminar el escritorio {numero}?").format(numero=numero),
								_("Confirmar Eliminación"),
								wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION
							)
							
							if dlg2.ShowModal() == wx.ID_YES:
								# Encontrar escritorio de respaldo
								respaldo = None
								for escritorio in escritorios:
									if escritorio.id != escritorio_eliminar.id:
										respaldo = escritorio
										break
								
								if respaldo:
									escritorio_eliminar.eliminar(respaldo=respaldo)
									wx.CallAfter(ui.message, _("Escritorio {numero} eliminado").format(numero=numero))
								else:
									wx.CallAfter(ui.message, _("Error: No se encontró escritorio de respaldo"))
							
							dlg2.Destroy()
						else:
							wx.CallAfter(ui.message, _("Número de escritorio inválido"))
					except ValueError:
						wx.CallAfter(ui.message, _("Debes introducir un número"))
					except Exception as e:
						wx.CallAfter(ui.message, _("Error: {error}").format(error=str(e)))
				
				dlg.Destroy()
			
			wx.CallAfter(mostrarDialogo)
		except Exception as e:
			ui.message(_("Error: {error}").format(error=str(e)))
	
	@scriptHandler.script(
		description=_("Mueve la ventana actual a un escritorio específico (abre diálogo)"),
		category=_("Escritorios Virtuales"),
		gesture="kb:NVDA+control+m"
	)
	def script_moverVentanaAEscritorio(self, gesture):
		"""Abre diálogo para mover ventana a escritorio específico"""
		if not LIBRERIA_DISPONIBLE or not self.gestor:
			ui.message(_("Gestor de escritorios no disponible"))
			return
		
		try:
			ventana = self.gestor.obtener_ventana_actual()
			if not ventana:
				ui.message(_("No hay ventana enfocada"))
				return
			
			escritorios = self.gestor.obtener_escritorios()
			
			def mostrarDialogo():
				dlg = wx.TextEntryDialog(
					None,
					_("Introduce el número del escritorio destino (1-{max}):").format(max=len(escritorios)),
					_("Mover Ventana a Escritorio")
				)
				
				if dlg.ShowModal() == wx.ID_OK:
					try:
						numero = int(dlg.GetValue())
						if 1 <= numero <= len(escritorios):
							escritorio_destino = escritorios[numero - 1]
							ventana.mover_a_escritorio(escritorio_destino)
							wx.CallAfter(ui.message, _("Ventana movida al escritorio {numero}").format(numero=numero))
						else:
							wx.CallAfter(ui.message, _("Número de escritorio inválido"))
					except ValueError:
						wx.CallAfter(ui.message, _("Debes introducir un número"))
					except Exception as e:
						wx.CallAfter(ui.message, _("Error: {error}").format(error=str(e)))
				
				dlg.Destroy()
			
			wx.CallAfter(mostrarDialogo)
		except Exception as e:
			ui.message(_("Error: {error}").format(error=str(e)))
	
	@scriptHandler.script(
		description=_("Cuenta las ventanas en el escritorio actual"),
		category=_("Escritorios Virtuales"),
		gesture="kb:NVDA+control+c"
	)
	def script_contarVentanas(self, gesture):
		"""Cuenta las ventanas en el escritorio actual"""
		if not LIBRERIA_DISPONIBLE or not self.gestor:
			ui.message(_("Gestor de escritorios no disponible"))
			return
		
		try:
			actual = self.gestor.obtener_escritorio_actual()
			ventanas = actual.obtener_ventanas()
			ventanas_especificas = [v for v in ventanas if not v.esta_anclada()]
			ventanas_ancladas = [v for v in ventanas if v.esta_anclada()]
			
			mensaje = _("Escritorio {numero}: {especificas} ventanas específicas, {ancladas} ancladas").format(
				numero=actual.numero,
				especificas=len(ventanas_especificas),
				ancladas=len(ventanas_ancladas)
			)
			ui.message(mensaje)
		except Exception as e:
			ui.message(_("Error: {error}").format(error=str(e)))
