# Changelog - Escritorios Virtuales para NVDA

Todos los cambios notables en este proyecto serán documentados en este archivo.

## [1.0.0] - 2025-10-24

### Añadido

#### Funcionalidades Completas de la Librería
- ✅ Navegación entre escritorios (anterior/siguiente)
- ✅ Ir a escritorio específico (con diálogo)
- ✅ Crear nuevos escritorios
- ✅ **Eliminar escritorio actual** (con confirmación)
- ✅ **Eliminar escritorio específico** (con diálogo y confirmación)
- ✅ **Contar ventanas en escritorio actual** (específicas y ancladas)
- ✅ Mover ventanas al escritorio anterior/siguiente
- ✅ **Mover ventana a escritorio específico** (con diálogo)
- ✅ Anclar/desanclar ventanas
- ✅ Información de ventana actual
- ✅ Listar todos los escritorios

#### Comandos de Teclado (14 comandos)
1. **NVDA+Ctrl+D** - Anuncia escritorio actual
2. **NVDA+Ctrl+Shift+D** - Lista todos los escritorios
3. **NVDA+Ctrl+C** - Contar ventanas en escritorio actual ⭐ NUEVO
4. **NVDA+Ctrl+←** - Escritorio anterior
5. **NVDA+Ctrl+→** - Escritorio siguiente
6. **NVDA+Ctrl+G** - Ir a escritorio específico (diálogo mejorado)
7. **NVDA+Ctrl+N** - Crear nuevo escritorio
8. **NVDA+Ctrl+Delete** - Eliminar escritorio actual ⭐ NUEVO
9. **NVDA+Ctrl+Shift+Delete** - Eliminar escritorio específico ⭐ NUEVO
10. **NVDA+Ctrl+W** - Info de ventana actual
11. **NVDA+Ctrl+M** - Mover ventana a escritorio específico ⭐ NUEVO
12. **NVDA+Ctrl+Shift+←** - Mover ventana al anterior
13. **NVDA+Ctrl+Shift+→** - Mover ventana al siguiente
14. **NVDA+Ctrl+P** - Anclar/desanclar ventana

#### Mejoras de Accesibilidad
- ✅ **Diálogos accesibles mejorados** - NVDA ahora lee correctamente los diálogos
- ✅ **Confirmaciones de eliminación** - Seguridad al eliminar escritorios
- ✅ **Mensajes claros** - Retroalimentación hablada mejorada
- ✅ **Manejo de errores robusto** - Mensajes de error informativos

#### Características Técnicas
- Integración completa con librería `escritorios-virtuales`
- Uso de `wx.CallAfter` para mejor accesibilidad de diálogos
- Validación de entrada en todos los diálogos
- Protección contra eliminación del último escritorio
- Navegación circular entre escritorios

#### Documentación
- README completo con todos los comandos
- Guía de instalación detallada
- Comandos rápidos de referencia
- Documentación del usuario en español
- Ejemplos de uso para cada comando
- Guía de empaquetado y distribución

### Corregido
- 🐛 **Diálogos silenciosos** - NVDA ahora lee correctamente los diálogos de entrada
- 🐛 **Accesibilidad de confirmaciones** - Diálogos de confirmación totalmente accesibles

### Mejorado
- 🔧 Mejor manejo de errores en todas las operaciones
- 🔧 Validación de entrada en diálogos numéricos
- 🔧 Mensajes más descriptivos y claros
- 🔧 Código más robusto y mantenible

## Funcionalidades por Categoría

### Información (4 comandos)
- Anuncia escritorio actual
- Lista todos los escritorios
- Contar ventanas
- Info de ventana actual

### Navegación (4 comandos)
- Escritorio anterior
- Escritorio siguiente
- Ir a escritorio específico
- (Navegación circular automática)

### Gestión de Escritorios (4 comandos)
- Crear nuevo escritorio
- Eliminar escritorio actual
- Eliminar escritorio específico
- Contar ventanas

### Gestión de Ventanas (4 comandos)
- Mover ventana al anterior
- Mover ventana al siguiente
- Mover ventana a específico
- Anclar/desanclar ventana

## Compatibilidad

- **NVDA:** 2019.3 - 2024.1 (y posteriores)
- **Windows:** 10 (build 10240+) y 11
- **Python:** 3.7+ (incluido con NVDA)
- **Librería:** escritorios-virtuales 1.0.0+

## Notas de Migración

### Desde versión anterior (si existiera)
No aplica - Esta es la primera versión estable.

### Nuevos usuarios
1. Instalar librería: `pip install escritorios-virtuales`
2. Instalar complemento: Abrir archivo `.nvda-addon`
3. Reiniciar NVDA
4. Probar con **NVDA+Ctrl+D**

## Problemas Conocidos

Ninguno reportado en esta versión.

## Próximas Versiones

### Planificado para v1.1.0
- [ ] Renombrar escritorios (si API lo permite)
- [ ] Atajos personalizables desde diálogo
- [ ] Más idiomas (inglés, etc.)
- [ ] Estadísticas de uso

### Planificado para v1.2.0
- [ ] Perfiles de escritorios
- [ ] Guardar/restaurar configuración
- [ ] Reglas automáticas de ventanas

## Agradecimientos

- Comunidad de NVDA por su excelente framework
- Usuarios que prueban y reportan problemas
- Desarrolladores de APIs de Windows

---

**Versión actual:** 1.0.0  
**Fecha:** 2024-10-24  
**Autor:** XeBoLaX  
**Licencia:** MIT
