# Modelo lógico | Kardex PRO

## Navegación funcional del sistema

La navegación está organizada según la secuencia real de trabajo de un Kardex:

1. **Inicio / Ruta del Kardex**: pantalla guía que muestra qué falta configurar y cuál es el siguiente paso recomendado.
2. **Catálogos base**: mantenimiento de productos, proveedores, solicitantes y personal.
3. **Administración**: usuarios, roles, PATH automático y diagnóstico de base.
4. **Registrar movimientos**: ingreso, salida, devolución y ajustes.
5. **Kardex consolidado**: vista operativa de una fila por producto/lote.
6. **Stock y alertas**: control de vencimiento, saldo y stock mínimo.
7. **Dashboard ejecutivo**: resumen visual de inventario.
8. **Reportes y exportación**: filtros, análisis y descarga en Excel.
9. **Importar Kardex anterior**: migración histórica desde el archivo previo con macros.

Esta estructura evita que el usuario registre movimientos sin haber creado antes los catálogos necesarios.


## Principio del sistema

El sistema no guarda el stock como una tabla manual. El stock se calcula desde la hoja `Movimientos`.

Cada movimiento suma o resta inventario según su tipo:

| Tipo de movimiento | Efecto |
|---|---:|
| Ingreso | Suma |
| Devolución | Suma |
| Ajuste entrada | Suma |
| Salida | Resta |
| Ajuste salida | Resta |

El stock y el Kardex consolidado se agrupan por:

```text
producto_id + producto + marca + lote + fecha_vencimiento + unidad
```

## Hojas de base

### Productos
Catálogo maestro. Define nombre, código, categoría, marca default, unidad default, stock mínimo y días para alerta de vencimiento.

### Proveedores
Catálogo de proveedores para registrar entradas.

### Solicitantes
Catálogo de unidades/personas que reciben productos en las salidas.

### Personal
Personal responsable de entregar o recibir productos.

### Usuarios
Control de acceso al sistema.

Campos principales:

```text
usuario_id, usuario, nombre, rol, password_hash, path_verificacion, activo, fecha_creacion
```

### Config
Parámetros globales del sistema:

```text
dias_alerta_global
moneda
institucion
path_verificacion
version_sistema
```

### Movimientos
Bitácora transaccional. Es la hoja principal para auditoría y reportes.

## Seguridad

El login solicita tres datos:

1. Usuario
2. Contraseña
3. PATH temporal generado automáticamente en la pantalla de login

El usuario inicial es:

```text
admin / admin123
```

El PATH no es fijo. El sistema muestra un código `KDX-XXXX-XXXX` que el usuario copia y pega para completar la validación. La contraseña se guarda como hash SHA-256, no como texto visible.

## Alertas

El sistema clasifica cada lote así:

| Estado | Regla |
|---|---|
| Sin stock | stock <= 0 |
| Vencido | fecha de vencimiento menor que hoy |
| Por vencer | días para vencer <= días de alerta del producto |
| Stock bajo | stock actual <= stock mínimo |
| Disponible | no cumple reglas críticas |

## Kardex consolidado

La hoja `Movimientos` mantiene todos los registros transaccionales para auditoría. Sobre esa hoja, el sistema calcula una vista `Kardex consolidado por lote` con una sola fila por producto/lote.

Cuando se registra una salida, el sistema no borra ni edita la entrada original; simplemente recalcula la vista consolidada para reflejar en la misma fila:

- Entrada total
- Salida acumulada
- Número de salidas
- Último entregado a
- Personal que entregó
- Fecha última salida
- Detalle de salidas
- Saldo actual

Esto permite tener una vista operativa limpia y una bitácora histórica completa al mismo tiempo.

## Reportes

El sistema puede exportar un Excel con:

- Kardex_Consolidado
- Stock_Actual
- Movimientos
- Alertas_Vencimiento
- Stock_Bajo
- Catalogo_Productos
- Catalogo_Proveedores
- Catalogo_Solicitantes
- Catalogo_Personal


## Seguridad de acceso

El sistema utiliza usuario y contraseña más un PATH temporal. En cada pantalla de login se genera un código en formato `KDX-XXXX-XXXX`; el usuario debe copiarlo y pegarlo en el campo de verificación. Este PATH no se configura manualmente en usuarios ni en administración, reduciendo el uso de claves fijas compartidas.