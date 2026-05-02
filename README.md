# 📨 MBox Manager

Aplicación de escritorio para gestionar y explorar archivos `.mbox` de gran tamaño. Permite cargar, filtrar, previsualizar y exportar correos, así como descargar sus archivos adjuntos, todo desde una interfaz gráfica moderna.

---

## ✨ Características

- **Carga eficiente** de archivos `.mbox` con barra de progreso en tiempo real (hilo separado)
- **Vista previa** del contenido de cada correo: remitente, destinatario, fecha, cuerpo y adjuntos
- **Filtros** por remitente, palabra clave (asunto/cuerpo) y rango de fechas
- **Descarga de adjuntos** del correo seleccionado, con selector de carpeta de destino
- **Exportación** de la lista de correos a Excel (`.xlsx`) o CSV
- Ordenación de columnas en la tabla haciendo clic en los encabezados
- Detección de adjuntos tanto `attachment` como `inline`
- Tema oscuro con interfaz basada en CustomTkinter

---

## 📦 Instalación

### Requisitos

- Python 3.10 o superior

### Dependencias

```bash
pip install customtkinter pandas openpyxl
```

### Clonar el repositorio

```bash
git clone https://github.com/juakolp25/mbox-manager.git
cd mbox-manager
```

---

## 🚀 Uso

```bash
python mbox_manager.py
```

1. Haz clic en **Abrir archivo .mbox** y selecciona tu archivo.
2. Pulsa **Cargar todos los correos** y espera a que se complete la carga.
3. Usa los filtros del panel izquierdo para buscar correos.
4. Haz clic en una fila para previsualizar el correo en el panel inferior.
5. Si el correo tiene adjuntos (indicados con 📎), pulsa **Descargar Adjuntos** y elige la carpeta de destino.
6. Exporta la lista filtrada a Excel o CSV con los botones de la sección **Exportar**.

---

## 📁 Estructura del proyecto

```
mbox-manager/
└── mbox_manager.py   # Aplicación completa (fichero único)
```

---

## 🔧 Detalles técnicos

| Aspecto | Solución |
|---|---|
| Carga de correos | Hilo separado con `threading` para no bloquear la UI |
| Comunicación UI ↔ hilo | `queue.Queue` con polling mediante `after()` |
| Serialización de mensajes | `bytes(msg)` con `policy=email.policy.compat32` para compatibilidad total |
| Detección de adjuntos | Por presencia de `filename` en la parte MIME (no solo por `Content-Disposition`) |
| Decodificación de cabeceras | `email.header.decode_header` con fallback a UTF-8 |
| Exportación | `pandas` + `openpyxl` para Excel, CSV con codificación UTF-8 BOM |

---

## 📋 Dependencias

| Paquete | Uso |
|---|---|
| `customtkinter` | Interfaz gráfica con tema oscuro |
| `pandas` | Exportación a Excel y CSV |
| `openpyxl` | Motor de escritura de archivos `.xlsx` |
| `mailbox`, `email` | Lectura y parseo de archivos `.mbox` (stdlib) |
| `tkinter` | Diálogos de archivo y tabla (stdlib) |

---

## 📄 Licencia

Este proyecto está bajo la licencia MIT. Consulta el archivo [LICENSE](LICENSE) para más detalles.
