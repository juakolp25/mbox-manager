"""
MBox Manager - Aplicación de escritorio para gestionar archivos .mbox grandes
Requiere: pip install customtkinter pandas openpyxl
"""
 
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
import mailbox
import email
import email.policy
from email.header import decode_header
from email.utils import parsedate_to_datetime
import threading
import pandas as pd
import os
import re
from datetime import datetime
import queue
import pathlib
 
 
# ─── Configuración de tema ───────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")
 
 
# ─── Utilidades ──────────────────────────────────────────────────────────────
 
def decode_mime_header(value):
    if not value:
        return ""
    try:
        parts = decode_header(value)
        decoded = []
        for part, enc in parts:
            if isinstance(part, bytes):
                enc = enc or "utf-8"
                try:
                    decoded.append(part.decode(enc, errors="replace"))
                except (LookupError, UnicodeDecodeError):
                    decoded.append(part.decode("utf-8", errors="replace"))
            else:
                decoded.append(str(part))
        return " ".join(decoded).strip()
    except Exception:
        return str(value)
 
 
def get_body(msg):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            cdispo = str(part.get("Content-Disposition", ""))
            if ctype == "text/plain" and "attachment" not in cdispo:
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    try:
                        body = payload.decode(charset, errors="replace")
                    except (LookupError, UnicodeDecodeError):
                        body = payload.decode("utf-8", errors="replace")
                    break
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            try:
                body = payload.decode(charset, errors="replace")
            except (LookupError, UnicodeDecodeError):
                body = payload.decode("utf-8", errors="replace")
    return body[:2000]
 
 
def parse_date(msg):
    date_str = msg.get("Date", "")
    if not date_str:
        return None
    try:
        return parsedate_to_datetime(date_str).replace(tzinfo=None)
    except Exception:
        return None
 
 
def get_attachments(msg):
    attachments = []
    if msg.is_multipart():
        for part in msg.walk():
            # Considerar adjunto si tiene nombre de archivo,
            # independientemente de si Content-Disposition dice 'attachment' o 'inline'
            filename = part.get_filename()
            if not filename:
                continue
            filename = decode_mime_header(filename)
            filename = re.sub(r'[\\/*?:"<>|]', "_", filename).strip()
            if not filename:
                filename = "adjunto_sin_nombre"
            payload = part.get_payload(decode=True)
            if payload:
                attachments.append({"filename": filename, "data": payload})
    return attachments
 
 
def get_desktop_path():
    return pathlib.Path.home() / "Desktop"
 
 
# ─── Aplicación ──────────────────────────────────────────────────────────────
 
class MBoxManagerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("MBox Manager  •  Gestión de correos")
        self.geometry("1280x780")
        self.minsize(900, 600)
 
        self.mbox_path = tk.StringVar(value="Ningún archivo seleccionado")
        self.results: list[dict] = []
        self._all_emails: list[dict] = []
        self.loading = False
        self.msg_queue: queue.Queue = queue.Queue()
        self._selected_idx: int = -1
 
        self._build_ui()
        self._poll_queue()
 
    # ── UI ───────────────────────────────────────────────────────────────────
 
    def _build_ui(self):
        # Barra superior
        header = ctk.CTkFrame(self, height=56, corner_radius=0, fg_color="#0f172a")
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
 
        ctk.CTkLabel(header, text="📨  MBox Manager",
                     font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
                     text_color="#38bdf8").pack(side="left", padx=20, pady=10)
 
        self.status_lbl = ctk.CTkLabel(header, text="",
                                       font=ctk.CTkFont(size=12), text_color="#94a3b8")
        self.status_lbl.pack(side="right", padx=20)
 
        # Layout principal
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=12, pady=12)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)
 
        left = ctk.CTkFrame(main, width=280, corner_radius=12)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        left.pack_propagate(False)
        self._build_left_panel(left)
 
        right = ctk.CTkFrame(main, corner_radius=12, fg_color="#0f172a")
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(0, weight=3)
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)
        self._build_table(right)
        self._build_detail(right)
 
    def _build_left_panel(self, parent):
        pad = {"padx": 16, "pady": 5}
 
        # ── ARCHIVO ──────────────────────────────────────────────────────────
        ctk.CTkLabel(parent, text="ARCHIVO", font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#64748b").pack(anchor="w", padx=16, pady=(14, 2))
 
        ctk.CTkButton(parent, text="📂  Abrir archivo .mbox",
                      command=self._open_file, height=34,
                      fg_color="#0284c7", hover_color="#0369a1"
                      ).pack(fill="x", **pad)
 
        ctk.CTkLabel(parent, textvariable=self.mbox_path,
                     font=ctk.CTkFont(size=10), text_color="#94a3b8",
                     wraplength=240, justify="left"
                     ).pack(fill="x", padx=16, pady=(0, 2))
 
        self.load_btn = ctk.CTkButton(parent, text="⚡  Cargar todos los correos",
                                      command=self._load_emails, height=34,
                                      fg_color="#059669", hover_color="#047857",
                                      state="disabled")
        self.load_btn.pack(fill="x", **pad)
 
        self.progress = ctk.CTkProgressBar(parent, height=5, corner_radius=4)
        self.progress.pack(fill="x", padx=16, pady=(3, 4))
        self.progress.set(0)
 
        ctk.CTkFrame(parent, height=1, fg_color="#1e293b").pack(fill="x", padx=12, pady=3)
 
        # ── ADJUNTOS (visible siempre, arriba de filtros) ─────────────────
        ctk.CTkLabel(parent, text="ADJUNTOS", font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#64748b").pack(anchor="w", padx=16, pady=(6, 2))
 
        self.attach_btn = ctk.CTkButton(
            parent, text="📎  Descargar Adjuntos",
            command=self._download_attachments, height=38,
            fg_color="#be185d", hover_color="#9d174d",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.attach_btn.pack(fill="x", **pad)
 
        self.attach_info_lbl = ctk.CTkLabel(
            parent,
            text="Selecciona un correo para ver sus adjuntos",
            font=ctk.CTkFont(size=10), text_color="#475569",
            justify="center", wraplength=240
        )
        self.attach_info_lbl.pack(pady=(2, 4), padx=16)
 
        ctk.CTkFrame(parent, height=1, fg_color="#1e293b").pack(fill="x", padx=12, pady=3)
 
        # ── FILTROS ──────────────────────────────────────────────────────────
        ctk.CTkLabel(parent, text="FILTROS", font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#64748b").pack(anchor="w", padx=16, pady=(6, 2))
 
        ctk.CTkLabel(parent, text="Remitente (De:)", font=ctk.CTkFont(size=11),
                     text_color="#cbd5e1").pack(anchor="w", padx=16)
        self.filter_from = ctk.CTkEntry(parent, placeholder_text="nombre@ejemplo.com", height=30)
        self.filter_from.pack(fill="x", **pad)
 
        ctk.CTkLabel(parent, text="Palabra clave (asunto/cuerpo)", font=ctk.CTkFont(size=11),
                     text_color="#cbd5e1").pack(anchor="w", padx=16)
        self.filter_keyword = ctk.CTkEntry(parent, placeholder_text="factura, reunión...", height=30)
        self.filter_keyword.pack(fill="x", **pad)
 
        # Fechas en una fila
        dates_frame = ctk.CTkFrame(parent, fg_color="transparent")
        dates_frame.pack(fill="x", padx=16, pady=3)
        dates_frame.columnconfigure(0, weight=1)
        dates_frame.columnconfigure(1, weight=1)
 
        ctk.CTkLabel(dates_frame, text="Desde", font=ctk.CTkFont(size=11),
                     text_color="#cbd5e1").grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(dates_frame, text="Hasta", font=ctk.CTkFont(size=11),
                     text_color="#cbd5e1").grid(row=0, column=1, sticky="w", padx=(8, 0))
 
        self.filter_date_from = ctk.CTkEntry(dates_frame, placeholder_text="YYYY-MM-DD", height=30)
        self.filter_date_from.grid(row=1, column=0, sticky="ew")
 
        self.filter_date_to = ctk.CTkEntry(dates_frame, placeholder_text="YYYY-MM-DD", height=30)
        self.filter_date_to.grid(row=1, column=1, sticky="ew", padx=(8, 0))
 
        ctk.CTkButton(parent, text="🔍  Aplicar filtros",
                      command=self._apply_filters, height=34,
                      fg_color="#7c3aed", hover_color="#6d28d9"
                      ).pack(fill="x", **pad)
 
        ctk.CTkButton(parent, text="✖  Limpiar filtros",
                      command=self._clear_filters, height=28,
                      fg_color="transparent", border_width=1,
                      border_color="#334155", text_color="#94a3b8",
                      hover_color="#1e293b"
                      ).pack(fill="x", **pad)
 
        ctk.CTkFrame(parent, height=1, fg_color="#1e293b").pack(fill="x", padx=12, pady=3)
 
        # ── EXPORTAR ─────────────────────────────────────────────────────────
        ctk.CTkLabel(parent, text="EXPORTAR", font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#64748b").pack(anchor="w", padx=16, pady=(6, 2))
 
        export_row = ctk.CTkFrame(parent, fg_color="transparent")
        export_row.pack(fill="x", padx=16, pady=4)
        export_row.columnconfigure(0, weight=1)
        export_row.columnconfigure(1, weight=1)
 
        ctk.CTkButton(export_row, text="📊 Excel",
                      command=lambda: self._export("xlsx"), height=32,
                      fg_color="#d97706", hover_color="#b45309"
                      ).grid(row=0, column=0, sticky="ew", padx=(0, 4))
 
        ctk.CTkButton(export_row, text="📄 CSV",
                      command=lambda: self._export("csv"), height=32,
                      fg_color="#0f766e", hover_color="#0d6b63"
                      ).grid(row=0, column=1, sticky="ew")
 
        self.count_lbl = ctk.CTkLabel(parent, text="0 correos encontrados",
                                      font=ctk.CTkFont(size=10), text_color="#475569")
        self.count_lbl.pack(pady=(6, 4))
 
    def _build_table(self, parent):
        frame = ctk.CTkFrame(parent, corner_radius=10, fg_color="#0f172a")
        frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 4))
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
 
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Dark.Treeview",
                        background="#1e293b", foreground="#e2e8f0",
                        fieldbackground="#1e293b", borderwidth=0,
                        rowheight=26, font=("Segoe UI", 10))
        style.configure("Dark.Treeview.Heading",
                        background="#0f172a", foreground="#7dd3fc",
                        font=("Segoe UI", 10, "bold"), borderwidth=0)
        style.map("Dark.Treeview",
                  background=[("selected", "#1d4ed8")],
                  foreground=[("selected", "white")])
 
        cols = ("Fecha", "De", "Para", "Asunto", "📎")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings",
                                 style="Dark.Treeview", selectmode="browse")
 
        self.tree.heading("Fecha",  text="Fecha",  command=lambda: self._sort_table("Fecha"))
        self.tree.heading("De",     text="De",     command=lambda: self._sort_table("De"))
        self.tree.heading("Para",   text="Para",   command=lambda: self._sort_table("Para"))
        self.tree.heading("Asunto", text="Asunto", command=lambda: self._sort_table("Asunto"))
        self.tree.heading("📎",     text="📎")
 
        self.tree.column("Fecha",  width=140, minwidth=100)
        self.tree.column("De",     width=200, minwidth=120)
        self.tree.column("Para",   width=180, minwidth=100)
        self.tree.column("Asunto", width=370, minwidth=150)
        self.tree.column("📎",     width=36,  minwidth=36, anchor="center")
 
        vsb = ttk.Scrollbar(frame, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
 
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
 
        self.tree.bind("<<TreeviewSelect>>", self._on_row_select)
 
    def _build_detail(self, parent):
        frame = ctk.CTkFrame(parent, corner_radius=10, fg_color="#1e293b")
        frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)
 
        hdr = ctk.CTkFrame(frame, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", padx=12, pady=4)
        hdr.columnconfigure(0, weight=1)
 
        ctk.CTkLabel(hdr, text="Vista previa del correo seleccionado",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#64748b").grid(row=0, column=0, sticky="w")
 
        self.attach_count_lbl = ctk.CTkLabel(hdr, text="",
                                             font=ctk.CTkFont(size=11, weight="bold"),
                                             text_color="#f472b6")
        self.attach_count_lbl.grid(row=0, column=1, sticky="e")
 
        self.detail_text = ctk.CTkTextbox(frame, wrap="word",
                                          font=ctk.CTkFont(family="Consolas", size=11),
                                          fg_color="#0f172a", text_color="#cbd5e1",
                                          border_width=0)
        self.detail_text.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
 
    # ── Archivo ──────────────────────────────────────────────────────────────
 
    def _open_file(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo .mbox",
            filetypes=[("Archivos MBox", "*.mbox"), ("Todos", "*.*")]
        )
        if path:
            self.mbox_path.set(path)
            self.load_btn.configure(state="normal")
            self._set_status(f"Archivo: {os.path.basename(path)} "
                             f"({os.path.getsize(path)/1_048_576:.1f} MB)")
            self._all_emails = []
            self.results = []
            self._selected_idx = -1
            self._refresh_table([])
 
    def _load_emails(self):
        if self.loading:
            return
        path = self.mbox_path.get()
        if not os.path.isfile(path):
            messagebox.showerror("Error", "El archivo no existe o no es válido.")
            return
        self.loading = True
        self.load_btn.configure(state="disabled", text="Cargando…")
        self.progress.set(0)
        threading.Thread(target=self._load_thread, args=(path,), daemon=True).start()
 
    def _load_thread(self, path):
        emails = []
        try:
            mbox  = mailbox.mbox(path, factory=None)
            total = sum(1 for _ in mbox)
            mbox  = mailbox.mbox(path, factory=None)
 
            for i, msg in enumerate(mbox):
                try:
                    frm     = decode_mime_header(msg.get("From", ""))
                    to      = decode_mime_header(msg.get("To", ""))
                    subject = decode_mime_header(msg.get("Subject", "(sin asunto)"))
                    dt      = parse_date(msg)
                    body    = get_body(msg)
 
                    has_attach   = False
                    attach_names = []
                    if msg.is_multipart():
                        for part in msg.walk():
                            fn = part.get_filename()
                            if fn:
                                has_attach = True
                                attach_names.append(decode_mime_header(fn))
 
                    # Serializar con bytes() que respeta el formato mbox original
                    # para que email.message_from_bytes pueda reconstruirlo bien
                    try:
                        raw = bytes(msg)
                    except Exception:
                        raw = msg.as_bytes()

                    emails.append({
                        "Fecha":        dt,
                        "De":           frm,
                        "Para":         to,
                        "Asunto":       subject,
                        "Cuerpo":       body,
                        "TieneAdjunto": has_attach,
                        "NombresAdj":   attach_names,
                        "_raw":         raw,
                    })
                except Exception:
                    pass
 
                if i % 200 == 0 or i == total - 1:
                    self.msg_queue.put(("progress", (i+1)/max(total,1),
                                        f"Leyendo… {i+1:,} / {total:,} correos"))
 
            self.msg_queue.put(("done", emails, total))
        except Exception as exc:
            self.msg_queue.put(("error", str(exc)))
 
    # ── Filtrado ─────────────────────────────────────────────────────────────
 
    def _apply_filters(self):
        if not self._all_emails:
            messagebox.showinfo("Info", "Primero carga un archivo .mbox.")
            return
 
        from_filter   = self.filter_from.get().strip().lower()
        keyword       = self.filter_keyword.get().strip().lower()
        date_from_str = self.filter_date_from.get().strip()
        date_to_str   = self.filter_date_to.get().strip()
 
        date_from = date_to = None
        try:
            if date_from_str:
                date_from = datetime.strptime(date_from_str, "%Y-%m-%d")
            if date_to_str:
                date_to = datetime.strptime(date_to_str, "%Y-%m-%d").replace(
                    hour=23, minute=59, second=59)
        except ValueError:
            messagebox.showerror("Error", "Formato de fecha inválido. Usa YYYY-MM-DD.")
            return
 
        filtered = []
        for mail in self._all_emails:
            if from_filter and from_filter not in mail["De"].lower():
                continue
            if keyword and keyword not in (mail["Asunto"] + " " + mail["Cuerpo"]).lower():
                continue
            if date_from and mail["Fecha"] and mail["Fecha"] < date_from:
                continue
            if date_to and mail["Fecha"] and mail["Fecha"] > date_to:
                continue
            filtered.append(mail)
 
        self.results = filtered
        self._refresh_table(filtered)
        self._set_status(f"Filtro aplicado: {len(filtered):,} resultado(s)")
 
    def _clear_filters(self):
        for e in (self.filter_from, self.filter_keyword,
                  self.filter_date_from, self.filter_date_to):
            e.delete(0, "end")
        self.results = list(self._all_emails)
        self._refresh_table(self.results)
        self._set_status(f"Filtros eliminados — {len(self.results):,} correos")
 
    # ── Tabla ────────────────────────────────────────────────────────────────
 
    def _refresh_table(self, data: list[dict]):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for mail in data:
            date_str = mail["Fecha"].strftime("%Y-%m-%d %H:%M") if mail["Fecha"] else ""
            self.tree.insert("", "end", values=(
                date_str,
                mail["De"][:60],
                mail["Para"][:50],
                mail["Asunto"][:100],
                "📎" if mail.get("TieneAdjunto") else "",
            ))
        self.count_lbl.configure(text=f"{len(data):,} correo(s) encontrado(s)")
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.configure(state="disabled")
        self.attach_count_lbl.configure(text="")
        self._selected_idx = -1
        self.attach_info_lbl.configure(text="Selecciona un correo para ver sus adjuntos")
 
    def _sort_table(self, col):
        reverse = getattr(self, f"_sort_{col}_rev", False)
        if col == "Fecha":
            self.results.sort(key=lambda x: x["Fecha"] or datetime.min, reverse=reverse)
        else:
            self.results.sort(key=lambda x: x[col].lower(), reverse=reverse)
        setattr(self, f"_sort_{col}_rev", not reverse)
        self._refresh_table(self.results)
 
    def _on_row_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        if idx >= len(self.results):
            return
 
        self._selected_idx = idx
        mail = self.results[idx]
        date_str = mail["Fecha"].strftime("%A %d %B %Y, %H:%M") if mail["Fecha"] else "Fecha desconocida"
 
        adj_list = mail.get("NombresAdj", [])
        if adj_list:
            adj_section = f"\n📎 Adjuntos ({len(adj_list)}):\n"
            adj_section += "\n".join(f"   • {n}" for n in adj_list)
            self.attach_count_lbl.configure(text=f"📎 {len(adj_list)} adjunto(s)")
            self.attach_info_lbl.configure(
                text=f"{len(adj_list)} adjunto(s) listo(s) — pulsa el botón rosa ↑"
            )
        else:
            adj_section = "\n(Sin archivos adjuntos)"
            self.attach_count_lbl.configure(text="")
            self.attach_info_lbl.configure(text="Este correo no tiene adjuntos")
 
        text = (
            f"De:       {mail['De']}\n"
            f"Para:     {mail['Para']}\n"
            f"Fecha:    {date_str}\n"
            f"Asunto:   {mail['Asunto']}"
            f"{adj_section}\n"
            f"{'─'*60}\n\n"
            f"{mail['Cuerpo']}"
        )
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.insert("end", text)
        self.detail_text.configure(state="disabled")
 
    # ── Descarga de adjuntos ─────────────────────────────────────────────────
 
    def _download_attachments(self):
        if self._selected_idx < 0 or self._selected_idx >= len(self.results):
            messagebox.showinfo("Sin selección",
                                "Primero selecciona un correo de la lista.")
            return
 
        mail = self.results[self._selected_idx]
 
        if not mail.get("TieneAdjunto"):
            messagebox.showinfo("Sin adjuntos",
                                "Este correo no contiene archivos adjuntos.")
            return
 
        try:
            # Usar policy=email.policy.compat32 asegura compatibilidad total
            # con mensajes mbox que pueden tener cabeceras no estándar
            msg_obj = email.message_from_bytes(
                mail["_raw"], policy=email.policy.compat32
            )
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo leer el mensaje:\n{exc}")
            return
 
        attachments = get_attachments(msg_obj)
        if not attachments:
            messagebox.showinfo("Sin adjuntos",
                                "Este correo no contiene archivos adjuntos.")
            return
 
        dest_folder_base = filedialog.askdirectory(
            title="Selecciona la carpeta donde guardar los adjuntos",
            initialdir=get_desktop_path()
        )
        if not dest_folder_base:
            return  # El usuario canceló
        dest_folder = pathlib.Path(dest_folder_base) / "Adjuntos_Extraidos"
        try:
            dest_folder.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo crear la carpeta destino:\n{exc}")
            return
 
        saved  = []
        errors = []
        for att in attachments:
            dest_path = dest_folder / att["filename"]
            if dest_path.exists():
                stem, suffix, counter = dest_path.stem, dest_path.suffix, 1
                while dest_path.exists():
                    dest_path = dest_folder / f"{stem}_{counter}{suffix}"
                    counter += 1
            try:
                dest_path.write_bytes(att["data"])
                saved.append(dest_path.name)
            except Exception as exc:
                errors.append(f"{att['filename']}: {exc}")
 
        if saved:
            resumen = (f"✅  {len(saved)} archivo(s) guardado(s) en:\n{dest_folder}\n\n"
                       + "\n".join(f"  • {n}" for n in saved))
            if errors:
                resumen += f"\n\n⚠️  {len(errors)} error(es):\n" + "\n".join(errors)
            messagebox.showinfo("Adjuntos descargados", resumen)
            self._set_status(f"✅ {len(saved)} adjunto(s) en Escritorio/Adjuntos_Extraidos")
        else:
            messagebox.showerror("Error al guardar",
                                 "No se pudo guardar ningún adjunto.\n\n" + "\n".join(errors))
 
    # ── Exportación ──────────────────────────────────────────────────────────
 
    def _export(self, fmt: str):
        if not self.results:
            messagebox.showinfo("Info", "No hay correos para exportar.")
            return
 
        ext   = "xlsx" if fmt == "xlsx" else "csv"
        ftype = [("Excel", "*.xlsx")] if fmt == "xlsx" else [("CSV", "*.csv")]
        out   = filedialog.asksaveasfilename(
            defaultextension=f".{ext}", filetypes=ftype, title="Guardar como…")
        if not out:
            return
 
        df = pd.DataFrame([{
            "Fecha":    r["Fecha"].strftime("%Y-%m-%d %H:%M") if r["Fecha"] else "",
            "De":       r["De"],
            "Para":     r["Para"],
            "Asunto":   r["Asunto"],
            "Cuerpo":   r["Cuerpo"],
            "Adjuntos": ", ".join(r.get("NombresAdj", [])),
        } for r in self.results])
 
        try:
            if fmt == "xlsx":
                with pd.ExcelWriter(out, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Correos")
                    ws = writer.sheets["Correos"]
                    for col_cells in ws.columns:
                        max_len = max(len(str(c.value or "")) for c in col_cells)
                        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len+4, 60)
            else:
                df.to_csv(out, index=False, encoding="utf-8-sig")
            messagebox.showinfo("Exportación completa",
                                f"✅ {len(self.results):,} correos exportados a:\n{out}")
        except Exception as exc:
            messagebox.showerror("Error al exportar", str(exc))
 
    # ── Cola de mensajes ──────────────────────────────────────────────────────
 
    def _poll_queue(self):
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                if msg[0] == "progress":
                    _, pct, label = msg
                    self.progress.set(pct)
                    self._set_status(label)
                elif msg[0] == "done":
                    _, emails, total = msg
                    self._all_emails = emails
                    self.results     = list(emails)
                    self.loading     = False
                    self.load_btn.configure(state="normal",
                                            text="⚡  Cargar todos los correos")
                    self.progress.set(1)
                    self._refresh_table(emails)
                    self._set_status(f"✅ {total:,} correos cargados correctamente")
                elif msg[0] == "error":
                    _, err = msg
                    self.loading = False
                    self.load_btn.configure(state="normal",
                                            text="⚡  Cargar todos los correos")
                    messagebox.showerror("Error al leer el archivo", err)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)
 
    def _set_status(self, text: str):
        self.status_lbl.configure(text=text)
 
 
# ─── Punto de entrada ─────────────────────────────────────────────────────────
 
if __name__ == "__main__":
    app = MBoxManagerApp()
    app.mainloop()