import os
import re
import imaplib
import email
from email.header import decode_header
import threading
from queue import Queue
from datetime import datetime, timedelta
from tkinter import filedialog, END

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
from ttkbootstrap.widgets.scrolled import ScrolledText
import fitz  # PyMuPDF para máxima velocidad en la unificación

# --- CONFIGURACIÓN DE ZOHO MAIL ---
IMAP_SERVER = "imap.zoho.com"
IMAP_PORT = 993


class ZohoPDFDownloaderApp(tb.Window):
    def __init__(self):
        super().__init__(
            title="Automatizador Ultrarrápido de PDFs / Zoho Mail",
            themename="superhero",
            size=(850, 830),
            resizable=(True, True)
        )

        self.mail = None
        self.log_queue = Queue()

        self._build_ui()
        self._bind_clipboard_fixes()
        self.after(100, self._process_queue)

    def _build_ui(self):
        # Header
        header_frame = tb.Frame(self, padding=15)
        header_frame.pack(fill=X)

        tb.Label(
            header_frame,
            text="Automatización de Correo Zoho & PDF (v4.0 C++ Engine)",
            font=("Helvetica", 16, "bold"),
            bootstyle="primary"
        ).pack(anchor=W)
        tb.Label(
            header_frame,
            text="Descarga acelerada y unificación en milisegundos con PyMuPDF",
            font=("Helvetica", 10),
            bootstyle="secondary"
        ).pack(anchor=W)

        # Credenciales Zoho
        cred_frame = tb.Labelframe(self, text=" Credenciales Zoho Mail Empresarial ", padding=15)
        cred_frame.pack(fill=X, padx=15, pady=5)

        tb.Label(cred_frame, text="Servidor IMAP Zoho:").grid(row=0, column=0, sticky=W, pady=3)
        self.cbo_server = tb.Combobox(
            cred_frame,
            state="readonly",
            values=["imap.zoho.com", "imap.zoho.eu", "imap.zoho.in"]
        )
        self.cbo_server.current(0)
        self.cbo_server.grid(row=0, column=1, padx=5, pady=3, sticky=EW)

        tb.Label(cred_frame, text="Correo Empresarial:").grid(row=1, column=0, sticky=W, pady=3)
        self.ent_email = tb.Entry(cred_frame, width=32)
        self.ent_email.grid(row=1, column=1, padx=5, pady=3, sticky=EW)

        tb.Label(cred_frame, text="Contraseña de App (Zoho):").grid(row=2, column=0, sticky=W, pady=3)
        self.ent_pass = tb.Entry(cred_frame, width=32, show="*")
        self.ent_pass.grid(row=2, column=1, padx=5, pady=3, sticky=EW)

        self.btn_connect = tb.Button(
            cred_frame,
            text="1. Conectar y Cargar Carpetas",
            bootstyle="info",
            command=self.start_connect_thread
        )
        self.btn_connect.grid(row=0, column=2, rowspan=3, padx=10, pady=3, sticky=NSEW)
        self.btn_demo = tb.Button(
            cred_frame,
            text="Vista demo",
            bootstyle="warning-outline",
            command=self._load_demo_data
        )
        self.btn_demo.grid(row=0, column=3, rowspan=3, padx=5, pady=3, sticky=NSEW)
        cred_frame.columnconfigure(1, weight=1)

        # Ubicación y Filtros
        process_frame = tb.Labelframe(self, text=" Ubicación y Filtros de Búsqueda ", padding=15)
        process_frame.pack(fill=X, padx=15, pady=5)

        tb.Label(process_frame, text="Carpeta / Subcarpeta:").grid(row=0, column=0, sticky=W, pady=5)
        self.cbo_folders = tb.Combobox(process_frame, state="readonly", width=45)
        self.cbo_folders.grid(row=0, column=1, padx=5, pady=5, sticky=EW)

        tb.Label(process_frame, text="Estado de lectura:").grid(row=1, column=0, sticky=W, pady=5)
        self.cbo_unread_filter = tb.Combobox(
            process_frame,
            state="readonly",
            values=["Solo no leídos (UNSEEN)", "Todos los correos (ALL)"]
        )
        self.cbo_unread_filter.current(0)
        self.cbo_unread_filter.grid(row=1, column=1, padx=5, pady=5, sticky=EW)

        # Rango de Fechas (Inicio y Fin)
        tb.Label(process_frame, text="Fecha Inicio:").grid(row=2, column=0, sticky=W, pady=5)
        self.dt_start = tb.DateEntry(process_frame, bootstyle="info", dateformat="%Y-%m-%d")
        self.dt_start.grid(row=2, column=1, padx=5, pady=5, sticky=EW)

        tb.Label(process_frame, text="Fecha Fin:").grid(row=3, column=0, sticky=W, pady=5)
        self.dt_end = tb.DateEntry(process_frame, bootstyle="info", dateformat="%Y-%m-%d")
        self.dt_end.grid(row=3, column=1, padx=5, pady=5, sticky=EW)

        self.chk_range_var = tb.BooleanVar(value=False)
        self.chk_range = tb.Checkbutton(
            process_frame,
            text="Aplicar rango de fechas",
            variable=self.chk_range_var,
            bootstyle="primary"
        )
        self.chk_range.grid(row=2, column=2, rowspan=2, padx=5, pady=5, sticky=W)

        tb.Label(process_frame, text="Carpeta destino local:").grid(row=4, column=0, sticky=W, pady=5)
        self.ent_output = tb.Entry(process_frame, width=35)
        self.ent_output.insert(0, os.path.abspath("Descargas_PDF"))
        self.ent_output.grid(row=4, column=1, padx=5, pady=5, sticky=EW)

        self.btn_browse = tb.Button(
            process_frame,
            text="Examinar",
            bootstyle="secondary-outline",
            command=self._browse_folder
        )
        self.btn_browse.grid(row=4, column=2, padx=5, pady=5)
        process_frame.columnconfigure(1, weight=1)

        # Acciones de Procesamiento
        self.btn_run = tb.Button(
            self,
            text="2. Descargar y Unificar PDFs",
            bootstyle="success",
            state=DISABLED,
            command=self.start_process_thread
        )
        self.btn_run.pack(fill=X, padx=15, pady=10)

        # Barra de progreso
        self.progress = tb.Progressbar(self, mode="indeterminate", bootstyle="success-striped")
        self.progress.pack(fill=X, padx=15, pady=5)

        # Consola de Estado / Logs
        log_frame = tb.Labelframe(self, text=" Consola de Estado ", padding=10)
        log_frame.pack(fill=BOTH, expand=True, padx=15, pady=10)

        self.txt_log = ScrolledText(log_frame, height=10, font=("Consolas", 9))
        self.txt_log.pack(fill=BOTH, expand=True)

    def _bind_clipboard_fixes(self):
        """Corrección de atajos de teclado para evitar congelamientos al pegar o cambiar de foco."""
        entries = [self.ent_email, self.ent_pass, self.ent_output]
        for entry in entries:
            entry.bind("<Control-v>", lambda e, widget=entry: self._safe_paste(widget))
            entry.bind("<Control-V>", lambda e, widget=entry: self._safe_paste(widget))
            entry.bind("<Shift-Insert>", lambda e, widget=entry: self._safe_paste(widget))

    def _safe_paste(self, widget):
        try:
            clipboard_text = self.clipboard_get()
            widget.insert(tb.INSERT, clipboard_text)
        except Exception:
            pass
        return "break"

    def log(self, message: str):
        self.log_queue.put(("LOG", message))

    def show_msg_box(self, title, msg, is_error=True):
        self.log_queue.put(("DIALOG", (title, msg, is_error)))

    def _process_queue(self):
        max_items = 20
        count = 0
        while not self.log_queue.empty() and count < max_items:
            msg_type, content = self.log_queue.get_nowait()
            if msg_type == "LOG":
                self.txt_log.insert(END, f"[{datetime.now().strftime('%H:%M:%S')}] {content}\n")
                self.txt_log.see(END)
            elif msg_type == "DIALOG":
                title, msg, is_error = content
                if is_error:
                    Messagebox.show_error(msg, title)
                else:
                    Messagebox.show_info(msg, title)
            count += 1

        self.after(100, self._process_queue)

    def _browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.ent_output.delete(0, END)
            self.ent_output.insert(0, folder)

    def _load_demo_data(self):
        """Carga datos ficticios para demostrar la interfaz sin conectarse a Zoho."""
        self.ent_email.delete(0, END)
        self.ent_email.insert(0, "demo.usuario@ejemplo.com")
        self.ent_pass.delete(0, END)
        self.ent_pass.insert(0, "••••••••")
        folders = ["INBOX", "INBOX/Facturas", "INBOX/Soportes", "Archivados/2026"]
        self.cbo_folders["values"] = folders
        self.cbo_folders.current(0)
        self.btn_run.config(state=NORMAL)
        demo_messages = [
            "MODO DEMOSTRACIÓN: no se realizó conexión con Zoho.",
            "Cuenta simulada: demo.usuario@ejemplo.com",
            "Se encontraron 4 carpetas/subcarpetas (datos ficticios).",
            "INBOX · 3 mensajes · 2 adjuntos PDF",
            "INBOX/Facturas · 5 mensajes · 5 adjuntos PDF",
            "INBOX/Soportes · 2 mensajes · 1 adjunto PDF",
            "Archivados/2026 · 8 mensajes · 4 adjuntos PDF",
        ]
        self.txt_log.delete("1.0", END)
        for message in demo_messages:
            self.txt_log.insert(END, f"[DEMO] {message}\n")
        self.txt_log.see(END)

    def decode_str(self, header_value):
        if not header_value:
            return ""
        decoded_list = decode_header(header_value)
        header_text = ""
        for text, encoding in decoded_list:
            if isinstance(text, bytes):
                header_text += text.decode(encoding if encoding else "utf-8", errors="ignore")
            else:
                header_text += text
        return header_text

    def _parse_folder_name(self, raw_folder_string):
        decoded = raw_folder_string.decode("utf-8", errors="ignore")
        match = re.search(r'\"([^\"]+)\"\s*$', decoded)
        if match:
            return match.group(1)
        
        parts = decoded.split(' "/" ')
        if len(parts) > 1:
            return parts[-1].strip('"')
        
        return decoded.split()[-1].strip('"')

    # --- CONTROL DE HILOS ---

    def start_connect_thread(self):
        server_val = self.cbo_server.get().strip()
        user_val = self.ent_email.get().strip()
        pass_val = self.ent_pass.get().strip()

        if not user_val or not pass_val:
            Messagebox.show_error("Por favor ingrese usuario y contraseña de aplicación.", "Campos vacíos")
            return

        self.btn_connect.config(state=DISABLED)
        self.progress.start()
        threading.Thread(
            target=self._connect_and_list_folders,
            args=(server_val, user_val, pass_val),
            daemon=True
        ).start()

    def _connect_and_list_folders(self, server, user, password):
        try:
            self.log(f"Conectando a servidor IMAP de Zoho ({server}:{IMAP_PORT})...")
            self.mail = imaplib.IMAP4_SSL(server, IMAP_PORT)
            self.mail.login(user, password)
            self.log("¡Autenticación en Zoho Mail exitosa!")

            self.log("Obteniendo listado de carpetas y subcarpetas...")
            status, folder_list = self.mail.list()

            folders = []
            if status == "OK":
                for f in folder_list:
                    folder_name = self._parse_folder_name(f)
                    if folder_name and folder_name not in folders:
                        folders.append(folder_name)

            folders = sorted(folders, key=lambda x: x.lower())

            self.cbo_folders["values"] = folders
            if folders:
                self.cbo_folders.current(0)

            self.log(f"Se encontraron {len(folders)} carpetas/subcarpetas (ordenadas A-Z).")
            self.btn_run.config(state=NORMAL)

        except Exception as e:
            self.log(f"Error al conectar con Zoho: {str(e)}")
            self.show_msg_box(
                "Error de Conexión",
                f"No se pudo conectar a Zoho:\n{str(e)}\n\nVerifique la Contraseña de Aplicación y que IMAP esté activo.",
                is_error=True
            )
        finally:
            self.progress.stop()
            self.btn_connect.config(state=NORMAL)

    def start_process_thread(self):
        selected_folder = self.cbo_folders.get()
        out_folder = self.ent_output.get().strip()

        if not selected_folder:
            Messagebox.show_error("Seleccione una carpeta de correo.", "Atención")
            return

        self.btn_run.config(state=DISABLED)
        self.progress.start()
        threading.Thread(
            target=self._download_and_merge,
            args=(selected_folder, out_folder),
            daemon=True
        ).start()

    def _clean_filename(self, text):
        if not text:
            return "Sin_Asunto"
        clean_text = re.sub(r'[\r\n\t\\/*?:"<>|]', "_", text)
        return clean_text.strip()

    def _build_search_query(self):
        query_parts = []

        unread_opt = self.cbo_unread_filter.get()
        if "UNSEEN" in unread_opt:
            query_parts.append("UNSEEN")
        else:
            query_parts.append("ALL")

        if self.chk_range_var.get():
            try:
                start_str = self.dt_start.entry.get()
                end_str = self.dt_end.entry.get()

                start_date = datetime.strptime(start_str, "%Y-%m-%d")
                end_date = datetime.strptime(end_str, "%Y-%m-%d")

                end_date_inclusive = end_date + timedelta(days=1)

                months_en = {
                    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
                    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
                }

                imap_start = f"{start_date.day:02d}-{months_en[start_date.month]}-{start_date.year}"
                imap_end = f"{end_date_inclusive.day:02d}-{months_en[end_date_inclusive.month]}-{end_date_inclusive.year}"

                query_parts.append(f'SINCE "{imap_start}" BEFORE "{imap_end}"')
            except ValueError:
                self.log("Formato de fecha inválido, omitiendo el filtro por rango de fechas.")

        return " ".join(query_parts)

    def _download_and_merge(self, folder, output_dir):
        try:
            os.makedirs(output_dir, exist_ok=True)
            indiv_dir = os.path.join(output_dir, "Archivos_Individuales")
            os.makedirs(indiv_dir, exist_ok=True)

            self.log(f"Accediendo a la carpeta '{folder}'...")
            status, _ = self.mail.select(f'"{folder}"', readonly=True)
            if status != "OK":
                self.log("No se pudo abrir la carpeta seleccionada.")
                return

            search_query = self._build_search_query()
            self.log(f"Ejecutando búsqueda IMAP en Zoho: {search_query}")

            status, messages = self.mail.search(None, search_query)
            if status != "OK" or not messages[0]:
                self.log("No se encontraron correos que coincidan con los filtros.")
                return

            mail_ids = messages[0].split()
            self.log(f"Total de correos a procesar: {len(mail_ids)}")

            downloaded_pdfs = []

            # DESCARGA
            for index, m_id in enumerate(mail_ids, 1):
                res, msg_data = self.mail.fetch(m_id, "(RFC822)")
                if res != "OK":
                    continue

                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        subject = self.decode_str(msg.get("Subject", "Sin_Asunto"))

                        for part in msg.walk():
                            if part.get_content_maintype() == 'multipart':
                                continue
                            if part.get('Content-Disposition') is None:
                                continue

                            filename = self.decode_str(part.get_filename())
                            if filename and filename.lower().endswith(".pdf"):
                                payload = part.get_payload(decode=True)
                                if payload:
                                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                    clean_subj = self._clean_filename(subject)[:30]
                                    clean_orig = self._clean_filename(filename)

                                    new_filename = f"{index:03d}_{timestamp}_{clean_subj}_{clean_orig}"
                                    file_path = os.path.join(indiv_dir, new_filename)

                                    with open(file_path, "wb") as f:
                                        f.write(payload)

                                    downloaded_pdfs.append(file_path)
                                    self.log(f"Descargado ({index}/{len(mail_ids)}): {new_filename}")

            if not downloaded_pdfs:
                self.log("No se encontraron archivos PDF adjuntos en los mensajes procesados.")
                return

            self.log(f"Se descargaron {len(downloaded_pdfs)} archivos PDF individuales.")

            # --- UNIFICACIÓN OPTIMIZADA CON PyMuPDF (C++) ---
            self.log("Iniciando la fusión ultra-rápida de los archivos PDF...")
            
            merged_pdf = fitz.open() # Instancia principal vacía
            
            for pdf_file in downloaded_pdfs:
                try:
                    with fitz.open(pdf_file) as doc:
                        merged_pdf.insert_pdf(doc)
                except Exception as pdf_err:
                    self.log(f"Omitiendo PDF corrupto ({os.path.basename(pdf_file)}): {str(pdf_err)}")

            combined_filename = f"UNIFICADO_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            combined_path = os.path.join(output_dir, combined_filename)

            merged_pdf.save(combined_path)
            merged_pdf.close()

            self.log("==========================================")
            self.log("¡PROCESO FINALIZADO EXITOSAMENTE!")
            self.log(f"1. PDF Individuales: {indiv_dir}")
            self.log(f"2. PDF Unificado: {combined_path}")
            self.log("==========================================")

            self.show_msg_box(
                "Éxito",
                f"Proceso finalizado correctamente.\n\nPDF Unificado guardado en:\n{combined_path}",
                is_error=False
            )

        except Exception as e:
            self.log(f"Error durante la descarga/unificación: {str(e)}")
            self.show_msg_box("Error", f"Ocurrió un error: {str(e)}", is_error=True)
        finally:
            self.progress.stop()
            self.btn_run.config(state=NORMAL)


if __name__ == "__main__":
    app = ZohoPDFDownloaderApp()
    app.mainloop()