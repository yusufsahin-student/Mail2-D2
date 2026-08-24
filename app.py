from __future__ import annotations

import json
import math
import os
import queue
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from mailbot_core import (
    MailBotError,
    RecipientPreview,
    SMTPConfig,
    SMTPMailer,
    create_previews,
    list_sheet_names,
    load_records,
    prepare_log_file,
    random_delay_seconds,
    smtp_error_text,
    write_log_row,
)


APP_TITLE = "FRC Sponsorluk Mail Botu"
SETTINGS_PATH = (
    Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    / "FRC_Sponsor_Mail_Botu"
    / "settings.json"
)


class MailBotApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1120x760")
        self.minsize(940, 650)

        self.previews: list[RecipientPreview] = []
        self.attachments: list[str] = []
        self.worker: threading.Thread | None = None
        self.stop_event = threading.Event()
        self.events: queue.Queue[tuple] = queue.Queue()
        self.current_log_path: Path | None = None

        self._create_variables()
        self._configure_style()
        self._build_ui()
        self._load_settings()
        self.after(100, self._poll_events)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _create_variables(self) -> None:
        self.data_file = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.email_column = tk.StringVar(value="email")
        self.subject_template = tk.StringVar(value="FRC takımımız için sponsorluk görüşmesi")
        self.smtp_host = tk.StringVar(value="smtp.gmail.com")
        self.smtp_port = tk.StringVar(value="587")
        self.smtp_security = tk.StringVar(value="STARTTLS")
        self.smtp_username = tk.StringVar()
        self.smtp_password = tk.StringVar()
        self.sender_address = tk.StringVar()
        self.sender_name = tk.StringVar()
        self.reply_to = tk.StringVar()
        self.test_address = tk.StringVar()
        self.delay_min = tk.StringVar(value="30")
        self.delay_max = tk.StringVar(value="60")
        self.skip_duplicates = tk.BooleanVar(value=True)
        self.status_text = tk.StringVar(value="Hazır")
        self.counter_text = tk.StringVar(value="0 hazır / 0 toplam")

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("Hint.TLabel", foreground="#555555")
        style.configure("Danger.TLabel", foreground="#9b1c1c")

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=14)
        root.pack(fill="both", expand=True)
        ttk.Label(root, text=APP_TITLE, style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            root,
            text="Excel satırlarını kişiselleştirilmiş sponsorluk e-postalarına dönüştürür.",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(2, 10))

        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True)
        self.content_tab = ttk.Frame(notebook, padding=12)
        self.smtp_tab = ttk.Frame(notebook, padding=12)
        self.preview_tab = ttk.Frame(notebook, padding=12)
        notebook.add(self.content_tab, text="1. Veri ve içerik")
        notebook.add(self.smtp_tab, text="2. SMTP ayarları")
        notebook.add(self.preview_tab, text="3. Önizleme ve gönderim")

        self._build_content_tab()
        self._build_smtp_tab()
        self._build_preview_tab()

        status_bar = ttk.Frame(root)
        status_bar.pack(fill="x", pady=(9, 0))
        ttk.Label(status_bar, textvariable=self.status_text).pack(side="left")
        ttk.Label(status_bar, textvariable=self.counter_text).pack(side="right")

    def _build_content_tab(self) -> None:
        tab = self.content_tab
        tab.columnconfigure(1, weight=1)
        tab.rowconfigure(5, weight=1)

        ttk.Label(tab, text="Excel / CSV dosyası").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        ttk.Entry(tab, textvariable=self.data_file).grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Button(tab, text="Seç…", command=self._choose_data_file).grid(row=0, column=2, padx=(8, 0), pady=5)

        ttk.Label(tab, text="Çalışma sayfası").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        self.sheet_combo = ttk.Combobox(tab, textvariable=self.sheet_name, state="readonly", width=28)
        self.sheet_combo.grid(row=1, column=1, sticky="w", pady=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _event: self._load_headers())

        ttk.Label(tab, text="E-posta sütunu").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=5)
        self.email_combo = ttk.Combobox(tab, textvariable=self.email_column, state="readonly", width=28)
        self.email_combo.grid(row=2, column=1, sticky="w", pady=5)

        ttk.Label(tab, text="Konu şablonu").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=5)
        ttk.Entry(tab, textvariable=self.subject_template).grid(row=3, column=1, columnspan=2, sticky="ew", pady=5)

        template_header = ttk.Frame(tab)
        template_header.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(8, 4))
        ttk.Label(template_header, text="E-posta metni").pack(side="left")
        ttk.Button(template_header, text="TXT şablonu yükle…", command=self._load_template_file).pack(side="right")

        self.body_text = scrolledtext.ScrolledText(tab, wrap="word", undo=True, font=("Segoe UI", 10))
        self.body_text.grid(row=5, column=0, columnspan=3, sticky="nsew")
        self.body_text.insert(
            "1.0",
            "Merhaba (isim) (soyisim),\n\n"
            "(şirket) ile FRC takımımız arasında olası bir sponsorluk görüşmesi yapmak istiyoruz. "
            "Takımımızın misyonu: (misyon)\n\n"
            "Uygun olduğunuzda görüşmekten memnuniyet duyarız.\n\n"
            "Saygılarımızla,\nFRC Takımı",
        )

        attachment_frame = ttk.Frame(tab)
        attachment_frame.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(8, 2))
        ttk.Label(attachment_frame, text="Ortak ekler:").pack(side="left")
        self.attachment_label = ttk.Label(attachment_frame, text="Ek yok", style="Hint.TLabel")
        self.attachment_label.pack(side="left", padx=8)
        ttk.Button(attachment_frame, text="Ek seç…", command=self._choose_attachments).pack(side="right")
        ttk.Button(attachment_frame, text="Temizle", command=self._clear_attachments).pack(side="right", padx=(0, 6))

        ttk.Label(
            tab,
            text="Şablonda Excel başlıklarını parantezle yazın: (isim), (şirket). Gerçek parantez için ((metin)) kullanın.",
            style="Hint.TLabel",
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(4, 0))

    def _build_smtp_tab(self) -> None:
        tab = self.smtp_tab
        tab.columnconfigure(1, weight=1)
        tab.columnconfigure(3, weight=1)

        fields = [
            ("SMTP sunucusu", self.smtp_host, 0, 0),
            ("Port", self.smtp_port, 0, 2),
            ("Kullanıcı adı", self.smtp_username, 1, 0),
            ("Parola / uygulama parolası", self.smtp_password, 2, 0),
            ("Gönderen e-posta", self.sender_address, 3, 0),
            ("Gönderen adı", self.sender_name, 4, 0),
            ("Yanıt adresi (isteğe bağlı)", self.reply_to, 5, 0),
        ]
        for label, variable, row, column in fields:
            ttk.Label(tab, text=label).grid(row=row, column=column, sticky="w", padx=(0, 8), pady=7)
            show = "•" if variable is self.smtp_password else ""
            span = 3 if column == 0 and row > 0 else 1
            ttk.Entry(tab, textvariable=variable, show=show).grid(
                row=row, column=column + 1, columnspan=span, sticky="ew", pady=7
            )

        ttk.Label(tab, text="Güvenlik").grid(row=1, column=2, sticky="w", padx=(14, 8), pady=7)
        ttk.Combobox(
            tab,
            textvariable=self.smtp_security,
            values=("STARTTLS", "SSL/TLS", "Yok"),
            state="readonly",
            width=16,
        ).grid(row=1, column=3, sticky="ew", pady=7)

        ttk.Separator(tab).grid(row=6, column=0, columnspan=4, sticky="ew", pady=14)
        ttk.Label(
            tab,
            text="Gmail için çoğunlukla smtp.gmail.com / 587 / STARTTLS; Microsoft 365 için smtp.office365.com / 587 / STARTTLS kullanılır.",
            wraplength=850,
        ).grid(row=7, column=0, columnspan=4, sticky="w", pady=5)
        ttk.Label(
            tab,
            text="Hesabınız iki aşamalı doğrulama veya uygulama parolası isteyebilir. Parola kaydedilmez.",
            style="Hint.TLabel",
        ).grid(row=8, column=0, columnspan=4, sticky="w", pady=5)

    def _build_preview_tab(self) -> None:
        tab = self.preview_tab
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)

        top = ttk.Frame(tab)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(top, text="Önizlemeyi oluştur / yenile", command=self._build_preview).pack(side="left")
        ttk.Label(top, text="Rastgele gecikme (sn):").pack(side="left", padx=(18, 5))
        ttk.Entry(top, textvariable=self.delay_min, width=7).pack(side="left")
        ttk.Label(top, text="–").pack(side="left", padx=4)
        ttk.Entry(top, textvariable=self.delay_max, width=7).pack(side="left")
        ttk.Checkbutton(top, text="Tekrarlanan adresleri atla", variable=self.skip_duplicates).pack(side="left", padx=14)

        tree_frame = ttk.Frame(tab)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(
            tree_frame,
            columns=("row", "email", "subject", "status"),
            show="headings",
            selectmode="browse",
        )
        self.tree.heading("row", text="Satır")
        self.tree.heading("email", text="E-posta")
        self.tree.heading("subject", text="Konu")
        self.tree.heading("status", text="Durum")
        self.tree.column("row", width=60, anchor="center", stretch=False)
        self.tree.column("email", width=220)
        self.tree.column("subject", width=270)
        self.tree.column("status", width=390)
        self.tree.tag_configure("invalid", foreground="#9b1c1c")
        self.tree.tag_configure("sent", foreground="#147a37")
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<Double-1>", self._show_selected_preview)

        send_box = ttk.LabelFrame(tab, text="Kontrollü gönderim", padding=10)
        send_box.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        send_box.columnconfigure(1, weight=1)
        ttk.Label(send_box, text="Test alıcısı").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(send_box, textvariable=self.test_address).grid(row=0, column=1, sticky="ew")
        self.test_button = ttk.Button(send_box, text="İlk geçerli satırı test gönder", command=self._send_test)
        self.test_button.grid(row=0, column=2, padx=(8, 0))
        self.send_button = ttk.Button(send_box, text="Tüm geçerli mailleri gönder", command=self._send_all)
        self.send_button.grid(row=1, column=1, sticky="e", pady=(9, 0))
        self.stop_button = ttk.Button(send_box, text="Durdur", command=self._stop_sending, state="disabled")
        self.stop_button.grid(row=1, column=2, padx=(8, 0), pady=(9, 0))
        self.progress = ttk.Progressbar(send_box, mode="determinate")
        self.progress.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(10, 0))

    def _choose_data_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Alıcı veri dosyasını seçin",
            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv"), ("Tüm dosyalar", "*.*")],
        )
        if not path:
            return
        self.data_file.set(path)
        try:
            names = list_sheet_names(path)
            self.sheet_combo["values"] = names
            self.sheet_name.set(names[0])
            self._load_headers()
            self.status_text.set("Veri dosyası yüklendi.")
        except MailBotError as exc:
            messagebox.showerror(APP_TITLE, str(exc))

    def _load_headers(self) -> None:
        path = self.data_file.get().strip()
        if not path:
            return
        try:
            headers, _records = load_records(path, self.sheet_name.get() or None)
            self.email_combo["values"] = headers
            preferred = next(
                (header for header in headers if header.casefold() in {"email", "e-mail", "eposta", "e-posta", "mail"}),
                None,
            )
            if self.email_column.get() not in headers:
                self.email_column.set(preferred or headers[0])
        except MailBotError as exc:
            messagebox.showerror(APP_TITLE, str(exc))

    def _load_template_file(self) -> None:
        path = filedialog.askopenfilename(title="Şablon seçin", filetypes=[("Metin dosyası", "*.txt"), ("Tüm dosyalar", "*.*")])
        if not path:
            return
        try:
            text = Path(path).read_text(encoding="utf-8-sig")
        except UnicodeDecodeError:
            text = Path(path).read_text(encoding="cp1254")
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"Şablon okunamadı: {exc}")
            return
        self.body_text.delete("1.0", "end")
        self.body_text.insert("1.0", text)

    def _choose_attachments(self) -> None:
        paths = filedialog.askopenfilenames(title="Tüm e-postalara eklenecek dosyaları seçin")
        if paths:
            self.attachments = list(paths)
            self._refresh_attachment_label()

    def _clear_attachments(self) -> None:
        self.attachments = []
        self._refresh_attachment_label()

    def _refresh_attachment_label(self) -> None:
        if not self.attachments:
            self.attachment_label.configure(text="Ek yok")
        elif len(self.attachments) == 1:
            self.attachment_label.configure(text=Path(self.attachments[0]).name)
        else:
            self.attachment_label.configure(text=f"{len(self.attachments)} dosya")

    def _get_preview_data(self) -> list[RecipientPreview]:
        path = self.data_file.get().strip()
        if not path:
            raise MailBotError("Önce Excel veya CSV dosyasını seçin.")
        headers, records = load_records(path, self.sheet_name.get() or None)
        if not records:
            raise MailBotError("Veri dosyasında başlık dışında alıcı satırı bulunamadı.")
        if self.email_column.get() not in headers:
            raise MailBotError("E-posta sütununu seçin.")
        return create_previews(
            records,
            self.email_column.get(),
            self.subject_template.get(),
            self.body_text.get("1.0", "end-1c"),
            skip_duplicate_addresses=self.skip_duplicates.get(),
        )

    def _build_preview(self, show_success: bool = True) -> bool:
        try:
            self.previews = self._get_preview_data()
        except MailBotError as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return False
        self._render_tree()
        valid_count = sum(item.valid for item in self.previews)
        invalid_count = len(self.previews) - valid_count
        self.counter_text.set(f"{valid_count} hazır / {len(self.previews)} toplam")
        self.status_text.set(f"Önizleme hazır: {valid_count} gönderilebilir, {invalid_count} atlanacak.")
        if show_success:
            messagebox.showinfo(
                APP_TITLE,
                f"Önizleme tamamlandı.\n\nGönderilebilir: {valid_count}\nAtlanacak: {invalid_count}\n\n"
                "Bir satırın metnini görmek için satıra çift tıklayın.",
            )
        return True

    def _render_tree(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for index, item in enumerate(self.previews):
            tag = "invalid" if not item.valid else ""
            self.tree.insert(
                "",
                "end",
                iid=str(index),
                values=(item.row_number, item.email, item.subject, item.status),
                tags=(tag,) if tag else (),
            )

    def _show_selected_preview(self, _event=None) -> None:
        selected = self.tree.selection()
        if not selected:
            return
        item = self.previews[int(selected[0])]
        window = tk.Toplevel(self)
        window.title(f"Satır {item.row_number} önizlemesi")
        window.geometry("760x560")
        window.transient(self)
        frame = ttk.Frame(window, padding=12)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=f"Alıcı: {item.email}").pack(anchor="w")
        ttk.Label(frame, text=f"Konu: {item.subject}", wraplength=710).pack(anchor="w", pady=(4, 8))
        text = scrolledtext.ScrolledText(frame, wrap="word", font=("Segoe UI", 10))
        text.pack(fill="both", expand=True)
        text.insert("1.0", item.body)
        text.configure(state="disabled")
        ttk.Button(frame, text="Kapat", command=window.destroy).pack(anchor="e", pady=(8, 0))

    def _smtp_config(self) -> SMTPConfig:
        try:
            port = int(self.smtp_port.get().strip())
        except ValueError as exc:
            raise MailBotError("SMTP portu sayı olmalıdır.") from exc
        config = SMTPConfig(
            host=self.smtp_host.get().strip(),
            port=port,
            security=self.smtp_security.get(),
            username=self.smtp_username.get().strip(),
            password=self.smtp_password.get(),
            sender_address=self.sender_address.get().strip(),
            sender_name=self.sender_name.get().strip(),
            reply_to=self.reply_to.get().strip(),
        )
        config.validate()
        return config

    def _send_test(self) -> None:
        if self._is_busy():
            return
        if not self._build_preview(show_success=False):
            return
        item = next((preview for preview in self.previews if preview.valid), None)
        if item is None:
            messagebox.showerror(APP_TITLE, "Test için gönderilebilir bir satır yok.")
            return
        test_address = self.test_address.get().strip()
        try:
            config = self._smtp_config()
            from mailbot_core import is_valid_email
            if not is_valid_email(test_address):
                raise MailBotError("Geçerli bir test alıcısı girin.")
        except MailBotError as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self._set_busy(True)
        self.status_text.set("Test e-postası gönderiliyor…")
        self.worker = threading.Thread(
            target=self._test_worker,
            args=(config, test_address, item, tuple(self.attachments)),
            daemon=True,
        )
        self.worker.start()

    def _test_worker(
        self,
        config: SMTPConfig,
        test_address: str,
        item: RecipientPreview,
        attachments: tuple[str, ...],
    ) -> None:
        try:
            with SMTPMailer(config) as mailer:
                mailer.send(test_address, "[TEST] " + item.subject, item.body, attachments)
            self.events.put(("test_done", test_address))
        except Exception as exc:
            self.events.put(("fatal", "Test gönderilemedi: " + smtp_error_text(exc)))

    def _send_all(self) -> None:
        if self._is_busy():
            return
        if not self._build_preview(show_success=False):
            return
        valid_items = [item for item in self.previews if item.valid]
        if not valid_items:
            messagebox.showerror(APP_TITLE, "Gönderilebilir alıcı bulunamadı.")
            return
        try:
            config = self._smtp_config()
            delay_min, delay_max = self._validated_delay_range()
        except MailBotError as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return

        invalid_count = len(self.previews) - len(valid_items)
        answer = messagebox.askyesno(
            APP_TITLE,
            f"{len(valid_items)} gerçek e-posta gönderilecek.\n"
            f"{invalid_count} sorunlu satır atlanacak.\n\n"
            f"Her gönderim arasında {delay_min:g}–{delay_max:g} saniye rastgele beklenecek.\n\n"
            "Önce test e-postasını kontrol ettiğinizden emin misiniz? Gönderim başlasın mı?",
            icon="warning",
        )
        if not answer:
            return

        self.stop_event.clear()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        data_parent = Path(self.data_file.get()).resolve().parent
        self.current_log_path = data_parent / "mail_bot_logs" / f"gonderim_{timestamp}.csv"
        try:
            prepare_log_file(self.current_log_path)
        except MailBotError as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self.progress.configure(maximum=len(valid_items), value=0)
        self._set_busy(True)
        self.status_text.set("Toplu gönderim başladı…")
        self.worker = threading.Thread(
            target=self._send_worker,
            args=(
                config,
                valid_items,
                delay_min,
                delay_max,
                self.current_log_path,
                tuple(self.attachments),
            ),
            daemon=True,
        )
        self.worker.start()

    def _send_worker(
        self,
        config: SMTPConfig,
        items: list[RecipientPreview],
        delay_min: float,
        delay_max: float,
        log_path: Path,
        attachments: tuple[str, ...],
    ) -> None:
        sent = 0
        failed = 0
        try:
            with SMTPMailer(config) as mailer:
                for position, item in enumerate(items, start=1):
                    if self.stop_event.is_set():
                        break
                    try:
                        mailer.send(item.email, item.subject, item.body, attachments)
                        sent += 1
                        result, detail = "GÖNDERİLDİ", ""
                    except Exception as exc:
                        failed += 1
                        result, detail = "HATA", smtp_error_text(exc)
                    write_log_row(
                        log_path,
                        row_number=item.row_number,
                        email=item.email,
                        subject=item.subject,
                        result=result,
                        detail=detail,
                    )
                    self.events.put(("row", item.row_number, result, detail, position, sent, failed))
                    if position < len(items):
                        delay = random_delay_seconds(delay_min, delay_max)
                        self.events.put(("delay", delay, sent, failed))
                        if self.stop_event.wait(delay):
                            break
            self.events.put(("batch_done", sent, failed, self.stop_event.is_set(), str(log_path)))
        except Exception as exc:
            self.events.put(("fatal", "Gönderim durdu: " + smtp_error_text(exc)))

    def _validated_delay_range(self) -> tuple[float, float]:
        try:
            minimum = float(self.delay_min.get().replace(",", "."))
            maximum = float(self.delay_max.get().replace(",", "."))
        except ValueError as exc:
            raise MailBotError("Rastgele gecikmenin alt ve üst değerleri sayı olmalıdır.") from exc
        if not math.isfinite(minimum) or not math.isfinite(maximum):
            raise MailBotError("Rastgele gecikme değerleri sonlu sayı olmalıdır.")
        if not 0.5 <= minimum <= 3600 or not 0.5 <= maximum <= 3600:
            raise MailBotError("Gecikme değerleri 0,5 ile 3600 saniye arasında olmalıdır.")
        if minimum > maximum:
            raise MailBotError("En düşük gecikme, en yüksek gecikmeden büyük olamaz.")
        return minimum, maximum

    def _stop_sending(self) -> None:
        self.stop_event.set()
        self.status_text.set("Durdurma isteği alındı; devam eden işlem tamamlanıyor…")
        self.stop_button.configure(state="disabled")

    def _poll_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                kind = event[0]
                if kind == "row":
                    _, row_number, result, detail, position, sent, failed = event
                    self.progress.configure(value=position)
                    for index, item in enumerate(self.previews):
                        if item.row_number == row_number:
                            item.status = "Gönderildi" if result == "GÖNDERİLDİ" else "Hata: " + detail
                            tags = ("sent",) if result == "GÖNDERİLDİ" else ("invalid",)
                            self.tree.item(str(index), values=(item.row_number, item.email, item.subject, item.status), tags=tags)
                            break
                    self.status_text.set(f"İşleniyor: {sent} gönderildi, {failed} hata.")
                elif kind == "delay":
                    _, delay, sent, failed = event
                    self.status_text.set(
                        f"Sonraki gönderim için {delay:.1f} saniye bekleniyor… "
                        f"({sent} başarılı, {failed} hata)"
                    )
                elif kind == "test_done":
                    self._set_busy(False)
                    self.status_text.set("Test e-postası gönderildi.")
                    messagebox.showinfo(APP_TITLE, f"Test e-postası {event[1]} adresine gönderildi.")
                elif kind == "batch_done":
                    _, sent, failed, stopped, log_path = event
                    self._set_busy(False)
                    state = "durduruldu" if stopped else "tamamlandı"
                    self.status_text.set(f"Gönderim {state}: {sent} başarılı, {failed} hata.")
                    messagebox.showinfo(
                        APP_TITLE,
                        f"Gönderim {state}.\n\nBaşarılı: {sent}\nHata: {failed}\n\nGünlük: {log_path}",
                    )
                elif kind == "fatal":
                    self._set_busy(False)
                    self.status_text.set("Gönderim hatası.")
                    messagebox.showerror(APP_TITLE, event[1])
        except queue.Empty:
            pass
        self.after(100, self._poll_events)

    def _set_busy(self, busy: bool) -> None:
        self.test_button.configure(state="disabled" if busy else "normal")
        self.send_button.configure(state="disabled" if busy else "normal")
        self.stop_button.configure(state="normal" if busy else "disabled")
        if not busy:
            self.worker = None

    def _is_busy(self) -> bool:
        return self.worker is not None and self.worker.is_alive()

    def _load_settings(self) -> None:
        try:
            data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return
        mapping = {
            "smtp_host": self.smtp_host,
            "smtp_port": self.smtp_port,
            "smtp_security": self.smtp_security,
            "smtp_username": self.smtp_username,
            "sender_address": self.sender_address,
            "sender_name": self.sender_name,
            "reply_to": self.reply_to,
            "delay_min": self.delay_min,
            "delay_max": self.delay_max,
        }
        for key, variable in mapping.items():
            if key in data:
                variable.set(data[key])
        if "delay_min" not in data and "send_interval" in data:
            self.delay_min.set(data["send_interval"])
        if "delay_max" not in data and "send_interval" in data:
            self.delay_max.set(data["send_interval"])

    def _save_settings(self) -> None:
        data = {
            "smtp_host": self.smtp_host.get(),
            "smtp_port": self.smtp_port.get(),
            "smtp_security": self.smtp_security.get(),
            "smtp_username": self.smtp_username.get(),
            "sender_address": self.sender_address.get(),
            "sender_name": self.sender_name.get(),
            "reply_to": self.reply_to.get(),
            "delay_min": self.delay_min.get(),
            "delay_max": self.delay_max.get(),
        }
        try:
            SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
            SETTINGS_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass

    def _on_close(self) -> None:
        if self._is_busy():
            if not messagebox.askyesno(APP_TITLE, "Gönderim sürüyor. Durdurup uygulamayı kapatmak istiyor musunuz?"):
                return
            self.stop_event.set()
        self._save_settings()
        self.destroy()


if __name__ == "__main__":
    MailBotApp().mainloop()
