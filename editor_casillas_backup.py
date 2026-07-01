# ============================================================
#  EDITOR DE CASILLAS (CANDADOS) — Código de respaldo
#  Eliminado de metatag_v8.py el 2026-06-30
#  Si necesitas restaurarlo, copia estas funciones a la clase
#  principal y agrega el botón en _build_control_panel()
# ============================================================

# ── Variables de instancia (agregar en __init__) ──
# self.editor_window  = None
# self.editor_entries = {}

# ── Botón (agregar en _build_control_panel, sección HERRAMIENTAS AVANZADAS) ──
# self._btn(parent, "✏️ Editor de Casillas (Candados)", self._open_editor, primary=False)

# ── Llamada en _navigate_to_row (después de _on_row_click) ──
# if self.editor_window and self.editor_window.winfo_exists():
#     self._populate_editor()

# ── Funciones ──

def _open_editor(self):
    if self.df is None:
        return messagebox.showinfo("Aviso", "Carga un archivo de datos primero.")
    if self.editor_window and self.editor_window.winfo_exists():
        self.editor_window.lift(); return

    self.editor_window = tk.Toplevel(self)
    self.editor_window.title("✏️ Editor y Autocompletado")
    self.editor_window.geometry(f"{int(450*self.current_scale)}x{int(650*self.current_scale)}")
    self.editor_window.configure(bg=C["surface"])

    hdr = tk.Frame(self.editor_window, bg=C["header_bg"])
    hdr.pack(fill="x")
    tk.Label(hdr, text="Valores de la Fila Actual", font=FONTS["H2"],
             bg=C["header_bg"], fg=C["header_fg"]).pack(pady=10)
    tk.Label(self.editor_window,
             text="Cierra el candado 🔒 para que el dato SIEMPRE se escriba y se copie.",
             bg=C["surface"], fg=C["text3"], font=FONTS["TINY"]).pack(pady=5)

    canvas = tk.Canvas(self.editor_window, bg=C["surface"], highlightthickness=0)
    vsb    = ttk.Scrollbar(self.editor_window, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas, bg=C["surface"])
    scroll_frame.bind("<Configure>",
                      lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw",
                         width=int(420*self.current_scale))
    canvas.configure(yscrollcommand=vsb.set)
    canvas.pack(side="left", fill="both", expand=True, padx=10, pady=5)
    vsb.pack(side="right", fill="y")

    def _on_editor_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    def _on_editor_mousewheel_linux(event):
        if event.num == 4: canvas.yview_scroll(-1, "units")
        elif event.num == 5: canvas.yview_scroll(1, "units")
    def _bind_editor_scroll(widget):
        widget.bind("<MouseWheel>", _on_editor_mousewheel)
        widget.bind("<Button-4>", _on_editor_mousewheel_linux)
        widget.bind("<Button-5>", _on_editor_mousewheel_linux)
        for child in widget.winfo_children():
            _bind_editor_scroll(child)

    self.editor_entries.clear()
    img_col = self.img_col_var.get()

    for col in self.df.columns:
        if col == img_col: continue
        row_frame = tk.Frame(scroll_frame, bg=C["surface"])
        row_frame.pack(fill="x", pady=4)
        is_locked = col in self.locked_columns
        btn_lock  = tk.Button(row_frame,
                              text="🔒" if is_locked else "🔓",
                              font=FONTS["BODY"], width=3, relief="flat",
                              bg=C["sel_bg"] if is_locked else C["btn_ghost_bg"],
                              fg=C["text"])
        btn_lock.pack(side="left", padx=5)

        def toggle_lock(c=col, btn=btn_lock):
            if c in self.locked_columns:
                self.locked_columns.remove(c)
                btn.configure(text="🔓", bg=C["btn_ghost_bg"])
            else:
                self.locked_columns.add(c)
                btn.configure(text="🔒", bg=C["sel_bg"])
            self._update_meta_preview()
        btn_lock.configure(command=toggle_lock)

        tk.Label(row_frame, text=col[:15], width=15, anchor="e",
                 bg=C["surface"], fg=C["text"], font=FONTS["LABEL_B"]).pack(side="left", padx=5)
        entry_var = tk.StringVar()
        entry = tk.Entry(row_frame, textvariable=entry_var,
                         bg=C["bg"], fg=C["text"], relief="solid", bd=1, font=FONTS["BODY"])
        entry.pack(side="left", fill="x", expand=True)

        def update_df(*args, c=col, v=entry_var):
            if self.current_row is not None and self.df is not None:
                self.df.at[self.current_row, c] = v.get()
                if self.original_df is not None:
                    idx_real = self.df.index[self.current_row]
                    self.original_df.at[idx_real, c] = v.get()
                self.grid.redraw()
                self._update_meta_preview()
        entry_var.trace_add("write", update_df)
        self.editor_entries[col] = entry_var

    _bind_editor_scroll(scroll_frame)
    self._populate_editor()

def _populate_editor(self):
    if not self.editor_window or not self.editor_window.winfo_exists(): return
    if self.current_row is None or self.df is None: return
    for col, var in self.editor_entries.items():
        if col in self.df.columns:
            val = str(self.df.at[self.current_row, col])
            if val == "nan": val = ""
            var.set(val)
