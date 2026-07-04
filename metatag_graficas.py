"""
MetaTag v8.9 — Módulo de gráficas y estadísticas
Extraído de metatag_v8.py para reducir la complejidad del archivo principal.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os, sys
from pathlib import Path

try:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.figure import Figure
    import numpy as np
    MATPLOTLIB_OK = True
except ImportError:
    MATPLOTLIB_OK = False


def _native_file_save(title="Guardar archivo", initialdir="", initialfile="",
                      default_ext=".png", zenity_filters=None, tk_filetypes=None):
    from tkinter import filedialog
    if sys.platform == "linux":
        import subprocess
        start_path = os.path.join(initialdir or os.path.expanduser("~"),
                                   initialfile or "")
        cmd = ["zenity", "--file-selection", "--save", "--confirm-overwrite",
               "--title", title, "--filename", start_path]
        for f in (zenity_filters or []):
            cmd += ["--file-filter", f]
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            if r.returncode == 0 and r.stdout.strip():
                path = r.stdout.strip().split("\n")[0]
                if not os.path.splitext(path)[1]:
                    path += default_ext
                return path
            if r.returncode == 1:
                return None
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass
        cmd = ["kdialog", "--getsavefilename", start_path, "--title", title]
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            if r.returncode == 0 and r.stdout.strip():
                path = r.stdout.strip().split("\n")[0]
                if not os.path.splitext(path)[1]:
                    path += default_ext
                return path
            if r.returncode == 1:
                return None
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass
    return filedialog.asksaveasfilename(
        title=title, initialdir=initialdir, initialfile=initialfile,
        defaultextension=default_ext, filetypes=tk_filetypes or [])


def show_stats(parent, df, img_col_var, current_scale, C, FONTS):
    """
    Abre la ventana de estadísticas / gráficas.

    Parameters
    ----------
    parent : tk.Tk | tk.Toplevel
        Ventana padre (MetaTagApp).
    df : pd.DataFrame
        DataFrame con los datos cargados.
    img_col_var : tk.StringVar
        Variable con el nombre de la columna de imagen.
    current_scale : float
        Factor de escala de la UI.
    C : dict
        Colores del tema activo.
    FONTS : dict
        Fuentes activas.
    """
    if df is None or len(df) == 0:
        return messagebox.showinfo("Sin datos", "Carga un archivo de datos primero.")
    if not MATPLOTLIB_OK:
        return messagebox.showerror("Falta Librería", "pip install matplotlib")

    plt.rcParams['lines.antialiased']  = True
    plt.rcParams['patch.antialiased']  = True
    plt.rcParams['path.simplify']      = True

    S_BG           = C["bg"]
    S_CARD         = C["surface"]
    S_BORDER       = C["border"]
    S_ACCENT       = C["accent"]
    S_ACCENT_LIGHT = C.get("accent_light", C["accent"])
    S_TEXT         = C["text"]
    S_TEXT_MUTE    = C["text2"]
    S_CHART_COLORS = C.get("chart_colors",
                            ["#A67C52","#D4A574","#7A4F2D","#3D1F0A","#B1A28F","#5C3518"])

    win = tk.Toplevel(parent)
    win.title("Análisis Cuantitativo Arqueológico (v8.9)")
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    ww = min(int(1200 * current_scale), sw - 40)
    wh = min(int(720  * current_scale), sh - 60)
    win.geometry(f"{ww}x{wh}")
    win.minsize(800, 500)
    win.configure(bg=S_BG)

    img_c  = img_col_var.get()
    validas = [c for c in df.columns if 1 < df[c].nunique() < 40 and c != img_c]
    if not validas:
        return messagebox.showinfo("Sin datos", "No hay categorías repetitivas para graficar.")

    top_frame = tk.Frame(win, bg=S_BG, pady=15, padx=20)
    top_frame.pack(fill="x")

    def make_selector(parent, label_text, icon, options, default, on_change):
        frame = tk.Frame(parent, bg=S_BG)
        tk.Label(frame, text=f"{icon} {label_text}", bg=S_BG, fg=S_TEXT_MUTE,
                 font=FONTS["TINY"]).pack(anchor="w")

        current = tk.StringVar(value=default)

        btn_frame = tk.Frame(frame, bg=S_BORDER, padx=1, pady=1)
        btn_frame.pack(fill="x", pady=(4, 0))

        inner = tk.Frame(btn_frame, bg=C["surface"], cursor="hand2")
        inner.pack(fill="x")

        display = tk.Label(inner, text=default, bg=C["surface"], fg=S_TEXT,
                           font=FONTS["BODY"], anchor="w", padx=12, pady=7)
        display.pack(side="left", fill="x", expand=True)

        arrow = tk.Label(inner, text="▾", bg=C["surface"], fg=C["accent"],
                         font=FONTS["LABEL_B"], padx=8)
        arrow.pack(side="right")

        popup_ref = [None]

        def toggle(e=None):
            if popup_ref[0] and popup_ref[0].winfo_exists():
                popup_ref[0].destroy()
                popup_ref[0] = None
                return

            m = tk.Toplevel(win)
            popup_ref[0] = m
            m.withdraw()
            m.overrideredirect(True)
            win.update_idletasks()
            x = btn_frame.winfo_rootx()
            y = btn_frame.winfo_rooty() + btn_frame.winfo_height() + 2
            min_w  = btn_frame.winfo_width()
            char_w = max((len(opt) for opt in options), default=10)
            popup_w = max(min_w, char_w * 9 + 32)
            popup_h = len(options) * 34 + 8
            screen_w = win.winfo_screenwidth()
            if x + popup_w > screen_w - 10:
                x = screen_w - popup_w - 10
            m.geometry(f"{popup_w}x{popup_h}+{x}+{y}")
            m.configure(bg=S_BORDER)

            container = tk.Frame(m, bg=C["surface"])
            container.pack(fill="both", expand=True, padx=1, pady=1)

            for opt in options:
                is_sel = (opt == current.get())
                bg = S_ACCENT if is_sel else C["surface"]
                fg = S_BG if is_sel else S_TEXT

                row = tk.Frame(container, bg=bg, cursor="hand2")
                row.pack(fill="x", padx=2, pady=1)

                lbl = tk.Label(row, text=f"  {opt}", bg=bg, fg=fg,
                               font=FONTS["BODY"], anchor="w", padx=8, pady=4)
                lbl.pack(fill="x")

                def select(val=opt, win_ref=m):
                    current.set(val)
                    display.configure(text=val)
                    on_change(val)
                    win_ref.destroy()
                    popup_ref[0] = None

                for w in (row, lbl):
                    w.bind("<Button-1>", lambda e, _s=select: _s())
                    w.bind("<Enter>", lambda e, r=row, l=lbl, v=opt:
                           (r.configure(bg=S_ACCENT_LIGHT), l.configure(bg=S_ACCENT_LIGHT, fg=S_TEXT)))
                    w.bind("<Leave>", lambda e, r=row, l=lbl, v=opt:
                           (r.configure(bg=S_ACCENT if v == current.get() else C["surface"]),
                            l.configure(bg=S_ACCENT if v == current.get() else C["surface"],
                                        fg=S_BG if v == current.get() else S_TEXT)))

            m.deiconify()

        def _close_popup(e=None):
            if popup_ref[0] and popup_ref[0].winfo_exists():
                popup_ref[0].destroy()
                popup_ref[0] = None

        win.bind("<FocusOut>",  _close_popup, add="+")
        win.bind("<Unmap>",     _close_popup, add="+")
        win.bind("<Configure>", _close_popup, add="+")

        for w in (inner, display, arrow):
            w.bind("<Button-1>", toggle)

        return frame, current

    def on_var_change(val):
        update_chart()

    def on_style_change(val):
        update_chart()

    CHART_OPTIONS = [
        "Dona HD",
        "Pastel Profesional",
        "Barras Horizontal",
        "Barras Vertical",
        "Lollipop",
    ]

    selector_var_frame, combo_var = make_selector(
        top_frame, "Variable a analizar", "📊", validas, validas[0], on_var_change)
    selector_style_frame, chart_type_var = make_selector(
        top_frame, "Estilo del Gráfico", "🎨",
        CHART_OPTIONS, "Dona HD", on_style_change)

    selector_var_frame.pack(side="left")
    selector_style_frame.pack(side="left", padx=(30, 0))

    def export_chart():
        if sys.platform == "win32":
            default_dir = os.path.join(os.path.expanduser("~"), "Pictures")
        elif sys.platform == "darwin":
            default_dir = os.path.join(os.path.expanduser("~"), "Pictures")
        else:
            default_dir = os.path.join(os.path.expanduser("~"), "Imágenes")
            if not os.path.isdir(default_dir):
                default_dir = os.path.expanduser("~")
        if not os.path.isdir(default_dir):
            default_dir = os.path.expanduser("~")

        path = _native_file_save(
            title="Exportar gráfica",
            initialdir=default_dir,
            initialfile=f"grafica_{combo_var.get()}.png",
            default_ext=".png",
            zenity_filters=["Imagen PNG | *.png",
                            "PDF vectorial | *.pdf",
                            "Imagen JPEG | *.jpg"],
            tk_filetypes=[("Imagen PNG (alta calidad)", "*.png"),
                          ("PDF vectorial", "*.pdf"),
                          ("Imagen JPEG", "*.jpg")])
        if not path:
            return
        try:
            fig.savefig(path, dpi=300, facecolor=fig.get_facecolor(),
                        bbox_inches="tight", pad_inches=0.3)
            messagebox.showinfo("Exportado", f"Gráfica guardada en:\n{path}", parent=win)
        except Exception as e:
            messagebox.showerror("Error al exportar", str(e), parent=win)

    export_btn_frame = tk.Frame(top_frame, bg=S_BG)
    export_btn_frame.pack(side="right", anchor="s")
    tk.Button(export_btn_frame, text="⬇ Exportar gráfica", command=export_chart,
              bg=C["accent"], fg=S_BG, font=FONTS["LABEL_B"], relief="flat",
              cursor="hand2", padx=14, pady=6).pack()

    body_paned = tk.PanedWindow(win, orient="horizontal",
                                bg=S_BORDER, sashwidth=4, bd=0)
    body_paned.pack(fill="both", expand=True, padx=20, pady=15)

    info_frame = tk.Frame(body_paned, bg=S_CARD,
                          highlightbackground=S_BORDER, highlightthickness=1)
    body_paned.add(info_frame, minsize=int(280*current_scale))
    info_frame.pack_propagate(False)

    header_insights = tk.Frame(info_frame, bg=C["header_bg"], pady=12)
    header_insights.pack(fill="x")
    tk.Label(header_insights, text="Insights Arqueológicos",
             bg=C["header_bg"], fg=C["header_fg"], font=FONTS["H2"]).pack()
    insight_text = tk.Text(info_frame, bg=S_CARD, fg=S_TEXT, font=FONTS["BODY"],
                           wrap="word", relief="flat", padx=18, pady=18)
    insight_text.pack(fill="both", expand=True)
    mode_footer = tk.Frame(info_frame, bg=C["accent_pale"],
                           highlightbackground=S_BORDER, highlightthickness=1)
    mode_footer.pack(fill="x", side="bottom")
    tk.Label(mode_footer,
             text="👆 Modo Exploración Visual Activo:\nHaz clic en una barra o sector para filtrar automáticamente tu tabla.",
             bg=C["accent_pale"], fg=S_ACCENT, font=FONTS["LABEL_B"],
             justify="center", wraplength=int(250*current_scale), pady=15).pack()

    chart_frame = tk.Frame(body_paned, bg=S_BG)
    body_paned.add(chart_frame, minsize=int(500*current_scale))
    try:
        _dpi_screen = win.winfo_fpixels("1i")
    except Exception:
        _dpi_screen = 96
    _fig_dpi = min(160, max(100, int(_dpi_screen * 1.25)))
    fig = Figure(figsize=(8, 6), dpi=_fig_dpi, facecolor=S_BG)
    canvas_widget = FigureCanvasTkAgg(fig, master=chart_frame)
    canvas_widget.get_tk_widget().pack(fill="both", expand=True)

    tk_chart_title = tk.Label(chart_frame, text="", bg=S_BG, fg=S_TEXT,
                               font=("Georgia", int(13*current_scale), "bold"), pady=10)
    tk_chart_title.pack(fill="x")

    def update_chart(*args):
        col   = combo_var.get()
        ctype = chart_type_var.get()
        if not isinstance(col, str) or col not in df.columns:
            return
        data   = df[col].replace("", pd.NA).dropna()
        counts = data.value_counts().sort_values(ascending=False)
        category_colors = {}
        fig.clear()
        fig.patch.set_facecolor(S_BG)
        total_count = len(data)
        n_cats = len(counts)
        _small_items_for_insight = []
        LABEL_THRESHOLD_PCT = 1.0

        if total_count == 0:
            ax = fig.add_subplot(111)
            ax.set_facecolor(S_BG)
            ax.text(0.5, 0.5, "No hay datos suficientes.",
                    ha="center", va="center", color=S_TEXT, fontsize=12)

        elif "Dona" in ctype or "Pastel" in ctype:
            ax = fig.add_axes([0.22, 0.08, 0.56, 0.84])
            ax.set_facecolor(S_BG)
            counts_sorted = counts.sort_values(ascending=True)
            wedge_w = 0.45 if "Dona" in ctype else 1.0
            colors_cycle = (S_CHART_COLORS * (n_cats // len(S_CHART_COLORS) + 1))
            for _i, _name in enumerate(counts_sorted.index):
                category_colors[_name] = colors_cycle[_i % len(colors_cycle)]

            pie_result = ax.pie(
                counts_sorted,
                labels=None,
                autopct=None,
                colors=colors_cycle[:n_cats],
                wedgeprops=dict(width=wedge_w, edgecolor=S_BG, linewidth=1.4,
                                antialiased=True),
                startangle=90)
            wedges = pie_result[0]

            for i, p in enumerate(wedges):
                pct = counts_sorted.values[i] / total_count * 100
                p._metatag_valor = f"{int(counts_sorted.values[i])} piezas ({pct:.1f}%)"
                p._metatag_label = counts_sorted.index[i]

            LABEL_THRESHOLD_PCT = 1.0

            left_items  = []
            right_items = []
            small_items = []

            for i, p in enumerate(wedges):
                ang = (p.theta2 - p.theta1) / 2.0 + p.theta1
                if ang % 90 == 0:
                    ang += 0.5
                pct = counts_sorted.values[i] / total_count * 100
                nombre = counts_sorted.index[i]
                n_abs  = int(counts_sorted.values[i])

                if pct < LABEL_THRESHOLD_PCT:
                    small_items.append((nombre, n_abs, pct))
                    continue

                ys_ang = np.sin(np.deg2rad(ang))
                xs_ang = np.cos(np.deg2rad(ang))
                label  = f"{nombre}  {pct:.1f}%"
                if xs_ang < 0:
                    left_items.append((ys_ang, label, xs_ang, i, pct))
                else:
                    right_items.append((ys_ang, label, xs_ang, i, pct))

            def place_labels_clean(items, side):
                if not items:
                    return
                ordered = sorted(items, key=lambda x: -x[0])
                n = len(ordered)

                y_top  =  1.15
                y_bot  = -1.15
                step   = (y_top - y_bot) / max(n, 1)
                total_h = step * (n - 1)
                y_start = (y_top + y_bot) / 2.0 + total_h / 2.0

                ys_placed = [y_start - k * step for k in range(n)]

                x_col = -1.80 if side == "left" else 1.80
                ha    = "right" if side == "left" else "left"
                bbox  = dict(boxstyle="round,pad=0.28",
                             fc=S_CARD, ec=S_BORDER, lw=0.7, alpha=0.90)

                for k, (y_orig, label, xs_ang, idx_w, pct) in enumerate(ordered):
                    y_lbl = ys_placed[k]
                    ang_mid = ((wedges[idx_w].theta2 - wedges[idx_w].theta1)
                               / 2.0 + wedges[idx_w].theta1)
                    r_arrow = 1.02
                    xa = r_arrow * np.cos(np.deg2rad(ang_mid))
                    ya = r_arrow * np.sin(np.deg2rad(ang_mid))
                    wedge_color = colors_cycle[idx_w % len(colors_cycle)]
                    ax.annotate(
                        label,
                        xy=(xa, ya),
                        xytext=(x_col, y_lbl),
                        horizontalalignment=ha,
                        fontsize=8,
                        color=S_TEXT,
                        arrowprops=dict(
                            arrowstyle="-",
                            color=wedge_color,
                            lw=1.0,
                            alpha=0.55,
                            shrinkA=0, shrinkB=4),
                        bbox=bbox,
                        va="center",
                        zorder=5,
                        annotation_clip=False)

            place_labels_clean(left_items,  "left")
            place_labels_clean(right_items, "right")

            tk_chart_title.configure(text=f"Distribución de {col}")

            if wedge_w < 1.0:
                top_val = counts_sorted.index[-1]
                top_pct = counts_sorted.max() / total_count * 100
                ax.text(0, 0.10, f"TOTAL\n{total_count}\npiezas",
                        ha="center", va="center",
                        color=S_TEXT, fontsize=13, fontweight="bold")
                ax.text(0, -0.30, f"Dominante:\n{top_val} ({top_pct:.1f}%)",
                        ha="center", va="center",
                        color=S_TEXT_MUTE, fontsize=8, fontstyle="italic")

            _small_items_for_insight = small_items

        elif "Horizontal" in ctype:
            ax = fig.add_subplot(111)
            ax.set_facecolor(S_BG)
            counts_asc = counts.sort_values(ascending=True)
            bar_h = max(0.35, min(0.65, 4.0 / max(n_cats, 1)))
            colors_list = (S_CHART_COLORS * (n_cats // len(S_CHART_COLORS) + 1))
            for _i, _name in enumerate(counts_asc.index):
                category_colors[_name] = colors_list[_i % len(colors_list)]
            bars = ax.barh(range(n_cats), counts_asc.values,
                           color=colors_list[:n_cats],
                           edgecolor=S_BG, linewidth=1.2, height=bar_h)
            ax.set_yticks(range(n_cats))
            ax.set_yticklabels(counts_asc.index, color=S_TEXT,
                               fontsize=max(7, min(10, 120 // max(n_cats, 1))))
            ax.set_xlabel("Cantidad de piezas", color=S_TEXT_MUTE, fontsize=9)
            ax.tick_params(colors=S_TEXT, length=0)
            for sp in ("top", "right", "left"):
                ax.spines[sp].set_visible(False)
            ax.spines["bottom"].set_color(S_BORDER)
            ax.xaxis.grid(True, linestyle=(0, (4, 4)), alpha=0.22,
                           linewidth=0.6, color=S_BORDER)
            ax.set_axisbelow(True)
            max_val = counts_asc.max()
            for bar, val in zip(bars, counts_asc.values):
                pct = val / total_count * 100
                ax.text(bar.get_width() + max_val * 0.015,
                        bar.get_y() + bar.get_height() / 2,
                        f"{val}  ({pct:.1f}%)",
                        va="center", ha="left",
                        color=S_TEXT,
                        fontsize=max(7, min(9, 100 // max(n_cats, 1))),
                        fontweight="bold")
            ax.set_xlim(0, max_val * 1.32)
            for bar, val in zip(bars, counts_asc.values):
                pct = val / total_count * 100
                bar._metatag_valor = f"{val} piezas ({pct:.1f}%)"
                bar._metatag_label = counts_asc.index[list(counts_asc.values).index(val)]
            tk_chart_title.configure(text=f"Frecuencia: {col}")

        elif "Vertical" in ctype:
            ax = fig.add_subplot(111)
            ax.set_facecolor(S_BG)
            colors_list = (S_CHART_COLORS * (n_cats // len(S_CHART_COLORS) + 1))
            for _i, _name in enumerate(counts.index):
                category_colors[_name] = colors_list[_i % len(colors_list)]
            bars = ax.bar(range(n_cats), counts.values,
                          color=colors_list[:n_cats],
                          edgecolor=S_BG, linewidth=1.2, width=0.65)
            rot = 30 if n_cats <= 8 else 45
            fs  = max(7, min(9, 80 // max(n_cats, 1)))
            ax.set_xticks(range(n_cats))
            ax.set_xticklabels(counts.index, color=S_TEXT,
                               fontsize=fs, rotation=rot, ha="right")
            ax.set_ylabel("Cantidad de piezas", color=S_TEXT_MUTE, fontsize=9)
            ax.tick_params(colors=S_TEXT, length=0)
            for sp in ("top", "right"):
                ax.spines[sp].set_visible(False)
            ax.spines["left"].set_color(S_BORDER)
            ax.spines["bottom"].set_color(S_BORDER)
            ax.yaxis.grid(True, linestyle=(0, (4, 4)), alpha=0.22,
                           linewidth=0.6, color=S_BORDER)
            ax.set_axisbelow(True)
            max_val = counts.max()
            for bar, val in zip(bars, counts.values):
                pct = val / total_count * 100
                ax.text(bar.get_x() + bar.get_width() / 2,
                        bar.get_height() + max_val * 0.012,
                        f"{val} ({pct:.1f}%)",
                        ha="center", va="bottom",
                        color=S_TEXT,
                        fontsize=max(6, min(8, 80 // max(n_cats, 1))),
                        fontweight="bold")
            ax.set_ylim(0, max_val * 1.28)
            for bar, val in zip(bars, counts.values):
                pct = val / total_count * 100
                bar._metatag_valor = f"{val} piezas ({pct:.1f}%)"
                bar._metatag_label = counts.index[list(counts.values).index(val)]
            fig.subplots_adjust(bottom=0.18)
            tk_chart_title.configure(text=f"Frecuencia: {col}")

        elif "Lollipop" in ctype:
            ax = fig.add_subplot(111)
            ax.set_facecolor(S_BG)
            counts_asc = counts.sort_values(ascending=True)
            y_pos   = list(range(n_cats))
            max_val = counts_asc.max()
            colors_list = (S_CHART_COLORS * (n_cats // len(S_CHART_COLORS) + 1))
            for _i, _name in enumerate(counts_asc.index):
                category_colors[_name] = colors_list[_i % len(colors_list)]
            for i, (y, val) in enumerate(zip(y_pos, counts_asc.values)):
                color = colors_list[i]
                ax.plot([0, val], [y, y], color=color, linewidth=1.8,
                        alpha=0.55, solid_capstyle="round")
                ax.scatter(val, y, color=color, s=90, zorder=5,
                           edgecolors=S_BG, linewidth=1.2)
                pct = val / total_count * 100
                x_offset = max_val * 0.02
                fs = max(7, min(9, 100 // max(n_cats, 1)))
                ax.text(val + x_offset, y,
                        f"{val}  ({pct:.1f}%)",
                        va="center", ha="left",
                        color=S_TEXT, fontsize=fs)
            ax.set_yticks(y_pos)
            ax.set_yticklabels(counts_asc.index, color=S_TEXT,
                               fontsize=max(7, min(10, 120 // max(n_cats, 1))))
            ax.tick_params(colors=S_TEXT, length=0)
            ax.set_xlabel("Cantidad de piezas", color=S_TEXT_MUTE, fontsize=9)
            for sp in ("top", "right", "left"):
                ax.spines[sp].set_visible(False)
            ax.spines["bottom"].set_color(S_BORDER)
            ax.xaxis.grid(True, linestyle=(0, (4, 4)), alpha=0.22,
                           linewidth=0.6, color=S_BORDER)
            ax.set_axisbelow(True)
            ax.set_xlim(-max_val * 0.02, max_val * 1.30)
            tk_chart_title.configure(text=f"Distribución: {col}")

        try:
            if "Dona" in ctype or "Pastel" in ctype:
                fig.subplots_adjust(left=0.02, right=0.98, top=0.96, bottom=0.04)
            else:
                fig.subplots_adjust(left=0.08, right=0.95, top=0.95, bottom=0.10)
        except Exception:
            fig.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.1)

        canvas_widget.draw()
        canvas_widget.get_tk_widget().update_idletasks()

        try:
            insight_text.configure(state="normal")
            insight_text.delete("1.0", "end")
            if total_count > 0:
                counts_desc = counts.sort_values(ascending=False)
                top_val     = counts_desc.index[0]
                top_count_v = int(counts_desc.iloc[0])
                low_val     = counts_desc.index[-1]
                low_count_v = int(counts_desc.iloc[-1])
                pct_top     = top_count_v / total_count * 100
                pct_low     = low_count_v / total_count * 100
                diversidad  = len(counts_desc)

                insight_text.tag_configure("h",
                    font=FONTS["LABEL_B"], foreground=S_ACCENT)
                insight_text.tag_configure("li",
                    font=FONTS["BODY"], spacing1=4)
                insight_text.tag_configure("small_h",
                    font=FONTS["LABEL_B"], foreground=S_TEXT_MUTE)
                insight_text.tag_configure("small_row",
                    font=FONTS["BODY"], spacing1=3,
                    foreground=S_TEXT_MUTE)

                insight_text.insert("end",
                    f"ANÁLISIS DE '{col.upper()}'\n\n", "h")
                insight_text.insert("end",
                    f"Se analizaron {total_count} elementos con esta "
                    f"característica registrada.\n\n")
                insight_text.insert("end",
                    "🥇 Categoría Dominante:\n", "li")
                insight_text.insert("end",
                    f" ▸ '{top_val}' lidera con {top_count_v} piezas.\n", "li")
                insight_text.insert("end",
                    f" ▸ Representa el {pct_top:.1f}% de la muestra.\n\n", "li")
                if diversidad > 1:
                    insight_text.insert("end",
                        "📉 Categoría Minoritaria:\n", "li")
                    insight_text.insert("end",
                        f" ▸ '{low_val}' es la menos frecuente, "
                        f"con {low_count_v} apariciones ({pct_low:.1f}%).\n\n", "li")
                insight_text.insert("end",
                    "⚖️ Diversidad Registrada:\n", "li")
                insight_text.insert("end",
                    f" ▸ Existen {diversidad} variaciones únicas "
                    f"documentadas.\n", "li")

                insight_text.insert("end", "\n🎨 Leyenda de Colores:\n", "h")
                for _i, (_name, _cnt) in enumerate(counts_desc.items()):
                    _color = category_colors.get(_name, S_TEXT_MUTE)
                    _tag = f"swatch_{_i}"
                    insight_text.tag_configure(_tag, foreground=_color)
                    _pct = _cnt / total_count * 100
                    insight_text.insert("end", "■ ", _tag)
                    insight_text.insert("end",
                        f"{_name}: {int(_cnt)} ({_pct:.1f}%)\n", "small_row")

                small = _small_items_for_insight
                if small:
                    insight_text.insert("end",
                        f"\n\n🔎 Categorías menores "
                        f"(< {LABEL_THRESHOLD_PCT:.0f}%, sin etiqueta "
                        f"en el gráfico):\n", "small_h")
                    for nombre, n_abs, pct in sorted(small,
                                                     key=lambda x: -x[2]):
                        insight_text.insert("end",
                            f" ▸ {nombre}: "
                            f"{n_abs} piezas ({pct:.1f}%)\n",
                            "small_row")

            else:
                insight_text.insert("end",
                    "Carga datos para visualizar el análisis.")
            insight_text.configure(state="disabled")
        except Exception:
            pass

    tooltip_annot = None

    def on_hover(event):
        nonlocal tooltip_annot
        if event.inaxes is None:
            if tooltip_annot is not None:
                tooltip_annot.set_visible(False)
                canvas_widget.draw_idle()
            return
        ax = event.inaxes
        found = False
        for patch in ax.patches:
            contains, _ = patch.contains(event)
            if contains:
                found = True
                valor = getattr(patch, "_metatag_valor", None)
                etiqueta = getattr(patch, "_metatag_label", None)
                if valor is None:
                    continue
                if tooltip_annot is None or tooltip_annot.axes != ax:
                    if tooltip_annot is not None:
                        tooltip_annot.remove()
                    tooltip_annot = ax.annotate(
                        "", xy=(0, 0), xytext=(15, 15), textcoords="offset points",
                        bbox=dict(boxstyle="round,pad=0.4", fc=S_ACCENT, ec="none", alpha=0.95),
                        color=S_BG, fontsize=9, fontweight="bold", zorder=10)
                tooltip_annot.xy = (event.xdata, event.ydata)
                tooltip_annot.set_text(f"{etiqueta}\n{valor}")
                tooltip_annot.set_visible(True)
                canvas_widget.draw_idle()
                break
        if not found and tooltip_annot is not None:
            tooltip_annot.set_visible(False)
            canvas_widget.draw_idle()

    canvas_widget.mpl_connect("motion_notify_event", on_hover)

    update_chart()
    combo_var.trace_add("write", update_chart)
    chart_type_var.trace_add("write", update_chart)
