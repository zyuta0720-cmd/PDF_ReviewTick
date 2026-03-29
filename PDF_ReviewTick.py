"""
PDF_ReviewTick.py
Copyright (c) 2026 zyutama
Released under the MIT license
https://opensource.org/licenses/mit-license.php

Requirements:
pip install PyMuPDF Pillow
Note: This software uses PyMuPDF (AGPL-3.0) and Pillow (HPND).
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import colorchooser
from tkinter import ttk
import fitz
import os
import io
import csv
from PIL import Image, ImageTk
import threading
import sys
import traceback
from datetime import datetime

class PDFDotAnnotator:
    """
    PDFドキュメント内で指定された複数のテキストを検索し、その位置の
    各文字の真上にフリーテキスト注釈（●）を追加するためのTkinterアプリケーション。
    """
    def __init__(self, master):
        self.master = master
        master.title("PDF検図補助ツール ReviewTick v1.0")
        master.geometry("1300x900") 
        
        self.s = ttk.Style()
        self.s.theme_use('clam')
        
        self.base_bg = '#f0f0f0'
        self.s.configure('TFrame', background=self.base_bg)
        self.s.configure('TButton', padding=6, relief="flat", background='#e1e1e1')
        self.s.map('TButton', background=[('active', '#cccccc')])
        
        # --- 状態管理変数 ---
        self.pdf_document = None
        self.current_page = 0
        self.total_pages = 0
        self.zoom = 1.5
        self.photo = None
        self.search_targets = []
        self.current_processing_page_num = 0
        self.opened_pdf_path = None
        self.thumbnail_widgets = []

        # デフォルトは赤
        self.search_color = (255, 0, 0) 
        self.process_thread = None

        # 進捗管理用StringVar
        self.progress_var = tk.StringVar(value="0")
        
        # スレッド間通信用のフラグとエラー保持
        self.process_error = None

        # 範囲・除外設定
        self.search_mode_var = tk.StringVar(value="all") # "all" or "rect"
        self.exclude_pages_var = tk.StringVar(value="")
        self.annot_type_var = tk.StringVar(value="dot") # "dot" or "check"
        self.annot_size_var = tk.DoubleVar(value=1.0)
        self.annot_offset_var = tk.DoubleVar(value=0.0)
        
        # 矩形選択用
        self.rect_id = None
        self.start_x = None
        self.start_y = None
        self.selection_pdf_rect = None # fitz.Rect (PDF座標系)
        
        # --- メインレイアウト ---
        self.main_frame = ttk.Frame(master)
        self.main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.control_frame = ttk.Frame(self.main_frame, width=300)
        self.control_frame.pack(side='left', fill='y', padx=(0, 10))
        self.control_frame.pack_propagate(False)
        
        self.pdf_display_frame = ttk.Frame(self.main_frame)
        self.pdf_display_frame.pack(side='right', fill='both', expand=True)

        # --- コントロールパネル要素 ---
        self._setup_control_panel()
        
        # --- PDF表示要素 ---
        self._setup_pdf_viewer()

        # イベントのバインド
        self.master.bind("<<UpdateProgress>>", self._on_progress_event)
        
        master.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _setup_control_panel(self):
        # ファイル操作グループ
        file_group = ttk.LabelFrame(self.control_frame, text="1. ファイル操作", padding="10")
        file_group.pack(fill='x', pady=5)
        ttk.Button(file_group, text="PDFを開く", command=self.open_pdf).pack(fill='x', pady=2)
        
        # 検索対象グループ
        search_group = ttk.LabelFrame(self.control_frame, text="2. 検索対象 (改行区切り)", padding="10")
        search_group.pack(fill='x', pady=5)
        self.search_text_area = scrolledtext.ScrolledText(search_group, height=10, width=30, wrap=tk.WORD, font=('TkDefaultFont', 10))
        self.search_text_area.pack(fill='x', pady=2)
        ttk.Button(search_group, text="検索対象を更新", command=self._update_search_targets).pack(fill='x', pady=2)

        # 検索設定グループ
        setting_group = ttk.LabelFrame(self.control_frame, text="3. 検索/注釈設定", padding="10")
        setting_group.pack(fill='x', pady=5)
        
        color_frame = ttk.Frame(setting_group)
        color_frame.pack(fill='x', pady=2)
        ttk.Label(color_frame, text="ドット色:").pack(side='left', padx=(0, 5))
        self.color_display = tk.Canvas(color_frame, width=20, height=20, bg=f'#{self.search_color[0]:02x}{self.search_color[1]:02x}{self.search_color[2]:02x}', relief="sunken")
        self.color_display.pack(side='left', padx=(0, 5))
        ttk.Button(color_frame, text="色を選択", command=self._choose_color).pack(side='left', fill='x', expand=True)

        # 注釈の形状選択
        type_frame = ttk.Frame(setting_group)
        type_frame.pack(fill='x', pady=5)
        ttk.Radiobutton(type_frame, text="各文字に ●", variable=self.annot_type_var, value="dot").pack(side='left', padx=(0, 10))
        ttk.Radiobutton(type_frame, text="右端に ✔", variable=self.annot_type_var, value="check").pack(side='left')

        # 調整スライダー
        adj_frame = ttk.Frame(setting_group)
        adj_frame.pack(fill='x', pady=5)
        
        ttk.Label(adj_frame, text="サイズ倍率:").grid(row=0, column=0, sticky='w')
        tk.Scale(adj_frame, from_=0.5, to=3.0, resolution=0.1, orient=tk.HORIZONTAL, variable=self.annot_size_var, bg=self.base_bg, highlightthickness=0).grid(row=0, column=1, sticky='ew')
        
        ttk.Label(adj_frame, text="上下位置調整:").grid(row=1, column=0, sticky='w')
        tk.Scale(adj_frame, from_=-10, to=10, resolution=0.5, orient=tk.HORIZONTAL, variable=self.annot_offset_var, bg=self.base_bg, highlightthickness=0).grid(row=1, column=1, sticky='ew')
        adj_frame.columnconfigure(1, weight=1)

        # 範囲・除外設定グループ
        range_group = ttk.LabelFrame(self.control_frame, text="4. 対象範囲/除外設定", padding="10")
        range_group.pack(fill='x', pady=5)

        ttk.Radiobutton(range_group, text="ページ全体を検索", variable=self.search_mode_var, value="all", command=self._on_search_mode_change).pack(anchor='w')
        ttk.Radiobutton(range_group, text="選択範囲のみ検索 (マウスで指定)", variable=self.search_mode_var, value="rect", command=self._on_search_mode_change).pack(anchor='w')
        
        self.rect_info_label = ttk.Label(range_group, text="範囲: 未設定 (全ページ共通)", foreground="blue")
        self.rect_info_label.pack(anchor='w', pady=(2, 5))
        ttk.Button(range_group, text="選択範囲をリセット", command=self._reset_selection).pack(fill='x', pady=2)

        ttk.Label(range_group, text="除外ページ (例: 1, 3, 5-8):").pack(anchor='w', pady=(5, 0))
        self.exclude_entry = ttk.Entry(range_group, textvariable=self.exclude_pages_var)
        self.exclude_entry.pack(fill='x', pady=2)

        # 実行ボタン
        ttk.Button(self.control_frame, text="5. 全ページを処理しPDFを保存", command=self.start_processing, style='Accent.TButton').pack(fill='x', pady=10)
        self.s.configure('Accent.TButton', background='#4CAF50', foreground='white')
        self.s.map('Accent.TButton', background=[('active', '#388E3C')])

        # 処理状況
        self.progress_bar = ttk.Progressbar(self.control_frame, orient='horizontal', mode='determinate')
        self.progress_bar.pack(fill='x', pady=5)
        self.status_label = ttk.Label(self.control_frame, text="ステータス: 準備完了", anchor='w')
        self.status_label.pack(fill='x')


    def _setup_pdf_viewer(self):
        # メインビューアとサイドバー（しおり/サムネイル）を分けるためのパンウィンドウ
        self.paned = ttk.PanedWindow(self.pdf_display_frame, orient=tk.HORIZONTAL)
        self.paned.pack(fill='both', expand=True)

        self.viewer_main_frame = ttk.Frame(self.paned)
        self.sidebar_frame = ttk.Frame(self.paned, width=200)
        self.paned.add(self.viewer_main_frame, weight=5)
        self.paned.add(self.sidebar_frame, weight=1)

        # ページナビゲーション
        nav_frame = ttk.Frame(self.viewer_main_frame)
        nav_frame.pack(fill='x', pady=(0, 5))
        
        ttk.Button(nav_frame, text="前へ", command=lambda: self._change_page(-1)).pack(side='left')
        self.page_label = ttk.Label(nav_frame, text="ページ 0 / 0", width=20, anchor='center')
        self.page_label.pack(side='left', padx=10)
        ttk.Button(nav_frame, text="次へ", command=lambda: self._change_page(1)).pack(side='left')

        # ズーム
        ttk.Label(nav_frame, text="ズーム:").pack(side='right', padx=(10, 5))
        ttk.Button(nav_frame, text="+", width=3, command=lambda: self._change_zoom(0.2)).pack(side='right')
        ttk.Button(nav_frame, text="-", width=3, command=lambda: self._change_zoom(-0.2)).pack(side='right', padx=(5, 0))

        # キャンバスとスクロールバーのコンテナ
        self.canvas_container = ttk.Frame(self.viewer_main_frame)
        self.canvas_container.pack(fill='both', expand=True)

        self.v_scroll = ttk.Scrollbar(self.canvas_container, orient=tk.VERTICAL)
        self.h_scroll = ttk.Scrollbar(self.canvas_container, orient=tk.HORIZONTAL)
        self.canvas = tk.Canvas(
            self.canvas_container, 
            bg='gray70', 
            highlightthickness=0,
            xscrollcommand=self.h_scroll.set,
            yscrollcommand=self.v_scroll.set
        )
        
        self.v_scroll.config(command=self.canvas.yview)
        self.h_scroll.config(command=self.canvas.xview)
        
        self.v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill='both', expand=True)

        # サイドバー（サムネイル一覧）の設定
        ttk.Label(self.sidebar_frame, text="しおり / ページ一覧", font=("", 10, "bold")).pack(pady=5)
        
        self.thumb_canvas = tk.Canvas(self.sidebar_frame, bg='#d0d0d0', highlightthickness=0)
        self.thumb_scroll = ttk.Scrollbar(self.sidebar_frame, orient=tk.VERTICAL, command=self.thumb_canvas.yview)
        self.thumb_canvas.config(yscrollcommand=self.thumb_scroll.set)
        
        self.thumb_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.thumb_canvas.pack(side=tk.LEFT, fill='both', expand=True)
        
        self.thumb_list_frame = ttk.Frame(self.thumb_canvas)
        self.thumb_canvas.create_window((0, 0), window=self.thumb_list_frame, anchor='nw', tags="frame")
        
        # サムネイルフレームのサイズ変更時にスクロール領域を更新
        self.thumb_list_frame.bind("<Configure>", lambda e: self.thumb_canvas.configure(scrollregion=self.thumb_canvas.bbox("all")))
        
        # マウスイベントのバインド
        self.canvas.bind("<ButtonPress-1>", self._on_mouse_down)
        self.canvas.bind("<B1-Motion>", self._on_mouse_move)
        self.canvas.bind("<ButtonRelease-1>", self._on_mouse_up)

    def _on_search_mode_change(self):
        if self.search_mode_var.get() == "all":
            self._reset_selection()
        self._display_page()

    def _reset_selection(self):
        self.selection_pdf_rect = None
        self.rect_info_label.config(text="範囲: 未設定 (全ページ共通)")
        if self.rect_id:
            self.canvas.delete(self.rect_id)
            self.rect_id = None
        self._display_page()

    def _on_mouse_down(self, event):
        if self.search_mode_var.get() != "rect": return
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.rect_id = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red', width=2)

    def _on_mouse_move(self, event):
        if self.search_mode_var.get() != "rect" or not self.rect_id: return
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect_id, self.start_x, self.start_y, cur_x, cur_y)

    def _on_mouse_up(self, event):
        if self.search_mode_var.get() != "rect" or not self.rect_id: return
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)
        
        # Canvas座標からPDF座標に変換
        x0, x1 = sorted([self.start_x / self.zoom, end_x / self.zoom])
        y0, y1 = sorted([self.start_y / self.zoom, end_y / self.zoom])
        
        self.selection_pdf_rect = fitz.Rect(x0, y0, x1, y1)
        self.rect_info_label.config(text=f"範囲: 設定済み ({int(x0)},{int(y0)}) - ({int(x1)},{int(y1)})")
        
        messagebox.showinfo("範囲設定", "検索範囲を設定しました。全ページに適用されます。")

    def _update_search_targets(self):
        """検索対象テキストリストを更新し、注釈色をチェックする"""
        text = self.search_text_area.get(1.0, tk.END).strip()
        self.search_targets = [t.strip() for t in text.split('\n') if t.strip()]
        messagebox.showinfo("更新完了", f"検索対象を{len(self.search_targets)}件に更新しました。")
        self._display_page()

    def _choose_color(self):
        """注釈色を選択するダイアログを表示"""
        color_code = colorchooser.askcolor(title="ドット色を選択")
        if color_code:
            self.search_color = tuple(int(c) for c in color_code[0])
            hex_color = color_code[1]
            self.color_display.config(bg=hex_color)

    def open_pdf(self):
        """PDFファイルを開き、初期化する"""
        if self.process_thread and self.process_thread.is_alive():
            messagebox.showwarning("処理中", "現在、PDF処理が進行中です。完了をお待ちください。")
            return

        filepath = filedialog.askopenfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if not filepath:
            return

        self._close_document()

        try:
            self.pdf_document = fitz.open(filepath)
            self.opened_pdf_path = filepath
            self.total_pages = self.pdf_document.page_count
            self.current_page = 0
            self._generate_thumbnails()
            self._display_page()
            self.status_label.config(text=f"ステータス: ファイル '{os.path.basename(filepath)}' を開きました。")
        except Exception as e:
            messagebox.showerror("エラー", f"PDFを開けませんでした: {e}")
            self.pdf_document = None
            self.total_pages = 0

    def _close_document(self):
        """開いているドキュメントを閉じる"""
        if self.pdf_document:
            self.pdf_document.close()
            self.pdf_document = None
        self.current_page = 0
        self.total_pages = 0
        self.canvas.delete("all")
        for widget in self.thumb_list_frame.winfo_children():
            widget.destroy()
        self.thumbnail_widgets = []
        self.thumb_canvas.config(scrollregion=(0,0,0,0))
        self.page_label.config(text="ページ 0 / 0")

    def _change_page(self, delta):
        """ページを移動する"""
        if not self.pdf_document:
            return

        new_page = self.current_page + delta
        if 0 <= new_page < self.total_pages:
            self.current_page = new_page
            self._display_page()

    def _change_zoom(self, delta):
        """ズームを変更する"""
        self.zoom = max(0.5, self.zoom + delta)
        self._display_page()

    def _display_page(self):
        """現在のページをキャンバスに表示する"""
        if not self.pdf_document or self.current_page >= self.total_pages:
            self.canvas.delete("all")
            return

        page = self.pdf_document.load_page(self.current_page)
        
        matrix = fitz.Matrix(self.zoom, self.zoom)
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        
        img_data = pix.tobytes("ppm")
        image = Image.open(io.BytesIO(img_data))
        
        self.photo = ImageTk.PhotoImage(image)
        
        # キャンバスのスクロール範囲を画像サイズに更新
        self.canvas.config(scrollregion=(0, 0, self.photo.width(), self.photo.height()))
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, image=self.photo, anchor="nw")
        
        # 選択範囲があれば再描画
        if self.search_mode_var.get() == "rect" and self.selection_pdf_rect:
            x0 = self.selection_pdf_rect.x0 * self.zoom
            y0 = self.selection_pdf_rect.y0 * self.zoom
            x1 = self.selection_pdf_rect.x1 * self.zoom
            y1 = self.selection_pdf_rect.y1 * self.zoom
            self.rect_id = self.canvas.create_rectangle(x0, y0, x1, y1, outline='red', width=2, dash=(4, 4))
        
        self.page_label.config(text=f"ページ {self.current_page + 1} / {self.total_pages}")
        self._update_thumbnail_highlight()

    def _generate_thumbnails(self):
        """サイドバーに全ページのサムネイルを生成する"""
        for widget in self.thumb_list_frame.winfo_children():
            widget.destroy()
        self.thumbnail_widgets = []

        if not self.pdf_document: return

        # サムネイルは固定の小さな倍率で生成
        thumb_zoom = 0.15
        for i in range(self.total_pages):
            page = self.pdf_document.load_page(i)
            pix = page.get_pixmap(matrix=fitz.Matrix(thumb_zoom, thumb_zoom))
            img = Image.open(io.BytesIO(pix.tobytes("ppm")))
            photo = ImageTk.PhotoImage(img)
            
            f = tk.Frame(self.thumb_list_frame, bg='#d0d0d0', pady=5)
            f.pack(fill='x', padx=5)
            
            l = tk.Label(f, image=photo, bg='#d0d0d0')
            l.image = photo # 参照保持
            l.pack()
            
            txt = tk.Label(f, text=f"P.{i+1}", bg='#d0d0d0', font=("", 8))
            txt.pack()
            
            # クリックイベントのバインド
            for w in (f, l, txt):
                w.bind("<Button-1>", lambda e, p=i: self._go_to_page(p))
            
            self.thumbnail_widgets.append(f)

    def _go_to_page(self, page_num):
        self.current_page = page_num
        self._display_page()

    def _update_thumbnail_highlight(self):
        """現在のページに対応するサムネイルを強調表示し、自動スクロールする"""
        if not self.thumbnail_widgets: return
        
        for i, f in enumerate(self.thumbnail_widgets):
            bg_color = '#4CAF50' if i == self.current_page else '#d0d0d0'
            f.config(bg=bg_color)
            for child in f.winfo_children():
                child.config(bg=bg_color)
            
            if i == self.current_page:
                # サムネイル一覧の中央付近に来るようにスクロール
                total = len(self.thumbnail_widgets)
                self.thumb_canvas.yview_moveto(max(0, (i - 2) / total) if total > 0 else 0)

    def _get_excluded_pages(self):
        """除外ページ設定をパースして整数のセットを返す"""
        exclude_str = self.exclude_pages_var.get().strip()
        if not exclude_str:
            return set()
        
        excluded = set()
        parts = [p.strip() for p in exclude_str.split(',')]
        for part in parts:
            try:
                if '-' in part:
                    start, end = map(int, part.split('-'))
                    excluded.update(range(start, end + 1))
                else:
                    excluded.add(int(part))
            except ValueError:
                continue
        return excluded

    # ---------------- 検索/注釈処理----------------
    def _process_page(self, page_index, base, page_width, page_height, settings):
        """
        単一のページを処理し、検索テキストに注釈を付け、PDFとCSVを更新する。
        """
        excluded_pages = settings["excluded_pages"]
        if (page_index + 1) in excluded_pages:
            return False, 0, []

        page = base.load_page(page_index)
        page_search_summary = []
        total_found_on_page = 0
        
        # 注釈の色のR, G, Bを0.0から1.0の範囲に正規化
        search_color = settings["search_color"]
        annot_color = (search_color[0] / 255, search_color[1] / 255, search_color[2] / 255)

        # 検索範囲の設定
        search_clip = None
        if settings["search_mode"] == "rect" and settings["selection_pdf_rect"]:
            search_clip = settings["selection_pdf_rect"]
        
        # 1. テキスト検索結果の準備 (PyMuPDFのRectとターゲット文字列のタプル)
        text_instances = []
        counts_on_page = {}

        # 診断用: ページ内にテキストが存在するか確認
        page_text = page.get_text().strip()
        if not page_text:
            print(f"DEBUG: ページ {page_index + 1} に検索可能なテキストが見当たりません。")

        # 通常モード: PDFのテキストレイヤーを直接検索
        for target in settings["search_targets"]:
            # ベースプログラムと同じシンプルな検索方法に戻す
            rects = page.search_for(target, clip=search_clip)
            counts_on_page[target] = len(rects)
            
            if rects:
                # コンソールに検出ログを表示（デバッグ用）
                print(f"ページ {page_index + 1}: '{target}' を {len(rects)} 件検出")
            else:
                # 検出されない場合のデバッグ情報を強化
                if target in page_text:
                    print(f"DEBUG: ページ {page_index + 1} 内に '{target}' は存在しますが、座標特定に失敗しました。")
            
            for rect in rects:
                text_instances.append((rect, target))
                total_found_on_page += 1

        # 2. 注釈の追加 (各文字の真上に固定サイズの●を配置)
        
        annot_type = settings["annot_type"]
        size_mult = settings["annot_size"]
        v_offset = settings["annot_offset"]
        
        current_target_counts = {}
        for inst_rect, target in text_instances:
            # 文字幅を均等に計算（最も確実な方法）
            char_width = inst_rect.width / len(target)
            h = inst_rect.height

            if annot_type == "dot":
                # 各文字の真上にドットを配置
                for i in range(len(target)):
                    # 半径の計算
                    r = max(h / 6.0, 1.5) * size_mult
                    
                    # 中心座標
                    cx = inst_rect.x0 + (i + 0.5) * char_width
                    cy = inst_rect.y0 - r - 1.0 + v_offset

                    dot_rect = fitz.Rect(cx - r, cy - r, cx + r, cy + r)
                    annot = page.add_circle_annot(dot_rect)
                    annot.set_colors(stroke=annot_color, fill=annot_color)
                    annot.set_opacity(1.0)
                    annot.update()
            
            else: # "check" (✔)
                # 文字列の右端にチェックマークを配置
                x_base = inst_rect.x1 + 1
                y_top = inst_rect.y0 + v_offset
                size = h * size_mult
                
                # チェックマークの3点
                p1 = fitz.Point(x_base + size * 0.2, y_top + size * 0.5)
                p2 = fitz.Point(x_base + size * 0.45, y_top + size * 0.85)
                p3 = fitz.Point(x_base + size * 0.85, y_top + size * 0.15)
                
                annot = page.add_polyline_annot([p1, p2, p3])
                annot.set_colors(stroke=annot_color)
                annot.set_border(width=max(1.0, (h/8) * size_mult)) 
                annot.set_opacity(1.0)
                annot.update()

            # 出現順序をカウントアップ
            current_target_counts[target] = current_target_counts.get(target, 0) + 1
            page_search_summary.append({
                "page": page_index + 1,
                "target": target,
                "occurrence_no": current_target_counts[target],
                "total_on_page": counts_on_page[target],
                "x": round(inst_rect.x0, 1),
                "y": round(inst_rect.y0, 1)
            })

        return total_found_on_page > 0, total_found_on_page, page_search_summary

    # ---------------- スレッド処理 ----------------
    def start_processing(self):
        """全ページ処理をスレッドで開始する"""
        
        if not self.pdf_document:
            messagebox.showwarning("ファイルなし", "先にPDFファイルを開いてください。")
            return
            
        # 実行直前に検索対象を最新の状態に更新する
        text = self.search_text_area.get(1.0, tk.END).strip()
        self.search_targets = [t.strip() for t in text.split('\n') if t.strip()]

        if not self.search_targets:
            messagebox.showwarning("検索対象なし", "検索対象のテキストを入力してください。")
            return
            
        if self.search_mode_var.get() == "rect" and not self.selection_pdf_rect:
            messagebox.showwarning("範囲未設定", "検索範囲が選択されていません。マウスで範囲を指定するか、ページ全体を選択してください。")
            return

        if not self.opened_pdf_path:
            messagebox.showerror("エラー", "PDFファイルのパスが取得できません。再度ファイルを開き直してください。")
            return

        # スレッド開始前に状態をリセット
        self.process_error = None

        # 出力ファイルパスの選択
        default_output_name = os.path.splitext(os.path.basename(self.opened_pdf_path))[0] + "_annotated.pdf"
        output_pdf_path = filedialog.asksaveasfilename(
            title="名前を付けて保存",
            defaultextension=".pdf",
            initialfile=default_output_name,

            filetypes=[("PDF files", "*.pdf")]
        )
        if not output_pdf_path:
            return

        csv_path = os.path.splitext(output_pdf_path)[0] + "_summary.csv"

        # ページサイズを取得
        first_page = self.pdf_document.load_page(0)
        page_width = first_page.rect.width
        page_height = first_page.rect.height
        
        # 処理開始
        self.progress_bar.config(mode='determinate', value=0, maximum=self.total_pages)
        self.status_label.config(text="ステータス: 処理開始...")
        
        # スレッドに渡す設定を事前にメインスレッドで取得 (Tkinter変数はスレッドセーフではないため)
        settings = {
            "search_targets": list(self.search_targets),
            "search_color": self.search_color,
            "search_mode": self.search_mode_var.get(),
            "selection_pdf_rect": self.selection_pdf_rect,
            "excluded_pages": self._get_excluded_pages(),
            "annot_type": self.annot_type_var.get(),
            "annot_size": self.annot_size_var.get(),
            "annot_offset": self.annot_offset_var.get(),
            "pdf_path": self.opened_pdf_path
        }

        # スレッド開始
        self.process_thread = threading.Thread(target=self._run_process_in_thread, args=(output_pdf_path, csv_path, page_width, page_height, settings))
        self.process_thread.daemon = True # アプリケーション終了時にスレッドも終了させる
        self.process_thread.start()

        
        self.master.after(100, self._check_thread)

    def _check_thread(self):
        """スレッドの完了をチェックし、GUIを更新する"""
        if self.process_thread and self.process_thread.is_alive():
            self.master.after(100, self._check_thread)
        else:
            self._process_completed()

    def _run_process_in_thread(self, output_pdf_path, csv_path, page_width, page_height, settings):
        """
        全ページ処理のメインロジック (スレッドで実行)
        """
        base = None
        try:
            # UI用のハンドルと分けるため、新しく開き直す（スレッドセーフのため）
            base = fitz.open(settings["pdf_path"])
            total_found = 0
            search_summary = []
            
            for i in range(self.total_pages):
                try:
                    # 進捗通知用の数値を更新 (Tkinter変数を介さず、直接数値をセット)
                    self.current_processing_page_num = i + 1
                    self.master.event_generate("<<UpdateProgress>>")
                    
                    found_any, found_count, page_summary = self._process_page(i, base, page_width, page_height, settings)
                    total_found += found_count
                    search_summary.extend(page_summary)
                except Exception as page_err:
                    print(f"ページ {i+1} の処理中にエラー: {page_err}")
                    continue # 特定のページのエラーで全体を止めない

            # PDFを保存（ファイルが他で開かれている場合の対策）
            try:
                # 同一ファイルへの上書き（Save）か、別名保存（Save As）かを判定
                if os.path.abspath(output_pdf_path) == os.path.abspath(settings["pdf_path"]):
                    # 同じファイルに保存する場合は増分保存（incremental）
                    base.saveIncr()
                else:
                    # 別のファイルとして保存する場合は、構造を最適化して圧縮保存
                    base.save(output_pdf_path, garbage=4, deflate=True, clean=True)
                print(f"PDF保存完了: {output_pdf_path}")
            except Exception as e:
                if any(err in str(e) for err in ["Permission denied", "code=2", "code=12"]):
                    raise Exception(f"PDF保存失敗: '{os.path.basename(output_pdf_path)}' が他のソフト（Acrobat等）で開かれています。\nファイルを閉じてから再試行してください。")
                raise e
            
            # CSVの追記保存 (実行履歴を蓄積する)
            file_exists = os.path.exists(csv_path)
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            try:
                # 既存のCSVがある場合は追記('a')、ない場合は新規作成
                # Excelで開きやすいように shift_jis を優先
                with open(csv_path, 'a', newline='', encoding='shift_jis') as f:
                    writer = csv.writer(f)
                    if not file_exists:
                        writer.writerow(["実行日時", "ページ", "検索対象", "出現番号", "ページ内合計", "X座標", "Y座標"])
                    for item in search_summary:
                        writer.writerow([timestamp, item['page'], item['target'], item['occurrence_no'], item['total_on_page'], item['x'], item['y']])
            except PermissionError:
                raise Exception(f"CSV保存失敗: '{os.path.basename(csv_path)}' がExcel等で開かれています。\nファイルを閉じてから再試行してください。")
            except UnicodeEncodeError:
                # shift_jisで書けない文字がある場合はUTF-8(BOM付き)で追記
                file_exists_utf8 = os.path.exists(csv_path)
                try:
                    with open(csv_path, 'a', newline='', encoding='utf_8_sig') as f:
                        writer = csv.writer(f)
                        if not file_exists_utf8:
                            writer.writerow(["Timestamp", "Page", "Target", "MatchNo", "TotalOnPage", "X", "Y"])
                        for item in search_summary:
                            writer.writerow([timestamp, item['page'], item['target'], item['occurrence_no'], item['total_on_page'], item['x'], item['y']])
                except PermissionError:
                    raise Exception(f"CSV保存失敗: '{os.path.basename(csv_path)}' がExcel等で開かれています。\nファイルを閉じてから再試行してください。")

            # 結果をスレッドオブジェクトに保存
            self.process_thread.result = (total_found, output_pdf_path, csv_path)
        except Exception as e:
            # 詳細なトレースバックをコンソールに出力
            err_detail = traceback.format_exc()
            print(f"重大な処理エラー:\n{err_detail}", file=sys.stderr)
            self.process_thread.result = e
        finally:
            if base:
                base.close()

    def _on_progress_event(self, event):
        """スレッドからの進捗更新通知を受けてUIを更新する"""
        page_num = self.current_processing_page_num
        if self.total_pages > 0:
            self.progress_bar.config(value=page_num)
            self.status_label.config(text=f"ステータス: ページ {page_num} / {self.total_pages} を処理中...")
        self.master.update_idletasks()

    def _process_completed(self):
        """処理完了後のGUI更新と結果表示"""
        # resultが設定されていない場合の安全策
        result = getattr(self.process_thread, 'result', None)
        
        if isinstance(result, Exception):
            self.progress_bar.config(value=0)
            self.status_label.config(text=f"エラー: {str(result)}")
            messagebox.showerror("処理失敗", str(result))
        elif result is None:
            self.progress_bar.config(value=0)
            self.status_label.config(text="エラー: 処理が中断されました。")
            messagebox.showerror("処理失敗", "予期せぬエラーにより処理が完了しませんでした。")
        else:
            self.progress_bar.config(value=self.total_pages)
            total_found, output_pdf_path, csv_path = result

            # 重要：処理結果を確認させるため、保存したファイルを現在のビューアで開き直す
            self._reload_annotated_pdf(output_pdf_path)

            # 保存先のフォルダを開くか確認
            if messagebox.askyesno("完了", f"処理が完了しました。合計{total_found}件の注釈を追加しました。\n\n保存先のフォルダを開きますか？"):
                output_dir = os.path.dirname(os.path.abspath(output_pdf_path))
                if os.path.exists(output_dir):
                    os.startfile(output_dir)

            self.status_label.config(text=f"ステータス: 処理完了。合計{total_found}件の注釈を追加しました。")
            messagebox.showinfo(
                "処理完了", 
                f"全ページの処理が完了しました。\n"
                f"注釈付きPDF: {output_pdf_path}\n"
                f"検索結果CSV: {csv_path}"
            )

    def _reload_annotated_pdf(self, new_path):
        """保存した注釈付きPDFを読み込んで表示を更新する"""
        try:
            if self.pdf_document:
                self.pdf_document.close()
            
            self.pdf_document = fitz.open(new_path)
            self.opened_pdf_path = new_path
            self.total_pages = self.pdf_document.page_count
            self._generate_thumbnails()
            self._display_page()
        except Exception as e:
            print(f"再読み込みエラー: {e}")

    # ---------------- 終了処理 ----------------
    def on_closing(self):
        """アプリケーション終了時の処理。ドキュメントを閉じ、ウィンドウを破壊します。"""
        if self.process_thread and self.process_thread.is_alive():
            messagebox.showwarning("処理中", "現在、PDF処理が進行中です。完了をお待ちください。")
            return
            
        self._close_document()
        if hasattr(self, "master"):
             self.master.destroy()


if __name__ == "__main__":
    
    # 必要なライブラリチェック
    try:
        if sys.platform.startswith('win'):
             try:
                 from ctypes import windll
                 windll.shcore.SetProcessDpiAwareness(1)
             except Exception:
                 pass
                 
    except Exception as e:
        root = tk.Tk()
        root.withdraw() 
        messagebox.showerror("起動エラー", f"プログラムの実行に必要なライブラリの初期化に失敗しました: {e}")
        sys.exit(1)
        
    root = tk.Tk()
    app = PDFDotAnnotator(root)
    
    root.mainloop()
