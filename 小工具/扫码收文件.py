# -*- coding: utf-8 -*-
"""扫码收文件（独立小工具）：手机扫二维码，把文件传到这台电脑。

同一 WiFi 下用，文件不过公网。收到的文件落在程序目录的 会务收件/。

源码运行：python 小工具/扫码收文件.py
"""

import _boot


def _size_text(n):
    try:
        n = float(n)
    except (TypeError, ValueError):
        return '-'
    for unit in ('B', 'KB', 'MB', 'GB'):
        if n < 1024 or unit == 'GB':
            return ('%.0f %s' % (n, unit)) if unit == 'B' \
                else ('%.1f %s' % (n, unit))
        n /= 1024.0
    return '-'


def _open(root):
    import os
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    import file_inbox
    import qrcode_lite
    import ui_theme

    p = ui_theme.PALETTE
    folder = os.path.join(_boot.APP_DIR, '会务收件')
    inbox = file_inbox.Inbox(folder=folder)

    win = tk.Toplevel(root)
    win.title('扫码收文件')
    win.geometry('880x640')
    win.minsize(820, 560)
    win.configure(bg=p['bg'])
    try:
        ui_theme.style_window(win)
    except Exception:
        pass

    keep = {'photo': None, 'rows': 0}

    body = tk.Frame(win, bg=p['bg'], padx=16, pady=14)
    body.pack(fill=tk.BOTH, expand=True)

    head = tk.Frame(body, bg=p['bg'])
    head.pack(fill=tk.X)
    tk.Label(head, text='扫码收文件', font=ui_theme.font(13, 'bold'),
             bg=p['bg'], fg=p['text']).pack(side=tk.LEFT)
    state_lbl = tk.Label(head, text='未启动', font=ui_theme.font(9),
                         bg=p['bg'], fg=p['muted'])
    state_lbl.pack(side=tk.LEFT, padx=(12, 0))
    tk.Label(head, text='手机和电脑连同一个 WiFi，用相机/微信扫一扫',
             font=ui_theme.font(9), bg=p['bg'],
             fg=p['faint']).pack(side=tk.RIGHT)

    top = tk.Frame(body, bg=p['bg'])
    top.pack(fill=tk.X, pady=(12, 10))

    qr_box = tk.Frame(top, bg=p['surface'], highlightthickness=1,
                      highlightbackground=p['border'])
    qr_box.pack(side=tk.LEFT, anchor='n')
    qr_label = tk.Label(qr_box, text='未启动', bg=p['surface'],
                        fg=p['faint'], font=ui_theme.font(9),
                        width=26, height=11)
    qr_label.pack(padx=10, pady=10)
    ttk.Button(qr_box, text='保存二维码图片', style='Ghost.TButton',
               command=lambda: _save_qr()).pack(padx=10, pady=(0, 10))

    right = tk.Frame(top, bg=p['bg'])
    right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(14, 0))

    def _row(label):
        row = tk.Frame(right, bg=p['bg'])
        row.pack(fill=tk.X, pady=3)
        tk.Label(row, text=label, font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted'], width=8, anchor='e').pack(
            side=tk.LEFT, padx=(0, 8))
        return row

    folder_var = tk.StringVar(value=inbox.folder)
    row = _row('保存到')
    tk.Entry(row, textvariable=folder_var, font=ui_theme.font(9),
             state='readonly', readonlybackground=p['surface'],
             relief=tk.FLAT, bd=0, highlightthickness=1,
             highlightbackground=p['border']).pack(
        side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
    ttk.Button(row, text='换位置', width=7,
               command=lambda: _choose_folder()).pack(side=tk.LEFT,
                                                      padx=(6, 0))
    ttk.Button(row, text='打开', width=5,
               command=lambda: _open_folder()).pack(side=tk.LEFT,
                                                    padx=(6, 0))

    url_var = tk.StringVar(value='未启动')
    row = _row('手机网址')
    tk.Entry(row, textvariable=url_var, font=ui_theme.font(9),
             state='readonly', readonlybackground=p['surface'],
             relief=tk.FLAT, bd=0, highlightthickness=1,
             highlightbackground=p['border']).pack(
        side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
    ttk.Button(row, text='复制', width=5,
               command=lambda: _copy_url()).pack(side=tk.LEFT, padx=(6, 0))

    row = _row('操作')
    run_btn = ttk.Button(row, text='开始接收', style='Accent.TButton',
                         command=lambda: _toggle())
    run_btn.pack(side=tk.LEFT)
    ttk.Button(row, text='打开文件夹', style='Ghost.TButton',
               command=lambda: _open_folder()).pack(side=tk.LEFT,
                                                    padx=(8, 0))
    ttk.Button(row, text='清空列表', style='Ghost.TButton',
               command=lambda: _clear()).pack(side=tk.LEFT, padx=(8, 0))

    bottom = tk.Frame(body, bg=p['bg'])
    bottom.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
    cols = ('name', 'size', 'time', 'from')
    tree = ttk.Treeview(bottom, columns=cols, show='headings', height=10)
    for key, text, width in (('name', '文件名', 380), ('size', '大小', 90),
                             ('time', '时间', 90), ('from', '来源', 150)):
        tree.heading(key, text=text)
        tree.column(key, width=width,
                    anchor='w' if key in ('name', 'from') else 'center')
    sb = ttk.Scrollbar(bottom, orient='vertical', command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    sb.pack(side=tk.RIGHT, fill=tk.Y)
    tree.bind('<Double-Button-1>', lambda e: _open_selected())

    # ---------- 行为 ----------
    def _refresh_qr():
        if not inbox.running:
            keep['photo'] = None
            qr_label.configure(image='', text='未启动', width=26, height=11)
            return
        try:
            matrix = qrcode_lite.make_matrix(inbox.url())
            img = qrcode_lite.to_image(matrix, scale=5, border=4)
            from PIL import ImageTk
            photo = ImageTk.PhotoImage(img)
        except Exception as exc:
            qr_label.configure(image='', text='二维码生成失败：%s' % exc,
                               width=26, height=11)
            keep['photo'] = None
            return
        keep['photo'] = photo
        # 关键：占位时 width/height 是字符单位，显示图片时会被当成像素，
        # 把二维码裁成 26x11 的一角（必须清零恢复自然尺寸）
        qr_label.configure(image=photo, text='', width=0, height=0)

    def _set_state(running):
        state_lbl.configure(text='接收中' if running else '未启动',
                            fg='#2e8b57' if running else p['muted'])
        run_btn.configure(text='停止接收' if running else '开始接收')
        url_var.set(inbox.url() if running else '未启动')

    def _toggle():
        if inbox.running:
            inbox.stop()
            _set_state(False)
            _refresh_qr()
            return
        try:
            inbox.start()
        except Exception as exc:
            messagebox.showerror('扫码收文件', '启动失败：\n%s' % exc,
                                 parent=win)
            return
        _set_state(True)
        _refresh_qr()

    def _choose_folder():
        path = filedialog.askdirectory(parent=win, title='选择保存位置',
                                       initialdir=inbox.folder)
        if not path:
            return
        was_running = inbox.running
        if was_running:
            inbox.stop()
        inbox.folder = path
        folder_var.set(path)
        if was_running:
            try:
                inbox.start()
            except Exception as exc:
                messagebox.showerror('扫码收文件', '重启失败：\n%s' % exc,
                                     parent=win)
        _set_state(inbox.running)
        _refresh_qr()

    def _open_folder():
        try:
            os.makedirs(inbox.folder, exist_ok=True)
            os.startfile(inbox.folder)      # noqa: S606  (Windows 打开目录)
        except Exception as exc:
            messagebox.showerror('扫码收文件', '打开失败：\n%s' % exc,
                                 parent=win)

    def _open_selected():
        sel = tree.selection()
        if not sel:
            return
        path = tree.item(sel[0], 'values')[0]
        full = os.path.join(inbox.folder, path)
        try:
            os.startfile(full)
        except Exception:
            _open_folder()

    def _copy_url():
        if not inbox.running:
            return
        try:
            win.clipboard_clear()
            win.clipboard_append(inbox.url())
            state_lbl.configure(text='接收中 · 网址已复制', fg='#2e8b57')
        except tk.TclError:
            pass

    def _save_qr():
        if not inbox.running:
            messagebox.showinfo('扫码收文件', '先点“开始接收”。', parent=win)
            return
        path = filedialog.asksaveasfilename(
            parent=win, title='保存二维码图片', defaultextension='.png',
            initialfile='扫码收文件二维码.png',
            filetypes=[('PNG 图片', '*.png')])
        if not path:
            return
        try:
            matrix = qrcode_lite.make_matrix(inbox.url())
            qrcode_lite.to_image(matrix, scale=8, border=4).save(path)
        except Exception as exc:
            messagebox.showerror('扫码收文件', '保存失败：\n%s' % exc,
                                 parent=win)

    def _clear():
        inbox.clear_list()
        for iid in tree.get_children():
            tree.delete(iid)
        keep['rows'] = 0

    def _sync_list():
        files = inbox.snapshot()
        if len(files) == keep['rows'] and files:
            return
        keep['rows'] = len(files)
        for iid in tree.get_children():
            tree.delete(iid)
        for item in files:
            tree.insert('', 'end', values=(
                item['name'], _size_text(item['size']), item['time'],
                item.get('who') or item.get('ip') or '-'))

    def _tick():
        if not win.winfo_exists():
            return
        try:
            _sync_list()
            if inbox.running:
                state_lbl.configure(
                    text='接收中 · 已收 %d 个文件' % len(inbox.snapshot()),
                    fg='#2e8b57')
        except tk.TclError:
            return
        try:
            win.after(800, _tick)
        except tk.TclError:
            pass

    def _on_close():
        try:
            inbox.stop()
        except Exception:
            pass
        try:
            win.destroy()
        except tk.TclError:
            pass

    win.protocol('WM_DELETE_WINDOW', _on_close)
    _boot.bind_close(win, root)

    _set_state(False)
    _toggle()          # 打开即开始接收，少一步操作
    _tick()
    win.deiconify()
    win.lift()


_boot.run('扫码收文件', _open)
