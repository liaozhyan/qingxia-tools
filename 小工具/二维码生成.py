# -*- coding: utf-8 -*-
"""二维码生成（独立小工具）：链接 / 文字 转二维码，可存成 PNG。

源码运行：python 小工具/二维码生成.py
"""

import _boot


def _open(root):
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    import ui_theme
    import qrcode_lite

    p = ui_theme.PALETTE
    win = tk.Toplevel(root)
    win.title('二维码生成')
    win.geometry('760x620')
    win.resizable(False, False)
    win.configure(bg=p['bg'])
    try:
        ui_theme.style_window(win)
    except Exception:
        pass
    _boot.bind_close(win, root)

    state = {'img': None, 'photo': None}

    body = tk.Frame(win, bg=p['bg'], padx=16, pady=14)
    body.pack(fill=tk.BOTH, expand=True)

    tk.Label(body, text='二维码生成', font=ui_theme.font(13, 'bold'),
             bg=p['bg'], fg=p['text']).pack(anchor='w')
    tk.Label(body, text='粘贴链接或文字后点“生成”。内容越长码越密，'
                        '超过约 134 字节会生成失败，建议用短链。',
             font=ui_theme.font(9), bg=p['bg'],
             fg=p['faint']).pack(anchor='w', pady=(2, 10))

    text = tk.Text(body, height=4, wrap='char', font=ui_theme.font(10),
                   relief=tk.FLAT, bd=0, highlightthickness=1,
                   highlightbackground=p['border'],
                   bg=p['surface'], fg=p['text'],
                   insertbackground=p['text'], padx=8, pady=6)
    text.pack(fill=tk.X)
    text.insert('1.0', 'https://')

    btns = tk.Frame(body, bg=p['bg'])
    btns.pack(fill=tk.X, pady=(10, 6))

    qr_box = tk.Frame(body, bg=p['surface'], highlightthickness=1,
                      highlightbackground=p['border'])
    qr_box.pack(fill=tk.BOTH, expand=True, pady=(4, 8))
    qr_label = tk.Label(qr_box, text='尚未生成', bg=p['surface'],
                        fg=p['faint'], font=ui_theme.font(10))
    qr_label.pack(expand=True)

    status = tk.Label(body, text='', font=ui_theme.font(9), bg=p['bg'],
                      fg=p['muted'], anchor='w')
    status.pack(fill=tk.X)

    def _render(_event=None):
        content = text.get('1.0', 'end').strip()
        if not content:
            messagebox.showinfo('二维码生成', '请先输入要生成的内容。',
                                parent=win)
            return
        try:
            matrix = qrcode_lite.make_matrix(content)
            img = qrcode_lite.to_image(matrix, scale=6, border=4)
            from PIL import ImageTk
            photo = ImageTk.PhotoImage(img)
        except Exception as exc:
            messagebox.showwarning(
                '二维码生成',
                '生成失败：\n%s\n\n内容太长时请缩短（约 134 字节以内），'
                '或改用短链。' % exc, parent=win)
            return
        state['img'] = img
        state['photo'] = photo          # 留引用，防被回收
        qr_label.configure(image=photo, text='')
        status.configure(text='已生成，可保存为 PNG 或直接截屏使用。')

    def _save():
        if state['img'] is None:
            messagebox.showinfo('二维码生成', '请先生成二维码。', parent=win)
            return
        path = filedialog.asksaveasfilename(
            parent=win, title='保存二维码图片', defaultextension='.png',
            initialfile='二维码.png', filetypes=[('PNG 图片', '*.png')])
        if not path:
            return
        try:
            state['img'].save(path)
        except Exception as exc:
            messagebox.showerror('二维码生成', '保存失败：\n%s' % exc,
                                 parent=win)
            return
        status.configure(text='已保存：%s' % path)

    ttk.Button(btns, text='生成二维码', style='Accent.TButton',
               command=_render).pack(side=tk.LEFT)
    ttk.Button(btns, text='保存图片', style='Ghost.TButton',
               command=_save).pack(side=tk.LEFT, padx=(8, 0))
    ttk.Button(btns, text='清空', style='Ghost.TButton',
               command=lambda: (text.delete('1.0', 'end'),
                                qr_label.configure(image='', text='尚未生成'),
                                status.configure(text=''))).pack(
        side=tk.LEFT, padx=(8, 0))

    text.bind('<Control-Return>', _render)
    _render()
    win.deiconify()
    win.lift()


_boot.run('二维码生成', _open)
