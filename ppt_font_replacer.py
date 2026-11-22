# -*- coding: utf-8 -*-
"""
PPT 字体统一替换工具
- 遍历所有幻灯片 / 母版 / 版式 / 备注页的全部文本
- 同时设置 Latin / 东亚(EA) / 复杂文种(CS) 三种字体，确保中文字体生效
- 只处理真实文本（文本框、占位符、表格、图表文字），图片内嵌文字不受影响
- 字体来源：选择一个字体文件夹，读取其中全部字体文件（须为同一套字体）
"""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from pptx import Presentation
from pptx.oxml.ns import qn

FONT_EXTENSIONS = ('.ttf', '.otf', '.ttc')


# ---------------- 字体名解析 ----------------

def get_font_family(font_path):
    """用 fontTools 读取字体内部的家族名（family name）。"""
    try:
        from fontTools.ttLib import TTFont, TTCollection
        ext = os.path.splitext(font_path)[1].lower()
        if ext == '.ttc':
            collection = TTCollection(font_path, fontNumber=0)
            font = collection.fonts[0]
        else:
            font = TTFont(font_path, fontNumber=0, lazy=True)
        name_table = font['name']
        # nameID 16 = Typographic Family, 1 = Family
        family = name_table.getDebugName(16) or name_table.getDebugName(1) or ''
        font.close()
        return family.strip() or None
    except Exception:
        return None


def scan_font_folder(folder):
    """
    扫描字体文件夹，返回 (font_name, file_count, error)。
    - 找出文件夹下全部字体文件，解析其家族名
    - 所有字体必须属于同一家族（Regular/Bold/Italic 等字重视为一套）
    - 出现多个不同家族时报错
    """
    files = [f for f in sorted(os.listdir(folder))
             if f.lower().endswith(FONT_EXTENSIONS)]
    if not files:
        return None, 0, '该文件夹下没有找到字体文件（.ttf / .otf / .ttc）'

    families = {}
    for f in files:
        fam = get_font_family(os.path.join(folder, f)) \
            or os.path.splitext(f)[0]
        families.setdefault(fam, []).append(f)

    if len(families) > 1:
        names = '、'.join(sorted(families.keys()))
        detail = '\n'.join(
            f'  {fam}：{len(fs)} 个文件' for fam, fs in sorted(families.items()))
        return None, len(files), (
            f'检测到多套不同字体：{names}\n\n{detail}\n\n'
            '请简化字体文件夹，只保留一套字体（该套字体的 '
            'Regular / Bold / Italic 等字重可以都保留），然后重新开始。')

    font_name = next(iter(families.keys()))
    return font_name, len(files), None


# ---------------- 字体替换核心 ----------------

def set_run_font(run, font_name):
    """同时设置 latin / ea / cs 字体，保证中英文都生效。"""
    rPr = run._r.get_or_add_rPr()
    for tag in ('latin', 'ea', 'cs'):
        el = rPr.find(qn('a:' + tag))
        if el is None:
            el = rPr.makeelement(qn('a:' + tag), {})
            rPr.append(el)
        el.set('typeface', font_name)


def replace_in_text_frame(tf, font_name, stats):
    for paragraph in tf.paragraphs:
        for run in paragraph.runs:
            set_run_font(run, font_name)
            stats['runs'] += 1
        # 段落末尾的空 run / endParaRPr
        end = paragraph._p.find(qn('a:endParaRPr'))
        if end is not None:
            for tag in ('latin', 'ea', 'cs'):
                el = end.find(qn('a:' + tag))
                if el is None:
                    el = end.makeelement(qn('a:' + tag), {})
                    end.append(el)
                el.set('typeface', font_name)


def walk_shapes(shapes, font_name, stats):
    for shape in shapes:
        # 组合图形递归
        if shape.shape_type == 6:  # GROUP
            walk_shapes(shape.shapes, font_name, stats)
            continue
        # 表格
        if getattr(shape, 'has_table', False) and shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    replace_in_text_frame(cell.text_frame, font_name, stats)
        # 图表标题
        if getattr(shape, 'has_chart', False) and shape.has_chart:
            chart = shape.chart
            try:
                if chart.has_title:
                    replace_in_text_frame(chart.chart_title.text_frame, font_name, stats)
            except Exception:
                pass
        # 文本框 / 占位符
        if getattr(shape, 'has_text_frame', False) and shape.has_text_frame:
            replace_in_text_frame(shape.text_frame, font_name, stats)


def replace_ppt_fonts(input_path, output_path, font_name):
    prs = Presentation(input_path)
    stats = {'runs': 0}

    # 幻灯片
    for slide in prs.slides:
        walk_shapes(slide.shapes, font_name, stats)
        # 备注页
        if slide.has_notes_slide:
            walk_shapes(slide.notes_slide.shapes, font_name, stats)
    # 母版与版式（覆盖背景文字框）
    for master in prs.slide_masters:
        walk_shapes(master.shapes, font_name, stats)
        for layout in master.slide_layouts:
            walk_shapes(layout.shapes, font_name, stats)

    prs.save(output_path)
    return stats


# ---------------- GUI ----------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('PPT 字体统一替换工具')
        self.geometry('640x440')
        self.resizable(False, False)

        self.ppt_path = tk.StringVar()
        self.font_dir = tk.StringVar()
        self.font_info = tk.StringVar(value='（未选择字体文件夹）')
        self._resolved_font_name = None

        self._build_ui()

    def _build_ui(self):
        pad = {'padx': 12, 'pady': 6}
        frm = ttk.Frame(self)
        frm.pack(fill='both', expand=True)

        ttk.Label(frm, text='第一步：选择 PPT 文件 (.pptx)').pack(anchor='w', **pad)
        row1 = ttk.Frame(frm)
        row1.pack(fill='x', **pad)
        ttk.Entry(row1, textvariable=self.ppt_path, width=55).pack(side='left')
        ttk.Button(row1, text='浏览…', command=self._pick_ppt).pack(side='left', padx=6)

        ttk.Separator(frm).pack(fill='x', pady=8)

        ttk.Label(frm, text='第二步：选择字体文件夹（放入一套字体的全部字重）').pack(anchor='w', **pad)
        row2 = ttk.Frame(frm)
        row2.pack(fill='x', **pad)
        ttk.Entry(row2, textvariable=self.font_dir, width=55).pack(side='left')
        ttk.Button(row2, text='浏览…', command=self._pick_font_dir).pack(side='left', padx=6)

        ttk.Label(frm, textvariable=self.font_info,
                  foreground='#555', wraplength=580).pack(anchor='w', padx=12)

        ttk.Separator(frm).pack(fill='x', pady=8)

        self.progress = ttk.Progressbar(frm, mode='determinate', length=580)
        self.progress.pack(padx=12, pady=6)

        ttk.Button(frm, text='开始统一替换字体',
                   command=self._run).pack(pady=12)

        self.log = tk.Text(frm, height=6, bg='#f7f7f7')
        self.log.pack(fill='x', padx=12, pady=6)

    def _log(self, msg):
        self.log.insert('end', msg + '\n')
        self.log.see('end')
        self.update_idletasks()

    def _pick_ppt(self):
        path = filedialog.askopenfilename(
            title='选择 PPT 文件',
            filetypes=[('PowerPoint 文件', '*.pptx')])
        if path:
            self.ppt_path.set(path)

    def _pick_font_dir(self):
        folder = filedialog.askdirectory(title='选择字体文件夹（一套字体）')
        if not folder:
            return
        font_name, count, err = scan_font_folder(folder)
        if err:
            self.font_dir.set('')
            self._resolved_font_name = None
            self.font_info.set('（未选择字体文件夹）')
            messagebox.showwarning('字体文件夹有问题', err)
            return
        self.font_dir.set(folder)
        self._resolved_font_name = font_name
        self.font_info.set(
            f'识别到字体：{font_name}（共 {count} 个字体文件，视为一套字体）')

    def _run(self):
        ppt = self.ppt_path.get().strip()
        font_dir = self.font_dir.get().strip()
        if not ppt or not os.path.isfile(ppt):
            messagebox.showwarning('提示', '请先选择 PPT 文件')
            return
        if not font_dir or not os.path.isdir(font_dir):
            messagebox.showwarning('提示', '请先选择字体文件夹')
            return
        if os.path.splitext(ppt)[1].lower() != '.pptx':
            messagebox.showwarning('提示', '只支持 .pptx 格式（请先另存为 .pptx）')
            return

        font_name = self._resolved_font_name
        if not font_name:
            font_name, _, err = scan_font_folder(font_dir)
            if err:
                messagebox.showwarning('字体文件夹有问题', err)
                return

        out_dir = os.path.dirname(ppt)
        base = os.path.splitext(os.path.basename(ppt))[0]
        out_path = os.path.join(out_dir, f'{base}_统一字体.pptx')

        self.progress['value'] = 30
        self._log(f'开始处理：{os.path.basename(ppt)}')
        self._log(f'目标字体：{font_name}')
        try:
            stats = replace_ppt_fonts(ppt, out_path, font_name)
            self.progress['value'] = 100
            self._log(f'完成！共替换 {stats["runs"]} 个文字片段。')
            self._log(f'输出文件：{out_path}')
            messagebox.showinfo('完成',
                                f'替换完成，共处理 {stats["runs"]} 个文字片段。\n\n{out_path}')
        except Exception as e:
            self.progress['value'] = 0
            self._log(f'出错：{e}')
            messagebox.showerror('错误', str(e))


if __name__ == '__main__':
    app = App()
    app.mainloop()
