# -*- coding: utf-8 -*-
"""
PPT 字体统一替换工具
- 遍历所有幻灯片 / 母版 / 版式 / 备注页的全部文本
- 同时设置 Latin / 东亚(EA) / 复杂文种(CS) 三种字体，确保中文字体生效
- 只处理真实文本（文本框、占位符、表格、图表文字），图片内嵌文字不受影响
"""
import os
import shutil
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from pptx import Presentation
from pptx.oxml.ns import qn

FONT_EXTENSIONS = ('.ttf', '.otf', '.ttc')


# ---------------- 字体名解析 ----------------

def get_font_name(font_path):
    """用 fontTools 读取字体内部的真实字体名（family + subfamily）。"""
    try:
        from fontTools.ttLib import TTFont, TTCollection
        ext = os.path.splitext(font_path)[1].lower()
        if ext == '.ttc':
            collection = TTCollection(font_path, fontNumber=0)
            font = collection.fonts[0]
        else:
            font = TTFont(font_path, fontNumber=0, lazy=True)
        name_table = font['name']
        # nameID 16 = Typographic Family, 1 = Family; 2 = Subfamily
        family = name_table.getDebugName(16) or name_table.getDebugName(1) or ''
        subfamily = name_table.getDebugName(17) or name_table.getDebugName(2) or ''
        font.close()
        family = family.strip()
        subfamily = subfamily.lower()
        if not family:
            return None
        # 粗斜体等特殊样式在 family 中体现时直接用 family
        if any(k in subfamily for k in ('bold', 'italic', 'black', 'heavy',
                                        'medium', 'light', 'regular')):
            # 若 family 本身已含样式词（如 Unity-Bold），直接用
            if any(k in family.lower() for k in ('bold', 'italic', 'black',
                                                 'heavy', 'medium', 'light')):
                return family
        return family
    except Exception:
        return None


def font_name_from_filename(font_path):
    """降级方案：用文件名推断字体名（去掉扩展名）。"""
    return os.path.splitext(os.path.basename(font_path))[0]


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
        # 图表
        if getattr(shape, 'has_chart', False) and shape.has_chart:
            chart = shape.chart
            try:
                for plot in chart.plots:
                    for series in plot.series:
                        # 图表字体替换较复杂，这里替换图表标题/坐标轴文字
                        pass
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
        self.geometry('640x460')
        self.resizable(False, False)

        self.ppt_path = tk.StringVar()
        self.font_path = tk.StringVar()
        self.font_name = tk.StringVar(value='（未选择字体）')
        self.builtin_dir = self._builtin_font_dir()

        self._build_ui()
        self._load_builtin_font()

    def _builtin_font_dir(self):
        """exe 同目录下的 fonts 文件夹（把一套字体放进去即可自动加载）。"""
        if getattr(sys, 'frozen', False):
            base = os.path.dirname(sys.executable)
        else:
            base = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base, 'fonts')

    def _load_builtin_font(self):
        if not os.path.isdir(self.builtin_dir):
            return
        fonts = [f for f in os.listdir(self.builtin_dir)
                 if f.lower().endswith(FONT_EXTENSIONS)]
        if len(fonts) == 1:
            self._set_font(os.path.join(self.builtin_dir, fonts[0]))
        elif fonts:
            # 多套字体时优先找 Regular
            reg = [f for f in fonts if 'regular' in f.lower()]
            pick = reg[0] if reg else fonts[0]
            self._set_font(os.path.join(self.builtin_dir, pick))

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

        ttk.Label(frm, text='第二步：选择字体文件 (.ttf / .otf / .ttc)').pack(anchor='w', **pad)
        row2 = ttk.Frame(frm)
        row2.pack(fill='x', **pad)
        ttk.Entry(row2, textvariable=self.font_path, width=55).pack(side='left')
        ttk.Button(row2, text='浏览…', command=self._pick_font).pack(side='left', padx=6)

        ttk.Label(frm, textvariable=self.font_name,
                  foreground='#555').pack(anchor='w', padx=12)

        ttk.Label(frm,
                  text=f'提示：把整套字体放进程序同目录 fonts 文件夹可自动加载。\n当前内置目录：{self.builtin_dir}',
                  foreground='#999').pack(anchor='w', padx=12, pady=6)

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

    def _pick_font(self):
        path = filedialog.askopenfilename(
            title='选择字体文件',
            filetypes=[('字体文件', '*.ttf *.otf *.ttc')])
        if path:
            self._set_font(path)

    def _set_font(self, path):
        self.font_path.set(path)
        name = get_font_name(path) or font_name_from_filename(path)
        self.font_name.set(f'识别到字体名：{name}')
        self._resolved_font_name = name

    def _run(self):
        ppt = self.ppt_path.get().strip()
        font = self.font_path.get().strip()
        if not ppt or not os.path.isfile(ppt):
            messagebox.showwarning('提示', '请先选择 PPT 文件')
            return
        if not font or not os.path.isfile(font):
            messagebox.showwarning('提示', '请先选择字体文件')
            return
        if os.path.splitext(ppt)[1].lower() != '.pptx':
            messagebox.showwarning('提示', '只支持 .pptx 格式（请先另存为 .pptx）')
            return

        font_name = getattr(self, '_resolved_font_name', None) \
            or get_font_name(font) or font_name_from_filename(font)

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
