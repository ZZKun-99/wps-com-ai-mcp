"""
WPS人机协作编辑 MCP 服务器 (wps-editor-mcp)
兼容Python 3.8+，不依赖官方MCP SDK

架构（基线 = 归档/v3，当前运行时版本 3.4.0）：
  server.py       地基层：COM 多级挂接链、僵尸实例识别、人编辑态救援、24 个老工具、JSON-RPC 主循环
  wps_session.py  薄壳层：默认只向 AI 广播 wps_docs / wps_do 两个极简工具，操作"人正在打开的活文档"
  详见 AGENTS.md（改动前必读）与 docs/现状报告.md

MCP协议实现基于JSON-RPC 2.0 over stdio
"""

__version__ = "3.4.0"

import os
import sys
import json
import time
from enum import Enum
from typing import Optional, Any, Dict, List, Tuple
from pathlib import Path


# ==================== stdio 编码加固（UTF-8 强制） ====================
# MCP 通过 stdio/管道通信，若宿主默认编码非 UTF-8（如 Windows 下 GBK/cp936），
# 中文请求/响应会出现乱码。这里把 stdin/stdout 尽可能强制为 UTF-8，
# 旧解释器（无 TextIOWrapper.reconfigure）自动跳过。
# 注意：本函数不接管写入逻辑，main() 中"每次响应后 sys.stdout.flush()"的行为保持不变。
def _force_utf8_stdio() -> None:
    """将 stdin/stdout 尽可能切换为 UTF-8（errors="replace"）。"""
    for stream in (sys.stdin, sys.stdout):
        if stream is None:
            continue
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except AttributeError:
            # Python < 3.7：TextIOWrapper 无 reconfigure，保持原样
            continue
        except (ValueError, OSError):
            # 流已被关闭或不可重配置（特殊重定向等），保持原样
            continue


_force_utf8_stdio()


class EditMode(str, Enum):
    FILE = "file"    # 直接文件解析
    COM = "com"      # WPS COM自动化


def get_mode(mode: Optional[str] = None) -> EditMode:
    """获取编辑模式，默认为file模式"""
    if mode is None:
        return EditMode.FILE
    # 容错大小写与空格：模型偶发传 "COM" / " com "，枚举值是大小写敏感的小写。
    return EditMode(str(mode).strip().lower())


# ==================== 文件格式解析模式 ====================

def _file_read_docx(file_path: str) -> dict:
    """读取Word文档内容"""
    from docx import Document
    doc = Document(file_path)
    paragraphs = [p.text for p in doc.paragraphs]
    tables = []
    for table in doc.tables:
        table_data = []
        for row in table.rows:
            table_data.append([cell.text for cell in row.cells])
        tables.append(table_data)
    return {
        "paragraphs": paragraphs,
        "tables": tables,
        "paragraph_count": len(paragraphs),
        "table_count": len(tables)
    }


def _file_write_docx(file_path: str, content: str, append: bool = False) -> str:
    """写入Word文档"""
    from docx import Document
    if append and os.path.exists(file_path):
        doc = Document(file_path)
    else:
        doc = Document()
    
    # 按换行符分割，每一行作为一个段落
    lines = content.split('\n')
    for line in lines:
        if line.strip():  # 非空行
            doc.add_paragraph(line)
        else:  # 空行也添加空段落
            doc.add_paragraph('')
    
    doc.save(file_path)
    return f"已写入: {file_path}"


def _file_read_xlsx(file_path: str, sheet_name: Optional[str] = None) -> dict:
    """读取Excel文件内容"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
    else:
        ws = wb.active
    data = []
    for row in ws.iter_rows(values_only=True):
        data.append(list(row))
    return {
        "sheet_name": ws.title,
        "data": data,
        "rows": ws.max_row,
        "columns": ws.max_column
    }


def _file_write_xlsx(file_path: str, sheet_name: Optional[str], data: list) -> str:
    """写入Excel文件"""
    from openpyxl import Workbook, load_workbook
    if os.path.exists(file_path):
        wb = load_workbook(file_path)
    else:
        wb = Workbook()
    if sheet_name:
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.create_sheet(sheet_name)
    else:
        ws = wb.active
    for row_idx, row_data in enumerate(data, 1):
        for col_idx, value in enumerate(row_data, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    wb.save(file_path)
    return f"已写入: {file_path}"


def _file_read_pptx(file_path: str) -> dict:
    """读取PowerPoint文件内容"""
    from pptx import Presentation
    prs = Presentation(file_path)
    slides_data = []
    for slide in prs.slides:
        slide_content = []
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                slide_content.append(shape.text)
        slides_data.append(slide_content)
    return {
        "slide_count": len(prs.slides),
        "slides": slides_data
    }


def _file_write_pptx(file_path: str, content: str) -> str:
    """写入PowerPoint文件"""
    from pptx import Presentation
    from pptx.util import Inches
    if os.path.exists(file_path):
        prs = Presentation(file_path)
    else:
        prs = Presentation()
    slide_layout = prs.slide_layouts[6]  # 空白布局
    slide = prs.slides.add_slide(slide_layout)
    left = Inches(1)
    top = Inches(1)
    width = Inches(8)
    height = Inches(1)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    textbox.text = content
    prs.save(file_path)
    return f"已写入: {file_path}"


# ==================== Excel表格操作 ====================

def _file_add_rows(file_path: str, sheet_name: Optional[str], count: int, after_row: Optional[int] = None) -> str:
    """添加行"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    if after_row is None:
        after_row = ws.max_row
    ws.insert_rows(after_row + 1, count)
    wb.save(file_path)
    return f"已在第{after_row}行后添加{count}行"


def _file_add_columns(file_path: str, sheet_name: Optional[str], count: int, after_col: Optional[int] = None) -> str:
    """添加列"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    if after_col is None:
        after_col = ws.max_column
    ws.insert_cols(after_col + 1, count)
    wb.save(file_path)
    return f"已在第{after_col}列后添加{count}列"


def _file_merge_cells(file_path: str, sheet_name: Optional[str], start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    """合并单元格"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    ws.merge_cells(start_row=start_row, start_column=start_col, end_row=end_row, end_column=end_col)
    wb.save(file_path)
    return f"已合并单元格: ({start_row},{start_col}) -> ({end_row},{end_col})"


def _file_unmerge_cells(file_path: str, sheet_name: Optional[str], start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    """取消合并单元格"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    ws.unmerge_cells(start_row=start_row, start_column=start_col, end_row=end_row, end_column=end_col)
    wb.save(file_path)
    return f"已取消合并: ({start_row},{start_col}) -> ({end_row},{end_col})"


def _file_set_column_width(file_path: str, sheet_name: Optional[str], column: str, width: float) -> str:
    """设置列宽"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    ws.column_dimensions[column].width = width
    wb.save(file_path)
    return f"已设置列{column}宽度为{width}"


def _file_set_row_height(file_path: str, sheet_name: Optional[str], row: int, height: float) -> str:
    """设置行高"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    ws.row_dimensions[row].height = height
    wb.save(file_path)
    return f"已设置行{row}高度为{height}"


# ==================== 图片插入 ====================

def _file_insert_image_docx(file_path: str, image_path: str, width: Optional[float] = None, height: Optional[float] = None) -> str:
    """在Word中插入图片"""
    from docx import Document
    from docx.shared import Inches
    doc = Document(file_path) if os.path.exists(file_path) else Document()
    if width and height:
        doc.add_picture(image_path, width=Inches(width), height=Inches(height))
    elif width:
        doc.add_picture(image_path, width=Inches(width))
    else:
        doc.add_picture(image_path)
    doc.save(file_path)
    return f"已插入图片: {image_path}"


def _file_insert_image_xlsx(file_path: str, sheet_name: Optional[str], image_path: str, anchor: str = "A1") -> str:
    """在Excel中插入图片"""
    from openpyxl import load_workbook
    from openpyxl.drawing.image import Image
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    img = Image(image_path)
    ws.add_image(img, anchor)
    wb.save(file_path)
    return f"已插入图片到{anchor}"


def _file_insert_image_pptx(file_path: str, image_path: str, slide_index: int = -1, left: float = 1, top: float = 1, width: float = 4) -> str:
    """在PowerPoint中插入图片"""
    from pptx import Presentation
    from pptx.util import Inches
    prs = Presentation(file_path) if os.path.exists(file_path) else Presentation()
    if slide_index == -1:
        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)
    else:
        slide = prs.slides[slide_index]
    slide.shapes.add_picture(image_path, Inches(left), Inches(top), Inches(width))
    prs.save(file_path)
    return f"已插入图片: {image_path}"


# ==================== 文档转换 ====================

def _set_slide_bg(slide, color_hex):
    """设置幻灯片背景颜色"""
    from pptx.dml.color import RGBColor
    
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(color_hex)


def _add_decorative_shapes(slide, accent_color_hex, slide_type="content"):
    """添加装饰形状"""
    from pptx.util import Inches
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.dml.color import RGBColor
    
    accent_color = RGBColor.from_string(accent_color_hex)
    
    if slide_type == "title":
        # Top accent bar
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), 
                                        Inches(13.333), Inches(0.12))
        shape.fill.solid()
        shape.fill.fore_color.rgb = accent_color
        shape.line.fill.background()
        
        # Bottom accent bar
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.2), 
                                        Inches(13.333), Inches(0.3))
        shape.fill.solid()
        shape.fill.fore_color.rgb = accent_color
        shape.line.fill.background()
        
        # Left decorative bar
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(2.8), 
                                        Inches(0.06), Inches(2))
        shape.fill.solid()
        shape.fill.fore_color.rgb = accent_color
        shape.line.fill.background()
    
    elif slide_type == "content":
        # Top thin line
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.35), 
                                        Inches(11.7), Inches(0.025))
        shape.fill.solid()
        shape.fill.fore_color.rgb = accent_color
        shape.line.fill.background()
        
        # Bottom bar
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.2), 
                                        Inches(13.333), Inches(0.3))
        shape.fill.solid()
        shape.fill.fore_color.rgb = accent_color
        shape.line.fill.background()
        
        # Accent dot
        shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.6), Inches(0.45), 
                                        Inches(0.12), Inches(0.12))
        shape.fill.solid()
        shape.fill.fore_color.rgb = accent_color
        shape.line.fill.background()


def _add_page_number(slide, page_num, total_pages, color_hex):
    """添加页码"""
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    
    left = Inches(12)
    top = Inches(7.22)
    width = Inches(1)
    height = Inches(0.25)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    p = tf.paragraphs[0]
    p.text = f"{page_num} / {total_pages}"
    p.font.size = Pt(9)
    p.font.color.rgb = RGBColor.from_string(color_hex)
    p.alignment = 2  # Right


def _doc_to_ppt(doc_path: str, ppt_path: str, template: str = "ocean") -> dict:
    """根据Word文档生成专业PPT"""
    from docx import Document
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.shapes import MSO_SHAPE
    
    # Read Word document
    doc = Document(doc_path)
    
    # Extract structure
    title = ""
    subtitle = ""
    sections = []
    current_section = None
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        
        if not title:
            title = text
            continue
        
        if not subtitle and len(text) < 100:
            subtitle = text
            continue
        
        is_header = False
        if para.runs and para.runs[0].font.bold:
            is_header = True
        elif text and text[0].isdigit() and '. ' in text[:5]:
            is_header = True
        elif text.startswith('关键词') or text.startswith('Keywords'):
            is_header = True
        
        if is_header:
            current_section = {"title": text, "content": []}
            sections.append(current_section)
        elif current_section:
            current_section["content"].append(text)
        else:
            sections.append({"title": "", "content": [text]})
    
    # Create PowerPoint
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    # Professional color themes
    themes = {
        "ocean": {
            "bg": "0A1628", "accent": "00B4D8", "accent2": "0077B6",
            "title": "FFFFFF", "text": "CAF0F8", "light": "90E0EF",
            "card": "1A2D4A"
        },
        "forest": {
            "bg": "1B2D1B", "accent": "52B788", "accent2": "40916C",
            "title": "FFFFFF", "text": "D8F3DC", "light": "95D5B2",
            "card": "2D4A2D"
        },
        "sunset": {
            "bg": "2D1B2E", "accent": "FF6B6B", "accent2": "EE5A24",
            "title": "FFFFFF", "text": "FFEAA7", "light": "FFB8B8",
            "card": "4A2D4A"
        },
        "royal": {
            "bg": "1A1A2E", "accent": "E94560", "accent2": "0F3460",
            "title": "FFFFFF", "text": "EAEAEA", "light": "A8A8A8",
            "card": "16213E"
        },
        "minimal": {
            "bg": "F8F9FA", "accent": "2D3436", "accent2": "636E72",
            "title": "2D3436", "text": "636E72", "light": "B2BEC3",
            "card": "FFFFFF"
        },
        "blue": {
            "bg": "0D1B2A", "accent": "00BFFF", "accent2": "1E90FF",
            "title": "FFFFFF", "text": "E0F0FF", "light": "87CEEB",
            "card": "1B2838"
        },
        "green": {
            "bg": "0D2818", "accent": "00FF7F", "accent2": "32CD32",
            "title": "FFFFFF", "text": "E0FFE0", "light": "90EE90",
            "card": "1A3D1A"
        }
    }
    
    theme = themes.get(template, themes["ocean"])
    slide_layout = prs.slide_layouts[6]  # Blank
    total_slides = len(sections) + 3  # Title + TOC + Content + Thank You
    
    # ============ SLIDE 1: Title ============
    slide = prs.slides.add_slide(slide_layout)
    _set_slide_bg(slide, theme["bg"])
    _add_decorative_shapes(slide, theme["accent"], "title")
    
    # Main title
    left, top = Inches(1.5), Inches(2.5)
    width, height = Inches(10), Inches(1.8)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = RGBColor.from_string(theme["title"])
    p.alignment = PP_ALIGN.LEFT
    
    # Subtitle
    if subtitle:
        left, top = Inches(1.5), Inches(4.5)
        width, height = Inches(10), Inches(1.2)
        textbox = slide.shapes.add_textbox(left, top, width, height)
        tf = textbox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        if len(subtitle) > 150:
            subtitle = subtitle[:150] + "..."
        p.text = subtitle
        p.font.size = Pt(20)
        p.font.color.rgb = RGBColor.from_string(theme["light"])
        p.alignment = PP_ALIGN.LEFT
    
    # ============ SLIDE 2: Table of Contents ============
    slide = prs.slides.add_slide(slide_layout)
    _set_slide_bg(slide, theme["bg"])
    _add_decorative_shapes(slide, theme["accent"], "content")
    _add_page_number(slide, 2, total_slides, theme["light"])
    
    # TOC Title
    left, top = Inches(1), Inches(0.8)
    width, height = Inches(11), Inches(0.7)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    p = tf.paragraphs[0]
    p.text = "CONTENTS"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = RGBColor.from_string(theme["accent"])
    
    # TOC items
    left, top = Inches(1), Inches(1.8)
    width, height = Inches(11), Inches(5)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    tf.word_wrap = True
    
    toc_items = [s for s in sections if s["title"]][:8]
    for i, section in enumerate(toc_items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        
        run1 = p.add_run()
        run1.text = f"  {i+1:02d}  "
        run1.font.size = Pt(26)
        run1.font.bold = True
        run1.font.color.rgb = RGBColor.from_string(theme["accent"])
        
        run2 = p.add_run()
        title_text = section["title"]
        if len(title_text) > 50:
            title_text = title_text[:50] + "..."
        run2.text = title_text
        run2.font.size = Pt(20)
        run2.font.color.rgb = RGBColor.from_string(theme["text"])
        
        p.space_after = Pt(14)
    
    # ============ CONTENT SLIDES ============
    for i, section in enumerate(sections):
        if not section["title"] and not section["content"]:
            continue
        
        slide_num = i + 3
        slide = prs.slides.add_slide(slide_layout)
        _set_slide_bg(slide, theme["bg"])
        _add_decorative_shapes(slide, theme["accent"], "content")
        _add_page_number(slide, slide_num, total_slides, theme["light"])
        
        # Section number badge
        if section["title"]:
            left, top = Inches(0.6), Inches(0.5)
            width, height = Inches(0.9), Inches(0.9)
            shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor.from_string(theme["accent"])
            shape.line.fill.background()
            
            textbox = slide.shapes.add_textbox(left, top, width, height)
            tf = textbox.text_frame
            p = tf.paragraphs[0]
            p.text = f"{i+1:02d}"
            p.font.size = Pt(28)
            p.font.bold = True
            p.font.color.rgb = RGBColor.from_string(theme["bg"])
            p.alignment = PP_ALIGN.CENTER
            p.space_before = Pt(12)
            
            # Section title
            left, top = Inches(1.8), Inches(0.6)
            width, height = Inches(10), Inches(0.8)
            textbox = slide.shapes.add_textbox(left, top, width, height)
            tf = textbox.text_frame
            p = tf.paragraphs[0]
            title_text = section["title"]
            if len(title_text) > 60:
                title_text = title_text[:60] + "..."
            p.text = title_text
            p.font.size = Pt(30)
            p.font.bold = True
            p.font.color.rgb = RGBColor.from_string(theme["title"])
        
        # Content card
        content_top = Inches(1.8) if section["title"] else Inches(0.5)
        
        left, top = Inches(0.6), content_top
        width, height = Inches(12), Inches(5.1)
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(theme["card"])
        shape.line.color.rgb = RGBColor.from_string(theme["accent"])
        shape.line.width = Pt(1)
        
        # Content bullets
        left, top = Inches(1), content_top + Inches(0.3)
        width, height = Inches(11.2), Inches(4.6)
        textbox = slide.shapes.add_textbox(left, top, width, height)
        tf = textbox.text_frame
        tf.word_wrap = True
        
        for j, line in enumerate(section["content"][:5]):
            if j == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            
            run1 = p.add_run()
            run1.text = "  ●  "
            run1.font.size = Pt(12)
            run1.font.color.rgb = RGBColor.from_string(theme["accent"])
            
            run2 = p.add_run()
            if len(line) > 120:
                line = line[:120] + "..."
            run2.text = line
            run2.font.size = Pt(17)
            run2.font.color.rgb = RGBColor.from_string(theme["text"])
            
            p.space_after = Pt(12)
    
    # ============ FINAL SLIDE: Thank You ============
    slide = prs.slides.add_slide(slide_layout)
    _set_slide_bg(slide, theme["bg"])
    _add_decorative_shapes(slide, theme["accent"], "title")
    
    left, top = Inches(2), Inches(2.5)
    width, height = Inches(9), Inches(1.5)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    p = tf.paragraphs[0]
    p.text = "THANK YOU"
    p.font.size = Pt(56)
    p.font.bold = True
    p.font.color.rgb = RGBColor.from_string(theme["title"])
    p.alignment = PP_ALIGN.CENTER
    
    left, top = Inches(3), Inches(4.2)
    width, height = Inches(7), Inches(0.8)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    p = tf.paragraphs[0]
    p.text = "感谢聆听 | Questions & Discussion"
    p.font.size = Pt(22)
    p.font.color.rgb = RGBColor.from_string(theme["light"])
    p.alignment = PP_ALIGN.CENTER
    
    # Decorative line
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5), Inches(5.3), 
                                    Inches(3), Inches(0.04))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(theme["accent"])
    shape.line.fill.background()
    
    # Save
    prs.save(ppt_path)
    
    return {
        "title": title,
        "slides_count": len(prs.slides),
        "sections_count": len(sections),
        "template": template,
        "output": ppt_path
    }


def _ppt_to_doc(ppt_path: str, doc_path: str) -> dict:
    """读取PPT生成Word总结文档"""
    from pptx import Presentation
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    
    prs = Presentation(ppt_path)
    
    doc = Document()
    
    # Title
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_para.add_run("Presentation Summary")
    run.font.size = Pt(28)
    run.font.bold = True
    
    doc.add_paragraph()
    
    # Metadata
    meta_para = doc.add_paragraph()
    meta_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = meta_para.add_run(f"Total Slides: {len(prs.slides)}")
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(100, 100, 100)
    
    doc.add_paragraph()
    doc.add_paragraph("=" * 50)
    doc.add_paragraph()
    
    # Extract content from each slide
    all_content = []
    
    for i, slide in enumerate(prs.slides):
        slide_content = {"index": i + 1, "texts": []}
        
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    if text:
                        slide_content["texts"].append(text)
        
        if slide_content["texts"]:
            all_content.append(slide_content)
    
    # Generate summary structure
    # Find title (usually first text of first slide)
    main_title = ""
    if all_content and all_content[0]["texts"]:
        main_title = all_content[0]["texts"][0]
    
    if main_title:
        h1 = doc.add_heading(main_title, level=1)
    
    doc.add_heading("Overview", level=2)
    doc.add_paragraph(f"This presentation contains {len(prs.slides)} slides covering the following topics:")
    
    # List all section titles
    doc.add_heading("Table of Contents", level=2)
    for i, slide_data in enumerate(all_content):
        if slide_data["texts"]:
            first_text = slide_data["texts"][0]
            if len(first_text) < 80:  # Likely a title
                doc.add_paragraph(f"Slide {slide_data['index']}: {first_text}", style='List Number')
    
    doc.add_paragraph()
    
    # Detailed content
    doc.add_heading("Detailed Content", level=2)
    
    for slide_data in all_content:
        if not slide_data["texts"]:
            continue
        
        # First text as slide heading
        slide_title = slide_data["texts"][0]
        doc.add_heading(f"Slide {slide_data['index']}: {slide_title}", level=3)
        
        # Rest as bullet points
        for text in slide_data["texts"][1:]:
            if len(text) > 5:  # Skip very short texts
                doc.add_paragraph(text, style='List Bullet')
        
        doc.add_paragraph()
    
    # Conclusion
    doc.add_heading("Key Takeaways", level=2)
    doc.add_paragraph("Based on the presentation content, the main points are:")
    
    # Extract key points from content
    key_points = []
    for slide_data in all_content:
        for text in slide_data["texts"]:
            if len(text) > 20 and len(text) < 150:
                key_points.append(text)
                if len(key_points) >= 5:
                    break
        if len(key_points) >= 5:
            break
    
    for point in key_points[:5]:
        doc.add_paragraph(point, style='List Bullet')
    
    # Save
    doc.save(doc_path)
    
    return {
        "slides_processed": len(all_content),
        "output": doc_path,
        "main_title": main_title
    }


# ==================== 批量操作 ====================

def _batch_read_docx(file_path: str, para_indices: list) -> list:
    """批量读取Word段落"""
    from docx import Document
    doc = Document(file_path)
    results = []
    for idx in para_indices:
        if 0 <= idx < len(doc.paragraphs):
            para = doc.paragraphs[idx]
            results.append({
                "index": idx,
                "text": para.text,
                "style": para.style.name if para.style else None
            })
        else:
            results.append({"index": idx, "error": "Index out of range"})
    return results


def _batch_read_xlsx(file_path: str, cells: list, sheet_name: str = None) -> list:
    """批量读取Excel单元格"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    results = []
    for cell_ref in cells:
        try:
            cell = ws[cell_ref]
            results.append({
                "cell": cell_ref,
                "value": cell.value,
                "type": type(cell.value).__name__ if cell.value is not None else "null"
            })
        except Exception as e:
            results.append({"cell": cell_ref, "error": str(e)})
    return results


def _batch_write_xlsx(file_path: str, data: list, sheet_name: str = None) -> str:
    """批量写入Excel单元格
    data格式: [{"cell": "A1", "value": "xxx"}, ...]
    """
    skipped = []          # V3.1：跳过的项（原来裸 except 把失败悄悄吞掉）
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    count = 0
    for item in data:
        try:
            cell_ref = item.get("cell")
            value = item.get("value")
            if cell_ref and value is not None:
                ws[cell_ref] = value
                count += 1
        except Exception as _err:      # V3.1：不再裸 except，至少把原因写出来
            print("[wps-editor] batch_write 跳过一项：" + repr(_err), file=sys.stderr)
            sys.stderr.flush()
            skipped.append(_err)
    wb.save(file_path)
    if skipped:
        return ("已批量写入%d个单元格；跳过%d项（原因见 stderr，多为 cell 地址非法）"
                % (count, len(skipped)))
    return f"已批量写入{count}个单元格"


def _batch_set_style_docx(file_path: str, operations: list) -> str:
    """批量设置Word段落样式
    operations格式: [{"para_index": 0, "font_name": "SimSun", "bold": true}, ...]
    """
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    from lxml import etree
    
    doc = Document(file_path)
    count = 0
    
    for op in operations:
        para_index = op.get("para_index")
        if para_index is None or para_index >= len(doc.paragraphs):
            continue
        
        para = doc.paragraphs[para_index]
        
        for run in para.runs:
            if "font_name" in op:
                rPr = run._element.get_or_add_rPr()
                for rf in rPr.findall(qn('w:rFonts')):
                    rPr.remove(rf)
                rFonts = etree.SubElement(rPr, qn('w:rFonts'))
                rFonts.set(qn('w:ascii'), op["font_name"])
                rFonts.set(qn('w:hAnsi'), op["font_name"])
                rFonts.set(qn('w:eastAsia'), op["font_name"])
                rFonts.set(qn('w:cs'), op["font_name"])
                rFonts.set(qn('w:hint'), 'eastAsia')
            
            if "font_size" in op:
                run.font.size = Pt(op["font_size"])
            if "bold" in op:
                run.font.bold = op["bold"]
            if "italic" in op:
                run.font.italic = op["italic"]
            if "underline" in op:
                run.font.underline = op["underline"]
            if "color" in op:
                r, g, b = int(op["color"][0:2], 16), int(op["color"][2:4], 16), int(op["color"][4:6], 16)
                run.font.color.rgb = RGBColor(r, g, b)
        
        count += 1
    
    doc.save(file_path)
    return f"已批量设置{count}个段落样式"


def _batch_set_style_xlsx(file_path: str, operations: list) -> str:
    """批量设置Excel单元格样式
    operations格式: [{"cell": "A1", "font_name": "SimSun", "bold": true}, ...]
    """
    skipped = []          # V3.1：跳过的项
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Color
    
    wb = load_workbook(file_path)
    ws = wb.active
    count = 0
    
    for op in operations:
        cell_ref = op.get("cell")
        if not cell_ref:
            continue
        
        try:
            cell = ws[cell_ref]
            
            font_kwargs = {}
            if "font_name" in op:
                font_kwargs['name'] = op["font_name"]
            if "font_size" in op:
                font_kwargs['size'] = op["font_size"]
            if "bold" in op:
                font_kwargs['bold'] = op["bold"]
            if "italic" in op:
                font_kwargs['italic'] = op["italic"]
            if "underline" in op and op["underline"]:
                font_kwargs['underline'] = 'single'
            if "color" in op:
                font_kwargs['color'] = Color(rgb=op["color"])
            
            if font_kwargs:
                cell.font = Font(**font_kwargs)
            
            if "bg_color" in op:
                cell.fill = PatternFill(start_color=op["bg_color"], end_color=op["bg_color"], fill_type="solid")
            
            count += 1
        except Exception as _err:      # V3.1：不再裸 except，至少把原因写出来
            print("[wps-editor] batch_set_style 跳过一项：" + repr(_err), file=sys.stderr)
            sys.stderr.flush()
            skipped.append(_err)
    
    wb.save(file_path)
    if skipped:
        return ("已批量设置%d个单元格样式；跳过%d项（原因见 stderr）" % (count, len(skipped)))
    return f"已批量设置{count}个单元格样式"


# ==================== 文档质量检查 ====================

def _check_docx_quality(file_path: str) -> dict:
    """检查Word文档质量"""
    from docx import Document
    doc = Document(file_path)
    
    issues = []
    stats = {
        "paragraphs": len(doc.paragraphs),
        "tables": len(doc.tables),
        "images": 0,
        "empty_paragraphs": 0,
        "long_paragraphs": 0,
        "short_paragraphs": 0
    }
    
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        
        if not text:
            stats["empty_paragraphs"] += 1
            continue
        
        if len(text) > 500:
            stats["long_paragraphs"] += 1
            issues.append(f"段落{i}: 内容过长({len(text)}字符)")
        
        if len(text) < 5 and text not in ["", " ", "。", "，", ".", ","]:
            stats["short_paragraphs"] += 1
            issues.append(f"段落{i}: 内容过短({text})")
    
    # Count images
    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            stats["images"] += 1
    
    # Quality score
    score = 100
    if stats["empty_paragraphs"] > 5:
        score -= 10
        issues.append(f"空段落过多({stats['empty_paragraphs']}个)")
    if stats["long_paragraphs"] > 3:
        score -= 10
        issues.append(f"长段落过多({stats['long_paragraphs']}个)")
    if stats["images"] == 0 and stats["paragraphs"] > 10:
        score -= 5
        issues.append("缺少图片")
    
    return {
        "score": max(0, score),
        "stats": stats,
        "issues": issues[:10],  # 最多返回10个问题
        "suggestions": _generate_suggestions(stats, issues)
    }


def _check_xlsx_quality(file_path: str) -> dict:
    """检查Excel文档质量"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path, data_only=True)
    
    issues = []
    stats = {
        "sheets": len(wb.sheetnames),
        "total_cells": 0,
        "empty_cells": 0,
        "merged_cells": 0,
        "formulas": 0
    }
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        stats["total_cells"] += ws.max_row * ws.max_column
        stats["merged_cells"] += len(ws.merged_cells.ranges)
        
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    stats["empty_cells"] += 1
                elif isinstance(cell.value, str) and cell.value.startswith("="):
                    stats["formulas"] += 1
    
    # Quality score
    score = 100
    empty_ratio = stats["empty_cells"] / max(stats["total_cells"], 1)
    if empty_ratio > 0.8:
        score -= 20
        issues.append(f"空单元格比例过高({empty_ratio:.1%})")
    if stats["formulas"] == 0 and stats["total_cells"] > 50:
        score -= 10
        issues.append("缺少公式")
    
    return {
        "score": max(0, score),
        "stats": stats,
        "issues": issues[:10],
        "suggestions": _generate_suggestions(stats, issues)
    }


def _generate_suggestions(stats: dict, issues: list) -> list:
    """生成改进建议"""
    suggestions = []
    
    if stats.get("empty_paragraphs", 0) > 5:
        suggestions.append("建议删除多余的空段落")
    if stats.get("long_paragraphs", 0) > 3:
        suggestions.append("建议将长段落拆分为多个短段落")
    if stats.get("images", 0) == 0:
        suggestions.append("建议添加图片以增强可读性")
    if stats.get("empty_cells", 0) > stats.get("total_cells", 0) * 0.5:
        suggestions.append("建议填充空单元格或删除空行/列")
    if stats.get("formulas", 0) == 0:
        suggestions.append("考虑使用公式进行数据计算")
    
    return suggestions[:5]  # 最多返回5个建议


# ==================== 模板管理 ====================

def _extract_docx_template(file_path: str) -> dict:
    """提取Word文档模板"""
    from docx import Document
    from docx.shared import Pt
    
    doc = Document(file_path)
    
    template = {
        "styles": [],
        "structure": [],
        "fonts": set(),
        "sizes": set()
    }
    
    # Extract paragraph styles
    for i, para in enumerate(doc.paragraphs):
        if para.text.strip():
            style_info = {
                "index": i,
                "text_preview": para.text[:50],
                "style_name": para.style.name if para.style else "Normal"
            }
            
            for run in para.runs:
                if run.font.name:
                    template["fonts"].add(run.font.name)
                if run.font.size:
                    template["sizes"].add(str(run.font.size))
            
            template["structure"].append(style_info)
    
    # Extract table info
    for i, table in enumerate(doc.tables):
        template["structure"].append({
            "type": "table",
            "index": i,
            "rows": len(table.rows),
            "columns": len(table.columns)
        })
    
    template["fonts"] = list(template["fonts"])
    template["sizes"] = list(template["sizes"])
    
    return template


def _apply_template(file_path: str, template: dict) -> str:
    """应用模板到文档"""
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    from lxml import etree
    
    doc = Document(file_path)
    
    # Apply font to all paragraphs
    target_font = template.get("font_name", "SimSun")
    target_size = template.get("font_size", 12)
    
    for para in doc.paragraphs:
        for run in para.runs:
            rPr = run._element.get_or_add_rPr()
            for rf in rPr.findall(qn('w:rFonts')):
                rPr.remove(rf)
            rFonts = etree.SubElement(rPr, qn('w:rFonts'))
            rFonts.set(qn('w:ascii'), target_font)
            rFonts.set(qn('w:hAnsi'), target_font)
            rFonts.set(qn('w:eastAsia'), target_font)
            rFonts.set(qn('w:cs'), target_font)
            rFonts.set(qn('w:hint'), 'eastAsia')
            run.font.size = Pt(target_size)
    
    doc.save(file_path)
    return f"已应用模板: {target_font} {target_size}pt"


# ==================== 样式设置 ====================

def _file_set_cell_style_xlsx(
    file_path: str, sheet_name: Optional[str], row: int, column: int,
    font_name: Optional[str] = None, font_size: Optional[int] = None,
    bold: Optional[bool] = None, italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    color: Optional[str] = None, bg_color: Optional[str] = None,
    align_h: Optional[str] = None, align_v: Optional[str] = None
) -> str:
    """设置Excel单元格样式"""
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Color
    
    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    cell = ws.cell(row=row, column=column)
    
    if font_name or font_size or bold is not None or italic is not None or underline is not None or color:
        font_kwargs = {}
        if font_name:
            font_kwargs['name'] = font_name
        if font_size:
            font_kwargs['size'] = font_size
        if bold is not None:
            font_kwargs['bold'] = bold
        if italic is not None:
            font_kwargs['italic'] = italic
        if underline is not None:
            # V3.1：原来只认布尔 true→'single'，传 "double" 会被丢掉；现在统一映射
            font_kwargs['underline'] = _underline_pair(underline)[1]
        if color:
            # V3.1：openpyxl 要 8 位 aRGB，原来直接 Color(rgb='FF0000') 会被 openpyxl 拒绝
            font_kwargs['color'] = Color(rgb=_argb_hex(color))
        try:
            cell.font = Font(**font_kwargs)
        except TypeError:
            font_kwargs.pop('underline', None)      # 极老 openpyxl 不接受 underline=None
            cell.font = Font(**font_kwargs)
    
    if bg_color:
        _fill = _argb_hex(bg_color)
        cell.fill = PatternFill(start_color=_fill, end_color=_fill, fill_type="solid")
    
    if align_h or align_v:
        align_kwargs = {}
        if align_h:
            align_kwargs['horizontal'] = align_h
        if align_v:
            align_kwargs['vertical'] = align_v
        cell.alignment = Alignment(**align_kwargs)
    
    wb.save(file_path)
    return f"已设置单元格({row},{column})样式"


def _file_set_para_style_docx(
    file_path: str, para_index: int,
    font_name: Optional[str] = None, font_size: Optional[int] = None,
    bold: Optional[bool] = None, italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    color: Optional[str] = None, align: Optional[str] = None
) -> str:
    """设置Word段落样式"""
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from lxml import etree
    
    doc = Document(file_path)
    if para_index >= len(doc.paragraphs):
        return f"段落索引{para_index}超出范围"
    
    para = doc.paragraphs[para_index]
    
    if align:
        align_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        para.alignment = align_map.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    
    for run in para.runs:
        if font_name:
            # 直接操作XML设置所有字体类型
            rPr = run._element.get_or_add_rPr()
            # 移除现有的rFonts
            for rf in rPr.findall(qn('w:rFonts')):
                rPr.remove(rf)
            # 创建新的rFonts，设置所有字体类型
            rFonts = etree.SubElement(rPr, qn('w:rFonts'))
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)
            rFonts.set(qn('w:eastAsia'), font_name)
            rFonts.set(qn('w:cs'), font_name)
            rFonts.set(qn('w:hint'), 'eastAsia')
        if font_size:
            run.font.size = Pt(font_size)
        if bold is not None:
            run.font.bold = bold
        if italic is not None:
            run.font.italic = italic
        if underline is not None:
            run.font.underline = underline
        if color:
            r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
            run.font.color.rgb = RGBColor(r, g, b)
    
    doc.save(file_path)
    return f"已设置段落{para_index}样式"


# ==================== WPS表格操作（COM） ====================

# 模块级缓存：按 app_type 记住已挂接/已新建的 COM 实例与归属标记
#   _APP_CACHE[app_type]    -> COM Application 对象
#   _APP_ATTACHED[app_type] -> True=挂接的是人类正在使用的实例 / False=本工具新建的实例
_APP_CACHE = {}
_APP_ATTACHED = {}
_APP_OWNED = _APP_ATTACHED  # 兼容旧名（语义即"归属标记"）

# 最近一次挂接尝试的方法清单 / 文档匹配信息，用于结构化报错与排查
_MATCH_ATTEMPTS = {}
_MATCH_INFO = {}

# OBJID_NATIVEOM：向窗口取"原生对象模型"（Word/Excel Window、PPT DocumentWindow）
_OBJID_NATIVEOM = 0xFFFFFFF0

# 各类 WPS 应用的挂接参数
#   prog_ids / clsids   ROT 挂接候选（版本化 ProgID 置前，避免命中 Microsoft Office）
#   top_classes         顶层窗口类名
#   child_classes       编辑区子窗口类名（OBJID_NATIVEOM 的有效宿主）
#   title_suffixes      窗口标题后缀（类名随版本变化时的兜底）
#   collection / item   已打开文档集合 / 文档对象名
_WPS_APP_SPECS = {
    "word": {
        "label": "WPS 文字",
        "prog_ids": [
            "Kwps.Application.6", "Kwps.Application",
            "KWPS.Application.6", "KWPS.Application",
            "WPS.Application", "Word.Application",
        ],
        "clsids": ["{000209FF-0000-4B30-A977-D214852036FF}"],
        "top_classes": ["OpusApp"],
        "child_classes": ["_WwG"],
        "title_suffixes": ["WPS 文字", "WPS文字", "WPS Writer"],
        "collection": "Documents",
        "item": "Document",
    },
    "excel": {
        "label": "WPS 表格",
        "prog_ids": [
            "Ket.Application.6", "Ket.Application",
            "KET.Application.6", "KET.Application",
            "ET.Application", "Excel.Application",
        ],
        "clsids": ["{00024500-0000-0000-C000-000000000046}"],
        "top_classes": ["XLMAIN"],
        "child_classes": ["EXCEL7", "_WwG"],
        "title_suffixes": ["WPS 表格", "WPS表格", "WPS Spreadsheets"],
        "collection": "Workbooks",
        "item": "Workbook",
    },
    "powerpoint": {
        "label": "WPS 演示",
        "prog_ids": [
            "Kwpp.Application.6", "Kwpp.Application",
            "KWPP.Application.6", "KWPP.Application",
            "WPP.Application", "PowerPoint.Application",
        ],
        "clsids": [],
        "top_classes": ["PPTFrameClass"],
        "child_classes": ["paneClassDC"],
        "title_suffixes": ["WPS 演示", "WPS演示", "WPS Presentation"],
        "collection": "Presentations",
        "item": "Presentation",
    },
}

# 窗口标题里的噪声标记（兼容模式 / 只读等），剥文档名时清掉
_TITLE_JUNK = (
    "[兼容模式]", "（兼容模式）", "[Compatibility Mode]",
    "(只读)", "（只读）", "(Read-Only)", "[Read-Only]",
)


class WpsAttachError(Exception):
    """挂接人类 WPS 实例 / 匹配已打开文档失败。

    异常文本即结构化 JSON，便于上游 Agent 直接解析并决定下一步行动（含求助话术与候选路径）。
    """

    def __init__(self, payload: dict):
        self.payload = payload
        super().__init__(json.dumps(payload, ensure_ascii=False, indent=2))


# ---------- COM 挂接：多级挂接链（ROT → 窗口句柄）；全部失败绝不静默新建 ----------

def _com_initialize():
    """确保当前线程已初始化 COM（幂等，失败不阻断）。"""
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass


def _com_to_plain(value, depth: int = 0):
    """把 COM 返回值转换成**可 JSON 序列化**的纯 Python 类型。

    WPS 表格的 Range.Value / NamedProperty 包装对象直接丢给 json.dumps 会报
    "Object of type NamedProperty is not JSON serializable"。纯化顺序：
      标量 → 原样；字节 → utf-8 解码；日期时间 → 字符串；序列/字典 → 逐元素递归；
      COM 包装对象 → 试取 .Value / .Value2 / .Text 再递归；兜底 str()。
    """
    import datetime as _dt
    if depth > 6:
        try:
            return str(value)
        except Exception:
            return None
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    if isinstance(value, (bytes, bytearray)):
        try:
            return value.decode("utf-8", "replace")
        except Exception:
            return str(value)
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
        return str(value)
    # comtypes lazybind 的 NamedProperty：属性/方法的惰性包装器。Excel 的 Range.Value 属于
    # "带可选参数的属性"，comtypes 无法直接 get，只会返回该包装器——必须 __call__() 一次
    # 才会真正 Invoke 取值，否则 json 序列化拿到的是包装对象本身（历史报错根源）。
    try:
        from comtypes.client.lazybind import NamedProperty as _NamedProperty
    except Exception:
        _NamedProperty = None
    if _NamedProperty is not None and isinstance(value, _NamedProperty):
        try:
            return _com_to_plain(value(), depth + 1)
        except Exception:
            return None
    if isinstance(value, (tuple, list)):
        return [_com_to_plain(item, depth + 1) for item in value]
    if isinstance(value, dict):
        return {str(k): _com_to_plain(v, depth + 1) for k, v in value.items()}
    for attr in ("Value", "Value2", "Text"):
        try:
            inner = getattr(value, attr)
        except Exception:
            continue
        if inner is None or inner is value or callable(inner):
            continue
        plain = _com_to_plain(inner, depth + 1)
        if plain is not None:
            return plain
    try:
        return str(value)
    except Exception:
        return None


def _safe_sheet_name(wb) -> Optional[str]:
    """安全读取工作簿当前激活工作表名（普通 Workbook 与编辑态 _DocShim 都兼容）。"""
    try:
        return str(wb.ActiveSheet.Name)
    except Exception:
        return None


def _com_oleobj(obj):
    """安全取 win32com 的 _oleobj_（comtypes 动态对象没有该属性）。"""
    try:
        return getattr(obj, "_oleobj_", None)
    except Exception:
        return None


def _com_try_attr(obj, name):
    """安全取 COM 属性：取到返回 (True, 值)；取不到返回 (False, None)。"""
    try:
        return True, getattr(obj, name)
    except Exception:
        return False, None


def _com_member_exists(obj, name: str) -> bool:
    """试探 COM 对象是否暴露某成员（GetIDsOfNames 优先，退化为属性访问）。"""
    if not name:
        return False
    oleobj = _com_oleobj(obj)
    if oleobj is not None:
        try:
            oleobj.GetIDsOfNames(name)
            return True
        except Exception:
            pass
    ok, _value = _com_try_attr(obj, name)
    return ok


def _window_class(hwnd) -> str:
    try:
        import win32gui
        return win32gui.GetClassName(hwnd)
    except Exception:
        return ""


def _window_title(hwnd) -> str:
    try:
        import win32gui
        return win32gui.GetWindowText(hwnd)
    except Exception:
        return ""


def _enum_top_windows(class_names=None, title_suffixes=None):
    """枚举顶层窗口，返回 [(hwnd, class_name, title)]；两类过滤都为空则返回全部。"""
    try:
        import win32gui
    except Exception:
        return []

    hits = []

    def _cb(hwnd, _param):
        cls = _window_class(hwnd)
        title = _window_title(hwnd)
        if class_names or title_suffixes:
            matched = bool(class_names) and cls in class_names
            if not matched and title_suffixes and title:
                tail = title.rstrip()
                for suffix in title_suffixes:
                    if tail.endswith(suffix):
                        matched = True
                        break
            if not matched:
                return True
        hits.append((hwnd, cls, title))
        return True

    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        pass
    return hits


def _enum_child_windows(hwnd, class_names=None):
    """枚举某窗口的所有子/后代窗口，返回 [(hwnd, class_name)]。"""
    try:
        import win32gui
    except Exception:
        return []

    hits = []

    def _cb(child, _param):
        cls = _window_class(child)
        if not class_names or cls in class_names:
            hits.append((child, cls))
        return True

    try:
        win32gui.EnumChildWindows(hwnd, _cb, None)
    except Exception:
        # 个别 pywin32 版本无 EnumChildWindows：退回 FindWindowEx 手工递归
        stack = [hwnd]
        while stack:
            parent = stack.pop()
            child = None
            try:
                while True:
                    child = win32gui.FindWindowEx(parent, child, None, None)
                    if not child:
                        break
                    cls = _window_class(child)
                    if not class_names or cls in class_names:
                        hits.append((child, cls))
                    stack.append(child)
            except Exception:
                continue
    return hits


def _native_om_from_hwnd(hwnd):
    """取窗口的原生对象模型：AccessibleObjectFromWindow(OBJID_NATIVEOM) → 后期绑定对象。

    两条调用路径互为兜底（oledll 自动校验 HRESULT / WinDLL 手工校验），任一路径成功即返回；
    全部失败返回 None（由上层继续走其他挂接手段，绝不新建实例）。
    """
    import ctypes
    from ctypes import byref, POINTER
    import ctypes.wintypes as wt

    try:
        import comtypes.client
        from comtypes.automation import IDispatch as _CIDispatch
    except Exception:
        return None

    _com_initialize()

    loaders = []
    try:
        loaders.append(("oledll", ctypes.oledll.oleacc))
    except Exception:
        pass
    try:
        loaders.append(("windll", ctypes.WinDLL("oleacc")))
    except Exception:
        pass

    for label, dll in loaders:
        try:
            func = dll.AccessibleObjectFromWindow
            func.restype = ctypes.c_long
        except Exception:
            continue
        ptr = POINTER(_CIDispatch)()
        try:
            hr = func(
                wt.HWND(hwnd),
                wt.DWORD(_OBJID_NATIVEOM),
                byref(_CIDispatch._iid_),
                byref(ptr),
            )
        except OSError:
            continue          # oledll 会对失败 HRESULT 抛 OSError
        except Exception:
            continue
        if label == "windll" and hr != 0:
            continue
        if not ptr:
            continue
        try:
            return comtypes.client.dynamic.Dispatch(ptr)
        except Exception:
            continue
    return None


def _application_from_native(obj, app_type: str):
    """从 OBJID_NATIVEOM 对象里取出 Application（兼容 Window / Document 两类形态）。"""
    if not obj:
        return None
    spec = _WPS_APP_SPECS.get(app_type, {})
    collection = spec.get("collection", "")
    item = spec.get("item", "")

    if collection and _com_member_exists(obj, collection):
        return obj                                     # 本身已是 Application
    ok, app = _com_try_attr(obj, "Application")         # Window 型
    if ok and app is not None:
        return app
    ok, parent = _com_try_attr(obj, "Parent")
    if ok and parent is not None and collection and _com_member_exists(parent, collection):
        return parent
    if item:
        ok, doc = _com_try_attr(obj, item)              # Document / Workbook / Presentation 型
        if ok and doc is not None:
            ok, app = _com_try_attr(doc, "Application")
            if ok and app is not None:
                return app
    return None


def _attach_by_rot(app_type: str):
    """ROT 挂接：按候选 ProgID / CLSID 依次尝试 GetActiveObject / GetObject。返回 (app, tried)。"""
    spec = _WPS_APP_SPECS.get(app_type, {})
    _com_initialize()
    try:
        import win32com.client
    except Exception as err:
        return None, ["导入 win32com.client 失败: %s" % err]

    candidates = list(spec.get("prog_ids", []))
    for clsid in spec.get("clsids", []):
        if clsid not in candidates:
            candidates.append(clsid)

    tried = []
    for name in candidates:
        for label, func in (
            ("win32com.client.GetActiveObject(%s)" % name,
             lambda n=name: win32com.client.GetActiveObject(n)),
            ("win32com.client.GetObject(Class=%s)" % name,
             lambda n=name: win32com.client.GetObject(Class=n)),
        ):
            try:
                app = func()
            except Exception:
                tried.append("%s -> 失败" % label)
                continue
            if app is None:
                tried.append("%s -> 失败" % label)
                continue
            tried.append("%s -> 成功" % label)
            return app, tried
    return None, tried


def _attach_by_window(app_type: str):
    """窗口句柄挂接：顶层窗口 → 编辑区子窗口 → OBJID_NATIVEOM → Application。返回 (app, tried)。"""
    spec = _WPS_APP_SPECS.get(app_type, {})
    _com_initialize()
    tried = []

    top_windows = _enum_top_windows(class_names=spec.get("top_classes"))
    tried.append("EnumWindows(class in %s) -> %d 个窗口"
                 % (spec.get("top_classes") or [], len(top_windows)))
    if not top_windows:
        top_windows = _enum_top_windows(title_suffixes=spec.get("title_suffixes"))
        tried.append("EnumWindows(title endswith %s) -> %d 个窗口"
                     % (spec.get("title_suffixes") or [], len(top_windows)))

    # 可见窗口优先：残留的隐藏自动化实例（实测存在）同样会被 EnumWindows 枚举到，
    # 排序后不会被先命中；同级内保持原 z 序（sort 稳定）。
    try:
        import win32gui
        top_windows.sort(key=lambda item: 0 if win32gui.IsWindowVisible(item[0]) else 1)
    except Exception:
        pass

    for hwnd, cls, _title in top_windows:
        children = _enum_child_windows(hwnd, spec.get("child_classes"))
        if not children:
            children = [(hwnd, cls)]                   # 少数版本直接在顶层窗口暴露原生对象
        for child_hwnd, child_cls in children:
            native = _native_om_from_hwnd(child_hwnd)
            if native is None:
                continue
            app = _application_from_native(native, app_type)
            if app is None:
                continue
            tried.append("AccessibleObjectFromWindow(OBJID_NATIVEOM)@hwnd=%s(%s) -> 成功"
                         % (child_hwnd, child_cls))
            return app, tried
    tried.append("AccessibleObjectFromWindow(OBJID_NATIVEOM) -> 失败")
    return None, tried


def _norm_com_path(path: str) -> str:
    """规范化绝对路径，用于与 COM 侧 FullName 比对"""
    try:
        return os.path.normcase(os.path.normpath(os.path.abspath(path)))
    except Exception:
        return os.path.normcase(os.path.normpath(str(path)))


def _com_dispatch_id(obj, name: str):
    """把成员名解析成 DISPID：先 GetIDsOfNames，失败再扫 ITypeInfo；取不到返回 None。

    本机 WPS 表格(ET)的 IDispatch::GetIDsOfNames 不含 SaveAs/ExportAsFixedFormat 名称，
    必须靠 ITypeInfo 扫描或直接用 DISPID 数字才能调到。
    """
    try:
        return int(obj._oleobj_.GetIDsOfNames(name))
    except Exception:
        pass
    try:
        type_info = obj._oleobj_.GetTypeInfo()
        attr = type_info.GetTypeAttr()
        for i in range(attr.cFuncs):
            try:
                func_desc = type_info.GetFuncDesc(i)
                names = type_info.GetNames(func_desc.memid)
                if names and str(names[0]).lower() == str(name).lower():
                    return int(func_desc.memid)
            except Exception:
                continue
    except Exception:
        pass
    return None


def _com_invoke(obj, dispids, *args):
    """对 COM 对象按 DISPID 列表依次直调 Invoke，成功即返回结果；全部失败抛出最后一个异常。

    dispids 元素可为 int（直接当 DISPID 用）或 str（成员名，先解析为 DISPID）。
    用于绕开 WPS 的名称查找缺陷（com_error -2147352573「找不到成员」）。
    """
    import pythoncom
    last_err = None
    for item in dispids:
        try:
            dispid = item if isinstance(item, int) else _com_dispatch_id(obj, item)
        except Exception as err:
            last_err = err
            continue
        if dispid is None:
            continue
        try:
            return obj._oleobj_.Invoke(dispid, 0, pythoncom.DISPATCH_METHOD, True, *args)
        except Exception as err:
            last_err = err
    if last_err is not None:
        raise last_err
    raise Exception(f"COM 调用失败：无法解析成员 {dispids!r}")


def _com_call(obj, name: str, *args, dispids=()):
    """调用 COM 方法：名称调用为首选，失败后按 DISPID 列表直调兜底，全部失败抛异常。

    dispids 为名称调用失败后额外尝试的数字 DISPID（本机 WPS 实测值）。
    """
    try:
        return getattr(obj, name)(*args)
    except Exception as name_err:
        last_err = None
        for dispid in [_com_dispatch_id(obj, name)] + list(dispids):
            if dispid is None:
                continue
            try:
                return _com_invoke(obj, (dispid,), *args)
            except Exception as err:
                last_err = err
        raise Exception(f"{name} 调用失败：名称调用报错 {name_err}；DISPID 直调报错 {last_err}")


def _pdf_written(output_path: str) -> bool:
    """判定 PDF 是否真实落盘"""
    try:
        return os.path.exists(output_path) and os.path.getsize(output_path) > 0
    except Exception:
        return False


def _com_save_as_workbook(wb, file_path: str):
    """工作簿另存为：名称调用优先，失败退回 DISPID 直调。

    1925 = Excel 经典 Workbook.SaveAs，3219457 = 本机 WPS 表格(ET)实测的 Workbook.SaveAs。
    """
    return _com_call(wb, "SaveAs", os.path.abspath(file_path), dispids=(1925, 3219457))


def _com_export_pdf_file(app, obj, src_path: str, output_path: str, attached: bool = False) -> str:
    """把已打开的文档对象导出为 PDF（多路兜底），返回命中的策略名；全部失败抛异常。

    策略顺序：
      1) ExportAsFixedFormat(path, 17)  —— Word 侧实测可用（wdExportFormatPDF）
      2) ExportAsFixedFormat(path, 2)   —— PowerPoint 侧 ppFixedFormatTypePDF
      3) ExportAsFixedFormat(0, path)   —— Excel 侧经典调用形态
      4) ExportPdf(path[, 0, 0])        —— 本机 WPS 表格实测可用（无副作用）
      5) SaveAs(path, 17/57/32)         —— 最后兜底，只在源文件的临时副本上执行，绝不污染原对象

    attached=True 表示 obj 是**人类正在编辑的文档**：一律跳过副本兜底
    （副本会把源文件复制到原目录并打开，等于在人类的 WPS 里另开窗口、污染目录）。
    """
    import shutil

    # V3.1：先写到同目录临时名，**成功后再原子替换**目标。
    # 原实现是"每次策略尝试前先 os.remove(output_path)"——只要所有策略都失败，
    # 人原来那份同名 PDF 就被删掉且没有替代品（回归：删了不补）。
    final_path = output_path
    output_path = _tmp_pdf_path(final_path)

    def attempt(label, fn):
        if os.path.exists(output_path):
            try:
                os.remove(output_path)        # 现在删的只是本次自己的临时文件
            except Exception:
                pass
        try:
            fn()
        except Exception as err:
            return f"{label}: {err}"
        if _pdf_written(output_path):
            try:
                os.replace(output_path, final_path)   # 同盘原子替换
            except Exception as err:
                return f"{label}: 已生成但替换目标失败: {err}"
            return None
        return f"{label}: 调用未报错但未生成 PDF"

    errors = []
    plans = [
        ("ExportAsFixedFormat(path,17)", lambda: _com_call(obj, "ExportAsFixedFormat", output_path, 17, dispids=(552,))),
        ("ExportAsFixedFormat(path,2)", lambda: _com_call(obj, "ExportAsFixedFormat", output_path, 2)),
        ("ExportAsFixedFormat(0,path)", lambda: _com_call(obj, "ExportAsFixedFormat", 0, output_path)),
        ("ExportPdf(path)", lambda: _com_call(obj, "ExportPdf", output_path, dispids=(3219476,))),
        ("ExportPdf(path,0,0)", lambda: _com_call(obj, "ExportPdf", output_path, 0, 0, dispids=(3219476,))),
    ]
    for label, fn in plans:
        err = attempt(label, fn)
        if err is None:
            return label
        errors.append(err)

    # 最后兜底：SaveAs + 导出格式常量，在源文件副本上执行，避免改动原文档/原工作簿
    src_abs = os.path.abspath(src_path) if src_path else ""
    if attached:
        errors.append("SaveAs@副本: 已挂接人类正在编辑的文档，跳过副本兜底（避免另开窗口/污染目录）")
    elif src_abs and os.path.exists(src_abs):
        suffix = Path(src_abs).suffix.lower()
        # V3.1：副本名带 pid，避免固定名互撞/残留（原来固定叫 __wps_pdf_src_copy<suffix>）
        copy_path = os.path.join(os.path.dirname(src_abs),
                                 "__wps_pdf_src_copy_%d%s" % (os.getpid(), suffix))
        if suffix in (".doc", ".docx"):
            kind, fmt, save_dispids = "documents", 17, (376, 568)          # wdFormatPDF
        elif suffix in (".ppt", ".pptx"):
            kind, fmt, save_dispids = "presentations", 32, (1905,)         # ppSaveAsPDF
        else:
            kind, fmt, save_dispids = "workbooks", 57, (1925, 3219457)     # xlTypePDF
        collection = {"documents": "Documents", "presentations": "Presentations",
                      "workbooks": "Workbooks"}[kind]
        copy_obj = None
        try:
            shutil.copy2(src_abs, copy_path)
            opener = getattr(app, collection)
            copy_obj = opener.Open(copy_path)
            _com_call(copy_obj, "SaveAs", output_path, fmt, dispids=save_dispids)
            if _pdf_written(output_path):
                try:
                    os.replace(output_path, final_path)
                except Exception as err:
                    errors.append(f"SaveAs(path,{fmt})@副本: 已生成但替换目标失败: {err}")
                else:
                    return f"SaveAs(path,{fmt})@副本"
            else:
                errors.append(f"SaveAs(path,{fmt})@副本: 调用未报错但未生成 PDF")
        except Exception as err:
            errors.append(f"SaveAs(path,{fmt})@副本: {err}")
        finally:
            if copy_obj is not None:
                try:
                    copy_obj.Close(False)
                except Exception:
                    pass
            try:
                if os.path.exists(copy_path):
                    os.remove(copy_path)
            except Exception:
                pass
    else:
        errors.append("SaveAs@副本: 源文件不存在，跳过")

    raise ValueError(
        "导出 PDF 全部策略失败：" + " | ".join(errors) +
        "\n提示：若人正停在某个单元格的编辑态，请让他按一下 Esc 或回车后重试"
        "（导出需要拿文档对象，编辑态下 WPS 会拒绝按名字解析 COM 成员）。")


def _app_main_hwnd(app):
    """尽力取 Application 的主窗口句柄；取不到返回 None（绝不抛错）。

    Excel/ET 用 Application.Hwnd；WPS 文字/演示退化为 Windows(1).Hwnd、ActiveWindow.Hwnd。
    """
    getters = (
        lambda: int(getattr(app, "Hwnd")),
        lambda: int(getattr(app, "Windows")(1).Hwnd),
        lambda: int(getattr(app, "ActiveWindow").Hwnd),
    )
    for getter in getters:
        try:
            hwnd = getter()
        except Exception:
            continue
        if hwnd:
            return hwnd
    return None


def _app_window_state(app) -> str:
    """判定 Application 主窗口可见性：'visible' / 'hidden' / 'unknown'。

    人类正在用的编辑器窗口一定可见（最小化仍算可见）；自动化残留/隐藏实例是隐藏窗口。
    只有在**明确读到隐藏**时才返回 'hidden'，读不到句柄一律 'unknown'，
    避免误伤取不到 Hwnd 的 WPS 文字/演示。
    """
    hwnd = _app_main_hwnd(app)
    if not hwnd:
        return "unknown"
    try:
        import win32gui
        if not win32gui.IsWindow(hwnd):
            return "unknown"
        return "visible" if win32gui.IsWindowVisible(hwnd) else "hidden"
    except Exception:
        return "unknown"


def _cache_alive(app) -> bool:
    """缓存实例可用性探测：任一轻量 COM 调用失败即判定缓存失效。"""
    _com_initialize()
    for member in ("Documents", "Workbooks", "Presentations"):
        ok, coll = _com_try_attr(app, member)
        if not ok or coll is None:
            continue
        try:
            int(coll.Count)
            return True
        except Exception:
            return False
    ok, _name = _com_try_attr(app, "Name")
    return bool(ok)


def _invalidate_app(app_type: str):
    """清空某类应用的缓存（不做 Quit，避免误关人类窗口）。"""
    _APP_CACHE.pop(app_type, None)
    _APP_ATTACHED.pop(app_type, None)


def _prepare_attached_app(app, app_type: str):
    """挂接人类实例后的轻量准备：只确保窗口可见，绝不改动人类的文档状态与告警设置。"""
    try:
        app.Visible = True
    except Exception:
        pass


def _dispatch_new_instance(app_type: str):
    """仅在调用方显式授权（allow_new_instance=True）时取实例：返回 (app, created_new)。

    ⚠️ 现状（2026-09-28 复核）：全文件 `_get_wps_app` 的 27 处调用**没有一处传 True**，
    所以这条分支目前不可达（评审把它列为死代码，属实）。**故意保留不删**：这是"绝不自己
    新建后台 WPS"这条既定原则的唯一逃生口，去掉它等于把未来的可能性也砍掉；删它属于
    净收益为负的清理。README 里"未打开则 Open 兜底"的说法与现状不符，以本条为准。

    created_new=True  —— 本工具新建的隐藏实例（归属 attached=False，用完可 Quit）；
    created_new=False —— Dispatch 实际**复用**了人类正在用的可见实例。

    实测：本机 WPS 表格的 Dispatch 会直接返回人类已打开的实例（不新建）。
    因此必须先判可见性，再决定要不要改 Visible / DisplayAlerts —— 否则会把人类的窗口隐藏掉。
    """
    spec = _WPS_APP_SPECS.get(app_type, {})
    _com_initialize()
    try:
        import win32com.client
    except Exception:
        return None, False
    for prog_id in spec.get("prog_ids", []):
        try:
            app = win32com.client.Dispatch(prog_id)
        except Exception:
            continue
        if _app_window_state(app) == "visible":
            return app, False          # 复用到了人类实例：不动它的任何属性
        try:
            app.DisplayAlerts = False
        except Exception:
            pass
        try:
            app.Visible = False
        except Exception:
            pass
        return app, True
    return None, False


def _build_attach_error(app_type: str, tried) -> dict:
    """挂接全链路失败的结构化错误：暂停任务、携带候选路径并向人类求助。"""
    spec = _WPS_APP_SPECS.get(app_type, {})
    label = spec.get("label", app_type)
    try:
        candidates = [h for h in _scan_open_documents_by_window() if h.get("app_type") == app_type]
    except Exception:
        candidates = []
    return {
        "ok": False,
        "reason": "no_live_instance",
        "app_type": app_type,
        "app_label": label,
        "tried_methods": list(tried or []),
        "doc_path_candidates": candidates,
        "attach_paused": True,
        "hint": ("未能挂接到人类正在使用的 %s 实例。本工具绝不新建实例、绝不另开窗口，"
                 "故已暂停本次任务，等待人类协助。" % label),
        "human_action_required": (
            "请人类协助确认：①WPS 是否已启动并打开了目标文档？"
            "②是否在使用 Microsoft Office 而非 WPS？"
            "③若不需要与编辑器实时同步，可改用 mode=\"file\"；"
            "④若已完成①②，请把当前编辑窗口标题或文件绝对路径告诉我。"
        ),
    }


def _get_wps_app(app_type: str, allow_new_instance: bool = False):
    """获取 WPS COM 应用对象，返回 (app, attached) 二元组。

    多级挂接链（全部失败绝不静默新建）：
      1) 模块缓存：先用轻量 COM 探测校验，探测抛错即清空缓存并重走完整挂接链；
      2) ROT：GetActiveObject / GetObject，逐个候选版本化 ProgID 与 CLSID；
      3) 窗口句柄：EnumWindows（可见窗口优先）→ 编辑区子窗口 → AccessibleObjectFromWindow(OBJID_NATIVEOM)。
    采信校验：候选实例的主窗口若**明确处于隐藏状态**，判定为自动化残留实例，不采信并继续下一级。
    全部失败：allow_new_instance=False（默认）→ 抛结构化 WpsAttachError 向人类求助；
              allow_new_instance=True（调用方显式授权）→ Dispatch 取实例（是否新建见 _dispatch_new_instance）。
    attached=True 表示挂接的是人类正在使用的实例：只允许就地 Save()，绝不 Quit / Close 人类文档。
    """
    _com_initialize()

    cached = _APP_CACHE.get(app_type)
    if cached is not None:
        # 人类实例不可能隐藏：若缓存里是隐藏实例（残留实例/被自动化隐藏），立即判失效重挂，
        # 否则一次误挂会被缓存固化，之后每次都"挂到空进程、读不到人类的文件"。
        stale = _APP_ATTACHED.get(app_type, False) and _app_window_state(cached) == "hidden"
        if not stale and _cache_alive(cached):
            return cached, _APP_ATTACHED.get(app_type, False)
        _invalidate_app(app_type)   # 缓存失效 → 重走完整挂接链

    tried = []
    app = None
    for attach in (_attach_by_rot, _attach_by_window):
        candidate, sub_tried = attach(app_type)
        tried.extend(sub_tried)
        if candidate is None:
            continue
        if _app_window_state(candidate) == "hidden":
            tried.append("%s -> 命中的实例主窗口隐藏，判定为自动化残留实例，不采信"
                         % getattr(attach, "__name__", str(attach)))
            continue
        app = candidate
        break

    if app is not None:
        _prepare_attached_app(app, app_type)
        _APP_CACHE[app_type] = app
        _APP_ATTACHED[app_type] = True
        _MATCH_ATTEMPTS[app_type] = tried
        return app, True

    if allow_new_instance:
        app, created_new = _dispatch_new_instance(app_type)
        if app is not None:
            _APP_CACHE[app_type] = app
            _APP_ATTACHED[app_type] = created_new   # 复用到人类实例时同样标记为 attached
            _MATCH_ATTEMPTS[app_type] = tried
            return app, created_new

    _MATCH_ATTEMPTS[app_type] = tried
    raise WpsAttachError(_build_attach_error(app_type, tried))


def _release_app(app, app_type: str, attached: bool):
    """释放应用对象（替代 app.Quit()）。

    attached=True  —— 挂接的是用户实例：只把窗口恢复可见，绝不 Quit、绝不 Close 用户文档。
    attached=False —— 本工具新建的实例：清缓存并退出（异常吞掉）。
    """
    if attached:
        try:
            app.Visible = True
        except Exception:
            pass
        return
    _invalidate_app(app_type)
    try:
        app.Quit()
    except Exception:
        pass


def _detect_app_type(app) -> Optional[str]:
    """判定 COM 实例所属应用：缓存归属优先，其次按集合成员探测。"""
    for app_type, cached in _APP_CACHE.items():
        if cached is app:
            return app_type
    for app_type, spec in _WPS_APP_SPECS.items():
        if _com_member_exists(app, spec["collection"]):
            return app_type
    return None


def _safe_com_str(obj, attr: str, default: str = "") -> str:
    ok, value = _com_try_attr(obj, attr)
    if not ok or value is None:
        return default
    try:
        return str(value)
    except Exception:
        return default


def _safe_com_bool(obj, attr: str, default=None):
    ok, value = _com_try_attr(obj, attr)
    if not ok:
        return default
    try:
        return bool(value)
    except Exception:
        return default


def _doc_window_title(obj, app_type: str) -> str:
    """取文档所在窗口标题（Word 用 ActiveWindow.Caption，Excel/PPT 用 Windows(1).Caption）。"""
    if app_type == "word":
        ok, window = _com_try_attr(obj, "ActiveWindow")
        if ok and window is not None:
            return _safe_com_str(window, "Caption")
        return ""
    ok, windows = _com_try_attr(obj, "Windows")
    if ok and windows is not None:
        try:
            return _safe_com_str(windows(1), "Caption")
        except Exception:
            return ""
    return ""


def _iter_open_docs(app, app_type: str):
    """遍历实例中已打开的文档，产出 (doc_obj, info)。单项读取失败只跳过该项。"""
    spec = _WPS_APP_SPECS[app_type]
    collection = getattr(app, spec["collection"])     # 失败由调用方处理（缓存失效 / 结构化报错）
    count = int(collection.Count)
    for index in range(1, count + 1):
        try:
            obj = collection(index)
        except Exception:
            continue
        if obj is None:
            continue
        info = {
            "app_type": app_type,
            "app_label": spec["label"],
            "index": index,
            "name": _safe_com_str(obj, "Name"),
            "full_name": _safe_com_str(obj, "FullName"),
            "window_title": _doc_window_title(obj, app_type),
            "saved": _safe_com_bool(obj, "Saved"),
            "read_only": _safe_com_bool(obj, "ReadOnly"),
        }
        yield obj, info


def _match_hint(app_type: str) -> dict:
    """把最近一次文档匹配信息附在返回值里（便于上游 Agent 判断是否复用到了人类文档）。"""
    info = _MATCH_INFO.get(app_type)
    if not info:
        return {}
    hint = {"matched_by": info.get("matched_by")}
    if info.get("full_name"):
        hint["matched_full_name"] = info["full_name"]
    return hint


# ==================== 名称解析被拒（人类正在单元格编辑态）时的救援 ====================
# 实测（2026-09-24，同一实例同一文件，中间只动选区）：
#   选区 F3（编辑框抱着） → ActiveWorkbook.Name 报 <unknown>.Name、Workbooks(1) 报 -2147418111
#   选区 F4（已退出编辑态）→ 立刻全部正常
# 即：人类把单元格置于编辑态时，WPS 表格会拒绝 IDispatch::GetIDsOfNames（按名字查 DISPID）。
# 故障态下的可用性分层（实测）：
#   ✅ Application 级：app.Name / Workbooks.Count / Windows(i).Caption
#   ✅ Sheet/Range **读**：app.ActiveSheet.Name / UsedRange / Cells(r,c).Value
#   ❌ Workbook 级：ActiveWorkbook(.Name/.FullName/.Saved) / Workbooks(1) / ActiveSheet.Parent
#   ❌ Sheet/Range **写**：Cells(r,c).Value = ... / Range(...).Value = ... / ActiveCell.Value = ...
#      → 四种写法全报 -2147352567「发生意外」
#
# ★ 由此定下路线（用户 2026-09-24 拍板）：**读走旁路，写等一会再 ESC**
#   读：无按键旁路（Caption 定位 + app.ActiveSheet）——秒回、完全不打扰人类、永不发 ESC
#   写：等 BUSY_RETRY_SECONDS（默认 5s，给人松手的机会）→ 仍不行就发 ESC 挤出编辑态 → 再写
#       （写绕不过去：编辑态下 WPS 拒绝一切写入）
#
# 为什么不用 comtypes 兜底：实测可行，但会多引一层依赖；而上面的路线已覆盖实际场景。

BUSY_RETRY_SECONDS = float(os.environ.get("WPS_MCP_BUSY_RETRY_SECONDS", "5"))
BUSY_RETRY_STEP = float(os.environ.get("WPS_MCP_BUSY_RETRY_STEP", "0.5"))
BUSY_ESC_ENABLED = os.environ.get("WPS_MCP_BUSY_ESC", "1") not in ("0", "false", "False")

# 会**写盘**的工具：这些在编辑态下必须走"等 BUSY_RETRY_SECONDS → 仍不行就发 ESC"；
# 其余一律按只读处理（走无按键旁路，秒回、完全不打扰人类）。
_WRITE_TOOLS = {
    "write_document", "edit_cell", "add_rows", "add_columns", "merge_cells", "unmerge_cells",
    "set_column_width", "set_row_height", "insert_image", "set_cell_style", "set_paragraph_style",
    "batch_write", "batch_set_style", "apply_template", "activate_sheet",
}
_CALL_MODE = "read"        # handle_tool_call 每次进入时重设


def _open_count(app, app_type: str) -> int:
    """集合里的文档数；失败返回 0（不抛错，用于区分"编辑态读不出名字"和"真的没开文档"）。"""
    try:
        return int(getattr(app, _WPS_APP_SPECS[app_type]["collection"]).Count)
    except Exception:
        return 0


def _docs_without_names(collected) -> bool:
    """能取到文档对象，却一个名字/路径/标题都读不出来 → 判定名称解析被服务端拒绝。"""
    if not collected:
        return False
    for _obj, info in collected:
        if info.get("name") or info.get("full_name") or info.get("window_title"):
            return False
    return True


def _caption_doc_name(app) -> str:
    """窗口标题里的文件名——编辑态下唯一还读得到的定位线索（Application 级不受影响）。"""
    try:
        caption = str(app.Windows(1).Caption)
    except Exception:
        return ""
    for junk in (" *", " - WPS 表格", " - WPS表格", "WPS 表格", "WPS表格"):
        caption = caption.replace(junk, "")
    return caption.strip()


class _DocShim(object):
    """无按键旁路用的最小文档替身：只包住 Application 级仍然可用的部分。

    ActiveSheet / Sheets() → 转给 Application（编辑态下可用）
    Save()                 → **不保存**（Workbook 对象拿不到），只立标记让上层如实告知人类
    Close()                → 一律拒绝，绝不关闭人类的文档
    """

    def __init__(self, app, app_type: str, name: str, full_name: str = ""):
        self._app = app
        self._app_type = app_type
        self._name = name
        # 旁路是靠"窗口标题 == 目标文件名 且 只开了一本"定到的，所以目标路径是可信的；
        # 如实回报它，上层（_match_hint / 测试）才能拿到完整路径而不是空串。
        self._full_name = full_name
        self.save_skipped = True

    @property
    def Name(self):
        return self._name

    @property
    def FullName(self):
        return self._full_name

    @property
    def Saved(self):
        return None

    @property
    def ActiveSheet(self):
        if self._app_type != "excel":
            raise RuntimeError("无按键旁路暂只支持 WPS 表格")
        return self._app.ActiveSheet

    def Sheets(self, key):
        if self._app_type != "excel":
            raise RuntimeError("无按键旁路暂只支持 WPS 表格")
        return self._app.Sheets(key)

    def Save(self):
        sys.stderr.write("[wps-editor] 编辑态旁路：跳过 wb.Save()，请人类自行 Ctrl+S\n")
        sys.stderr.flush()
        return False

    def Close(self, *args, **kwargs):
        raise RuntimeError("编辑态旁路下不允许关闭人类的文档")


def _doc_busy_error(app_type: str, path, collected) -> dict:
    """名称解析被拒时的**准确**报错（不再误报"目标文档不在其中"）。"""
    payload = _doc_not_open_error(app_type, path, collected)
    payload["reason"] = "com_name_resolution_refused"
    payload["hint"] = ("已挂接到人类 %s 实例，但 WPS 拒绝按名字解析 COM 成员——实测这是人类正处于"
                       "**单元格编辑态**（编辑框抱着某个单元格）时的表现，**不是**文件没打开。"
                       % _WPS_APP_SPECS[app_type]["label"])
    payload["human_action_required"] = (
        ("写操作在编辑态下会被 WPS 拒绝（实测 Cells(...).Value= 必失败）。工具已等待 %.0f 秒，%s。"
         "请按一下 Esc 或点击其它单元格后重试。"
         % (BUSY_RETRY_SECONDS,
            "并已用 ESC 尝试退出编辑态但仍未成功" if BUSY_ESC_ENABLED else "ESC 已关闭（WPS_MCP_BUSY_ESC=0）"))
        if _CALL_MODE == "write" else
        ("请在 WPS 里按一下 Esc 或点击任意其它单元格退出单元格编辑状态，然后让我重试——"
         "实测退出后同一秒即可恢复。"))
    payload["retry_budget_seconds"] = BUSY_RETRY_SECONDS
    payload["call_mode"] = _CALL_MODE
    return payload


def _esc_exit_edit_mode(app) -> bool:
    """最后手段：发一次 ESC 退出单元格编辑态。

    两道保险：① 只有显式允许才发；② 只有 WPS 正处于前台窗口时才发
    （否则按键可能落到别的程序里）。ESC 会丢掉编辑框里没提交的输入，故放最后。
    """
    if not BUSY_ESC_ENABLED:
        return False
    try:
        import win32gui
        if win32gui.GetForegroundWindow() != int(app.Hwnd):
            return False
    except Exception:
        return False
    try:
        app.SendKeys("{ESC}")
        return True
    except Exception:
        return False


def _bypass_shim(app, app_type: str, target: str):
    """无按键旁路：窗口标题定位 + Application 级读写。成立返回 _DocShim，否则 None。

    只在**只读**场景使用：实测编辑态下"读得到、写不进去"，所以它救得了读、救不了写。
    """
    want = os.path.basename(target)
    if _caption_doc_name(app) == want and _open_count(app, app_type) == 1:
        _MATCH_INFO[app_type] = {"matched_by": "caption_bypass_unsaved", "full_name": target}
        return _DocShim(app, app_type, want, target)
    return None


def _retry_locate(app, app_type: str, target: str, seconds: float):
    """退避重试：不碰键盘，等人类自己松手（松手后同一秒即可定位成功）。"""
    deadline = time.time() + max(seconds, 0.0)
    while True:
        time.sleep(BUSY_RETRY_STEP)
        try:
            fresh = list(_iter_open_docs(app, app_type))
        except Exception:
            fresh = []
        for obj, info in fresh:
            if info.get("full_name") and _norm_com_path(info["full_name"]) == target:
                _MATCH_INFO[app_type] = {"matched_by": "retry_after_name_refused",
                                         "full_name": info["full_name"]}
                return obj
        if time.time() >= deadline:
            return None


def _locate_with_rescue(app, app_type: str, target: str, collected):
    """编辑态下的救援路线（读 / 写分叉）。命中返回文档对象，否则 None。

    读（_CALL_MODE != "write"）：**旁路优先**，秒回且完全不打扰人类；旁路不成立才退避重试。
                                读**永不发 ESC**——没必要为了读去打断人类。
    写（_CALL_MODE == "write"）：旁路救不了写（实测 `Value=` 必失败），所以
        ① 等 BUSY_RETRY_SECONDS（给人松手的机会）；
        ② 仍不行 → 发 ESC 把人类挤出编辑态（仅当 WPS 处于前台窗口）；
        ③ 再试一次。
    """
    if _CALL_MODE != "write":
        shim = _bypass_shim(app, app_type, target)
        if shim is not None:
            return shim
        return _retry_locate(app, app_type, target, BUSY_RETRY_SECONDS)

    doc = _retry_locate(app, app_type, target, BUSY_RETRY_SECONDS)
    if doc is not None:
        return doc
    if _esc_exit_edit_mode(app):
        time.sleep(0.5)
        return _retry_locate(app, app_type, target, 1.0)
    return None


def _doc_not_open_error(app_type: str, path, collected) -> dict:
    """目标文档不在人类实例中：结构化报错 + 已有文档清单 + 人类求助话术。"""
    spec = _WPS_APP_SPECS[app_type]
    open_docs = [
        {
            "app_type": info["app_type"],
            "name": info["name"],
            "full_name": info["full_name"],
            "window_title": info["window_title"],
            "saved": info["saved"],
            "read_only": info["read_only"],
        }
        for _obj, info in collected
    ]
    if path is None:
        reason = "path_required" if len(open_docs) > 1 else "no_open_document"
    else:
        reason = "target_not_open"
    human_action = "请人类确认：您当前在 WPS 中编辑的是哪一份文件？"
    if len(open_docs) > 1:
        human_action += "（当前实例已打开 %d 份文档，请告知目标文件名，或先关闭无关文档）" % len(open_docs)
    elif not open_docs:
        human_action += "（当前实例未打开任何文档，请人类先打开目标文件）"
    else:
        human_action += "若目标文件尚未打开，请人类先在 WPS 中打开该文件，我再继续。"
    return {
        "ok": False,
        "reason": reason,
        "app_type": app_type,
        "app_label": spec["label"],
        "requested_path": os.path.abspath(path) if path else None,
        "open_documents": open_docs,
        "attached_instance": True,
        "tried_methods": _MATCH_ATTEMPTS.get(app_type, []),
        "hint": ("已挂接到人类正在使用的 %s 实例，但目标文档不在其中。为不打扰人类环境，"
                 "本工具不会新建实例、不会另开窗口、不会创建副本。" % spec["label"]),
        "human_action_required": human_action,
    }


def _get_open_document(app, path: Optional[str] = None, app_type: Optional[str] = None):
    """在**人类已打开的文档**里复用目标文档：只复用，绝不 Open 新文档、绝不新开窗口。

    命中规则：
      1) 给定 path：按规范化绝对路径比对 FullName；
      2) 未给 path 且实例中恰好只有一个已打开文档：直接复用该文档（matched_by=single_open_doc）。
    未命中抛结构化 WpsAttachError（reason=target_not_open / path_required / no_open_document），
    并附带当前实例全部已打开文档的 FullName + 窗口标题清单与人类求助提示。
    缓存实例一次 COM 调用抛错即视为失效：清缓存 → 重走完整挂接链 → 重试一次。
    """
    if app_type is None:
        app_type = _detect_app_type(app)
    if app_type is None or app_type not in _WPS_APP_SPECS:
        raise WpsAttachError({
            "ok": False,
            "reason": "app_type_unknown",
            "error": "无法判定当前 COM 实例属于 word/excel/powerpoint 中的哪一类",
            "hint": "请先调用 list_open_documents 确认人类正打开的 WPS 应用与文件路径。",
            "human_action_required": "请人类告知正在编辑的 WPS 应用类型与文件路径。",
        })

    target = _norm_com_path(os.path.abspath(path)) if path else None
    try:
        collected = list(_iter_open_docs(app, app_type))
    except Exception as first_err:
        # 缓存实例已失效（COM 调用报错）：清缓存 → 重走完整挂接链 → 重试一次
        _invalidate_app(app_type)
        try:
            app, _attached = _get_wps_app(app_type, allow_new_instance=False)
            collected = list(_iter_open_docs(app, app_type))
            _MATCH_INFO[app_type] = {"matched_by": "reattach", "reattach_reason": str(first_err)}
        except WpsAttachError:
            raise
        except Exception as second_err:
            raise WpsAttachError({
                "ok": False,
                "reason": "enumerate_failed",
                "app_type": app_type,
                "error": str(second_err),
                "first_error": str(first_err),
                "tried_methods": _MATCH_ATTEMPTS.get(app_type, []),
                "hint": "已挂接到 WPS 实例但无法枚举已打开文档，请人类确认 WPS 状态后重试。",
                "human_action_required": "请人类确认 WPS 是否处于可用状态（未模态弹窗阻塞、未处于兼容性故障）。",
            })

    if target:
        for obj, info in collected:
            if info["full_name"] and _norm_com_path(info["full_name"]) == target:
                _MATCH_INFO[app_type] = {
                    "matched_by": _MATCH_INFO.get(app_type, {}).get("matched_by", "path"),
                    "full_name": info["full_name"],
                }
                return obj
        # 名称解析被拒（多为人类正处于单元格编辑态）→ 三级救援
        if _docs_without_names(collected) or (not collected and _open_count(app, app_type) > 0):
            rescued = _locate_with_rescue(app, app_type, target, collected)
            if rescued is not None:
                return rescued
            raise WpsAttachError(_doc_busy_error(app_type, path, collected))
    elif len(collected) == 1:
        obj, info = collected[0]
        _MATCH_INFO[app_type] = {"matched_by": "single_open_doc", "full_name": info["full_name"]}
        return obj

    raise WpsAttachError(_doc_not_open_error(app_type, path, collected))


def _get_open_workbook(app, path: Optional[str] = None):
    """复用人类已打开的 WPS 表格工作簿（只复用不新开）。"""
    return _get_open_document(app, path, "excel")


def _get_open_presentation(app, path: Optional[str] = None):
    """复用人类已打开的 WPS 演示文稿（只复用不新开）。"""
    return _get_open_document(app, path, "powerpoint")


# ---------- 路径兜底（只读）：COM 不可用时也能拿到人类正在编辑的文件路径 ----------

def _title_to_doc_name(title: str, suffixes) -> str:
    """从 WPS 窗口标题里剥出文档名（如 "报告.docx - WPS 文字" → "报告.docx"）。"""
    if not title:
        return ""
    text = title.strip()
    for suffix in suffixes or []:
        stripped = False
        for sep in (" - ", " — ", " – ", " -- ", " | ", "-", " "):
            if text.endswith(sep + suffix):
                text = text[: -(len(sep) + len(suffix))].strip()
                stripped = True
                break
        if stripped:
            break
    for junk in _TITLE_JUNK:
        text = text.replace(junk, "")
    return text.strip().strip("*-").strip()


def _guess_paths_by_name(doc_name: str, limit: int = 20) -> list:
    """在常见目录里按文件名检索候选绝对路径；同时利用 ~$ 锁文件反推目标文件（只读）。"""
    if not doc_name:
        return []
    stem = Path(doc_name).stem or doc_name
    home = os.path.expanduser("~")
    roots = []
    for sub in ("Desktop", "Documents", "Downloads", "桌面", "文档", "下载",
                "OneDrive/Desktop", "OneDrive/文档", "OneDrive/桌面"):
        candidate_dir = os.path.join(home, *sub.split("/"))
        if os.path.isdir(candidate_dir):
            roots.append(candidate_dir)

    found = []
    seen = set()

    def _add(path_str):
        key = os.path.normcase(os.path.normpath(path_str))
        if key in seen or len(found) >= limit:
            return
        seen.add(key)
        found.append(path_str)

    for root in roots:
        visited = 0
        for dirpath, dirnames, filenames in os.walk(root):
            visited += 1
            if visited > 300 or len(found) >= limit:
                break
            if dirpath[len(root):].count(os.sep) >= 3:
                dirnames[:] = []
            for filename in filenames:
                if filename.startswith("~$"):
                    if stem and stem in filename:
                        _add(os.path.join(dirpath, doc_name))   # 锁文件反推同目录目标文件
                    continue
                if stem and stem in filename:
                    _add(os.path.join(dirpath, filename))
            if len(found) >= limit:
                break
        if len(found) >= limit:
            break

    recent_dir = os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows", "Recent")
    if stem and os.path.isdir(recent_dir) and len(found) < limit:
        try:
            for filename in os.listdir(recent_dir):
                if stem in filename:
                    _add(os.path.join(recent_dir, filename))
        except Exception:
            pass
    return found


def _top_windows_for(app_type: str) -> list:
    """某类应用的顶层候选窗口（类名命中 + 标题后缀命中，去重）。"""
    spec = _WPS_APP_SPECS.get(app_type, {})
    windows = _enum_top_windows(class_names=spec.get("top_classes"))
    seen = set(item[0] for item in windows)
    for item in _enum_top_windows(title_suffixes=spec.get("title_suffixes")):
        if item[0] not in seen:
            seen.add(item[0])
            windows.append(item)
    return windows


def _scan_open_documents_by_window() -> list:
    """COM 全链路不可用时的兜底（只读）：窗口标题 → 文档名 → 锁文件/同名文件检索出候选路径。"""
    hits = []
    for app_type, spec in _WPS_APP_SPECS.items():
        for hwnd, cls, title in _top_windows_for(app_type):
            doc_name = _title_to_doc_name(title, spec["title_suffixes"])
            hits.append({
                "app_type": app_type,
                "app_label": spec["label"],
                "window_title": title,
                "window_class": cls,
                "hwnd": int(hwnd),
                "document_name": doc_name,
                "candidate_paths": _guess_paths_by_name(doc_name),
                "resolved_by": "window_title_scan",
            })
    return hits


def _list_open_documents(app_type: Optional[str] = None) -> dict:
    """只读：列出人类当前在 WPS 中打开的文档（含未保存文档的路径）。

    主路径：挂接人类实例 → 遍历 Documents / Workbooks / Presentations（FullName / 标题 / 未保存标记）。
    兜底路径：COM 全链路不可用时枚举 WPS 窗口标题，并结合 ~$ 锁文件、桌面/文档/下载/最近使用同名文件推测候选路径。
    全程只读：不新建实例、不新开窗口、不修改任何文档、不改变人类编辑状态。
    """
    targets = [app_type] if app_type in _WPS_APP_SPECS else list(_WPS_APP_SPECS.keys())
    documents = []
    attach_errors = []
    window_scan = []

    def _scan():
        if not window_scan:
            try:
                window_scan.extend(_scan_open_documents_by_window())
            except Exception:
                pass
        return window_scan

    for target in targets:
        app = None
        try:
            app, _attached = _get_wps_app(target)          # 默认绝不新建实例
        except WpsAttachError as err:
            attach_errors.append(err.payload)
        except Exception as err:
            attach_errors.append({"ok": False, "reason": "attach_failed",
                                  "app_type": target, "error": str(err)})
        if app is None:
            for item in _scan():
                if item.get("app_type") == target:
                    item = dict(item)
                    item["attached"] = False
                    documents.append(item)
            continue
        try:
            for _obj, info in _iter_open_docs(app, target):
                info = dict(info)
                info["attached"] = True
                documents.append(info)
        except Exception as err:
            attach_errors.append({"ok": False, "reason": "enumerate_failed", "app_type": target,
                                  "error": str(err), "tried_methods": _MATCH_ATTEMPTS.get(target, [])})
            _invalidate_app(target)
            for item in _scan():
                if item.get("app_type") == target:
                    item = dict(item)
                    item["attached"] = False
                    documents.append(item)

    window_hits = [doc for doc in documents if not doc.get("attached")]
    result = {
        "ok": True,
        "count": len(documents),
        "documents": documents,
        "attach_errors": attach_errors,
        "fallback_candidates": window_hits,
    }
    if window_hits:
        result["fallback_note"] = ("fallback_candidates 由窗口标题 + ~$ 锁文件/同名文件推测得到，"
                                   "路径可能不准，且不含未保存的改动内容。")
    if not documents:
        result["human_action_required"] = ("未发现人类打开的 WPS 文档。请人类先打开需要协作的文件"
                                           "（若已打开，请把窗口标题或文件绝对路径告诉我）。")
    return result


def _com_read_docx(file_path: str) -> dict:
    """通过COM读取Word：默认挂接人类实例、复用其已打开的文档（只读，不保存、不新开窗口）"""
    app, attached = _get_wps_app("word")
    try:
        if not attached:
            app.Visible = False
        doc = _get_open_document(app, os.path.abspath(file_path), "word")
        content = doc.Content.Text
        if not attached:
            doc.Close()
        result = {"content": content}
        result.update(_match_hint("word"))
        return result
    finally:
        _release_app(app, "word", attached)


def _com_write_docx(file_path: str, content: str) -> str:
    """通过COM写入Word：挂接人类实例 → 复用其已打开的文档 → 全量替换内容 → 就地 Save。

    先挂接后落盘：连不上人类实例时直接结构化报错求助，绝不"无路径新建文档 + SaveAs"
    （会在人类窗口多出空白文档）；attached=True 时只允许 Save()，不得 SaveAs / Close / Quit。
    """
    app, attached = _get_wps_app("word")
    try:
        if not attached:
            app.Visible = False
        target = os.path.abspath(file_path)
        if not attached and not os.path.exists(target):
            from docx import Document
            Document().save(target)  # 仅自建实例场景落盘（挂接人类时绝不凭空造文件）
        doc = _get_open_document(app, target, "word")
        doc.Content.Text = content
        doc.Save()
        if not attached:
            doc.Close()
        return f"已保存: {target}"
    finally:
        _release_app(app, "word", attached)


def _com_read_xlsx(file_path: str, sheet_name: Optional[str] = None) -> dict:
    """通过COM读取Excel：默认挂接人类实例、复用其已打开的工作簿（只读，不保存、不新开窗口）。

    读取一律走"按名引用目标工作表"，**不会切换**人类当前激活的工作表；
    返回体额外给出 active_sheet_before / active_sheet_after / switched，
    便于上游确认"读另一张表时人类看到的表没有被切走"，且读到的是内存中的实时值。
    """
    app, attached = _get_wps_app("excel")
    try:
        if not attached:
            app.Visible = False
        wb = _get_open_workbook(app, os.path.abspath(file_path))
        active_before = _safe_sheet_name(wb)
        if sheet_name:
            ws = wb.Sheets(sheet_name)
        else:
            ws = wb.ActiveSheet
        _ok, used_range = _com_try_attr(ws, "UsedRange")
        _ok2, _raw = _com_try_attr(used_range, "Value") if used_range is not None else (False, None)
        data = _com_to_plain(_raw)
        sheet_title = _safe_com_str(ws, "Name", sheet_name or "")
        active_after = _safe_sheet_name(wb)
        saved = _safe_com_bool(wb, "Saved")
        if not attached:
            wb.Close()
        result = {
            "sheet_name": sheet_title,
            "requested_sheet": sheet_name,
            "active_sheet_before": active_before,
            "active_sheet_after": active_after,
            "switched": bool(active_before and active_after and active_before != active_after),
            "live": bool(attached),
            "saved": saved,
            "data": data,
        }
        result.update(_match_hint("excel"))
        return result
    finally:
        _release_app(app, "excel", attached)


def _com_attr(obj, name, default=None):
    """安全读 COM 对象属性。

    comtypes 找不到成员名时抛的是 NameError（"Name xxx not found"），而非 AttributeError，
    getattr(obj, name, default) 的默认值兜不住，必须显式捕获。
    """
    try:
        return getattr(obj, name, default)
    except Exception:
        return default


def _com_set_value(target, value):
    """给 COM 对象的「默认属性」（Excel Range.Value，DISPID=0）赋值。

    comtypes lazybind 的 __setattr__ 只接受 cParams==1 的 propput，本机 WPS 类型库下
    对 Range.Value 会直接抛 AttributeError('Value')；此处退化为直接以 DISPID 0 +
    PROPERTYPUT / PROPERTYPUTREF 调用底层 _invoke。win32com 等其他绑定不受影响。
    """
    try:
        target.Value = value
        return
    except Exception as first_err:
        direct_err = first_err
    inner = _com_attr(target, "_comobj")
    if not hasattr(inner, "_invoke"):
        disp = _com_attr(target, "disp")
        inner = _com_attr(disp, "_comobj")
    if inner is None or not hasattr(inner, "_invoke"):
        raise AttributeError("无法定位 COM 底层对象以写入 Value：%s" % direct_err)
    # DISPID_VALUE=0。invkind 先试 PROPERTYPUT(8)，失败回落 PROPERTYPUTREF(4)：
    # 本机 WPS 的 Range.Value 只接受 4（传 8 报 DISP_E_MEMBERNOTFOUND），
    # 而标准 Excel 走 8，故两路都保留。
    last_err = None
    for invkind in (8, 4):
        try:
            inner._invoke(0, invkind, 0, value)
            return
        except Exception as exc:
            last_err = exc
    raise last_err


def _col_to_index(col) -> Optional[int]:
    """列标识转 1-based 列号：支持 'G'/'g'、'AA'、7、'7'；None 原样返回。"""
    if col is None:
        return None
    if isinstance(col, bool):
        return None
    if isinstance(col, int):
        return int(col)
    text = str(col).strip().upper()
    if not text:
        return None
    if text.isdigit():
        return int(text)
    idx = 0
    for ch in text:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"非法列标识：{col!r}")
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def _col_letter(col_index: int) -> str:
    """1-based 列号转列字母（7 → 'G'）。"""
    n = int(col_index)
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


def _com_read_row(file_path: str, row: int, sheet_name: Optional[str] = None,
                  start_col=None, end_col=None) -> dict:
    """通过COM**只读取目标行**：挂接人类实例 → 复用其已打开工作簿 → 取该行在列范围内的值。

    与 read_document（整表 UsedRange）互补：本工具只取一行，返回体量与该行列数同阶，
    适合"只看某一行/某一记录"。严格只读：不 Activate、不 Save、不 Close（挂接态），
    人类当前激活的工作表与未保存修改均不受影响。

    列范围默认取目标表 UsedRange 的列区间；可用 start_col / end_col 覆盖（'G' 或 7）。
    """
    app, attached = _get_wps_app("excel")
    try:
        if not attached:
            app.Visible = False
        wb = _get_open_workbook(app, os.path.abspath(file_path))
        active_before = _safe_sheet_name(wb)
        if sheet_name:
            ws = wb.Sheets(sheet_name)
        else:
            ws = wb.ActiveSheet

        used = _com_attr(ws, "UsedRange")
        first_row = int(_com_to_plain(_com_attr(used, "Row", 1)) or 1)
        first_col = int(_com_to_plain(_com_attr(used, "Column", 1)) or 1)
        n_rows = int(_com_to_plain(_com_attr(_com_attr(used, "Rows"), "Count", 1)) or 1)
        n_cols = int(_com_to_plain(_com_attr(_com_attr(used, "Columns"), "Count", 1)) or 1)
        last_row = first_row + n_rows - 1
        last_col = first_col + n_cols - 1

        c1 = _col_to_index(start_col) or first_col
        c2 = _col_to_index(end_col) or last_col
        if c2 < c1:
            c1, c2 = c2, c1
        row_no = int(row)

        # 实测（本机 WPS）：Range(Cells(r,c1), Cells(r,c2)).Value 一次取整行最快且稳定；
        # 单格时 COM 返回标量而非二维元组，故统一展平成一维后与列号对齐。
        cell_l = _com_call(ws, "Cells", row_no, c1)
        cell_r = _com_call(ws, "Cells", row_no, c2)
        rng = _com_call(ws, "Range", cell_l, cell_r)
        raw = _com_to_plain(_com_attr(rng, "Value"))
        if isinstance(raw, list):
            if len(raw) == 1 and isinstance(raw[0], list):
                flat = list(raw[0])
            else:
                flat = list(raw)
        else:
            flat = [raw]
        flat = flat[: c2 - c1 + 1]
        while len(flat) < (c2 - c1 + 1):
            flat.append(None)

        columns = [
            {"column": _col_letter(ci), "column_index": ci, "value": v}
            for ci, v in zip(range(c1, c2 + 1), flat)
        ]
        non_empty = {
            _col_letter(ci): v
            for ci, v in zip(range(c1, c2 + 1), flat)
            if v is not None and v != ""
        }

        sheet_title = _safe_com_str(ws, "Name", sheet_name or "")
        active_after = _safe_sheet_name(wb)
        saved = _safe_com_bool(wb, "Saved")
        if not attached:
            wb.Close()

        result = {
            "sheet_name": sheet_title,
            "requested_sheet": sheet_name,
            "row": row_no,
            "row_in_used_range": bool(first_row <= row_no <= last_row),
            "used_range": {
                "first_row": first_row,
                "last_row": last_row,
                "first_column": _col_letter(first_col),
                "last_column": _col_letter(last_col),
                "rows": n_rows,
                "columns": n_cols,
            },
            "column_range": [_col_letter(c1), _col_letter(c2)],
            "values": flat,
            "columns": columns,
            "non_empty": non_empty,
            "active_sheet_before": active_before,
            "active_sheet_after": active_after,
            "switched": bool(active_before and active_after and active_before != active_after),
            "live": bool(attached),
            "saved": saved,
        }
        result.update(_match_hint("excel"))
        return result
    finally:
        _release_app(app, "excel", attached)


def _file_read_row(file_path: str, row: int, sheet_name: Optional[str] = None,
                   start_col=None, end_col=None) -> dict:
    """按行读取磁盘上的 Excel（openpyxl，读的是已保存版本，不含未保存修改）。"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    # 口径与 COM 模式对齐：列范围缺省取"数据区首列 .. 数据区末列"，行范围亦取数据区，
    # 避免同一文件在 file/com 两种模式下 values 的对齐基准不同（com 用 UsedRange，file 原先固定从 A 起）。
    min_col = max(int(getattr(ws, "min_column", 1) or 1), 1)
    max_col = max(int(getattr(ws, "max_column", 1) or 1), 1)
    min_row = max(int(getattr(ws, "min_row", 1) or 1), 1)
    max_row = max(int(getattr(ws, "max_row", 1) or 1), 1)
    c1 = _col_to_index(start_col) or min_col
    c2 = _col_to_index(end_col) or max(max_col, c1)
    if c2 < c1:
        c1, c2 = c2, c1
    row_no = int(row)
    flat = [ws.cell(row=row_no, column=ci).value for ci in range(c1, c2 + 1)]
    return {
        "sheet_name": ws.title,
        "requested_sheet": sheet_name,
        "row": row_no,
        "row_in_used_range": bool(min_row <= row_no <= max_row),
        "used_range": {
            "first_row": min_row,
            "last_row": max_row,
            "first_column": _col_letter(min_col),
            "last_column": _col_letter(max_col),
            "rows": max_row - min_row + 1,
            "columns": max_col - min_col + 1,
        },
        "column_range": [_col_letter(c1), _col_letter(c2)],
        "values": flat,
        "columns": [
            {"column": _col_letter(ci), "column_index": ci, "value": v}
            for ci, v in zip(range(c1, c2 + 1), flat)
        ],
        "non_empty": {
            _col_letter(ci): v
            for ci, v in zip(range(c1, c2 + 1), flat)
            if v is not None and v != ""
        },
        "live": False,
    }


def _com_activate_sheet(file_path: str, sheet_name: Optional[str] = None,
                        sheet_index: Optional[int] = None) -> dict:
    """通过COM切换（激活）Excel工作表：挂接人类实例 → 复用其已打开的工作簿 → 激活目标表。

    这是**会改变人类所见界面**的交互动作（人类当前工作表会跟着切），故列为写类工具；
    attached=True 时只 Activate，绝不 Save / Close / Quit，也绝不新建工作表。
    """
    if not sheet_name and not sheet_index:
        raise ValueError("必须提供 sheet_name 或 sheet_index 之一")
    app, attached = _get_wps_app("excel")
    try:
        if not attached:
            app.Visible = False
        wb = _get_open_workbook(app, os.path.abspath(file_path))
        active_before = _safe_sheet_name(wb)
        if sheet_index:
            ws = wb.Sheets(int(sheet_index))
        else:
            ws = wb.Sheets(sheet_name)
        target_name = _safe_com_str(ws, "Name", sheet_name or "")
        try:
            _com_call(ws, "Activate")
        except Exception:
            # 个别环境对 Sheet.Activate 名称解析不友好，退化为 Select 再试一次
            _com_call(ws, "Select")
        active_after = _safe_sheet_name(wb)
        if not attached:
            wb.Close()
        result = {
            "ok": True,
            "target_sheet": target_name,
            "active_sheet_before": active_before,
            "active_sheet_after": active_after,
            "switched": bool(active_after and active_after != active_before),
            "live": bool(attached),
        }
        result.update(_match_hint("excel"))
        return result
    finally:
        _release_app(app, "excel", attached)


def _com_write_xlsx(file_path: str, sheet_name: Optional[str], data: list) -> str:
    """通过COM写入Excel：挂接人类实例 → 复用其已打开的工作簿 → 清空后按二维数组一次性写入，就地 Save。

    先挂接后落盘：连不上人类实例时直接结构化报错求助，绝不"无路径新建工作簿 + SaveAs"
    （会在人类窗口多出空白工作簿）；attached=True 时只允许 Save()，不得 SaveAs / Close / Quit。
    """
    app, attached = _get_wps_app("excel")
    try:
        if not attached:
            app.Visible = False
        target = os.path.abspath(file_path)
        if not attached and not os.path.exists(target):
            from openpyxl import Workbook
            nb = Workbook()
            nb.active.title = sheet_name or "Sheet"  # 表名与后续 COM 定位保持一致
            nb.save(target)  # 仅自建实例场景落盘（挂接人类时绝不凭空造文件）
        wb = _get_open_workbook(app, target)
        ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
        try:
            ws.Cells.ClearContents()
        except Exception:
            ws.Cells.Clear()  # 本机 WPS 表格 IDispatch 不暴露 ClearContents，退回 Clear
        rows = [list(row) for row in (data or [])]
        if rows:
            width = max(len(row) for row in rows)
            matrix = tuple(tuple(row + [None] * (width - len(row))) for row in rows)
            _com_set_value(ws.Range(ws.Cells(1, 1), ws.Cells(len(rows), width)), matrix)
        wb.Save()
        if not attached:
            wb.Close()
        return f"已保存: {target}"
    finally:
        _release_app(app, "excel", attached)


def _com_read_pptx(file_path: str) -> dict:
    """通过COM读取PowerPoint：默认挂接人类实例、复用其已打开的演示文稿（只读，不保存、不新开窗口）"""
    app, attached = _get_wps_app("powerpoint")
    try:
        if not attached:
            app.Visible = False
        prs = _get_open_presentation(app, os.path.abspath(file_path))
        slides_data = []
        for slide in prs.Slides:
            slide_content = []
            for shape in slide.Shapes:
                if shape.HasTextFrame:
                    slide_content.append(shape.TextFrame.TextRange.Text)
            slides_data.append(slide_content)
        if not attached:
            prs.Close()
        result = {"slides": slides_data}
        result.update(_match_hint("powerpoint"))
        return result
    finally:
        _release_app(app, "powerpoint", attached)


def _com_write_pptx(file_path: str, content: str) -> str:
    """通过COM写入PowerPoint：挂接人类实例 → 复用其已打开的演示文稿 → 全量替换，就地 Save。

    先挂接后落盘：连不上人类实例时直接结构化报错求助，绝不"无路径新建演示文稿 + SaveAs"
    （会在人类窗口多出空白演示文稿）；attached=True 时只允许 Save()。
    """
    app, attached = _get_wps_app("powerpoint")
    try:
        if not attached:
            app.Visible = False
        target = os.path.abspath(file_path)
        if not attached and not os.path.exists(target):
            from pptx import Presentation
            Presentation().save(target)  # 仅自建实例场景落盘（挂接人类时绝不凭空造文件）
        prs = _get_open_presentation(app, target)
        while prs.Slides.Count > 0:
            prs.Slides(1).Delete()
        slide = prs.Slides.Add(1, 12)  # ppLayoutBlank
        shape = slide.Shapes.AddTextbox(1, 100, 100, 600, 100)
        shape.TextFrame.TextRange.Text = content
        prs.Save()
        if not attached:
            prs.Close()
        return f"已保存: {target}"
    finally:
        _release_app(app, "powerpoint", attached)


def _com_insert_image_word(file_path: str, image_path: str) -> str:
    """通过COM在Word中插入图片（复用人类已打开文档，就地 Save）"""
    app, attached = _get_wps_app("word")
    try:
        if not attached:
            app.Visible = False
        doc = _get_open_document(app, os.path.abspath(file_path), "word")
        selection = app.Selection
        selection.InlineShapes.AddPicture(os.path.abspath(image_path))
        doc.Save()
        if not attached:
            doc.Close()
        return f"已插入图片: {image_path}"
    finally:
        _release_app(app, "word", attached)


def _com_insert_image_excel(file_path: str, sheet_name: Optional[str], image_path: str, range_addr: str) -> str:
    """通过COM在Excel中插入图片（复用人类已打开工作簿，就地 Save）"""
    app, attached = _get_wps_app("excel")
    try:
        if not attached:
            app.Visible = False
        wb = _get_open_workbook(app, os.path.abspath(file_path))
        ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
        rng = ws.Range(range_addr)
        ws.Shapes.AddPicture(os.path.abspath(image_path), True, True, rng.Left, rng.Top, rng.Width, rng.Height)
        wb.Save()
        if not attached:
            wb.Close()
        return f"已插入图片到{range_addr}"
    finally:
        _release_app(app, "excel", attached)


def _com_run_macro(file_path: str, macro_name: str, args: Optional[list] = None) -> str:
    """通过COM运行WPS宏（复用人类已打开文档）。

    V3.1 三处收紧：
      ① **按扩展名分流应用**：原实现写死 `_get_wps_app("word")`，对 .xlsx 必然挂不上 → 报错。
      ② **默认禁用**（`WPS_MCP_ALLOW_MACRO`，默认 0）：`Application.Run` 能调用文档/加载项里
         任何已存在的宏（不能注入新代码，但"能跑哪段宏"完全由文件决定），属于代码执行面。
         要用于是刻意开启；报错信息里带了开启方法，不是把你锁在门外。
      ③ **不无条件 Save**：原来无论宏做了什么都会 `doc.Save()`，会把人类未保存的编辑一起存掉
         （等于替人按了 Ctrl+S）。现在跟随 `WPS_MCP_AUTOSAVE`，默认只改内存、由人自己 Ctrl+S。
    """
    if os.environ.get("WPS_MCP_ALLOW_MACRO", "0") in ("0", "false", "False", ""):
        raise ValueError(
            "run_macro 默认被禁用（宏能执行文档/加载项里任意已存在的宏，属代码执行面）。"
            "确认要用，请在启动 MCP 的环境变量里设 WPS_MCP_ALLOW_MACRO=1 后重启服务。")
    if not str(macro_name or "").strip():
        raise ValueError("run_macro 需要非空的 macro_name（宏名，通常是 '模块名.过程名'）")
    ext = Path(file_path).suffix.lower()
    app_type = ("word" if ext in (".doc", ".docx")
                else "powerpoint" if ext in (".ppt", ".pptx")
                else "excel" if ext in (".xls", ".xlsx") else None)
    if app_type is None:
        raise ValueError("run_macro 不支持的扩展名：%s（支持 doc/docx/xls/xlsx/ppt/pptx）" % (ext,))
    app, attached = _get_wps_app(app_type)
    try:
        if not attached:
            app.Visible = False
        doc = _get_open_document(app, os.path.abspath(file_path), app_type)
        if args:
            result = app.Run(macro_name, *args)
        else:
            result = app.Run(macro_name)
        saved = False
        if os.environ.get("WPS_MCP_AUTOSAVE", "0") not in ("0", "false", "False", ""):
            doc.Save()
            saved = True
        if not attached:
            doc.Close()
        return "宏执行结果: %s（%s）" % (
            result, "已保存" if saved else "未自动保存，改动只在内存里，请人自行 Ctrl+S")
    finally:
        _release_app(app, app_type, attached)


def _com_export_pdf(file_path: str, output_path: Optional[str] = None) -> str:
    """通过COM导出PDF：复用人类已打开的文档导出，attached 情况下绝不 SaveAs 改动原文档"""
    ext = Path(file_path).suffix.lower()
    if ext in ['.doc', '.docx']:
        app, attached = _get_wps_app("word")
        try:
            if not attached:
                app.Visible = False
            doc = _get_open_document(app, os.path.abspath(file_path), "word")
            if not output_path:
                output_path = str(Path(file_path).with_suffix('.pdf'))
            _com_export_pdf_file(app, doc, os.path.abspath(file_path), str(output_path), attached=attached)
            if not attached:
                doc.Close()
            return f"已导出PDF: {output_path}"
        finally:
            _release_app(app, "word", attached)
    elif ext in ['.xls', '.xlsx']:
        app, attached = _get_wps_app("excel")
        try:
            if not attached:
                app.Visible = False
            wb = _get_open_workbook(app, os.path.abspath(file_path))
            if not output_path:
                output_path = str(Path(file_path).with_suffix('.pdf'))
            _com_export_pdf_file(app, wb, os.path.abspath(file_path), str(output_path), attached=attached)
            if not attached:
                wb.Close()
            return f"已导出PDF: {output_path}"
        finally:
            _release_app(app, "excel", attached)
    elif ext in ['.ppt', '.pptx']:
        app, attached = _get_wps_app("powerpoint")
        try:
            if not attached:
                app.Visible = False
            prs = _get_open_presentation(app, os.path.abspath(file_path))
            if not output_path:
                output_path = str(Path(file_path).with_suffix('.pdf'))
            try:
                _com_export_pdf_file(app, prs, os.path.abspath(file_path), str(output_path), attached=attached)
            except Exception:
                if attached:
                    raise  # 挂接人类实例：绝不 SaveAs，避免改动人类正在编辑的演示文稿
                prs.SaveAs(os.path.abspath(output_path), 32)  # ppSaveAsPDF
            if not attached:
                prs.Close()
            return f"已导出PDF: {output_path}"
        finally:
            _release_app(app, "powerpoint", attached)
    else:
        raise ValueError(f"不支持的文件格式: {ext}")


# ==================== V3 薄壳接线（新增；老地基与老工具一个都没动） ====================
# 详细设计见同目录 wps_session.py。这里只做三件事：
#   ① tools/list 默认只广播 2 个极简工具（分层加载，省常驻 token）
#   ② wps_do / wps_docs 的分发
#   ③ 按 action 判定本次是读还是写，喂给已有的 _CALL_MODE 救援分叉
# 注意：server.py 是以 `python server.py` 运行的，本模块在 sys.modules 里的名字是 "__main__"，
# 子模块若 `import server` 会**再执行一遍**整个文件（两份 _APP_CACHE）。所以把本模块对象当 host 传进去。
_WPS_SESSION = None
_THIN_APP_TYPE = {"excel": "excel", "word": "word", "ppt": "powerpoint"}
_THIN_TOOL_NAMES = ("wps_docs", "wps_do")


def _host_module():
    """取本模块对象。非注册方式加载（测试用 importlib）时 sys.modules 里没有，退化为 __main__。"""
    return sys.modules.get(__name__) or sys.modules.get("__main__")


def _thin_session():
    """惰性加载薄壳模块。加载失败不致命：tools/list 会退回全量工具，老路照跑。"""
    global _WPS_SESSION
    if _WPS_SESSION is None:
        here = os.path.dirname(os.path.abspath(__file__))
        if here not in sys.path:
            sys.path.insert(0, here)
        import wps_session
        _WPS_SESSION = wps_session
    return _WPS_SESSION


def _advertised_tools():
    """分层加载开关：默认只广播薄工具；WPS_MCP_LEGACY_TOOLS=1 退回全量工具。"""
    if os.environ.get("WPS_MCP_LEGACY_TOOLS", "0") not in ("0", "false", "False", ""):
        return TOOLS
    try:
        return _thin_session().THIN_TOOLS
    except Exception as err:
        sys.stderr.write("[wps-editor] 薄工具加载失败，本次回退全量工具：%r\n" % (err,))
        sys.stderr.flush()
        return TOOLS


def _handle_thin(name, arguments):
    """wps_docs / wps_do 的分发入口。"""
    session = _thin_session()
    arguments = arguments or {}
    app, action = "", ""
    try:
        if name == "wps_docs":
            app = (arguments.get("app") or "all").strip().lower()
            return _list_open_documents(None if app in ("", "all") else _THIN_APP_TYPE.get(app, app))

        global _CALL_MODE
        app = (arguments.get("app") or "").strip().lower()
        action = (arguments.get("action") or "").strip().lower()
        if session.is_write(app, action):
            _CALL_MODE = "write"
        inner = dict(arguments.get("args") or {})
        # 容错：模型偶尔把参数平铺在顶层（ref/values/sheet 直接挂在 arguments 下）
        for key, value in arguments.items():
            if key not in ("app", "action", "args", "path") and key not in inner:
                inner[key] = value
        return session.handle(_host_module(), app, action, inner, arguments.get("path"))
    except Exception as err:
        # 关键：JSON-RPC 里未捕获的异常会变成裸错误（-32000 …发生意外），人读不懂、AI 学不到。
        # 统一翻成结构化人话 + 下一步动作后再返回。
        sys.stderr.write("[wps-editor] 薄工具出错：%r\n" % (err,))
        sys.stderr.flush()
        return session.error_payload(_host_module(), app, action, err)


# ==================== V3.1 写盘前的「人正开着」冲突闸 + 颜色规范化 ====================
# 背景（2026-09-28 复核评审意见后落地）：
#   ① write_document / batch_write / batch_set_style / apply_template 走的是**文件模式**，
#      直接改写磁盘文件；而人可能正在 WPS 里编辑**同一份文件的内存副本**。两边互不知道，
#      谁后保存谁覆盖 → 静默丢数据。所以写盘前先查一次"这份文件是不是人正开着"。
#   ② set_cell_style 的颜色：COM 用 BGR 整数（且 int("#FF0000",16) 必抛）、openpyxl 用
#      8 位 aRGB 字符串，两边红蓝极易颠倒；下划线 COM 要枚举、openpyxl 要 'single'/'double'。
#      统一从一个 '#RRGGBB' 入口解析，再各按各家规矩转换。

def _live_conflict(file_path: str):
    """目标文件是否正在人类的 WPS 实例里打开。命中返回文档信息 dict，否则 None。

    只读、绝不新建实例、绝不改状态。**一律不抛错**：挂不上人类实例（没开 WPS / 没权限）
    就当作没冲突——那种情况下文件模式写盘恰恰是对的做法。
    """
    try:
        target = _norm_com_path(os.path.abspath(file_path))
    except Exception:
        return None
    ext = Path(str(file_path)).suffix.lower()
    app_type = ("word" if ext in (".doc", ".docx")
                else "powerpoint" if ext in (".ppt", ".pptx") else "excel")
    try:
        app, _attached = _get_wps_app(app_type)
    except Exception:
        return None
    try:
        for _obj, info in _iter_open_docs(app, app_type):
            if info.get("full_name") and _norm_com_path(info["full_name"]) == target:
                return {"app_type": app_type, "name": info.get("name"),
                        "full_name": info.get("full_name"),
                        "window_title": info.get("window_title"),
                        "saved": info.get("saved")}
    except Exception:
        return None
    return None


def _guard_file_write(tool: str, file_path: str):
    """文件模式写盘前的统一闸：命中"人正开着"就抛人话错误，别把人未保存的副本覆盖掉。"""
    hit = _live_conflict(file_path)
    if not hit:
        return
    raise ValueError(
        "%s 走的是「文件模式」——直接改写磁盘上的 %s，而这份文件**人正开着**"
        "（窗口标题 %r，未保存改动=%s）。两边互不知道，谁后保存谁覆盖，会静默丢数据。请二选一：\n"
        "  ① 让人先 Ctrl+S 保存并关闭这份文件，再重试；\n"
        "  ② 要改「人正打开的那一份」，改用 wps_do 写活文档（内存实时值、可 Ctrl+Z）：\n"
        "     wps_do(app=\"excel\", action=\"set\", args={\"ref\": \"3:3\", \"values\": [\"a\",\"b\"]})\n"
        "     （整行 ref=\"3:3\"、整列 ref=\"C:C\"、一片 ref=\"A2:C3\"；写前进 action=\"where\" 看一眼）"
        % (tool, hit["full_name"] or file_path, hit["window_title"], hit["saved"]))


def _parse_hex_color(value: str):
    """把颜色统一解析成 (r, g, b)。

    接受 '#RRGGBB' / 'RRGGBB' / '#AARRGGBB' / 'AARRGGBB'（大小写均可）。
    原实现直接 int(color, 16)：带 '#' 必抛 ValueError，不带又会被 COM 当 BGR 解释 → 写错颜色。
    """
    text = str(value or "").strip()
    if text.startswith("#"):
        text = text[1:]
    elif text[:2].lower() == "0x":
        text = text[2:]
    text = text.strip()
    # 注意：这里**不能**用 lstrip("0x")——它会把 '0000FF' 的前导 0 一起吃掉，
    # 于是 '#0000FF' 会变成 'FF'（本次自测就踩到了）。
    if len(text) == 6:
        text = "FF" + text
    if len(text) != 8:
        raise ValueError(
            "颜色要写成 '#RRGGBB'（如 '#FF0000' 红）或 8 位 'AARRGGBB'，收到 %r" % (value,))
    try:
        n = int(text, 16)
    except Exception:
        raise ValueError("颜色不是合法的十六进制：%r（应形如 '#FF0000'）" % (value,))
    return ((n >> 16) & 0xFF, (n >> 8) & 0xFF, n & 0xFF)


def _com_color(value: str) -> int:
    """'#RRGGBB' → COM 要的 BGR 整数（Excel/WPS 的 Color 属性是 BGR）。"""
    r, g, b = _parse_hex_color(value)
    return r | (g << 8) | (b << 16)


def _argb_hex(value: str) -> str:
    """'#RRGGBB' → openpyxl 要的 8 位 aRGB（不补不透明位会被 openpyxl 拒绝）。"""
    r, g, b = _parse_hex_color(value)
    return "FF%02X%02X%02X" % (r, g, b)


def _underline_pair(value):
    """下划线：JSON 布尔/字符串 → (COM 枚举值, openpyxl 值)。

    COM 要 xlUnderlineStyle（-4142 无 / 2 单线 / -4119 双线），openpyxl 要 None/'single'/'double'。
    原实现把 JSON 的 True 直接丢给 COM，等于传了 1（非法枚举）。
    """
    if value is None:
        return None, None
    text = str(value).strip().lower()
    if value is False or text in ("false", "0", "none", ""):
        return -4142, None
    if text in ("double", "2"):
        return -4119, "double"
    return 2, "single"


def _tmp_pdf_path(path: str) -> str:
    """给 PDF 产出一个同目录临时名（必须以 .pdf 结尾，导出接口靠后缀判格式）。"""
    d, name = os.path.split(str(path))
    stem = name[:-4] if name.lower().endswith(".pdf") else name
    return os.path.join(d, "%s.__wps_tmp_%d.pdf" % (stem, os.getpid()))


# ==================== 工具定义 ====================

TOOLS = [
    {
        "name": "list_open_documents",
        "description": ("只读：列出人类当前在 WPS 中打开的文档（word/excel/powerpoint），"
                        "返回每个文档的名称、绝对路径 FullName、窗口标题、是否已保存/只读。"
                        "用于确认人类正在编辑的文件路径后再读写；全程只读，"
                        "不新建实例、不新开窗口、不修改任何文档。"
                        "COM 挂接失败时会返回结构化求助信息与窗口标题推测的候选路径。"),
        "inputSchema": {
            "type": "object",
            "properties": {
                "app_type": {"type": "string", "enum": ["word", "excel", "powerpoint", "all"],
                             "description": "限定应用类型，默认 all（三类全部探测）", "default": "all"}
            },
            "required": []
        }
    },
    {
        "name": "read_document",
        "description": "读取Office文档内容（Word/Excel/PowerPoint）",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文件路径"},
                "mode": {"type": "string", "enum": ["file", "com"], "description": "编辑模式", "default": "file"},
                "sheet_name": {"type": "string", "description": "Excel工作表名称"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "write_document",
        "description": ("写入Office文档内容。**注意：语义是整篇替换**（Excel 会先清空整表再写、"
                        "并立刻保存）。目标文件正被人打开时必须显式传 replace_all=true，"
                        "否则会被拒绝；只想改几个格子请用 wps_do 的 get/set。"),
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文件路径"},
                "content": {"type": "string", "description": "写入内容（Excel 传二维数组的 JSON）"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"},
                "sheet_name": {"type": "string", "description": "Excel工作表名称"},
                "append": {"type": "boolean",
                           "description": "仅 Word+file 支持追加；其余组合传 true 会被拒绝",
                           "default": False},
                "replace_all": {"type": "boolean",
                                "description": "目标正被人打开时，必须显式传 true 才允许整篇替换",
                                "default": False}
            },
            "required": ["file_path", "content"]
        }
    },
    {
        "name": "edit_cell",
        "description": "编辑Excel单元格",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "row": {"type": "integer", "description": "行号"},
                "column": {"type": "integer", "description": "列号"},
                "value": {"type": "string", "description": "单元格值"},
                "sheet_name": {"type": "string", "description": "工作表名称"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "row", "column", "value"]
        }
    },
    {
        "name": "get_sheet_names",
        "description": "获取Excel工作表名称列表",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "activate_sheet",
        "description": "切换（激活）Excel工作表：把人类 WPS 表格当前显示的工作表切到指定表；会改变人类所见界面",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "sheet_name": {"type": "string", "description": "目标工作表名称"},
                "sheet_index": {"type": "integer", "description": "目标工作表序号（从1开始），与 sheet_name 二选一"},
                "mode": {"type": "string", "enum": ["com"], "default": "com"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "read_row",
        "description": "按行读取Excel：只取目标行的值，不读整表。默认走COM读人类已打开表格的实时值，且不切换其当前工作表（适合“只看某一行/某一条记录”）；列范围默认取该表数据区，可用 start_col/end_col 限定。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "row": {"type": "integer", "description": "要读取的行号（从1开始）"},
                "sheet_name": {"type": "string", "description": "工作表名称，缺省用当前激活表"},
                "start_col": {"type": ["string", "integer"], "description": "起始列，支持 'G' 或 7，缺省用数据区首列"},
                "end_col": {"type": ["string", "integer"], "description": "结束列，支持 'H' 或 8，缺省用数据区末列"},
                "mode": {"type": "string", "enum": ["com", "file"], "description": "com=读已打开表格的实时值（默认）；file=读磁盘已保存版本", "default": "com"}
            },
            "required": ["file_path", "row"]
        }
    },
    {
        "name": "add_rows",
        "description": "在Excel中添加行",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "count": {"type": "integer", "description": "添加行数"},
                "after_row": {"type": "integer", "description": "在第几行后添加"},
                "sheet_name": {"type": "string", "description": "工作表名称"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "count"]
        }
    },
    {
        "name": "add_columns",
        "description": "在Excel中添加列",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "count": {"type": "integer", "description": "添加列数"},
                "after_col": {"type": "integer", "description": "在第几列后添加"},
                "sheet_name": {"type": "string", "description": "工作表名称"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "count"]
        }
    },
    {
        "name": "merge_cells",
        "description": "合并Excel单元格",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "start_row": {"type": "integer", "description": "起始行"},
                "start_col": {"type": "integer", "description": "起始列"},
                "end_row": {"type": "integer", "description": "结束行"},
                "end_col": {"type": "integer", "description": "结束列"},
                "sheet_name": {"type": "string", "description": "工作表名称"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "start_row", "start_col", "end_row", "end_col"]
        }
    },
    {
        "name": "unmerge_cells",
        "description": "取消合并Excel单元格",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "start_row": {"type": "integer", "description": "起始行"},
                "start_col": {"type": "integer", "description": "起始列"},
                "end_row": {"type": "integer", "description": "结束行"},
                "end_col": {"type": "integer", "description": "结束列"},
                "sheet_name": {"type": "string", "description": "工作表名称"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "start_row", "start_col", "end_row", "end_col"]
        }
    },
    {
        "name": "set_column_width",
        "description": "设置Excel列宽",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "column": {"type": "string", "description": "列字母（如A、B）"},
                "width": {"type": "number", "description": "宽度"},
                "sheet_name": {"type": "string", "description": "工作表名称"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "column", "width"]
        }
    },
    {
        "name": "set_row_height",
        "description": "设置Excel行高",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "row": {"type": "integer", "description": "行号"},
                "height": {"type": "number", "description": "高度"},
                "sheet_name": {"type": "string", "description": "工作表名称"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "row", "height"]
        }
    },
    {
        "name": "insert_image",
        "description": "插入图片到Office文档",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Office文件路径"},
                "image_path": {"type": "string", "description": "图片文件路径"},
                "width": {"type": "number", "description": "宽度（英寸）"},
                "height": {"type": "number", "description": "高度（英寸）"},
                "anchor": {"type": "string", "description": "Excel锚点单元格（如A1）", "default": "A1"},
                "slide_index": {"type": "integer", "description": "PowerPoint幻灯片索引（-1新建）", "default": -1},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "image_path"]
        }
    },
    {
        "name": "set_cell_style",
        "description": "设置Excel单元格样式",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "row": {"type": "integer", "description": "行号"},
                "column": {"type": "integer", "description": "列号"},
                "font_name": {"type": "string", "description": "字体名称"},
                "font_size": {"type": "integer", "description": "字体大小"},
                "bold": {"type": "boolean", "description": "是否加粗"},
                "italic": {"type": "boolean", "description": "是否斜体"},
                "underline": {"type": "boolean", "description": "是否下划线"},
                "color": {"type": "string", "description": "字体颜色（如FF0000）"},
                "bg_color": {"type": "string", "description": "背景颜色"},
                "align_h": {"type": "string", "description": "水平对齐（left/center/right）"},
                "align_v": {"type": "string", "description": "垂直对齐（top/center/bottom）"},
                "sheet_name": {"type": "string", "description": "工作表名称"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "row", "column"]
        }
    },
    {
        "name": "set_paragraph_style",
        "description": "设置Word段落样式",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Word文件路径"},
                "para_index": {"type": "integer", "description": "段落索引"},
                "font_name": {"type": "string", "description": "字体名称（如SimHei、SimSun、KaiTi、FangSong）"},
                "font_size": {"type": "integer", "description": "字体大小"},
                "bold": {"type": "boolean", "description": "是否加粗"},
                "italic": {"type": "boolean", "description": "是否斜体"},
                "underline": {"type": "boolean", "description": "是否下划线"},
                "color": {"type": "string", "description": "字体颜色（如FF0000红色、0000FF蓝色）"},
                "align": {"type": "string", "description": "对齐方式（left/center/right/justify）"},
                "mode": {"type": "string", "enum": ["file", "com"], "default": "file"}
            },
            "required": ["file_path", "para_index"]
        }
    },
    {
        "name": "export_pdf",
        "description": "导出Office文档为PDF",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Office文件路径"},
                "output_path": {"type": "string", "description": "输出PDF路径"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "run_macro",
        "description": "运行WPS宏",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Office文件路径"},
                "macro_name": {"type": "string", "description": "宏名称"},
                "args": {"type": "array", "description": "宏参数", "items": {}}
            },
            "required": ["file_path", "macro_name"]
        }
    },
    {
        "name": "doc_to_ppt",
        "description": "根据Word文档自动生成专业PPT演示文稿",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_path": {"type": "string", "description": "Word文档路径"},
                "ppt_path": {"type": "string", "description": "输出PPT路径"},
                "template": {"type": "string", "description": "模板样式", "enum": ["ocean", "forest", "sunset", "royal", "minimal", "blue", "green"], "default": "ocean"}
            },
            "required": ["doc_path", "ppt_path"]
        }
    },
    {
        "name": "ppt_to_doc",
        "description": "读取PPT生成Word总结文档",
        "inputSchema": {
            "type": "object",
            "properties": {
                "ppt_path": {"type": "string", "description": "PPT文件路径"},
                "doc_path": {"type": "string", "description": "输出Word路径"}
            },
            "required": ["ppt_path", "doc_path"]
        }
    },
    {
        "name": "batch_read",
        "description": "批量读取文档内容（多个段落或单元格）",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文件路径"},
                "indices": {"type": "array", "description": "Word段落索引列表或Excel单元格列表（如['A1','B2']）", "items": {}},
                "sheet_name": {"type": "string", "description": "Excel工作表名称"}
            },
            "required": ["file_path", "indices"]
        }
    },
    {
        "name": "batch_write",
        "description": "批量写入Excel单元格",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel文件路径"},
                "data": {"type": "array", "description": "数据列表，格式: [{\"cell\":\"A1\",\"value\":\"xxx\"}]", "items": {}},
                "sheet_name": {"type": "string", "description": "工作表名称"}
            },
            "required": ["file_path", "data"]
        }
    },
    {
        "name": "batch_set_style",
        "description": "批量设置样式（Word段落或Excel单元格）",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文件路径"},
                "operations": {"type": "array", "description": "操作列表", "items": {}},
                "sheet_name": {"type": "string", "description": "Excel工作表名称"}
            },
            "required": ["file_path", "operations"]
        }
    },
    {
        "name": "check_quality",
        "description": "检查文档质量并提供改进建议",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文件路径"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "extract_template",
        "description": "提取文档模板（字体、样式、结构）",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文件路径"}
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "apply_template",
        "description": "应用模板到文档",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "文件路径"},
                "font_name": {"type": "string", "description": "目标字体"},
                "font_size": {"type": "integer", "description": "目标字号"}
            },
            "required": ["file_path"]
        }
    }
]


def handle_tool_call(name: str, arguments: dict) -> Any:
    """处理工具调用"""
    # 决定本次调用的"意图"：读还是写。它决定遇到编辑态时走哪条救援路线——
    #   读：优先无按键旁路（秒回、完全不打扰人类；编辑态下读是能读到的）
    #   写：等 BUSY_RETRY_SECONDS（默认 5s，给人松手的机会）→ 仍不行就发 ESC 挤出编辑态再写
    # 依据：实测编辑态下 WPS **读可以、写一律拒绝**，所以写拿不到"不打扰"的方案。
    global _CALL_MODE
    _CALL_MODE = "write" if name in _WRITE_TOOLS else "read"

    # V3 薄壳：两个入口工具，内部的读写意图由 _handle_thin 按 action 重设 _CALL_MODE。
    if name in _THIN_TOOL_NAMES:
        return _handle_thin(name, arguments)

    if name == "list_open_documents":
        app_type = arguments.get("app_type") or "all"
        if app_type not in _WPS_APP_SPECS:
            app_type = None
        return _list_open_documents(app_type)

    if name == "read_document":
        file_path = os.path.abspath(arguments["file_path"])
        mode = arguments.get("mode", "file")
        sheet_name = arguments.get("sheet_name")
        ext = Path(file_path).suffix.lower()
        edit_mode = get_mode(mode)
        
        if ext in ['.docx', '.doc']:
            if edit_mode == EditMode.COM:
                return _com_read_docx(file_path)
            return _file_read_docx(file_path)
        elif ext in ['.xlsx', '.xls']:
            if edit_mode == EditMode.COM:
                return _com_read_xlsx(file_path, sheet_name)
            return _file_read_xlsx(file_path, sheet_name)
        elif ext in ['.pptx', '.ppt']:
            if edit_mode == EditMode.COM:
                return _com_read_pptx(file_path)
            return _file_read_pptx(file_path)
        else:
            raise ValueError(f"不支持的文件格式: {ext}")
    
    elif name == "write_document":
        file_path = os.path.abspath(arguments["file_path"])
        content = arguments["content"]
        mode = arguments.get("mode", "file")
        sheet_name = arguments.get("sheet_name")
        append = arguments.get("append", False)
        ext = Path(file_path).suffix.lower()
        edit_mode = get_mode(mode)
        
        os.makedirs(os.path.dirname(file_path) or '.', exist_ok=True)
        
        # ---- V3.1 闸①：append 只在「Word + 文件模式」下真的生效（_file_write_docx）；
        #      xlsx/pptx 两种模式都没有追加语义，原实现**静默忽略** → 调用方以为在追加，
        #      实际是整篇清空重写。这是比"没有备份"更直接会伤数据的一条。
        if append and not (ext in ['.docx', '.doc'] and edit_mode != EditMode.COM):
            raise ValueError(
                "append 只在「Word + mode=\"file\"」下生效；%s + mode=%r 没有追加语义，"
                "原实现会静默清空重写。要追加请改用 Word+file，或把完整内容一次写全。"
                % (ext or "该格式", mode))
        # ---- V3.1 闸②：目标正被人开着时，file 与 com 都会打爆它（file 覆盖磁盘、
        #      com 清空内存并立刻 Save）。除非调用方显式声明"我就是要整篇替换"。
        _hit = _live_conflict(file_path)
        if _hit and not arguments.get("replace_all"):
            raise ValueError(
                "write_document 会把 %s **整篇替换**（清空后重写）并立刻保存，而这份文件"
                "**人正开着**（窗口标题 %r，未保存改动=%s）——人没保存的内容会被一起盖掉。\n"
                "确认要整篇替换就显式传 \"replace_all\": true；否则请改用 wps_do 只改要改的格子：\n"
                "  wps_do(app=\"excel\", action=\"set\", args={\"ref\": \"B5\", \"values\": \"x\"})\n"
                "  （整行 ref=\"3:3\"、一片 ref=\"A2:C3\"；写前先 action=\"where\" 看一眼）"
                % (_hit["full_name"] or file_path, _hit["window_title"], _hit["saved"]))
        
        if ext in ['.docx', '.doc']:
            if edit_mode == EditMode.COM:
                return _com_write_docx(file_path, content)
            return _file_write_docx(file_path, content, append)
        elif ext in ['.xlsx', '.xls']:
            data = json.loads(content)
            if edit_mode == EditMode.COM:
                return _com_write_xlsx(file_path, sheet_name, data)
            return _file_write_xlsx(file_path, sheet_name, data)
        elif ext in ['.pptx', '.ppt']:
            if edit_mode == EditMode.COM:
                return _com_write_pptx(file_path, content)
            return _file_write_pptx(file_path, content)
        else:
            raise ValueError(f"不支持的文件格式: {ext}")
    
    elif name == "edit_cell":
        file_path = os.path.abspath(arguments["file_path"])
        row = arguments["row"]
        column = arguments["column"]
        value = arguments["value"]
        sheet_name = arguments.get("sheet_name")
        mode = arguments.get("mode", "file")
        edit_mode = get_mode(mode)
        
        if edit_mode == EditMode.COM:
            app, attached = _get_wps_app("excel")
            try:
                if not attached:
                    app.Visible = False
                wb = _get_open_workbook(app, file_path)
                ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
                _com_set_value(ws.Cells(row, column), value)
                wb.Save()
                if not attached:
                    wb.Close()
                if _com_attr(wb, "save_skipped", False):
                    return ("已更新单元格 [%d,%d] = %s（注意：WPS 处于单元格编辑态，本轮走的是"
                            "无按键旁路，**未自动保存**，请你自己按 Ctrl+S）" % (row, column, value))
                return f"已更新单元格 [{row},{column}] = {value}"
            finally:
                _release_app(app, "excel", attached)
        else:
            from openpyxl import load_workbook
            wb = load_workbook(file_path)
            ws = wb[sheet_name] if sheet_name else wb.active
            ws.cell(row=row, column=column, value=value)
            wb.save(file_path)
            return f"已更新单元格 [{row},{column}] = {value}"
    
    elif name == "get_sheet_names":
        file_path = os.path.abspath(arguments["file_path"])
        mode = arguments.get("mode", "file")
        edit_mode = get_mode(mode)
        
        if edit_mode == EditMode.COM:
            app, attached = _get_wps_app("excel")
            try:
                if not attached:
                    app.Visible = False
                wb = _get_open_workbook(app, file_path)
                names = [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)]
                if not attached:
                    wb.Close()
                return names
            finally:
                _release_app(app, "excel", attached)
        else:
            from openpyxl import load_workbook
            wb = load_workbook(file_path, read_only=True)
            try:
                names = list(wb.sheetnames)
            finally:
                # V3.1：read_only 模式下 openpyxl 会**一直持有归档句柄**，
                # 不 close 会把文件锁住（实测：之后改名报 WinError 32「另一个程序正在使用此文件」）。
                try:
                    wb.close()
                except Exception:
                    pass
            return names

    elif name == "activate_sheet":
        file_path = os.path.abspath(arguments["file_path"])
        sheet_name = arguments.get("sheet_name")
        sheet_index = arguments.get("sheet_index")
        mode = arguments.get("mode", "com")
        if get_mode(mode) != EditMode.COM:
            raise ValueError(
                "activate_sheet 仅支持 mode=com：切换的是人类正在操作的 WPS 表格，"
                "文件模式不会影响人类界面"
            )
        return _com_activate_sheet(file_path, sheet_name, sheet_index)
    
    elif name == "read_row":
        file_path = os.path.abspath(arguments["file_path"])
        row = int(arguments["row"])
        if row < 1:
            raise ValueError(
                f"read_row 的行号必须是从 1 开始的正整数，收到 row={row}；"
                "行号 0/负数在 COM 与文件模式下均非有效行"
            )
        sheet_name = arguments.get("sheet_name")
        start_col = arguments.get("start_col")
        end_col = arguments.get("end_col")
        mode = arguments.get("mode", "com")
        if get_mode(mode) == EditMode.COM:
            return _com_read_row(file_path, row, sheet_name, start_col, end_col)
        return _file_read_row(file_path, row, sheet_name, start_col, end_col)

    elif name == "add_rows":
        file_path = os.path.abspath(arguments["file_path"])
        count = arguments["count"]
        after_row = arguments.get("after_row")
        sheet_name = arguments.get("sheet_name")
        mode = arguments.get("mode", "file")
        
        if mode == "com":
            app, attached = _get_wps_app("excel")
            try:
                if not attached:
                    app.Visible = False
                wb = _get_open_workbook(app, file_path)
                ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
                if after_row:
                    ws.Rows(f"{after_row + 1}:{after_row + count}").Insert()
                else:
                    ws.Rows(f"{ws.UsedRange.Rows.Count + 1}:{ws.UsedRange.Rows.Count + count}").Insert()
                wb.Save()
                if not attached:
                    wb.Close()
                return f"已添加{count}行"
            finally:
                _release_app(app, "excel", attached)
        else:
            return _file_add_rows(file_path, sheet_name, count, after_row)
    
    elif name == "add_columns":
        file_path = os.path.abspath(arguments["file_path"])
        count = arguments["count"]
        after_col = arguments.get("after_col")
        sheet_name = arguments.get("sheet_name")
        mode = arguments.get("mode", "file")
        
        if mode == "com":
            app, attached = _get_wps_app("excel")
            try:
                if not attached:
                    app.Visible = False
                wb = _get_open_workbook(app, file_path)
                ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
                # V3.1：原来用 chr(64+n) 与 chr(ord(letter)+k)，>26 会算出 '['、跨 Z 会断。
                # 统一走 _col_letter()（1-based 列号 → 列字母），并先校验 range 不越界。
                if after_col:
                    first = int(after_col) + 1
                else:
                    first = int(ws.UsedRange.Columns.Count) + 1
                last = first + int(count) - 1
                if int(count) < 1:
                    raise ValueError("add_columns 的 count 必须 >= 1，收到 %r" % (count,))
                col_letter = _col_letter(first)
                end_letter = _col_letter(last)
                ws.Columns(f"{col_letter}:{end_letter}").Insert()
                wb.Save()
                if not attached:
                    wb.Close()
                return f"已添加{count}列（{col_letter}:{end_letter}）"
            finally:
                _release_app(app, "excel", attached)
        else:
            return _file_add_columns(file_path, sheet_name, count, after_col)
    
    elif name == "merge_cells":
        file_path = os.path.abspath(arguments["file_path"])
        start_row = arguments["start_row"]
        start_col = arguments["start_col"]
        end_row = arguments["end_row"]
        end_col = arguments["end_col"]
        sheet_name = arguments.get("sheet_name")
        mode = arguments.get("mode", "file")
        
        if mode == "com":
            app, attached = _get_wps_app("excel")
            try:
                if not attached:
                    app.Visible = False
                wb = _get_open_workbook(app, file_path)
                ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
                ws.Range(ws.Cells(start_row, start_col), ws.Cells(end_row, end_col)).Merge()
                wb.Save()
                if not attached:
                    wb.Close()
                return f"已合并单元格"
            finally:
                _release_app(app, "excel", attached)
        else:
            return _file_merge_cells(file_path, sheet_name, start_row, start_col, end_row, end_col)
    
    elif name == "unmerge_cells":
        file_path = os.path.abspath(arguments["file_path"])
        start_row = arguments["start_row"]
        start_col = arguments["start_col"]
        end_row = arguments["end_row"]
        end_col = arguments["end_col"]
        sheet_name = arguments.get("sheet_name")
        mode = arguments.get("mode", "file")
        
        if mode == "com":
            app, attached = _get_wps_app("excel")
            try:
                if not attached:
                    app.Visible = False
                wb = _get_open_workbook(app, file_path)
                ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
                ws.Range(ws.Cells(start_row, start_col), ws.Cells(end_row, end_col)).UnMerge()
                wb.Save()
                if not attached:
                    wb.Close()
                return f"已取消合并"
            finally:
                _release_app(app, "excel", attached)
        else:
            return _file_unmerge_cells(file_path, sheet_name, start_row, start_col, end_row, end_col)
    
    elif name == "set_column_width":
        file_path = os.path.abspath(arguments["file_path"])
        column = arguments["column"]
        width = arguments["width"]
        sheet_name = arguments.get("sheet_name")
        mode = arguments.get("mode", "file")
        
        if mode == "com":
            app, attached = _get_wps_app("excel")
            try:
                if not attached:
                    app.Visible = False
                wb = _get_open_workbook(app, file_path)
                ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
                ws.Columns(column).ColumnWidth = width
                wb.Save()
                if not attached:
                    wb.Close()
                return f"已设置列{column}宽度"
            finally:
                _release_app(app, "excel", attached)
        else:
            return _file_set_column_width(file_path, sheet_name, column, width)
    
    elif name == "set_row_height":
        file_path = os.path.abspath(arguments["file_path"])
        row = arguments["row"]
        height = arguments["height"]
        sheet_name = arguments.get("sheet_name")
        mode = arguments.get("mode", "file")
        
        if mode == "com":
            app, attached = _get_wps_app("excel")
            try:
                if not attached:
                    app.Visible = False
                wb = _get_open_workbook(app, file_path)
                ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
                ws.Rows(row).RowHeight = height
                wb.Save()
                if not attached:
                    wb.Close()
                return f"已设置行{row}高度"
            finally:
                _release_app(app, "excel", attached)
        else:
            return _file_set_row_height(file_path, sheet_name, row, height)
    
    elif name == "insert_image":
        file_path = os.path.abspath(arguments["file_path"])
        image_path = os.path.abspath(arguments["image_path"])
        width = arguments.get("width")
        height = arguments.get("height")
        anchor = arguments.get("anchor", "A1")
        slide_index = arguments.get("slide_index", -1)
        mode = arguments.get("mode", "file")
        ext = Path(file_path).suffix.lower()
        
        if mode == "com":
            if ext in ['.doc', '.docx']:
                return _com_insert_image_word(file_path, image_path)
            elif ext in ['.xls', '.xlsx']:
                sheet_name = arguments.get("sheet_name")
                return _com_insert_image_excel(file_path, sheet_name, image_path, anchor)
            elif ext in ['.ppt', '.pptx']:
                return _file_insert_image_pptx(file_path, image_path, slide_index, width or 1, 1, width or 4)
        else:
            if ext in ['.docx', '.doc']:
                return _file_insert_image_docx(file_path, image_path, width, height)
            elif ext in ['.xlsx', '.xls']:
                sheet_name = arguments.get("sheet_name")
                return _file_insert_image_xlsx(file_path, sheet_name, image_path, anchor)
            elif ext in ['.pptx', '.ppt']:
                return _file_insert_image_pptx(file_path, image_path, slide_index, width or 1, 1, width or 4)
            else:
                raise ValueError(f"不支持的文件格式: {ext}")
    
    elif name == "set_cell_style":
        file_path = os.path.abspath(arguments["file_path"])
        row = arguments["row"]
        column = arguments["column"]
        font_name = arguments.get("font_name")
        font_size = arguments.get("font_size")
        bold = arguments.get("bold")
        italic = arguments.get("italic")
        underline = arguments.get("underline")
        color = arguments.get("color")
        bg_color = arguments.get("bg_color")
        align_h = arguments.get("align_h")
        align_v = arguments.get("align_v")
        sheet_name = arguments.get("sheet_name")
        mode = arguments.get("mode", "file")
        
        if mode == "com":
            app, attached = _get_wps_app("excel")
            try:
                if not attached:
                    app.Visible = False
                wb = _get_open_workbook(app, file_path)
                ws = wb.Sheets(sheet_name) if sheet_name else wb.ActiveSheet
                cell = ws.Cells(row, column)
                if font_name:
                    cell.Font.Name = font_name
                if font_size:
                    cell.Font.Size = font_size
                if bold is not None:
                    cell.Font.Bold = bold
                if italic is not None:
                    cell.Font.Italic = italic
                if underline is not None:
                    # V3.1：JSON 布尔不能直接喂 COM，要映射成 xlUnderlineStyle 枚举
                    cell.Font.Underline = _underline_pair(underline)[0]
                if color:
                    # V3.1：COM 的 Color 是 BGR；原来是 int(color,16)，带 '#' 必抛
                    cell.Font.Color = _com_color(color)
                if bg_color:
                    cell.Interior.Color = _com_color(bg_color)
                if align_h:
                    align_map = {"left": -4131, "center": -4108, "right": -4152}
                    cell.HorizontalAlignment = align_map.get(align_h, -4131)
                if align_v:
                    align_map = {"top": -4160, "center": -4108, "bottom": -4107}
                    cell.VerticalAlignment = align_map.get(align_v, -4160)
                wb.Save()
                if not attached:
                    wb.Close()
                return f"已设置单元格样式"
            finally:
                _release_app(app, "excel", attached)
        else:
            return _file_set_cell_style_xlsx(
                file_path, sheet_name, row, column,
                font_name, font_size, bold, italic, underline, color, bg_color, align_h, align_v
            )
    
    elif name == "set_paragraph_style":
        file_path = os.path.abspath(arguments["file_path"])
        para_index = arguments["para_index"]
        font_name = arguments.get("font_name")
        font_size = arguments.get("font_size")
        bold = arguments.get("bold")
        italic = arguments.get("italic")
        underline = arguments.get("underline")
        color = arguments.get("color")
        align = arguments.get("align")
        mode = arguments.get("mode", "file")
        
        if mode == "com":
            app, attached = _get_wps_app("word")
            try:
                if not attached:
                    app.Visible = False
                doc = _get_open_document(app, file_path)
                para = doc.Paragraphs(para_index + 1)
                if font_name:
                    para.Range.Font.Name = font_name
                if font_size:
                    para.Range.Font.Size = font_size
                if bold is not None:
                    para.Range.Font.Bold = bold
                if italic is not None:
                    para.Range.Font.Italic = italic
                if underline is not None:
                    para.Range.Font.Underline = underline
                if color:
                    para.Range.Font.Color = int(color, 16)
                if align:
                    align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
                    para.Alignment = align_map.get(align, 0)
                doc.Save()
                if not attached:
                    doc.Close()
                return f"已设置段落样式"
            finally:
                _release_app(app, "word", attached)
        else:
            return _file_set_para_style_docx(file_path, para_index, font_name, font_size, bold, italic, underline, color, align)
    
    elif name == "export_pdf":
        file_path = os.path.abspath(arguments["file_path"])
        output_path = arguments.get("output_path")
        if output_path:
            output_path = os.path.abspath(output_path)
        return _com_export_pdf(file_path, output_path)
    
    elif name == "run_macro":
        file_path = os.path.abspath(arguments["file_path"])
        macro_name = arguments["macro_name"]
        args = arguments.get("args")
        return _com_run_macro(file_path, macro_name, args)
    
    elif name == "doc_to_ppt":
        doc_path = os.path.abspath(arguments["doc_path"])
        ppt_path = os.path.abspath(arguments["ppt_path"])
        template = arguments.get("template", "default")
        os.makedirs(os.path.dirname(ppt_path) or '.', exist_ok=True)
        return _doc_to_ppt(doc_path, ppt_path, template)
    
    elif name == "ppt_to_doc":
        ppt_path = os.path.abspath(arguments["ppt_path"])
        doc_path = os.path.abspath(arguments["doc_path"])
        os.makedirs(os.path.dirname(doc_path) or '.', exist_ok=True)
        return _ppt_to_doc(ppt_path, doc_path)
    
    elif name == "batch_read":
        file_path = os.path.abspath(arguments["file_path"])
        indices = arguments["indices"]
        sheet_name = arguments.get("sheet_name")
        ext = Path(file_path).suffix.lower()
        
        if ext in ['.docx', '.doc']:
            return _batch_read_docx(file_path, indices)
        elif ext in ['.xlsx', '.xls']:
            return _batch_read_xlsx(file_path, indices, sheet_name)
        else:
            raise ValueError(f"batch_read不支持: {ext}")
    
    elif name == "batch_write":
        file_path = os.path.abspath(arguments["file_path"])
        data = arguments["data"]
        sheet_name = arguments.get("sheet_name")
        # V3.1：这些工具只有文件模式一条路（openpyxl/python-docx 直接写盘）。
        # 目标正被人开着时必须拒绝，否则会覆盖掉人未保存的内存副本。
        _guard_file_write("batch_write", file_path)
        return _batch_write_xlsx(file_path, data, sheet_name)
    
    elif name == "batch_set_style":
        file_path = os.path.abspath(arguments["file_path"])
        operations = arguments["operations"]
        sheet_name = arguments.get("sheet_name")
        ext = Path(file_path).suffix.lower()
        _guard_file_write("batch_set_style", file_path)   # V3.1：只有文件模式，写盘前查冲突
        
        if ext in ['.docx', '.doc']:
            return _batch_set_style_docx(file_path, operations)
        elif ext in ['.xlsx', '.xls']:
            return _batch_set_style_xlsx(file_path, operations)
        else:
            raise ValueError(f"batch_set_style不支持: {ext}")
    
    elif name == "check_quality":
        file_path = os.path.abspath(arguments["file_path"])
        ext = Path(file_path).suffix.lower()
        
        if ext in ['.docx', '.doc']:
            return _check_docx_quality(file_path)
        elif ext in ['.xlsx', '.xls']:
            return _check_xlsx_quality(file_path)
        else:
            raise ValueError(f"check_quality不支持: {ext}")
    
    elif name == "extract_template":
        file_path = os.path.abspath(arguments["file_path"])
        return _extract_docx_template(file_path)
    
    elif name == "apply_template":
        file_path = os.path.abspath(arguments["file_path"])
        template = {
            "font_name": arguments.get("font_name", "SimSun"),
            "font_size": arguments.get("font_size", 12)
        }
        _guard_file_write("apply_template", file_path)    # V3.1：只有文件模式，写盘前查冲突
        return _apply_template(file_path, template)
    
    else:
        raise ValueError(f"未知工具: {name}")


# ==================== MCP协议实现 ====================

class MCPServer:
    """简化版MCP服务器，基于JSON-RPC 2.0 over stdio"""
    
    def __init__(self):
        self.server_info = {
            "name": "wps-editor",
            "version": "3.4.0"
        }
        self.capabilities = {
            "tools": {}
        }
    
    def create_response(self, id: Any, result: Any) -> dict:
        return {
            "jsonrpc": "2.0",
            "id": id,
            "result": result
        }
    
    def create_error(self, id: Any, code: int, message: str) -> dict:
        return {
            "jsonrpc": "2.0",
            "id": id,
            "error": {
                "code": code,
                "message": message
            }
        }
    
    def handle_request(self, request: dict) -> dict:
        method = request.get("method")
        params = request.get("params", {})
        id = request.get("id")
        
        # V3.1：JSON-RPC 2.0 规定**通知不应回包**。原来只有 notifications/initialized 被特判，
        # 其余通知（如 notifications/cancelled）会落到 else，发出一条 id:null 的错误响应，
        # 严格客户端会因此报错。这里先把通知统一吞掉，再补上 ping。
        if id is None and str(method or "").startswith("notifications/"):
            return None
        
        try:
            if method == "initialize":
                return self.create_response(id, {
                    "protocolVersion": "2024-11-05",
                    "capabilities": self.capabilities,
                    "serverInfo": self.server_info
                })
            
            elif method == "ping":
                return self.create_response(id, {})
            
            elif method == "tools/list":
                return self.create_response(id, {
                    "tools": _advertised_tools()
                })
            
            elif method == "tools/call":
                tool_name = params.get("name")
                arguments = params.get("arguments", {})
                result = handle_tool_call(tool_name, arguments)
                return self.create_response(id, {
                    "content": [
                        {
                            "type": "text",
                            "text": json.dumps(result, ensure_ascii=False, indent=2)
                        }
                    ]
                })
            
            elif method == "notifications/initialized":
                return None
            
            else:
                return self.create_error(id, -32601, f"方法不存在: {method}")
        
        except WpsAttachError as e:
            # 挂接失败：以「工具结果 + isError」形式回传结构化求助（候选路径 / 已尝试方法），
            # 不占用 JSON-RPC 错误通道，便于上游 Agent 直接解析 human_action_required 并转达人类。
            return self.create_response(id, {
                "content": [
                    {
                        "type": "text",
                        "text": json.dumps(e.payload, ensure_ascii=False, indent=2)
                    }
                ],
                "isError": True
            })
        except ValueError as e:
            # V3.1：参数/前置条件错误也走"人话结果"而不是裸 JSON-RPC 错误。
            # 之前这类错误直接抛成 -32000，人只看到一串异常文本，AI 也学不到该怎么改。
            return self.create_response(id, {
                "content": [
                    {
                        "type": "text",
                        "text": json.dumps({
                            "ok": False,
                            "error": str(e),
                            "hint": "这是参数/前置条件错误（不是崩溃）。按上面的提示修正后重试；"
                                    "薄工具可用 wps_do(action=\"help\") 看正确写法。",
                        }, ensure_ascii=False, indent=2)
                    }
                ],
                "isError": True
            })
        except Exception as e:
            return self.create_error(id, -32000, str(e))
    
    def run(self):
        """运行MCP服务器（stdio模式）"""
        # 启动就把「这轮跑的是哪份代码、广播几个工具」打到 stderr：
        # 过去反复踩过"客户端跑的还是旧代码"这个坑，这里给一个一眼可辨的自证。
        try:
            mode = "%d 个薄工具（分层加载；WPS_MCP_LEGACY_TOOLS=1 可退回全量 %d 个）" % (
                len(_advertised_tools()), len(TOOLS))
        except Exception as err:
            mode = "工具清单异常：%r" % (err,)
        gates = "宏=%s 写后自动保存=%s" % (
            "允许" if os.environ.get("WPS_MCP_ALLOW_MACRO", "0") not in ("0", "false", "False", "") else "禁用",
            "开" if os.environ.get("WPS_MCP_AUTOSAVE", "0") not in ("0", "false", "False", "") else "关")
        sys.stderr.write("WPS Editor MCP Server v3.4.0 启动中... [%s | %s]\n" % (mode, gates))
        sys.stderr.flush()
        
        for line in sys.stdin:
            line = line.strip()
            if not line:
                continue
            
            try:
                request = json.loads(line)
                response = self.handle_request(request)
                if response:
                    output = json.dumps(response, ensure_ascii=False)
                    sys.stdout.write(output + "\n")
                    sys.stdout.flush()
            except json.JSONDecodeError as e:
                sys.stderr.write(f"JSON解析错误: {e}\n")
                sys.stderr.flush()
            except Exception as e:
                sys.stderr.write(f"处理请求错误: {e}\n")
                sys.stderr.flush()


# 入口点
def main():
    server = MCPServer()
    server.run()


if __name__ == "__main__":
    main()
