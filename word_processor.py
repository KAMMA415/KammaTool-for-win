"""
Word 文档处理模块
提供 Word 文档的批量处理功能
"""

import os
import tempfile
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


class WordProcessor:
    """Word文档处理器"""
    
    def __init__(self):
        self.supported_extensions = ['.docx', '.doc']

    def _safe_save(self, doc, file_path: Path) -> bool:
        """安全保存：先写临时文件，成功后原子替换原文件，失败时原文件完全不受影响。"""
        file_path = Path(file_path)
        tmp_path = None
        try:
            fd, tmp_str = tempfile.mkstemp(suffix='.tmp', dir=file_path.parent)
            os.close(fd)
            tmp_path = Path(tmp_str)
            doc.save(tmp_path)
            tmp_path.replace(file_path)  # 原子替换
            return True
        except Exception as e:
            print(f"保存文件失败 {file_path.name}: {e}")
            if tmp_path and tmp_path.exists():
                try:
                    tmp_path.unlink()
                except:
                    pass
            raise
    
    def get_word_files(self, directory):
        """获取目录下所有Word文件"""
        directory = Path(directory)
        files = []
        for ext in self.supported_extensions:
            files.extend(directory.glob(f'*{ext}'))
        return sorted(files)
    
    def format_document(self, file_path):
        """
        统一格式化Word文档
        - 设置统一字体：宋体 12pt
        - 设置段落行距
        """
        try:
            if str(file_path).lower().endswith('.doc'):
                print(f"不支持 .doc 格式: {file_path}")
                return False
            doc = Document(file_path)
            
            # 遍历所有段落
            for para in doc.paragraphs:
                # 设置段落格式
                para_format = para.paragraph_format
                para_format.line_spacing = 1.5
                
                # 设置字体
                for run in para.runs:
                    run.font.name = '宋体'
                    run.font.size = Pt(12)
            
            # 遍历所有表格
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            for run in para.runs:
                                run.font.name = '宋体'
                                run.font.size = Pt(12)
            
            # 安全保存（先写临时文件再替换）
            self._safe_save(doc, file_path)
            return True
        except Exception as e:
            print(f"处理文件失败 {file_path}: {e}")
            return False
    
    def merge_word_files(self, directory, output_path):
        """
        合并目录下所有Word文件
        """
        try:
            directory = Path(directory)
            files = self.get_word_files(directory)
            
            if not files:
                print("未找到Word文件")
                return False
            
            # 创建新文档
            merged_doc = Document()
            
            # 逐个添加文档
            for file in files:
                if file.suffix.lower() == '.doc':
                    print(f"合并功能不支持 .doc 格式，跳过: {file.name}")
                    continue
                print(f"正在合并: {file.name}")
                try:
                    doc = Document(file)
                    
                    # 添加分隔标题
                    merged_doc.add_heading(f'文件: {file.name}', level=1)
                    
                    # 复制段落
                    for element in doc.element.body:
                        merged_doc.element.body.append(element)
                    
                    # 添加分页符
                    merged_doc.add_page_break()
                    
                except Exception as e:
                    print(f"合并文件失败 {file.name}: {e}")
                    continue
            
            # 保存合并后的文档
            merged_doc.save(output_path)
            print(f"合并完成: {output_path}")
            return True
            
        except Exception as e:
            print(f"合并过程发生错误: {e}")
            return False
    
    def convert_to_txt(self, file_path):
        """
        将Word文档转换为TXT文本
        """
        try:
            if str(file_path).lower().endswith('.doc'):
                print(f"转换功能不支持 .doc 格式，跳过: {file_path}")
                return False
            file_path = Path(file_path)
            doc = Document(file_path)
            
            # 提取所有文本
            full_text = []
            for para in doc.paragraphs:
                full_text.append(para.text)
            
            # 保存为txt
            txt_path = file_path.with_suffix('.txt')
            with open(txt_path, 'w', encoding='utf-8-sig') as f:
                f.write('\n'.join(full_text))
            
            print(f"转换完成: {txt_path}")
            return True
            
        except Exception as e:
            print(f"转换失败 {file_path}: {e}")
            return False
    
    def replace_text(self, file_path, old_text, new_text):
        """
        批量替换Word文档中的文本
        """
        try:
            if str(file_path).lower().endswith('.doc'):
                print(f"替换功能不支持 .doc 格式，跳过: {file_path}")
                return False
            doc = Document(file_path)
            
            # 替换段落中的文本
            for para in doc.paragraphs:
                if old_text in para.text:
                    for run in para.runs:
                        run.text = run.text.replace(old_text, new_text)
            
            # 替换表格中的文本
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            if old_text in para.text:
                                for run in para.runs:
                                    run.text = run.text.replace(old_text, new_text)
            
            self._safe_save(doc, file_path)
            return True
            
        except Exception as e:
            print(f"替换文本失败 {file_path}: {e}")
            return False
    
    def clear_formatting(self, file_path):
        """
        清除Word文档格式（保留文本）
        """
        try:
            if str(file_path).lower().endswith('.doc'):
                print(f"清除格式功能不支持 .doc 格式，跳过: {file_path}")
                return False
            doc = Document(file_path)
            
            # 清除段落格式
            for para in doc.paragraphs:
                for run in para.runs:
                    run.font.name = '宋体'
                    run.font.size = Pt(12)
                    run.font.bold = False
                    run.font.italic = False
                    run.font.underline = False
                    run.font.color.rgb = RGBColor(0, 0, 0)
            
            self._safe_save(doc, file_path)
            return True
            
        except Exception as e:
            print(f"清除格式失败 {file_path}: {e}")
            return False
    
    def apply_unified_styles(self, file_path):
        """
        应用统一样式到Word文档
        按照详细规范配置标题、标题1-5和正文样式
        
        样式配置：
        - 标题(Title): 微软雅黑、二号(22pt)、加粗、居中、单倍行距
        - 标题1: 微软雅黑、三号(16pt)、加粗、左对齐、单倍行距
        - 标题2: 微软雅黑、四号(14pt)、左对齐、单倍行距
        - 标题3: 微软雅黑、四号(14pt)、左对齐、单倍行距
        - 标题4: 微软雅黑、五号(10.5pt)、加粗、左对齐、单倍行距
        - 标题5: 微软雅黑、五号(10.5pt)、加粗、左对齐、单倍行距
        - 正文(Normal): 宋体、五号(10.5pt)、两端对齐、单倍行距
        """
        try:
            if str(file_path).lower().endswith('.doc'):
                print(f"应用统一样式功能不支持 .doc 格式，跳过: {file_path}")
                return False
            doc = Document(file_path)
            styles = doc.styles
            
            # 样式配置映射
            style_configs = {
                'Title': {
                    'font_name': '微软雅黑',
                    'font_size': Pt(22),  # 二号
                    'bold': True,
                    'alignment': WD_ALIGN_PARAGRAPH.CENTER,
                },
                'Heading 1': {
                    'font_name': '微软雅黑',
                    'font_size': Pt(16),  # 三号
                    'bold': True,
                    'alignment': WD_ALIGN_PARAGRAPH.LEFT,
                },
                'Heading 2': {
                    'font_name': '微软雅黑',
                    'font_size': Pt(14),  # 四号
                    'bold': False,
                    'alignment': WD_ALIGN_PARAGRAPH.LEFT,
                },
                'Heading 3': {
                    'font_name': '微软雅黑',
                    'font_size': Pt(14),  # 四号
                    'bold': False,
                    'alignment': WD_ALIGN_PARAGRAPH.LEFT,
                },
                'Heading 4': {
                    'font_name': '微软雅黑',
                    'font_size': Pt(10.5),  # 五号
                    'bold': True,
                    'alignment': WD_ALIGN_PARAGRAPH.LEFT,
                },
                'Heading 5': {
                    'font_name': '微软雅黑',
                    'font_size': Pt(10.5),  # 五号
                    'bold': True,
                    'alignment': WD_ALIGN_PARAGRAPH.LEFT,
                },
                'Normal': {
                    'font_name': '宋体',
                    'font_size': Pt(10.5),  # 五号
                    'bold': False,
                    'alignment': WD_ALIGN_PARAGRAPH.JUSTIFY,  # 两端对齐
                },
            }
            
            # 第一步：配置样式定义
            print("  第1步：配置样式定义...")
            for style_name, config in style_configs.items():
                try:
                    style = styles[style_name]
                    
                    # 设置字体
                    style.font.name = config['font_name']
                    style.font.size = config['font_size']
                    style.font.bold = config['bold']
                    
                    # 设置段落格式
                    style.paragraph_format.alignment = config['alignment']
                    style.paragraph_format.line_spacing = 1.0  # 单倍行距
                    
                    # 设置缩进为0
                    style.paragraph_format.left_indent = Pt(0)
                    style.paragraph_format.right_indent = Pt(0)
                    
                    # 段前段后间距设置为自动（None）
                    style.paragraph_format.space_before = None
                    style.paragraph_format.space_after = None
                    
                    print(f"    [OK] 已配置样式: {style_name}")
                    
                except KeyError:
                    print(f"    [WARN] 样式不存在: {style_name}")
                    continue
                except Exception as e:
                    print(f"    [FAIL] 配置样式失败 {style_name}: {e}")
                    continue
            
            # 第二步：遍历段落，清除直接格式化并应用样式
            print("  第2步：清除段落直接格式化...")
            para_count = 0
            for para in doc.paragraphs:
                # 获取段落的样式名称
                style_name = para.style.name if para.style else 'Normal'
                
                # 只处理我们配置的样式
                if style_name in style_configs:
                    config = style_configs[style_name]
                    
                    # 清除所有run的直接格式化，让样式定义生效
                    for run in para.runs:
                        # 设置字体（从样式配置）
                        run.font.name = config['font_name']
                        run.font.size = config['font_size']
                        run.font.bold = config['bold']
                        
                        # 清除其他格式
                        run.font.italic = False
                        run.font.underline = False
                        run.font.color.rgb = RGBColor(0, 0, 0)
                        
                        # 清除文本突出显示颜色（高亮）和其他格式
                        run.font.highlight_color = None
                    
                    # 设置段落格式（从样式配置）
                    para.paragraph_format.alignment = config['alignment']
                    para.paragraph_format.line_spacing = 1.0
                    # 正文样式保留原有缩进，其他样式清零缩进
                    if style_name != "Normal":
                        para.paragraph_format.left_indent = Pt(0)
                        para.paragraph_format.right_indent = Pt(0)
                    para.paragraph_format.right_indent = Pt(0)
                    para.paragraph_format.space_after = None
                    
                    para_count += 1
            
            print(f"    [OK] 已处理 {para_count} 个段落")
            
            
            # 第三步：处理表格样式和内容
            print("  第3步：处理表格样式和内容...")
            table_count = 0
            for table in doc.tables:
                # 设置表格样式为普通表格 网格型（Light Grid）
                try:
                    table.style = 'Light Grid'
                except Exception as e:
                    # 如果样式设置失败，打印警告但继续处理
                    print(f"      [WARN] 表格样式设置失败: {e}")
                    pass
                
                # 处理表格中的所有单元格
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            # 表格中的文本统一使用宋体、五号、不加粗
                            for run in para.runs:
                                run.font.name = '宋体'
                                run.font.size = Pt(10.5)  # 五号
                                run.font.bold = False
                                run.font.italic = False
                                run.font.underline = False
                                run.font.color.rgb = RGBColor(0, 0, 0)
                                
                                # 清除文本突出显示颜色（高亮）
                                run.font.highlight_color = None
                            
                            # 段落格式不修改（保留原有对齐方式和缩进）
                            table_count += 1
            
            print(f"    [OK] 已处理 {table_count} 个表格单元格段落")
            
            # 安全保存文档
            self._safe_save(doc, file_path)
            print(f"[OK] 统一样式完成: {file_path}")
            return True
            
        except Exception as e:
            print(f"应用统一样式失败 {file_path}: {e}")
            import traceback
            traceback.print_exc()
            return False

    def unify_table_format(self, file_path):
        """
        统一表格为：表布局——自动调整——根据窗口自动调整表格
        支持 .doc 和 .docx
        """
        import win32com.client
        import os
        try:
            abs_path = os.path.abspath(file_path)
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            try:
                doc = word.Documents.Open(abs_path)
                for table in doc.Tables:
                    table.AutoFitBehavior(2) # 2 = wdAutoFitWindow
                doc.Save()
                doc.Close()
                return True
            finally:
                word.Quit()
        except Exception as e:
            print(f"统一表格格式失败 {file_path}: {e}")
            import traceback
            traceback.print_exc()
            return False
