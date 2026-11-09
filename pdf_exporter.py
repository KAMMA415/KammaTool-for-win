"""
PDF导出模块
支持将Excel文件导出为PDF
"""
import os
import shutil
import tempfile
import time
from pathlib import Path
from typing import List, Optional
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border
try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    print("警告: win32com未安装，PDF导出功能可能不可用")
try:
    import PyPDF2
    PYPDF2_AVAILABLE = True
except ImportError:
    PYPDF2_AVAILABLE = False
    print("警告: PyPDF2未安装，PDF合并功能可能不可用")


class PDFExporter:
    """PDF导出器"""
    
    def __init__(self):
        self.processor = None  # 将在GUI中设置
    
    def export_to_pdf_merged(self, excel_files: List[Path], output_pdf: Path) -> bool:
        """
        合并导出PDF功能
        
        功能：
        a. 文件类型：Microsoft Excel 工作表（.xlsx新版本、.xls老版本）
        b. 目标文件：目录中的所有Excel文件或单独一个文件
        c. 目的：
           1) 将所有文件及所有SHEET页，输出到同一个PDF文件中
           2) 使用EXCEL原文件名作为输出的PDF文件名（由GUI层处理）
           3) 输出PDF的页面顺序，使用EXCEL文件中SHEET页的自然顺序
        """
        if not WIN32COM_AVAILABLE:
            print("错误: win32com未安装，无法导出PDF")
            return False
        
        if not PYPDF2_AVAILABLE:
            print("错误: PyPDF2未安装，无法合并PDF")
            return False
        
        excel_app = None
        temp_dir = None
        temp_files = []
        temp_pdf_files = []  # 存储临时PDF文件
        
        try:
            # 启动Excel应用程序
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            excel_app.ScreenUpdating = False
            
            processed_files = []
            failed_files = []
            
            # 创建临时目录用于存储临时PDF文件
            temp_dir = tempfile.mkdtemp(prefix="excel_pdf_merge_")
            
            # 遍历所有Excel文件，按文件顺序处理
            for file_path in excel_files:
                # 检查文件是否存在和格式
                if not file_path.exists():
                    failed_files.append(f"{file_path.name} (文件不存在)")
                    continue
                
                file_ext = file_path.suffix.lower()
                if file_ext not in ['.xlsx', '.xls']:
                    print(f"跳过非Excel文件: {file_path.name}")
                    continue
                
                wb = None
                try:
                    # 打开Excel文件
                    temp_dir_list = [temp_dir]
                    wb = self._open_workbook(excel_app, file_path, temp_dir_list, temp_files)
                    temp_dir = temp_dir_list[0]
                    
                    if wb is None:
                        failed_files.append(f"{file_path.name} (无法打开)")
                        continue
                    
                    # 获取Sheet数量
                    sheet_count = wb.Worksheets.Count
                    print(f"处理文件: {file_path.name}，包含 {sheet_count} 个Sheet")
                    
                    # 为当前文件生成临时PDF文件名
                    temp_pdf_file = Path(temp_dir) / f"temp_{len(temp_pdf_files):04d}.pdf"
                    temp_pdf_files.append(temp_pdf_file)
                    
                    # 激活工作簿
                    wb.Activate()
                    
                    # 直接使用Excel的"发布为PDF或XPS"功能导出当前文件的所有Sheet
                    print(f"  正在使用Excel默认功能导出PDF: {temp_pdf_file.name}")
                    wb.ExportAsFixedFormat(
                        Type=0,  # xlTypePDF = 0，表示PDF格式
                        Filename=str(temp_pdf_file.resolve()),
                        Quality=0,  # xlQualityStandard = 0，标准质量
                        IncludeDocProperties=True,  # 包含文档属性
                        IgnorePrintAreas=False,  # 不忽略打印区域
                        OpenAfterPublish=False  # 发布后不打开文件
                    )
                    
                    print(f"  [OK] 已导出: {file_path.name} (共 {sheet_count} 个Sheet)")
                    
                    wb.Close(SaveChanges=False)
                    processed_files.append(file_path.name)
                    
                except Exception as e:
                    if wb:
                        try:
                            wb.Close(SaveChanges=False)
                        except:
                            pass
                    error_msg = str(e)
                    failed_files.append(f"{file_path.name} ({error_msg})")
                    print(f"处理文件 {file_path.name} 时出错: {error_msg}")
                    import traceback
                    traceback.print_exc()
                    continue
            
            excel_app.Quit()
            
            # 检查是否有PDF文件生成
            if len(temp_pdf_files) == 0:
                print("错误: 没有成功生成任何PDF文件")
                return False
            
            # 合并所有PDF文件
            print(f"\n正在合并 {len(temp_pdf_files)} 个PDF文件...")
            try:
                pdf_merger = PyPDF2.PdfMerger()
                
                for pdf_file in temp_pdf_files:
                    if pdf_file.exists():
                        print(f"  合并: {pdf_file.name}")
                        pdf_merger.append(str(pdf_file.resolve()))
                
                # 保存合并后的PDF
                output_path = str(output_pdf.resolve())
                print(f"正在保存合并后的PDF到: {output_path}")
                pdf_merger.write(output_path)
                pdf_merger.close()
                
                print(f"PDF合并成功: {output_path}")
            except Exception as e:
                print(f"合并PDF失败: {e}")
                import traceback
                traceback.print_exc()
                return False
            
            # 清理临时文件
            if temp_dir and os.path.exists(temp_dir):
                try:
                    for pdf_file in temp_pdf_files:
                        if pdf_file.exists():
                            pdf_file.unlink()
                    for tf in temp_files:
                        if tf.exists():
                            tf.unlink()
                    os.rmdir(temp_dir)
                except Exception as e:
                    print(f"清理临时文件时出错: {e}")
            
            if failed_files:
                print(f"\n警告: {len(failed_files)} 个文件处理失败:")
                for failed in failed_files:
                    print(f"  - {failed}")
            
            print(f"\n成功处理 {len(processed_files)} 个文件")
            return True
            
        except Exception as e:
            print(f"合并导出PDF失败: {e}")
            import traceback
            traceback.print_exc()
            
            # 清理资源
            if excel_app:
                try:
                    excel_app.Quit()
                except:
                    pass
            if temp_dir and os.path.exists(temp_dir):
                try:
                    for pdf_file in temp_pdf_files:
                        if pdf_file.exists():
                            try:
                                pdf_file.unlink()
                            except:
                                pass
                    for tf in temp_files:
                        if tf.exists():
                            try:
                                tf.unlink()
                            except:
                                pass
                    os.rmdir(temp_dir)
                except:
                    pass
            return False
    
    def _open_workbook(self, excel_app, file_path: Path, temp_dir_list: list, temp_files: List[Path]):
        """
        打开Excel工作簿，支持临时文件方案处理特殊字符
        
        参数:
            temp_dir_list: [temp_dir] 列表，用于在方法内部更新temp_dir
        """
        file_path_str = str(file_path.resolve())
        last_error = None
        
        # 尝试直接打开（多种方式）
        try:
            return excel_app.Workbooks.Open(file_path_str, ReadOnly=True, UpdateLinks=0, CorruptLoad=0)
        except Exception as e1:
            last_error = e1
            try:
                return excel_app.Workbooks.Open(file_path_str, UpdateLinks=0, CorruptLoad=0, ReadOnly=False)
            except Exception as e2:
                last_error = e2
                try:
                    return excel_app.Workbooks.Open(file_path_str)
                except Exception as e3:
                    last_error = e3
        
        # 如果直接打开失败，使用临时文件
        if not temp_dir_list or temp_dir_list[0] is None:
            temp_dir = tempfile.mkdtemp(prefix="excel_pdf_")
            if temp_dir_list:
                temp_dir_list[0] = temp_dir
            else:
                temp_dir_list.append(temp_dir)
        else:
            temp_dir = temp_dir_list[0]
        
        temp_index = len(temp_files)
        file_ext = file_path.suffix.lower()
        temp_file = Path(temp_dir) / f"temp_{temp_index:04d}{file_ext}"
        shutil.copy2(file_path, temp_file)
        temp_files.append(temp_file)
        
        temp_path_str = str(temp_file.resolve())
        try:
            return excel_app.Workbooks.Open(temp_path_str, ReadOnly=True, UpdateLinks=0, CorruptLoad=0)
        except:
            try:
                return excel_app.Workbooks.Open(temp_path_str)
            except Exception as e:
                raise Exception(f"无法打开文件: {e}")
    
    def export_to_pdf_split(self, excel_files: List[Path], output_dir: Path) -> bool:
        """拆分导出每个Sheet为单独的PDF文件"""
        if not WIN32COM_AVAILABLE:
            print("错误: win32com未安装，无法导出PDF")
            return False

        # 确保输出目录存在（B14修复）
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"无法创建输出目录 {output_dir}: {e}")
            return False

        # 创建临时目录用于处理包含特殊字符的文件

        temp_dir = None
        temp_files = []
        
        try:
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            
            used_names = {}  # 用于跟踪已使用的文件名
            
            for file_path in excel_files:
                wb = None
                temp_file_path = None
                try:
                    # 检查文件是否存在
                    if not file_path.exists():
                        print(f"文件不存在: {file_path}")
                        continue
                    
                    # 尝试多种方式打开文件
                    file_path_str = str(file_path.resolve())
                    open_success = False
                    last_error = None
                    
                    # 方式1: 使用ReadOnly模式（推荐）
                    try:
                        wb = excel_app.Workbooks.Open(
                            file_path_str,
                            ReadOnly=True,
                            UpdateLinks=0,
                            CorruptLoad=0
                        )
                        open_success = True
                    except Exception as e1:
                        last_error = e1
                        # 方式2: 不使用ReadOnly
                        try:
                            wb = excel_app.Workbooks.Open(
                                file_path_str,
                                UpdateLinks=0,
                                CorruptLoad=0,
                                ReadOnly=False
                            )
                            open_success = True
                        except Exception as e2:
                            last_error = e2
                            # 方式3: 最简单的打开方式
                            try:
                                wb = excel_app.Workbooks.Open(file_path_str)
                                open_success = True
                            except Exception as e3:
                                last_error = e3
                    
                    # 如果所有直接打开方式都失败，尝试使用临时文件
                    if not open_success:
                        try:
                            if temp_dir is None:
                                temp_dir = tempfile.mkdtemp(prefix="excel_pdf_")
                            temp_index = len(temp_files)
                            temp_file_name = f"temp_{temp_index:04d}.xlsx"
                            temp_file_path = Path(temp_dir) / temp_file_name
                            shutil.copy2(file_path, temp_file_path)
                            temp_files.append(temp_file_path)
                            
                            temp_file_path_str = str(temp_file_path.resolve())
                            try:
                                wb = excel_app.Workbooks.Open(
                                    temp_file_path_str,
                                    ReadOnly=True,
                                    UpdateLinks=0,
                                    CorruptLoad=0
                                )
                                open_success = True
                            except:
                                try:
                                    wb = excel_app.Workbooks.Open(temp_file_path_str)
                                    open_success = True
                                except:
                                    raise last_error
                        except Exception as e_temp:
                            raise last_error
                    
                    if not open_success:
                        raise last_error
                    
                    for sheet in wb.Worksheets:
                        # 生成文件名
                        base_name = sheet.Name
                        # 清理文件名（移除非法字符）
                        safe_name = self._sanitize_filename(base_name)
                        
                        # 处理重名
                        if safe_name in used_names:
                            used_names[safe_name] += 1
                            safe_name = f"{safe_name}_{used_names[safe_name]}"
                        else:
                            used_names[safe_name] = 0
                        
                        pdf_path = output_dir / f"{safe_name}.pdf"
                        
                        # 导出当前Sheet为PDF
                        sheet.ExportAsFixedFormat(
                            Type=0,  # xlTypePDF
                            Filename=str(pdf_path.resolve()),
                            Quality=0,
                            IncludeDocProperties=True,
                            IgnorePrintAreas=False,
                            OpenAfterPublish=False
                        )
                    
                    # 关闭工作簿
                    if wb:
                        wb.Close(SaveChanges=False)
                        wb = None
                        
                except Exception as e:
                    error_msg = str(e)
                    error_code = self._extract_error_code(e)
                    
                    # 提供更准确的错误信息
                    if "-2147352567" in error_code or "-2147352567" in error_msg:
                        error_msg = f"无法打开文件（错误代码: {error_code}）。可能原因：1) 文件被其他程序打开 2) 文件路径包含特殊字符 3) 文件格式问题。请检查文件是否可正常打开。"
                    elif "被另一个进程使用" in error_msg or "正在使用" in error_msg or "is being used" in error_msg.lower():
                        error_msg = f"文件可能被其他程序打开（如Excel），请先关闭该文件"
                    elif "找不到文件" in error_msg or "文件不存在" in error_msg or "cannot find" in error_msg.lower():
                        error_msg = f"文件不存在或路径错误"
                    elif "权限" in error_msg or "permission" in error_msg.lower():
                        error_msg = f"文件访问权限不足"
                    else:
                        error_msg = f"处理失败: {error_msg} (错误代码: {error_code})"
                    
                    print(f"处理文件 {file_path} 时出错: {error_msg}")
                    
                    # 确保工作簿被关闭
                    try:
                        if wb:
                            wb.Close(SaveChanges=False)
                    except:
                        pass
                    continue
            
            excel_app.Quit()
            
            # 清理临时文件
            if temp_dir and os.path.exists(temp_dir):
                try:
                    for temp_file in temp_files:
                        if temp_file.exists():
                            try:
                                temp_file.unlink()
                            except:
                                pass
                    os.rmdir(temp_dir)
                except:
                    pass
            
            return True
        except Exception as e:
            print(f"拆分导出PDF失败: {e}")
            try:
                excel_app.Quit()
            except:
                pass
            # 清理临时文件
            if temp_dir and os.path.exists(temp_dir):
                try:
                    for temp_file in temp_files:
                        if temp_file.exists():
                            try:
                                temp_file.unlink()
                            except:
                                pass
                    os.rmdir(temp_dir)
                except:
                    pass
            return False
    
    def _extract_error_code(self, error) -> str:
        """提取COM错误的错误代码"""
        try:
            # COM错误通常是元组格式: (错误代码, 错误描述, ...)
            if isinstance(error, tuple) and len(error) > 0:
                return str(error[0])
            elif hasattr(error, 'args') and len(error.args) > 0:
                return str(error.args[0])
            elif hasattr(error, 'hresult'):
                return str(error.hresult)
            else:
                return str(error)
        except:
            return str(error)
    
    def _sanitize_filename(self, filename: str) -> str:
        """清理文件名，移除非法字符"""
        illegal_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
        for char in illegal_chars:
            filename = filename.replace(char, '_')
        # 限制长度
        if len(filename) > 200:
            filename = filename[:200]
        return filename


class ExcelSplitter:
    """Excel拆分器"""
    
    def __init__(self):
        pass
    
    def split_to_excel_files(self, excel_files: List[Path], output_dir: Path) -> bool:
        """将每个Sheet导出为单独的Excel文件"""
        from excel_processor import ExcelProcessor
        processor = ExcelProcessor()
        
        used_names = {}  # 用于跟踪已使用的文件名
        
        for file_path in excel_files:
            wb = processor.load_workbook(file_path)
            if not wb:
                continue
            
            for sheet_name in wb.sheetnames:
                # 生成文件名
                base_name = sheet_name
                # 清理文件名
                safe_name = self._sanitize_filename(base_name)
                
                # 处理重名
                if safe_name in used_names:
                    used_names[safe_name] += 1
                    safe_name = f"{safe_name}_{used_names[safe_name]}"
                else:
                    used_names[safe_name] = 0
                
                # 创建新工作簿
                new_wb = openpyxl.Workbook()
                new_ws = new_wb.active
                new_ws.title = sheet_name
                
                # 复制数据
                source_sheet = wb[sheet_name]
                for row in source_sheet.iter_rows():
                    for cell in row:
                        new_cell = new_ws.cell(row=cell.row, column=cell.column)
                        new_cell.value = cell.value
                        # 安全地复制样式（避免StyleProxy错误）
                        try:
                            if cell.has_style:
                                # 复制字体（创建新对象而不是直接赋值）
                                if cell.font and cell.font.name:
                                    try:
                                        font_color = cell.font.color
                                        if hasattr(font_color, 'rgb') and font_color.rgb:
                                            color_value = font_color.rgb
                                        elif hasattr(font_color, 'theme'):
                                            color_value = font_color.theme
                                        else:
                                            color_value = None
                                        
                                        new_cell.font = Font(
                                            name=cell.font.name or 'Calibri',
                                            size=cell.font.size or 11,
                                            bold=cell.font.bold or False,
                                            italic=cell.font.italic or False,
                                            vertAlign=getattr(cell.font, 'vertAlign', None),
                                            underline=cell.font.underline or None,
                                            strike=cell.font.strike or False,
                                            color=color_value
                                        )
                                    except:
                                        pass
                                
                                # 复制对齐
                                if cell.alignment:
                                    try:
                                        new_cell.alignment = Alignment(
                                            horizontal=cell.alignment.horizontal or 'general',
                                            vertical=cell.alignment.vertical or 'bottom',
                                            text_rotation=cell.alignment.text_rotation or 0,
                                            wrap_text=cell.alignment.wrap_text or False,
                                            shrink_to_fit=cell.alignment.shrink_to_fit or False,
                                            indent=cell.alignment.indent or 0
                                        )
                                    except:
                                        pass
                                
                                # 复制填充
                                if cell.fill and hasattr(cell.fill, 'fill_type') and cell.fill.fill_type:
                                    try:
                                        new_cell.fill = PatternFill(
                                            fill_type=cell.fill.fill_type,
                                            start_color=getattr(cell.fill, 'start_color', None),
                                            end_color=getattr(cell.fill, 'end_color', None)
                                        )
                                    except:
                                        pass
                                
                                # 复制边框（简化处理）
                                if cell.border:
                                    try:
                                        # 只复制基本边框属性
                                        new_cell.border = Border(
                                            left=getattr(cell.border, 'left', None),
                                            right=getattr(cell.border, 'right', None),
                                            top=getattr(cell.border, 'top', None),
                                            bottom=getattr(cell.border, 'bottom', None)
                                        )
                                    except:
                                        pass
                                
                                # 复制数字格式
                                if cell.number_format:
                                    try:
                                        new_cell.number_format = str(cell.number_format)
                                    except:
                                        pass
                        except Exception as e:
                            # 如果样式复制失败，至少保留值
                            print(f"复制样式时出错（行{cell.row}，列{cell.column}）: {e}")
                            pass
                
                # 保存
                output_file = output_dir / f"{safe_name}.xlsx"
                processor.save_workbook(new_wb, output_file)
        
        return True
    
    def _sanitize_filename(self, filename: str) -> str:
        """清理文件名，移除非法字符"""
        illegal_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
        for char in illegal_chars:
            filename = filename.replace(char, '_')
        # 限制长度
        if len(filename) > 200:
            filename = filename[:200]
        return filename

