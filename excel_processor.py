"""
Excel处理核心模块
支持.xlsx和.xls格式的Excel文件处理
"""
import os
import re
import warnings
from pathlib import Path
from typing import List, Tuple, Optional
import zipfile
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from copy import copy
import pandas as pd

# 抑制 openpyxl 的样式警告
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


class ExcelProcessor:
    """Excel文件处理器"""
    
    def __init__(self):
        self.supported_extensions = ['.xlsx', '.xls']
    
    def get_excel_files(self, path: str) -> List[Path]:
        """获取目录中的所有Excel文件"""
        path_obj = Path(path)
        excel_files = []
        
        if path_obj.is_file():
            # 过滤临时文件
            if path_obj.name.startswith('~$'):
                return []
            if path_obj.suffix.lower() in self.supported_extensions:
                excel_files.append(path_obj)
        elif path_obj.is_dir():
            for ext in self.supported_extensions:
                # 过滤临时文件 (exclude files starting with ~$)
                files = [f for f in path_obj.glob(f'*{ext}') if not f.name.startswith('~$')]
                excel_files.extend(files)
        
        return excel_files
    
    def load_workbook(self, file_path: Path) -> Optional[Workbook]:
        """加载Excel工作簿"""
        try:
            # 1. 检查文件大小
            if file_path.stat().st_size == 0:
                print(f"跳过空文件: {file_path}")
                return None

            return openpyxl.load_workbook(file_path)
            
        except zipfile.BadZipFile:
            print(f"错误：{file_path.name} 不是有效的Excel文件（可能已加密或损坏）")
            return None
        except Exception as e:
            print(f"加载文件失败 {file_path}: {e}")
            return None
    
    def save_workbook(self, wb: Workbook, file_path: Path):
        """保存Excel工作簿（保存失败时抛出异常，让调用方感知）"""
        try:
            wb.save(file_path)
        except PermissionError as e:
            msg = f"保存失败：{file_path.name} 被其他程序占用，请关闭 Excel 后重试"
            print(msg)
            raise PermissionError(msg) from e
        except Exception as e:
            print(f"保存文件失败 {file_path}: {e}")
            raise
    
    def process_excel_file(self, file_path: Path, processor_func):
        """处理Excel文件的通用包装函数，确保资源正确释放
        
        Args:
            file_path: Excel文件路径
            processor_func: 处理函数，接收(workbook)作为参数
            
        Returns:
            bool: 是否成功处理
        """
        wb = None
        try:
            wb = self.load_workbook(file_path)
            if not wb:
                return False
            
            # 执行处理逻辑
            processor_func(wb)
            
            # 保存并返回成功
            self.save_workbook(wb, file_path)
            return True
            
        except Exception as e:
            print(f"处理文件异常 {file_path}: {e}")
            import traceback
            traceback.print_exc()
            return False
        finally:
            # 确保workbook被关闭
            if wb:
                try:
                    wb.close()
                except:
                    pass
    
    
    # ========== 功能1: 统一格式 ==========
    def format_cells(self, file_path: Path):
        """统一格式化所有单元格"""
        def process(wb):
            for sheet in wb.worksheets:
                # 设置字体
                font = Font(name='微软雅黑', size=10, color='000000')
                # 设置对齐
                alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
                # 无填充
                fill = PatternFill(fill_type=None)
                
                # 格式化所有单元格
                for row in sheet.iter_rows():
                    for cell in row:
                        cell.font = font
                        cell.alignment = alignment
                        cell.fill = fill
                
                # 设置列宽
                for col in range(1, sheet.max_column + 1):
                    sheet.column_dimensions[get_column_letter(col)].width = 8
                
                # 设置行高
                for row in range(1, sheet.max_row + 1):
                    sheet.row_dimensions[row].height = 14
        
        return self.process_excel_file(file_path, process)
    
    
    # ========== 功能2: 删除行或列 ==========
    def delete_rows_or_columns(self, file_path: Path, row_info: Tuple[int, int] = None, col_info: Tuple[int, int] = None):
        """删除指定的行或列
        
        Args:
            row_info: (起始行号, 数量) 元组，如 (3, 5) 表示从第3行开始删除5行
            col_info: (起始列号, 数量) 元组，如 (2, 3) 表示从第2列开始删除3列
        """
        wb = self.load_workbook(file_path)
        if not wb:
            return False
        
        try:
            for sheet in wb.worksheets:
                # 删除行（从大到小排序，避免索引变化）
                if row_info:
                    start_row, count = row_info
                    # 从大到小删除，避免索引变化
                    for i in range(count - 1, -1, -1):
                        row_num = start_row + i
                        if 1 <= row_num <= sheet.max_row:
                            sheet.delete_rows(row_num)
                
                # 删除列（从大到小排序）
                if col_info:
                    start_col, count = col_info
                    # 从大到小删除，避免索引变化
                    for i in range(count - 1, -1, -1):
                        col_num = start_col + i
                        if 1 <= col_num <= sheet.max_column:
                            sheet.delete_cols(col_num)
            
            self.save_workbook(wb, file_path)
            return True
        except Exception as e:
            print(f"删除失败: {e}")
            return False
    
    # ========== 功能3: 添加行或列 ==========
    def insert_rows_or_columns(self, file_path: Path, row_info: Tuple[int, int] = None, col_info: Tuple[int, int] = None):
        """在指定位置插入行或列
        
        Args:
            row_info: (起始行号, 数量) 元组，如 (3, 5) 表示在第3行位置插入5行
            col_info: (起始列号, 数量) 元组，如 (2, 3) 表示在第2列位置插入3列
        """
        wb = self.load_workbook(file_path)
        if not wb:
            return False
        
        try:
            for sheet in wb.worksheets:
                # 插入行
                if row_info:
                    start_row, count = row_info
                    if 1 <= start_row <= sheet.max_row + 1:
                        # 在指定位置插入多行
                        sheet.insert_rows(start_row, count)
                
                # 插入列
                if col_info:
                    start_col, count = col_info
                    if 1 <= start_col <= sheet.max_column + 1:
                        # 在指定位置插入多列
                        sheet.insert_cols(start_col, count)
            
            self.save_workbook(wb, file_path)
            return True
        except Exception as e:
            print(f"添加失败: {e}")
            return False
    
    # ========== 功能4: 改SHEET名 ==========
    def rename_sheets_by_a1(self, file_path: Path):
        """根据A1单元格内容重命名Sheet"""
        wb = self.load_workbook(file_path)
        if not wb:
            return False
        
        try:
            for sheet in wb.worksheets:
                a1_value = sheet['A1'].value
                if a1_value:
                    new_name = str(a1_value).strip()
                    # Excel Sheet名称限制：最大31字符，不能包含某些特殊字符
                    new_name = new_name[:31].replace('/', '_').replace('\\', '_').replace('?', '_').replace('*', '_').replace('[', '_').replace(']', '_').replace(':', '_')
                    if new_name:
                        sheet.title = new_name
            
            self.save_workbook(wb, file_path)
            return True
        except Exception as e:
            print(f"重命名失败: {e}")
            return False
    
    # ========== 功能5: 收集目录 ==========
    def collect_directory(self, directory: Path, output_file: Path):
        """收集目录中所有Excel文件的文件名和Sheet名"""
        excel_files = self.get_excel_files(str(directory))
        # 排除输出文件自身（避免循环引用和文件占用冲突）
        output_name = output_file.name.lower()
        excel_files = [f for f in excel_files if f.name.lower() != output_name]

        wb = Workbook()
        ws = wb.active
        ws.title = "目录"
        
        # 设置表头
        ws['A1'] = '文件名'
        ws['B1'] = 'Sheet名称'
        
        row = 2
        for file_path in excel_files:
            file_wb = self.load_workbook(file_path)
            if file_wb:
                for sheet_name in file_wb.sheetnames:
                    ws[f'A{row}'] = file_path.name
                    ws[f'B{row}'] = sheet_name
                    row += 1
        
        wb.save(output_file)
        return True
    
    def collect_all_filenames(self, directory: Path, output_file: Path):
        """收集目录及子目录中所有文件名,按目录组织输出到Excel
        
        功能描述:
        1. 读取当前目录中所有目录及子目录的文件名
        2. 将每个目录中的文件名输出到EXCEL中
        3. 输出原则:
           a. 第一列: 目录名(相对路径)
           b. 后续列: 该目录中的文件名列表(每个文件名记录在目录名相同行后面的不同列中)
        4. 输出文件为: all_name.xlsx
        
        Args:
            directory: 要扫描的根目录
            output_file: 输出的Excel文件路径
            
        Returns:
            bool: 是否成功
        """
        try:
            # 创建新的工作簿
            wb = Workbook()
            ws = wb.active
            ws.title = "所有文件名"
            
            # 设置表头
            ws['A1'] = '目录名'
            ws['B1'] = '文件列表'
            
            # 设置表头样式
            header_font = Font(name='微软雅黑', size=11, bold=True)
            header_alignment = Alignment(horizontal='center', vertical='center')
            ws['A1'].font = header_font
            ws['B1'].font = header_font
            ws['A1'].alignment = header_alignment
            ws['B1'].alignment = header_alignment
            
            # 收集所有目录和文件
            # 使用字典存储: {目录路径: [文件名列表]}
            dir_files = {}
            # 获取输出文件的文件名（用于在扫描时排除自身）
            output_filename = output_file.name.lower()

            # 递归遍历所有目录和文件
            for root, dirs, files in os.walk(directory):
                root_path = Path(root)

                # 计算相对路径
                try:
                    relative_path = root_path.relative_to(directory)
                    # 如果是根目录本身,使用 "." 表示
                    dir_name = str(relative_path) if str(relative_path) != '.' else '根目录'
                except ValueError:
                    dir_name = str(root_path)

                # 过滤掉临时文件、隐藏文件，以及输出文件本身（避免循环引用）
                filtered_files = [
                    f for f in files
                    if not f.startswith('~$')
                    and not f.startswith('.')
                    and f.lower() != output_filename
                ]

                # 只记录包含文件的目录
                if filtered_files:
                    dir_files[dir_name] = sorted(filtered_files)

            # 写入Excel
            current_row = 2
            for dir_name in sorted(dir_files.keys()):
                files = dir_files[dir_name]

                # 第一列写入目录名
                ws.cell(row=current_row, column=1, value=dir_name)

                # 从第二列开始写入文件名列表
                for col_idx, filename in enumerate(files, start=2):
                    ws.cell(row=current_row, column=col_idx, value=filename)

                current_row += 1

            # 自动调整列宽
            for col in range(1, ws.max_column + 1):
                max_length = 0
                column = get_column_letter(col)

                for cell in ws[column]:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass

                # 设置列宽(最小8,最大50)
                adjusted_width = min(max(max_length + 2, 12), 50)
                ws.column_dimensions[column].width = adjusted_width

            # 保存文件（直接调用 wb.save，异常会被外层 except 捕获并返回 False）
            wb.save(output_file)
            print(f"[OK] 文件名收集完成，共处理 {len(dir_files)} 个目录，输出：{output_file.name}")
            return True
            
        except Exception as e:
            print(f"收集文件名失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    # ========== 功能9: 拆分单元格 ==========
    def unmerge_cells(self, file_path: Path):
        """取消合并单元格并填充内容"""
        wb = self.load_workbook(file_path)
        if not wb:
            return False
        
        try:
            for sheet in wb.worksheets:
                # 获取所有合并单元格
                merged_ranges = list(sheet.merged_cells.ranges)
                
                for merged_range in merged_ranges:
                    # 获取合并区域的值（左上角单元格的值）
                    top_left_cell = sheet[merged_range.min_row][merged_range.min_col - 1]
                    value = top_left_cell.value
                    
                    # 取消合并
                    sheet.unmerge_cells(str(merged_range))
                    
                    # 如果只跨行不跨列，填充内容
                    if merged_range.min_col == merged_range.max_col:
                        col = merged_range.min_col
                        for row in range(merged_range.min_row, merged_range.max_row + 1):
                            sheet.cell(row=row, column=col).value = value
            
            self.save_workbook(wb, file_path)
            return True
        except Exception as e:
            print(f"拆分单元格失败: {e}")
            return False

    def unmerge_v2_deduplicate(self, file_path: Path) -> dict:
        """拆分V2去重：严格逻辑实现
        
        Returns:
            dict: 处理统计信息 {sheet_name: {'unmerge': count, 'fill': count, 'delete': count}}
        """
        wb = self.load_workbook(file_path)
        if not wb:
            return None
        
        stats = {}
        
        try:
            for sheet in wb.worksheets:
                sheet_stats = {'unmerge': 0, 'fill': 0, 'delete': 0}
                stats[sheet.title] = sheet_stats
                
                # 1. 预先扫描并缓存所有合并单元格
                # copying the range objects because unmerge will remove them from the collection
                cached_merged_ranges = [
                    copy(r) for r in sheet.merged_cells.ranges
                ]
                
                # 2. 执行 unmerge (全部取消)
                # 直接清空 merged_cells list 或逐个 unmerge
                # 逐个 unmerge 更安全，虽然效率低一点但能保证状态正确
                # 注意：openpyxl 的 merged_cells 是一个特殊的 list，遍历移除可能导致问题，也就是为什么上面缓存了
                # 这里我们直接对 sheet 执行 unmerge
                for merged_range in cached_merged_ranges:
                    try:
                        sheet.unmerge_cells(str(merged_range))
                        sheet_stats['unmerge'] += 1
                    except KeyError:
                        pass
                
                # 3. 单列跨行补齐并记录补齐列
                # key: (min_row, max_row), value: list of col_indices
                fill_intervals = {} 
                
                for merged_range in cached_merged_ranges:
                    min_col, min_row, max_col, max_row = merged_range.bounds
                    
                    # 判断规则
                    width = max_col - min_col + 1
                    height = max_row - min_row + 1
                    
                    if width == 1 and height > 1:
                        # 规则: 列宽=1 且 行高>1 -> 补齐
                        # 获取左上角的值
                        top_left_cell = sheet.cell(row=min_row, column=min_col)
                        top_left_value = top_left_cell.value
                        
                        # 向下填充
                        for row in range(min_row + 1, max_row + 1):
                            cell = sheet.cell(row=row, column=min_col)
                            cell.value = top_left_value
                            # 尝试复制样式 (基本属性)
                            if top_left_cell.has_style:
                                cell._style = copy(top_left_cell._style)
                                
                        sheet_stats['fill'] += 1
                        
                        # 记录补齐列区间
                        interval_key = (min_row, max_row)
                        if interval_key not in fill_intervals:
                            fill_intervals[interval_key] = []
                        fill_intervals[interval_key].append(min_col)
                        
                    elif width > 1:
                        # 规则: 列宽>1 -> 不做任何复制或填充
                        pass
                
                # 4. 按补齐列规则执行行去重
                # 识别区间：同一连续行区间内；至少3列及以上属于补齐列
                rows_to_delete = set()
                
                for (start_row, end_row), cols in fill_intervals.items():
                    if len(cols) >= 3:
                        # 满足去重条件：3列及以上，且高度一致（key就是start, end）
                        # 规则：仅保留第1行，删除其余重复行
                        # 需要验证签名吗？因为我们是刚刚填充的，这些列的值在 start_row 到 end_row 肯定是一样的。
                        # 用户规则："若区间内第2行至第N行的行签名与第1行完全一致"
                        # "仅使用‘补齐列’的值生成行签名"
                        # 由于这些列刚刚被填充为相同的值，所以行签名必然一致。
                        # 我们可以直接标记删除。
                        
                        for r in range(start_row + 1, end_row + 1):
                            rows_to_delete.add(r)
                
                # 5. 执行删除 (自下向上)
                sorted_rows = sorted(list(rows_to_delete), reverse=True)
                for row_idx in sorted_rows:
                    sheet.delete_rows(row_idx)
                    sheet_stats['delete'] += 1

            self.save_workbook(wb, file_path)
            return stats
        except Exception as e:
            print(f"拆分V2去重失败: {e}")
            return None


    
    # ========== 功能10: 合并EXCEL ==========
    def merge_excel_files(self, directory: Path, output_file: Path):
        """合并目录中所有Excel文件到一个文件"""
        excel_files = [f for f in self.get_excel_files(str(directory)) 
                      if f.name.lower() != 'all.xlsx']
        
        wb = Workbook()
        ws = wb.active
        ws.title = 'a'
        
        seen_rows = set()
        current_row = 1
        
        for file_path in excel_files:
            file_wb = self.load_workbook(file_path)
            if not file_wb:
                continue
            
            for sheet in file_wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    # 转换为元组用于去重
                    row_tuple = tuple(row)
                    if row_tuple not in seen_rows:
                        seen_rows.add(row_tuple)
                        for col_idx, value in enumerate(row, start=1):
                            ws.cell(row=current_row, column=col_idx).value = value
                        current_row += 1
        
        self.save_workbook(wb, output_file)
        return True
    
    # ========== 功能11: 合并EXCEL-列变行 ==========
    def merge_excel_transpose(self, directory: Path, output_file: Path):
        """合并Excel文件并转置（列变行）"""
        excel_files = [f for f in self.get_excel_files(str(directory)) 
                      if f.name.lower() != 'all.xlsx']
        
        wb = Workbook()
        ws = wb.active
        ws.title = 'a'
        
        seen_rows = set()
        current_row = 1
        
        for file_path in excel_files:
            file_wb = self.load_workbook(file_path)
            if not file_wb:
                continue
            
            for sheet in file_wb.worksheets:
                # 读取数据
                data = []
                for row in sheet.iter_rows(values_only=True):
                    data.append(list(row))
                
                if not data:
                    continue
                
                # 转置：列变行
                transposed = list(zip(*data)) if data else []
                
                for col_data in transposed:
                    # 第一列是Sheet名称
                    row_data = [sheet.title] + list(col_data)
                    row_tuple = tuple(row_data)
                    
                    if row_tuple not in seen_rows:
                        seen_rows.add(row_tuple)
                        for col_idx, value in enumerate(row_data, start=1):
                            ws.cell(row=current_row, column=col_idx).value = value
                        current_row += 1
        
        self.save_workbook(wb, output_file)
        return True
    
    # ========== 功能12: 线性文本 ==========
    def convert_to_linear_text(self, file_path: Path, output_file: Path):
        """将Excel转换为线性文本格式"""
        wb = self.load_workbook(file_path)
        if not wb:
            return False
        
        try:
            lines = []
            
            for sheet in wb.worksheets:
                # 获取表头（第一行）
                headers = []
                for col in range(1, sheet.max_column + 1):
                    cell_value = sheet.cell(row=1, column=col).value
                    if cell_value:
                        headers.append(str(cell_value).replace('\n', ';').replace(',', '_'))
                    else:
                        headers.append('')
                
                # 处理每一行数据
                for row_idx in range(2, sheet.max_row + 1):
                    # 获取产品名称（假设第一列）
                    product_name = sheet.cell(row=row_idx, column=1).value
                    if not product_name:
                        continue
                    
                    product_name = str(product_name).replace('\n', ';').replace(',', '_')
                    
                    # 构建行文本
                    row_parts = []
                    for col_idx in range(1, sheet.max_column + 1):
                        cell_value = sheet.cell(row=row_idx, column=col_idx).value
                        if cell_value:
                            cell_str = str(cell_value).replace('\n', ';').replace(',', '_')
                            header = headers[col_idx - 1] if col_idx <= len(headers) else ''
                            row_parts.append(f"{product_name}+{header}: {cell_str}")
                    
                    if row_parts:
                        lines.append('，'.join(row_parts))
            
            # 保存为UTF-8编码的文本文件
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write('\n'.join(lines))
            
            return True
        except Exception as e:
            print(f"转换文本失败: {e}")
            return False

    def convert_excel_to_txt(self, file_path: Path) -> bool:
        """
        将Excel文件转换为TXT文本文件（TSV格式）
        使用制表符作为分隔符，避免与内容冲突
        
        Args:
            file_path: Excel文件路径
            
        Returns:
            bool: 转换成功返回True，失败返回False
        """
        wb = self.load_workbook(file_path)
        if not wb:
            return False
        
        try:
            # 输出文件名：原文件名_sheet名.txt
            
            for sheet in wb.worksheets:
                lines = []
                
                # 遍历所有行
                for row in sheet.iter_rows(values_only=True):
                    # 处理每个单元格的值
                    processed_cells = []
                    for cell_value in row:
                        if cell_value is None:
                            processed_cells.append('')
                        else:
                            # 转换为字符串并处理特殊字符
                            cell_str = str(cell_value)
                            # 替换换行符为空格
                            cell_str = cell_str.replace('\n', ' ').replace('\r', ' ')
                            # 替换制表符为空格（避免破坏分隔符）
                            cell_str = cell_str.replace('\t', ' ')
                            processed_cells.append(cell_str)
                    
                    # 使用制表符连接
                    lines.append('\t'.join(processed_cells))
                
                # 生成输出文件名
                # 注意：不使用with_suffix()，因为当文件名包含多个点号时
                # (如 file.1.15_part01)，with_suffix会错误地将最后一个点后的内容当作后缀替换掉
                if len(wb.worksheets) == 1:
                    output_file = file_path.parent / f"{file_path.stem}.txt"
                else:
                    output_file = file_path.parent / f"{file_path.stem}_{sheet.title}.txt"
                
                # 保存为UTF-8编码（带BOM，确保Windows中文显示正确）
                with open(output_file, 'w', encoding='utf-8-sig') as f:
                    f.write('\n'.join(lines))
                
                print(f"已转换: {output_file.name}")
            
            return True
            
        except Exception as e:
            print(f"转换失败 {file_path.name}: {e}")
            return False
        
    # ========== 功能13: 单元格加序号 ==========
    def add_numbering_to_cells(self, file_path: Path):
        """为每一个非单行内容的单元格加序号
        
        功能描述：
        1. 遍历Excel文件的所有单元格，为每一个单元格中非一行内容的单元格加序号，都要用"1、 2、 3、 ……"格式
        2. 原有单元格中，存在"1、2、3、……"不做修改
        3. 原有单元格中，存在"1. 2. 3. ……"将"."换为"、"
        4. 遇到空行，则跳过此行，为下一行继续添加序号
        """
        wb = self.load_workbook(file_path)
        if not wb:
            return False
        
        try:
            for sheet in wb.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value is None:
                            continue
                        
                        text = str(cell.value)
                        lines = text.split('\n')
                        
                        # 如果只有一行内容，不需要处理
                        if len(lines) <= 1:
                            continue
                        
                        # 检查是否已有完整的中文序号格式（所有非空行都以"数字、"开头）
                        has_chinese_numbering = True
                        for line in lines:
                            line_stripped = line.strip()
                            if line_stripped:  # 非空行
                                if not re.match(r'^\d+、', line_stripped):
                                    has_chinese_numbering = False
                                    break
                        
                        # 如果已有完整的中文序号格式，不做修改
                        if has_chinese_numbering:
                            continue
                        
                        # 先处理英文序号格式，将"."换为"、"
                        processed_lines = []
                        for line in lines:
                            # 替换行首的英文序号格式为中文格式（"数字. "或"数字." -> "数字、"）
                            line = re.sub(r'^(\d+)\.\s*', r'\1、', line)
                            processed_lines.append(line)
                        
                        # 为多行内容添加序号，遇到空行跳过，但序号继续递增
                        numbered_lines = []
                        line_number = 1
                        
                        for line in processed_lines:
                            line_stripped = line.strip()
                            if line_stripped:  # 非空行
                                # 检查是否已经有序号格式（中文或英文）
                                match = re.match(r'^(\d+)[、.]\s*', line_stripped)
                                if match:
                                    # 如果已有序号，确保是中文格式
                                    line = re.sub(r'^(\d+)\.\s*', r'\1、', line)
                                    numbered_lines.append(line)
                                    # 从已有序号中提取数字，更新line_number（确保序号连续）
                                    existing_num = int(match.group(1))
                                    # 如果已有序号大于当前序号，更新当前序号
                                    if existing_num >= line_number:
                                        line_number = existing_num + 1
                                else:
                                    # 没有序号，添加序号（格式：数字、 内容）
                                    numbered_lines.append(f"{line_number}、 {line_stripped}")
                                    line_number += 1
                            else:
                                # 空行，跳过，不添加序号，但保留空行
                                numbered_lines.append(line)
                        
                        cell.value = '\n'.join(numbered_lines)
            
            self.save_workbook(wb, file_path)
            return True
        except Exception as e:
            print(f"加序号失败: {e}")
            return False
    
    # ========== 功能14: A1内容加列 ==========
    def add_a1_column(self, file_path: Path):
        """将A1内容复制到新列"""
        wb = self.load_workbook(file_path)
        if not wb:
            return False
        
        try:
            for sheet in wb.worksheets:
                a1_value = sheet['A1'].value
                if not a1_value:
                    continue
                
                # 找到最后一个有内容的列
                last_col = 1
                for col in range(1, sheet.max_column + 1):
                    for row in range(1, sheet.max_row + 1):
                        if sheet.cell(row=row, column=col).value:
                            last_col = max(last_col, col)
                            break
                
                # 在新列填充A1内容
                new_col = last_col + 1
                for row in range(1, sheet.max_row + 1):
                    # 检查该行是否有内容
                    has_content = False
                    for col in range(1, last_col + 1):
                        if sheet.cell(row=row, column=col).value:
                            has_content = True
                            break
                    
                    if has_content:
                        sheet.cell(row=row, column=new_col).value = a1_value
            
            self.save_workbook(wb, file_path)
            return True
        except Exception as e:
            print(f"A1加列失败: {e}")
            return False
    
    # ========== 功能15: 全角转半角 ==========
    def convert_fullwidth_to_halfwidth(self, file_path: Path):
        """全角转半角
        
        功能描述：
        1. 数字、英文：都转化为英文的半角数字、英文
        2. 删除："　"全角空格和" "半角空格
        3. 删除：制表符
        4. 将所有的"——"替换成"-"
        5. 所有的全角括号"（）"都替换为半角的"()"
        """
        wb = self.load_workbook(file_path)
        if not wb:
            return False
        
        try:
            # 创建全角到半角的映射表
            fullwidth_to_halfwidth = {}
            
            # 全角数字转半角（０-９ -> 0-9）
            for i in range(10):
                fullwidth_to_halfwidth[chr(0xFF10 + i)] = chr(0x30 + i)  # ０-９ -> 0-9
            
            # 全角英文大写转半角（Ａ-Ｚ -> A-Z）
            for i in range(26):
                fullwidth_to_halfwidth[chr(0xFF21 + i)] = chr(0x41 + i)  # Ａ-Ｚ -> A-Z
            
            # 全角英文小写转半角（ａ-ｚ -> a-z）
            for i in range(26):
                fullwidth_to_halfwidth[chr(0xFF41 + i)] = chr(0x61 + i)  # ａ-ｚ -> a-z
            
            # 全角空格和半角空格
            fullwidth_to_halfwidth['　'] = ''  # 全角空格删除
            fullwidth_to_halfwidth[' '] = ''   # 半角空格删除
            
            # 全角括号
            fullwidth_to_halfwidth['（'] = '('
            fullwidth_to_halfwidth['）'] = ')'
            
            for sheet in wb.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value is None:
                            continue
                        
                        # 如果是数字类型，先转换为字符串处理
                        original_type = type(cell.value)
                        text = str(cell.value)
                        
                        # 使用映射表进行转换
                        result = []
                        for char in text:
                            if char in fullwidth_to_halfwidth:
                                replacement = fullwidth_to_halfwidth[char]
                                if replacement:  # 如果不是删除操作
                                    result.append(replacement)
                                # 如果是删除操作（全角空格或半角空格），不添加任何字符
                            else:
                                result.append(char)
                        
                        text = ''.join(result)
                        
                        # 删除制表符
                        text = text.replace('\t', '')
                        
                        # 将所有的"——"替换成"-"
                        text = text.replace('——', '-')
                        
                        # 更新单元格值
                        # 如果原值是数字类型，尝试转换回数字
                        if original_type in (int, float):
                            try:
                                # 尝试转换为数字
                                if '.' in text or 'e' in text.lower() or 'E' in text:
                                    cell.value = float(text)
                                else:
                                    cell.value = int(text)
                            except (ValueError, OverflowError):
                                # 转换失败，保持字符串
                                cell.value = text
                        else:
                            cell.value = text
            
            self.save_workbook(wb, file_path)
            return True
        except Exception as e:
            print(f"全角转半角失败: {e}")
            return False
    
    # ========== 功能16: 删除指定SHEET页 ==========
    def delete_sheet(self, file_path: Path, sheet_name: str) -> bool:
        """删除指定的SHEET页
        
        Args:
            file_path: Excel文件路径
            sheet_name: 要删除的SHEET页名称
        
        Returns:
            bool: 是否成功删除
        """
        try:
            file_ext = file_path.suffix.lower()
            
            # 对于.xlsx文件，使用openpyxl处理
            if file_ext == '.xlsx':
                wb = self.load_workbook(file_path)
                if not wb:
                    return False
                
                # 检查SHEET是否存在
                if sheet_name not in wb.sheetnames:
                    print(f"文件 {file_path.name} 中未找到名为 '{sheet_name}' 的SHEET页")
                    return False
                
                # 检查是否只有一个SHEET，Excel文件至少需要一个SHEET
                if len(wb.sheetnames) <= 1:
                    print(f"文件 {file_path.name} 只有一个SHEET页，无法删除")
                    return False
                
                # 删除SHEET
                wb.remove(wb[sheet_name])
                self.save_workbook(wb, file_path)
                return True
            
            # 对于.xls文件，使用pandas处理
            elif file_ext == '.xls':
                try:
                    # 读取所有sheet
                    excel_file = pd.ExcelFile(file_path)
                    all_sheets = excel_file.sheet_names
                    
                    # 检查SHEET是否存在
                    if sheet_name not in all_sheets:
                        print(f"文件 {file_path.name} 中未找到名为 '{sheet_name}' 的SHEET页")
                        return False
                    
                    # 检查是否只有一个SHEET
                    if len(all_sheets) <= 1:
                        print(f"文件 {file_path.name} 只有一个SHEET页，无法删除")
                        return False
                    
                    # 读取所有需要保留的sheet
                    sheets_to_keep = [s for s in all_sheets if s != sheet_name]
                    
                    # 创建新的ExcelWriter，保存为.xlsx格式（pandas保存.xls可能有问题）
                    # 使用临时文件名，避免与原文件冲突
                    temp_output_path = file_path.parent / f"{file_path.stem}_temp_{file_path.suffix}.xlsx"
                    with pd.ExcelWriter(temp_output_path, engine='openpyxl') as writer:
                        for sheet in sheets_to_keep:
                            df = pd.read_excel(file_path, sheet_name=sheet)
                            df.to_excel(writer, sheet_name=sheet, index=False)
                    
                    # 备份原.xls文件
                    backup_path = file_path.with_suffix('.xls.bak')
                    if backup_path.exists():
                        backup_path.unlink()
                    file_path.rename(backup_path)
                    
                    # 将临时文件重命名为.xlsx格式
                    final_output_path = file_path.with_suffix('.xlsx')
                    if final_output_path.exists():
                        final_output_path.unlink()
                    temp_output_path.rename(final_output_path)
                    
                    print(f"注意：文件 {file_path.name} 已转换为.xlsx格式（原文件备份为.xls.bak）")
                    
                    return True
                except Exception as e:
                    print(f"处理.xls文件失败 {file_path}: {e}")
                    return False
            
            else:
                print(f"不支持的文件格式: {file_ext}")
                return False
                
        except Exception as e:
            print(f"删除SHEET页失败 {file_path}: {e}")
            return False
    
    # ========== 功能18: 清除公式 ==========
    def clear_formulas(self, file_path: Path) -> bool:
        """清除公式，将公式转化为数字（结果不变）
        
        功能描述：
        1. 遍历Excel文件的所有单元格
        2. 如果单元格包含公式，则获取公式的计算结果（数值）
        3. 用计算结果替换公式，保持结果不变
        
        Args:
            file_path: Excel文件路径
        
        Returns:
            bool: 是否成功清除公式
        """
        try:
            file_ext = file_path.suffix.lower()
            
            # 对于.xlsx文件，使用openpyxl处理
            if file_ext == '.xlsx':
                # 使用data_only=True加载工作簿，这样公式会被转换为计算结果（数值）
                # 注意：如果文件之前没有被Excel打开并保存过，某些公式可能无法获取计算结果
                try:
                    # 使用 self.load_workbook 的逻辑检查文件有效性，但这里我们需要 data_only=True
                    # 所以手动复用检查逻辑
                    if file_path.stat().st_size == 0:
                        print(f"Skipping empty file: {file_path}")
                        return False
                        
                    wb = openpyxl.load_workbook(file_path, data_only=True)
                except zipfile.BadZipFile:
                    print(f"Error processing {file_path.name}: File is not a valid Excel file (possibly encrypted or corrupted).")
                    return False
                except Exception as e:
                    print(f"加载文件失败 {file_path}: {e}")
                    return False
                
                if not wb:
                    return False
                
                # 此时所有公式都已被转换为计算结果（如果文件之前被Excel计算过）
                # 直接保存即可，公式会被替换为计算结果
                self.save_workbook(wb, file_path)
                return True
            
            # 对于.xls文件，使用pandas处理
            elif file_ext == '.xls':
                try:
                    # 读取所有sheet
                    excel_file = pd.ExcelFile(file_path)
                    all_sheets = excel_file.sheet_names
                    
                    # 使用临时文件名，避免与原文件冲突
                    temp_output_path = file_path.parent / f"{file_path.stem}_temp_{file_path.suffix}.xlsx"
                    
                    # 读取所有sheet并保存为.xlsx格式（pandas会自动计算公式）
                    with pd.ExcelWriter(temp_output_path, engine='openpyxl') as writer:
                        for sheet_name in all_sheets:
                            # 读取数据（pandas会自动计算公式）
                            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                            # 保存为Excel（此时公式已被计算结果替换）
                            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                    
                    # 备份原.xls文件
                    backup_path = file_path.with_suffix('.xls.bak')
                    if backup_path.exists():
                        backup_path.unlink()
                    file_path.rename(backup_path)
                    
                    # 将临时文件重命名为.xlsx格式
                    final_output_path = file_path.with_suffix('.xlsx')
                    if final_output_path.exists():
                        final_output_path.unlink()
                    temp_output_path.rename(final_output_path)
                    
                    print(f"注意：文件 {file_path.name} 已转换为.xlsx格式（原文件备份为.xls.bak）")
                    
                    return True
                except Exception as e:
                    print(f"处理.xls文件失败 {file_path}: {e}")
                    return False
            
            else:
                print(f"不支持的文件格式: {file_ext}")
                return False
                
        except Exception as e:
            print(f"清除公式失败 {file_path}: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    # ========== 功能20: 删表及合并 ==========
    def delete_sheets(self, file_path: Path, sheet_names: List[str]) -> bool:
        """删除指定的一个或多个SHEET页
        
        Args:
            file_path: Excel文件路径
            sheet_names: 要删除的SHEET页名称列表
        
        Returns:
            bool: 是否成功删除
        """
        try:
            file_ext = file_path.suffix.lower()
            
            # 对于.xlsx文件，使用openpyxl处理
            if file_ext == '.xlsx':
                wb = self.load_workbook(file_path)
                if not wb:
                    return False
                
                # 检查要删除的SHEET是否存在
                existing_sheets = [s for s in sheet_names if s in wb.sheetnames]
                if not existing_sheets:
                    print(f"文件 {file_path.name} 中未找到任何指定的SHEET页")
                    return False
                
                # 检查删除后是否还有SHEET，Excel文件至少需要一个SHEET
                if len(wb.sheetnames) <= len(existing_sheets):
                    print(f"文件 {file_path.name} 删除后没有剩余的SHEET页，无法删除")
                    return False
                
                # 删除所有指定的SHEET
                for sheet_name in existing_sheets:
                    wb.remove(wb[sheet_name])
                    print(f"已删除SHEET页: {sheet_name}")
                
                self.save_workbook(wb, file_path)
                return True
            
            # 对于.xls文件，使用pandas处理
            elif file_ext == '.xls':
                try:
                    # 读取所有sheet
                    excel_file = pd.ExcelFile(file_path)
                    all_sheets = excel_file.sheet_names
                    
                    # 检查要删除的SHEET是否存在
                    existing_sheets = [s for s in sheet_names if s in all_sheets]
                    if not existing_sheets:
                        print(f"文件 {file_path.name} 中未找到任何指定的SHEET页")
                        return False
                    
                    # 检查删除后是否还有SHEET
                    if len(all_sheets) <= len(existing_sheets):
                        print(f"文件 {file_path.name} 删除后没有剩余的SHEET页，无法删除")
                        return False
                    
                    # 读取所有需要保留的sheet
                    sheets_to_keep = [s for s in all_sheets if s not in existing_sheets]
                    
                    # 创建新的ExcelWriter，保存为.xlsx格式（pandas保存.xls可能有问题）
                    # 使用临时文件名，避免与原文件冲突
                    temp_output_path = file_path.parent / f"{file_path.stem}_temp_{file_path.suffix}.xlsx"
                    with pd.ExcelWriter(temp_output_path, engine='openpyxl') as writer:
                        for sheet in sheets_to_keep:
                            df = pd.read_excel(file_path, sheet_name=sheet)
                            df.to_excel(writer, sheet_name=sheet, index=False)
                    
                    # 备份原.xls文件
                    backup_path = file_path.with_suffix('.xls.bak')
                    if backup_path.exists():
                        backup_path.unlink()
                    file_path.rename(backup_path)
                    
                    # 将临时文件重命名为.xlsx格式
                    final_output_path = file_path.with_suffix('.xlsx')
                    if final_output_path.exists():
                        final_output_path.unlink()
                    temp_output_path.rename(final_output_path)
                    
                    print(f"已删除SHEET页: {', '.join(existing_sheets)}")
                    print(f"注意：文件 {file_path.name} 已转换为.xlsx格式（原文件备份为.xls.bak）")
                    
                    return True
                except Exception as e:
                    print(f"处理.xls文件失败 {file_path}: {e}")
                    return False
            
            else:
                print(f"不支持的文件格式: {file_ext}")
                return False
                
        except Exception as e:
            print(f"删除SHEET页失败 {file_path}: {e}")
            return False
    
    # ========== 功能21: 归属确定 ==========
    def set_ownership(self, file_path: Path, cell_addresses: List[str]) -> bool:
        """归属确定 - 从指定单元格读取内容，复制到最后一个有内容的列的后面一列
        
        Args:
            file_path: Excel文件路径
            cell_addresses: 单元格地址列表（如['A1', 'B2', 'C3']）
        
        Returns:
            bool: 是否成功处理
        """
        try:
            file_ext = file_path.suffix.lower()
            
            # 对于.xlsx文件，使用openpyxl处理
            if file_ext == '.xlsx':
                wb = self.load_workbook(file_path)
                if not wb:
                    return False
                
                try:
                    for sheet in wb.worksheets:
                        # 从指定单元格读取内容
                        cell_values = []
                        for cell_address in cell_addresses:
                            try:
                                cell = sheet[cell_address]
                                cell_value = cell.value
                                # 将单元格值转换为字符串，None转为空字符串
                                if cell_value is None:
                                    cell_values.append("")
                                else:
                                    cell_values.append(str(cell_value).strip())
                            except Exception as e:
                                print(f"读取单元格 {cell_address} 失败: {e}")
                                cell_values.append("")
                        
                        # 用"-"连接所有单元格的值
                        final_value = "-".join(cell_values)
                        
                        # 找到最后一个有内容的列
                        last_col = 0
                        for col in range(1, sheet.max_column + 1):
                            for row in range(1, sheet.max_row + 1):
                                cell_value = sheet.cell(row=row, column=col).value
                                if cell_value is not None and str(cell_value).strip():
                                    last_col = max(last_col, col)
                                    break
                        
                        # 如果没有找到有内容的列，跳过
                        if last_col == 0:
                            continue
                        
                        # 新列位置
                        new_col = last_col + 1
                        
                        # 找到所有有内容的行（从第1列到last_col列中至少有一列有内容）
                        rows_with_content = []
                        for row in range(1, sheet.max_row + 1):
                            has_content = False
                            for col in range(1, last_col + 1):
                                cell_value = sheet.cell(row=row, column=col).value
                                if cell_value is not None and str(cell_value).strip():
                                    has_content = True
                                    break
                            if has_content:
                                rows_with_content.append(row)
                        
                        # 如果没有有内容的行，跳过
                        if not rows_with_content:
                            continue
                        
                        # 将读取的内容填入新列的所有有内容的行（所有行填入相同的值）
                        for row in rows_with_content:
                            sheet.cell(row=row, column=new_col).value = final_value
                    
                    self.save_workbook(wb, file_path)
                    return True
                except Exception as e:
                    print(f"处理失败 {file_path}: {e}")
                    return False
            
            # 对于.xls文件，使用pandas处理
            elif file_ext == '.xls':
                try:
                    # 读取所有sheet
                    excel_file = pd.ExcelFile(file_path)
                    all_sheets = excel_file.sheet_names
                    
                    # 使用临时文件名，避免与原文件冲突
                    temp_output_path = file_path.parent / f"{file_path.stem}_temp_{file_path.suffix}.xlsx"
                    
                    with pd.ExcelWriter(temp_output_path, engine='openpyxl') as writer:
                        for sheet_name in all_sheets:
                            # 读取数据
                            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                            
                            # 从指定单元格读取内容
                            cell_values = []
                            for cell_address in cell_addresses:
                                try:
                                    # 将单元格地址转换为行列索引（如A1 -> row=0, col=0）
                                    from openpyxl.utils import coordinate_from_string, column_index_from_string
                                    col_letter, row_num = coordinate_from_string(cell_address)
                                    col_idx = column_index_from_string(col_letter) - 1  # pandas使用0-based索引
                                    row_idx = row_num - 1  # pandas使用0-based索引
                                    
                                    # 检查索引是否有效
                                    if 0 <= row_idx < len(df) and 0 <= col_idx < len(df.columns):
                                        cell_value = df.iloc[row_idx, col_idx]
                                        if pd.isna(cell_value):
                                            cell_values.append("")
                                        else:
                                            cell_values.append(str(cell_value).strip())
                                    else:
                                        print(f"单元格 {cell_address} 超出范围")
                                        cell_values.append("")
                                except Exception as e:
                                    print(f"读取单元格 {cell_address} 失败: {e}")
                                    cell_values.append("")
                            
                            # 用"-"连接所有单元格的值
                            final_value = "-".join(cell_values)
                            
                            # 找到最后一个有内容的列
                            last_col = 0
                            for col_idx in range(len(df.columns)):
                                if df.iloc[:, col_idx].notna().any() and (df.iloc[:, col_idx].astype(str).str.strip() != '').any():
                                    last_col = max(last_col, col_idx)
                            
                            # 如果没有找到有内容的列，直接保存
                            if last_col == 0:
                                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                                continue
                            
                            # 新列位置
                            new_col = last_col + 1
                            
                            # 找到所有有内容的行
                            rows_with_content = []
                            for row_idx in range(len(df)):
                                has_content = False
                                for col_idx in range(last_col + 1):
                                    cell_value = df.iloc[row_idx, col_idx]
                                    if pd.notna(cell_value) and str(cell_value).strip():
                                        has_content = True
                                        break
                                if has_content:
                                    rows_with_content.append(row_idx)
                            
                            # 将读取的内容填入新列的所有有内容的行（所有行填入相同的值）
                            if rows_with_content:
                                # 确保新列存在
                                while len(df.columns) <= new_col:
                                    df[len(df.columns)] = None
                                
                                for row_idx in rows_with_content:
                                    df.iloc[row_idx, new_col] = final_value
                            
                            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                    
                    # 备份原.xls文件
                    backup_path = file_path.with_suffix('.xls.bak')
                    if backup_path.exists():
                        backup_path.unlink()
                    file_path.rename(backup_path)
                    
                    # 将临时文件重命名为.xlsx格式
                    final_output_path = file_path.with_suffix('.xlsx')
                    if final_output_path.exists():
                        final_output_path.unlink()
                    temp_output_path.rename(final_output_path)
                    
                    print(f"注意：文件 {file_path.name} 已转换为.xlsx格式（原文件备份为.xls.bak）")
                    
                    return True
                except Exception as e:
                    print(f"处理.xls文件失败 {file_path}: {e}")
                    return False
            
            else:
                print(f"不支持的文件格式: {file_ext}")
                return False
                
        except Exception as e:
            print(f"归属确定失败 {file_path}: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    # ========== 功能23: 多维导入模板 ==========
    def multi_dimension_import_template(self, template_file: Path, target_files: List[Path]) -> bool:
        """多维导入模板
        
        严格按照以下规则执行：
        1. 源文件模板是唯一的主表，首行字段名称及顺序不可变
        2. 以"列名"作为唯一匹配依据，而非列顺序
        3. 仅当输入Sheet中的列名与源文件模板中的列名完全一致时，才允许写入该列
        4. 输入文件中存在、但源文件模板中不存在的列，必须忽略
        5. 源文件模板中存在、但输入文件中不存在的列，该列对应单元格写入为空
        6. 从每个Sheet的第二行开始逐行读取数据
        7. 所有数据合并完成后，以整行数据内容作为唯一性判断标准，删除重复行
        
        Args:
            template_file: 模板源文件路径
            target_files: 目标文件列表（目录中的所有Excel文件或单独一个文件）
        
        Returns:
            bool: 是否成功处理
        """
        try:
            # 【一、加载源文件模板】
            template_wb = self.load_workbook(template_file)
            if not template_wb:
                print(f"无法加载模板文件: {template_file}")
                return False
            
            # 使用第一个sheet作为模板
            template_sheet = template_wb.active
            
            # 【二、读取源文件模板的列名（第一行）】
            # 获取第一行作为列名
            template_column_names = []  # 保持顺序的列名列表
            template_column_name_to_index = {}  # 列名到列索引的映射
            
            print(f"正在读取模板文件列名...")
            for col_idx in range(1, template_sheet.max_column + 1):
                cell_value = template_sheet.cell(row=1, column=col_idx).value
                if cell_value is not None:
                    col_name = str(cell_value).strip()
                    if col_name:
                        template_column_names.append(col_name)
                        template_column_name_to_index[col_name] = col_idx
                    else:
                        # 空列名，也记录位置
                        template_column_names.append("")
                        template_column_name_to_index[""] = col_idx
                else:
                    # None值，也记录位置
                    template_column_names.append("")
                    template_column_name_to_index[""] = col_idx
            
            print(f"模板列名数量: {len(template_column_names)}")
            print(f"模板列名: {template_column_names}")
            
            if not template_column_names:
                print("模板文件没有列名（第一行为空）")
                template_wb.close()
                return False
            
            # 获取模板的列数
            template_max_col = len(template_column_names)
            
            # 【三、收集所有要追加的数据】
            all_new_rows = []  # 使用列表保持顺序，后续去重
            
            # 【四、遍历所有输入Excel文件】
            for target_file in target_files:
                print(f"\n正在处理目标文件: {target_file.name}")
                target_wb = self.load_workbook(target_file)
                if not target_wb:
                    print(f"无法加载目标文件: {target_file}")
                    continue
                
                # 【五、遍历每个文件中的所有Sheet】
                for target_sheet in target_wb.worksheets:
                    print(f"  处理Sheet: {target_sheet.title}")
                    
                    # 检查是否有数据行（至少需要第一行作为列名）
                    if target_sheet.max_row < 2:
                        print(f"    [WARN] Sheet中没有数据行（至少需要第一行作为列名，第二行开始是数据）")
                        continue
                    
                    # 【六、读取输入Sheet的列名（第一行）】
                    target_column_name_to_index = {}  # 输入文件的列名到列索引的映射
                    
                    for col_idx in range(1, target_sheet.max_column + 1):
                        cell_value = target_sheet.cell(row=1, column=col_idx).value
                        if cell_value is not None:
                            col_name = str(cell_value).strip()
                            if col_name:
                                target_column_name_to_index[col_name] = col_idx
                        else:
                            # None值，使用空字符串作为列名
                            target_column_name_to_index[""] = col_idx
                    
                    print(f"    输入Sheet列名数量: {len(target_column_name_to_index)}")
                    print(f"    输入Sheet列名: {list(target_column_name_to_index.keys())[:10]}")
                    
                    # 【七、建立列名映射关系】
                    # 找出输入文件中与模板列名匹配的列
                    matched_columns = {}  # 模板列名 -> 输入文件列索引
                    for template_col_name in template_column_names:
                        if template_col_name in target_column_name_to_index:
                            matched_columns[template_col_name] = target_column_name_to_index[template_col_name]
                    
                    print(f"    匹配的列数: {len(matched_columns)}/{len(template_column_names)}")
                    
                    if not matched_columns:
                        print(f"    [WARN] 没有找到匹配的列名")
                        continue
                    
                    # 【八、从第二行开始逐行读取数据】
                    matched_row_count = 0
                    for target_row_idx in range(2, target_sheet.max_row + 1):
                        # 构建新行：根据模板的列顺序
                        new_row = [None] * template_max_col
                        
                        # 按列名映射到源文件模板对应的列下
                        for template_col_idx, template_col_name in enumerate(template_column_names):
                            if template_col_name in matched_columns:
                                # 输入文件中有该列，读取数据
                                target_col_idx = matched_columns[template_col_name]
                                cell_value = target_sheet.cell(row=target_row_idx, column=target_col_idx).value
                                new_row[template_col_idx] = cell_value
                            # 如果输入文件中没有该列，new_row[template_col_idx]保持为None（空值）
                        
                        # 将新行转换为元组并添加到列表
                        row_tuple = tuple(new_row)
                        all_new_rows.append(row_tuple)
                        matched_row_count += 1
                    
                    print(f"    从该Sheet提取了 {matched_row_count} 行数据")
                
                # 关闭目标文件
                target_wb.close()
            
            print(f"\n总共收集到 {len(all_new_rows)} 行新数据")
            
            # 【九、获取模板中已有的行（用于去重）】
            existing_rows = set()
            # 从第二行开始读取（第一行是列名）
            for row_idx in range(2, template_sheet.max_row + 1):
                row_data = []
                for col_idx in range(1, template_max_col + 1):
                    cell_value = template_sheet.cell(row=row_idx, column=col_idx).value
                    row_data.append(cell_value)
                existing_rows.add(tuple(row_data))
            
            print(f"模板中已有 {len(existing_rows)} 行数据（不包括列名行）")
            
            # 【十、所有数据合并完成后，以整行数据内容作为唯一性判断标准，删除重复行】
            # 合并所有行（包括模板中已有的和新收集的）
            all_rows = list(existing_rows) + all_new_rows
            
            # 使用字典去重，保持第一次出现的行
            unique_rows_dict = {}
            for row_tuple in all_rows:
                if row_tuple not in unique_rows_dict:
                    unique_rows_dict[row_tuple] = row_tuple
            
            unique_rows = list(unique_rows_dict.values())
            print(f"去重后共有 {len(unique_rows)} 行数据（原有{len(existing_rows)}行，新增{len(unique_rows) - len(existing_rows)}行）")
            
            # 【十一、只追加不在已有行中的新行】
            rows_to_add = []
            for row_tuple in unique_rows:
                if row_tuple not in existing_rows:
                    rows_to_add.append(row_tuple)
            
            print(f"将追加 {len(rows_to_add)} 行新数据到模板")
            
            if len(rows_to_add) == 0:
                print("没有新数据需要追加")
                template_wb.close()
                return True
            
            # 【十二、将映射后的完整行追加到源文件模板的末尾】
            new_row_num = template_sheet.max_row + 1
            for row_tuple in rows_to_add:
                for col_idx, cell_value in enumerate(row_tuple, start=1):
                    if col_idx <= template_max_col:
                        template_sheet.cell(row=new_row_num, column=col_idx).value = cell_value
                new_row_num += 1
            
            # 【十三、保存模板文件】
            print(f"正在保存模板文件...")
            self.save_workbook(template_wb, template_file)
            template_wb.close()
            
            print(f"多维导入模板完成，共处理 {len(target_files)} 个目标文件，追加 {len(rows_to_add)} 行新数据")
            return True
            
        except Exception as e:
            print(f"多维导入模板失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    # ========== 功能: EXCEL拆分(GPT20M) ==========
    def split_excel_by_size(self, file_path: Path, max_size_mb: int = 20) -> bool:
        """将Excel文件按大小拆分为多个不超过指定大小的文件
        
        功能描述：
        1. 检查文件大小，如果 ≤ max_size_mb，则跳过
        2. 如果 > max_size_mb，按行拆分为多个文件
        3. 每个拆分文件保留原始表头（第一行）
        4. 拆分文件命名：原文件名_part01.xlsx, 原文件名_part02.xlsx...
        5. 拆分文件与原文件位于同一目录
        6. 不覆盖原文件
        
        Args:
            file_path: Excel文件路径
            max_size_mb: 最大文件大小（MB），默认20MB
            
        Returns:
            bool: 是否成功处理（跳过或拆分成功都返回True）
        """
        try:
            # 1. 检查文件大小
            file_size_bytes = file_path.stat().st_size
            file_size_mb = file_size_bytes / (1024 * 1024)
            max_size_bytes = max_size_mb * 1024 * 1024
            
            print(f"\n检查文件: {file_path.name}")
            print(f"文件大小: {file_size_mb:.2f} MB")
            
            # 如果文件小于等于限制大小，跳过
            if file_size_bytes <= max_size_bytes:
                print(f"[OK] 文件大小在限制内，无需拆分")
                return True
            
            print(f"[INFO] 文件超过 {max_size_mb}MB，开始拆分...")
            
            # 2. 加载工作簿
            wb = self.load_workbook(file_path)
            if not wb:
                print("[FAIL] 无法加载文件")
                return False
            
            try:
                # 只处理第一个工作表（根据需求，可以扩展为处理所有工作表）
                sheet = wb.active
                total_rows = sheet.max_row
                total_cols = sheet.max_column
                
                print(f"工作表: {sheet.title}")
                print(f"总行数: {total_rows}, 总列数: {total_cols}")
                
                # 3. 检查是否为空或仅有表头
                if total_rows <= 1:
                    print("[OK] 文件仅包含表头或为空，无需拆分")
                    return True
                
                # 4. 读取表头（第一行）
                header_row = []
                header_styles = []
                for col in range(1, total_cols + 1):
                    cell = sheet.cell(row=1, column=col)
                    header_row.append(cell.value)
                    # 保存样式
                    if cell.has_style:
                        header_styles.append({
                            'font': copy(cell.font),
                            'alignment': copy(cell.alignment),
                            'fill': copy(cell.fill),
                            'border': copy(cell.border),
                            'number_format': cell.number_format
                        })
                    else:
                        header_styles.append(None)
                
                # 5. 初始估算每批行数
                # 估算公式：(总行数 - 表头) * (目标大小 / 文件大小)
                data_rows = total_rows - 1
                estimated_rows_per_file = max(1, int((data_rows * max_size_bytes) / file_size_bytes))
                
                print(f"初始估算每批行数: {estimated_rows_per_file}")
                
                # 6. 开始拆分
                part_num = 1
                current_row = 2  # 从第二行开始（第一行是表头）
                rows_per_batch = estimated_rows_per_file
                
                output_dir = file_path.parent
                base_name = file_path.stem
                
                while current_row <= total_rows:
                    # 计算本批次的结束行
                    end_row = min(current_row + rows_per_batch - 1, total_rows)
                    
                    if end_row < current_row:
                        break
                    
                    # 生成输出文件名
                    if part_num == 1:
                        output_file = output_dir / f"{base_name}_part{part_num:02d}{file_path.suffix}"
                    else:
                        output_file = output_dir / f"{base_name}_part{part_num:02d}{file_path.suffix}"
                    
                    print(f"\n创建拆分文件 {part_num}: {output_file.name}")
                    print(f"  行范围: {current_row} - {end_row} (共 {end_row - current_row + 1} 行数据)")
                    
                    # 使用try-except包装整个拆分文件创建过程
                    try:
                        # 创建新工作簿
                        new_wb = Workbook()
                        new_sheet = new_wb.active
                        new_sheet.title = sheet.title
                        
                        # 写入表头
                        for col, (value, style) in enumerate(zip(header_row, header_styles), start=1):
                            cell = new_sheet.cell(row=1, column=col)
                            cell.value = value
                            if style:
                                cell.font = style['font']
                                cell.alignment = style['alignment']
                                cell.fill = style['fill']
                                cell.border = style['border']
                                cell.number_format = style['number_format']
                        
                        # 复制数据行（添加进度提示）
                        print(f"  正在复制数据...")
                        new_row_idx = 2
                        total_data_rows = end_row - current_row + 1
                        progress_interval = max(1, total_data_rows // 10)  # 每10%显示一次进度
                        
                        for idx, row_idx in enumerate(range(current_row, end_row + 1)):
                            # 显示进度
                            if idx % progress_interval == 0:
                                progress_pct = int((idx / total_data_rows) * 100)
                                print(f"    进度: {progress_pct}% ({idx}/{total_data_rows} 行)")
                            
                            for col in range(1, total_cols + 1):
                                source_cell = sheet.cell(row=row_idx, column=col)
                                target_cell = new_sheet.cell(row=new_row_idx, column=col)
                                
                                # 复制值
                                target_cell.value = source_cell.value
                                
                                # 复制样式（仅在有样式时复制，提高性能）
                                if source_cell.has_style:
                                    target_cell.font = copy(source_cell.font)
                                    target_cell.alignment = copy(source_cell.alignment)
                                    target_cell.fill = copy(source_cell.fill)
                                    target_cell.border = copy(source_cell.border)
                                    target_cell.number_format = source_cell.number_format
                            
                            new_row_idx += 1
                        
                        print(f"  数据复制完成！")
                        
                        # 复制列宽
                        for col in range(1, total_cols + 1):
                            col_letter = get_column_letter(col)
                            if col_letter in sheet.column_dimensions:
                                new_sheet.column_dimensions[col_letter].width = sheet.column_dimensions[col_letter].width
                        
                        # 保存文件
                        print(f"  正在保存文件...")
                        self.save_workbook(new_wb, output_file)
                        new_wb.close()
                        print(f"  文件保存完成！")
                        
                        # 检查生成文件的大小
                        output_size_bytes = output_file.stat().st_size
                        output_size_mb = output_size_bytes / (1024 * 1024)
                        print(f"  文件大小: {output_size_mb:.2f} MB")
                        
                        # 如果文件超过限制
                        if output_size_bytes > max_size_bytes:
                            print(f"  [WARN] 超过限制，需要减少行数")
                            
                            # 删除刚生成的文件
                            try:
                                output_file.unlink()
                            except:
                                pass
                            
                            # 减少20%行数重试
                            rows_per_batch = max(1, int(rows_per_batch * 0.8))
                            print(f"  调整为每批 {rows_per_batch} 行，重试...")
                            continue
                        
                        # 如果文件远小于限制（< 15MB），下次可以增加行数
                        if output_size_mb < (max_size_mb * 0.75):
                            print(f"  提示: 文件较小，下批将增加行数")
                            rows_per_batch = int(rows_per_batch * 1.1)
                        
                        # 成功，移动到下一批
                        current_row = end_row + 1
                        part_num += 1
                        
                    except Exception as e:
                        print(f"  [FAIL] 创建拆分文件时发生错误: {e}")
                        import traceback
                        traceback.print_exc()
                        # 尝试清理可能创建的文件
                        try:
                            if output_file.exists():
                                output_file.unlink()
                        except:
                            pass
                        # 继续处理下一批，减少行数
                        rows_per_batch = max(1, int(rows_per_batch * 0.8))
                        print(f"  调整为每批 {rows_per_batch} 行，重试...")
                        continue
                
                # 检查是否真的拆分了
                if part_num == 1:
                    print("\n[OK] 无需拆分或拆分失败")
                    return True
                
                print(f"\n[OK] 拆分完成！共生成 {part_num - 1} 个文件")
                print(f"[OK] 原文件保留: {file_path.name}")
                return True
                
            finally:
                wb.close()
                
        except Exception as e:
            print(f"[FAIL] EXCEL拆分失败: {e}")
            import traceback
            traceback.print_exc()
            return False

