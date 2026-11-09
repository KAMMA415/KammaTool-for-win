
import os
import shutil
import threading
from datetime import datetime
from pathlib import Path
import win32com.client

class FileProcessor:
    """文件操作处理器"""
    
    def __init__(self, msg_queue, config_manager):
        self.msg_queue = msg_queue
        self.config = config_manager
        
    def log(self, message):
        """发送日志消息"""
        self.msg_queue.put(('log', [message]))
        
    def show_message(self, type_, title, message):
        """发送弹窗消息"""
        self.msg_queue.put(('show_message', [type_, title, message]))

    def _put_append(self, widget, text):
        """向日志控件追加文本（widget 为 None 时降级为 log）"""
        if widget is None:
            self.log(text.rstrip())
        else:
            self.msg_queue.put(('append_text', [widget, text]))

    def _put_progress(self, widget, val, max_val=None):
        """更新进度条"""
        self.msg_queue.put(('update_progress', [widget, val, max_val]))

    def _put_stop_progress(self, widget):
        """停止进度条"""
        self.msg_queue.put(('stop_progress', [widget]))

    def batch_rename(self, folder, prefix, replace_dict=None):
        """批量重命名
        - 自动跳过 Office 临时文件（~$ 开头）
        - 递归处理目标目录下的所有文件和所有子目录名称
        - 自动跳过已含前缀的文件/目录
        - 自动跳过新名与旧名相同的文件/目录
        - 自动处理目标名称重名冲突（加 _1、_2… 后缀）
        """
        try:
            if not os.path.isdir(folder):
                self.show_message('error', "错误", f"目录不存在：{folder}")
                return False

            entries = []
            base_folder = os.path.abspath(folder)
            for root, dirs, files in os.walk(folder, topdown=True):
                for file_name in files:
                    path = os.path.join(root, file_name)
                    depth = len(os.path.relpath(path, base_folder).split(os.sep))
                    entries.append((path, file_name, False, depth))
                for dir_name in dirs:
                    path = os.path.join(root, dir_name)
                    depth = len(os.path.relpath(path, base_folder).split(os.sep))
                    entries.append((path, dir_name, True, depth))

            entries.sort(key=lambda item: (-item[3], item[2], item[1].lower()))
            count = 0
            file_count = 0
            dir_count = 0
            skip_count = 0
            conflict_count = 0

            for old_path, name, is_dir, _depth in entries:
                item_type = "目录" if is_dir else "文件"

                # 跳过 Office 临时文件（~$ 开头）
                if name.startswith('~$'):
                    self.log(f"跳过临时项目: {name}")
                    skip_count += 1
                    continue

                if not replace_dict and prefix and name.startswith(prefix):
                    self.log(f"跳过已含前缀的{item_type}: {name}")
                    skip_count += 1
                    continue

                # 计算新名称
                if replace_dict:
                    new_name = name
                    for old_str, new_str in replace_dict.items():
                        new_name = new_name.replace(old_str, new_str)
                    if prefix:
                        new_name = prefix + new_name
                else:
                    new_name = prefix + name

                # 跳过无变化的文件/目录
                if new_name == name:
                    skip_count += 1
                    continue

                # 处理目标名称冲突
                parent_folder = os.path.dirname(old_path)
                new_path = os.path.join(parent_folder, new_name)
                if os.path.exists(new_path):
                    if is_dir:
                        base, ext = new_name, ""
                    else:
                        base, ext = os.path.splitext(new_name)
                    suffix = 1
                    while os.path.exists(os.path.join(parent_folder, f"{base}_{suffix}{ext}")):
                        suffix += 1
                    new_name = f"{base}_{suffix}{ext}"
                    new_path = os.path.join(parent_folder, new_name)
                    self.log(f"重名冲突，自动调整为: {new_name}")
                    conflict_count += 1

                try:
                    os.rename(old_path, new_path)
                    self.log(f"✓ {item_type} {name} → {new_name}")
                    count += 1
                    if is_dir:
                        dir_count += 1
                    else:
                        file_count += 1
                except Exception as e:
                    self.log(f"✗ 重命名失败 {item_type} {name}: {e}")
                    skip_count += 1

            msg = f"批量重命名完成：成功 {count} 个"
            if count:
                msg += f"（文件 {file_count} 个，目录 {dir_count} 个）"
            if skip_count:
                msg += f"，跳过 {skip_count} 个"
            if conflict_count:
                msg += f"（{conflict_count} 个自动处理重名）"
            self.show_message('info', "完成", msg)
            self.log(msg)
            return True
        except Exception as e:
            self.log(f"重命名失败: {str(e)}")
            self.show_message('error', "错误", f"重命名失败: {str(e)}")
            return False

    @staticmethod
    def preview_filename_deletion(filename, delete_text):
        """返回删除指定内容后的文件名；None 表示不匹配，空串表示结果不安全。"""
        if not delete_text or filename.startswith('~$'):
            return None

        stem, extension = os.path.splitext(filename)
        if delete_text not in stem:
            return None

        new_stem = stem.replace(delete_text, '')
        if (not new_stem.strip() or new_stem in ('.', '..')
                or new_stem.endswith((' ', '.'))):
            return ""

        reserved_names = {'CON', 'PRN', 'AUX', 'NUL'}
        reserved_names.update(f'COM{index}' for index in range(1, 10))
        reserved_names.update(f'LPT{index}' for index in range(1, 10))
        if new_stem.split('.', 1)[0].upper() in reserved_names:
            return ""

        return new_stem + extension

    def delete_filename_content(self, folder, delete_text):
        """递归删除文件名主名称中的指定内容，不修改目录名和文件扩展名。"""
        try:
            if not os.path.isdir(folder):
                self.show_message('error', "错误", f"目录不存在：{folder}")
                return False
            if delete_text == "":
                self.show_message('error', "错误", "删除内容不能为空")
                return False

            files = []
            for root, _dirs, filenames in os.walk(folder, topdown=True):
                for filename in filenames:
                    files.append(os.path.join(root, filename))
            files.sort(key=lambda path: path.lower())

            count = 0
            invalid_count = 0
            failed_count = 0
            conflict_count = 0

            for old_path in files:
                name = os.path.basename(old_path)
                new_name = self.preview_filename_deletion(name, delete_text)
                if new_name is None:
                    continue
                if new_name == "":
                    self.log(f"跳过（删除后文件名无效）: {name}")
                    invalid_count += 1
                    continue

                parent_folder = os.path.dirname(old_path)
                new_path = os.path.join(parent_folder, new_name)
                if os.path.exists(new_path):
                    base, extension = os.path.splitext(new_name)
                    suffix = 1
                    while os.path.exists(os.path.join(parent_folder, f"{base}_{suffix}{extension}")):
                        suffix += 1
                    new_name = f"{base}_{suffix}{extension}"
                    new_path = os.path.join(parent_folder, new_name)
                    self.log(f"重名冲突，自动调整为: {new_name}")
                    conflict_count += 1

                try:
                    os.rename(old_path, new_path)
                    self.log(f"✓ 文件 {name} → {new_name}")
                    count += 1
                except Exception as e:
                    self.log(f"✗ 删除文件名内容失败 {name}: {e}")
                    failed_count += 1

            msg = f"删除文件名指定内容完成：成功 {count} 个文件"
            if invalid_count:
                msg += f"，保护跳过 {invalid_count} 个"
            if failed_count:
                msg += f"，失败 {failed_count} 个"
            if conflict_count:
                msg += f"（{conflict_count} 个自动处理重名）"
            self.show_message('info', "完成", msg)
            self.log(msg)
            return True
        except Exception as e:
            self.log(f"删除文件名指定内容失败: {str(e)}")
            self.show_message('error', "错误", f"删除文件名指定内容失败: {str(e)}")
            return False

    def convert_to_pdf(self, target_dir, include_subdirs, overwrite, result_widget, progress_widget):
        """文件转PDF"""
        try:
            supported_extensions = {
                '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.txt'
            }
            
            files_to_convert = []
            if include_subdirs:
                for root, dirs, files in os.walk(target_dir):
                    for file in files:
                        if os.path.splitext(file.lower())[1] in supported_extensions:
                            files_to_convert.append(os.path.join(root, file))
            else:
                for file in os.listdir(target_dir):
                    file_path = os.path.join(target_dir, file)
                    if os.path.isfile(file_path) and os.path.splitext(file.lower())[1] in supported_extensions:
                        files_to_convert.append(file_path)
            
            if not files_to_convert:
                self._put_append(result_widget, "未找到支持的文件类型。\n")
                return
            
            n = len(files_to_convert)
            self._put_append(result_widget, f"找到 {n} 个文件需要转换:\n\n")
            self._put_progress(progress_widget, 0, n)
            
            success_count = 0
            failed_count = 0
            
            for i, file_path in enumerate(files_to_convert):
                try:
                    file_name = os.path.basename(file_path)
                    self._put_append(result_widget, f"[{i+1}/{n}] 转换: {file_name}...")
                    
                    pdf_path = os.path.splitext(file_path)[0] + '.pdf'
                    
                    if os.path.exists(pdf_path) and not overwrite:
                        self._put_append(result_widget, " 跳过（文件已存在）\n")
                        continue
                    
                    success = self._convert_single_file(file_path, pdf_path)
                    
                    if success:
                        self._put_append(result_widget, " 成功\n")
                        success_count += 1
                    else:
                        self._put_append(result_widget, " 失败\n")
                        failed_count += 1
                        self.log(f"转换失败的文件: {file_path}")
                        
                except Exception as e:
                    self._put_append(result_widget, f" 错误: {str(e)}\n")
                    failed_count += 1
                
                self._put_progress(progress_widget, i + 1, None)
                
            self._put_append(result_widget, f"\n转换完成！\n成功: {success_count}\n失败: {failed_count}\n")
            self.log(f"PDF转换完成，成功{success_count}个，失败{failed_count}个")
            
        except Exception as e:
            self._put_append(result_widget, f"转换过程出错: {str(e)}\n")
            self.log(f"PDF转换失败: {str(e)}")

    def _convert_single_file(self, source_path, pdf_path):
        """单文件转换核心逻辑"""
        ext = os.path.splitext(source_path)[1].lower()
        try:
            if ext in ['.doc', '.docx']:
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                try:
                    doc = word.Documents.Open(source_path)
                    doc.SaveAs(pdf_path, FileFormat=17) # wdFormatPDF = 17
                    doc.Close()
                    return True
                finally:
                    word.Quit()
            elif ext in ['.xls', '.xlsx']:
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                try:
                    wb = excel.Workbooks.Open(source_path)
                    # 导出为PDF (0=xlTypePDF)
                    wb.ExportAsFixedFormat(0, pdf_path)
                    wb.Close(False)
                    return True
                finally:
                    excel.Quit()
            elif ext in ['.ppt', '.pptx']:
                ppt = win32com.client.DispatchEx("PowerPoint.Application")
                try:
                    pres = ppt.Presentations.Open(source_path, WithWindow=False)
                    pres.SaveAs(pdf_path, 32) # ppSaveAsPDF = 32
                    pres.Close()
                    return True
                finally:
                    ppt.Quit()
            elif ext == '.txt':
                # 简单TXT转PDF逻辑，使用reportlab
                from reportlab.pdfgen import canvas
                from reportlab.lib.pagesizes import A4
                from reportlab.pdfbase import pdfmetrics
                from reportlab.pdfbase.ttfonts import TTFont
                
                c = canvas.Canvas(pdf_path, pagesize=A4)
                try:
                    pdfmetrics.registerFont(TTFont('SimHei', 'simhei.ttf'))
                    c.setFont('SimHei', 10)
                except:
                    c.setFont('Helvetica', 10)
                    
                y = 800
                with open(source_path, 'r', encoding='utf-8', errors='ignore') as f:
                    for line in f:
                        c.drawString(50, y, line.strip())
                        y -= 15
                        if y < 50:
                            c.showPage()
                            y = 800
                            try:
                                c.setFont('SimHei', 10)
                            except:
                                c.setFont('Helvetica', 10)
                c.save()
                return True
        except Exception as e:
            print(f"转换错误 {source_path}: {e}")
            return False
        return False

    def merge_pdfs(self, target_dir, include_subdirs, result_widget, progress_widget):
        """合并PDF"""
        try:
            from PyPDF2 import PdfReader, PdfWriter
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import A4
            import io
            
            pdf_files = []
            if include_subdirs:
                for root, dirs, files in os.walk(target_dir):
                    for file in files:
                        if file.lower().endswith('.pdf'):
                            pdf_files.append(os.path.join(root, file))
            else:
                for file in os.listdir(target_dir):
                    if file.lower().endswith('.pdf'):
                        pdf_files.append(os.path.join(target_dir, file))
            
            dir_name = os.path.basename(target_dir)
            output_filename = f"{dir_name}-产品知识库.pdf"
            output_path = os.path.join(target_dir, output_filename)
            
            # 排除自身
            pdf_files = [f for f in pdf_files if os.path.basename(f) != output_filename]
            pdf_files.sort(key=lambda x: os.path.basename(x).lower())
            
            if not pdf_files:
                self._put_append(result_widget, "未找到PDF文件。\n")
                return

            n = len(pdf_files)
            self._put_append(result_widget, f"找到 {n} 个PDF文件需要合并。\n")
            self._put_progress(progress_widget, 0, n)
            
            writer = PdfWriter()
            
            for i, pdf_file in enumerate(pdf_files):
                try:
                    filename = os.path.basename(pdf_file)
                    self._put_append(result_widget, f"合并: {filename}\n")
                    
                    # 生成分隔页
                    packet = io.BytesIO()
                    can = canvas.Canvas(packet, pagesize=A4)
                    can.setFont("Helvetica", 20)
                    can.drawString(100, 400, f"File: {filename}")
                    can.save()
                    packet.seek(0)
                    new_pdf = PdfReader(packet)
                    writer.add_page(new_pdf.pages[0])
                    
                    reader = PdfReader(pdf_file)
                    for page in reader.pages:
                        writer.add_page(page)
                        
                except Exception as e:
                    self._put_append(result_widget, f"错误 {filename}: {str(e)}\n")
                
                self._put_progress(progress_widget, i + 1, None)
                
            with open(output_path, 'wb') as f:
                writer.write(f)
                
            self._put_append(result_widget, f"\n成功合并到: {output_path}\n")
            self.log(f"合并PDF完成: {output_path}")
            
        except Exception as e:
            self._put_append(result_widget, f"合并失败: {str(e)}\n")
            self.log(f"合并PDF失败: {str(e)}")

    def organize_files(self, source_dir, days=None, time_types=None, file_types_cfg=None, 
                       log_widget=None, progress_widget=None, 
                       start_date=None, end_date=None):
        """文件整理
        
        Args:
            source_dir: 源目录
            days: 天数模式，过去X天（可选，与start_date/end_date互斥）
            start_date: 开始日期（datetime对象，可选）
            end_date: 结束日期（datetime对象，可选）
            time_types: 时间类型字典 {'create_time', 'modify_time', 'access_time'}
            file_types_cfg: 文件类型配置
            log_widget: 日志控件
            progress_widget: 进度条控件
        """
        try:
            # 定义文件类型扩展名映射
            FILE_TYPE_EXTENSIONS = {
                'office': ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx'],
                'wps': ['.wps', '.et', '.dps'],  # WPS文字、表格、演示
                'pdf': ['.pdf'],
                'image': ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.webp', '.ico']
            }

            time_types = time_types or {'modify_time': True}
            valid_time_keys = ('create_time', 'modify_time', 'access_time')
            if not any(time_types.get(key) for key in valid_time_keys):
                time_types = {'modify_time': True}
            
            # 定义目录黑名单
            SKIP_DIRECTORIES = ['Hope', '图片', '归档附件', '文档', '知识库']
            
            # 兼容中文 Windows（桌面路径可能是 Desktop 或 桌面）
            home = os.path.expanduser("~")
            desktop = os.path.join(home, "Desktop")
            if not os.path.isdir(desktop):
                desktop = os.path.join(home, "桌面")
            if not os.path.isdir(desktop):
                desktop = home  # 最终 fallback 到用户目录
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            target_base = os.path.join(desktop, f"文件整理_{timestamp}")
            os.makedirs(target_base, exist_ok=True)
            
            self._put_append(log_widget, f"扫描目录: {source_dir}\n")
            self._put_append(log_widget, f"目标位置: {target_base}\n")
            
            # 显示文件类型筛选信息
            active_types = [name for name, enabled in (file_types_cfg or {}).items() if enabled]
            if active_types:
                self._put_append(log_widget, f"文件类型: {', '.join(active_types)}\n")

            time_type_labels = {
                'create_time': '创建时间',
                'modify_time': '修改时间',
                'access_time': '访问时间'
            }
            active_time_types = [time_type_labels[key] for key in valid_time_keys if time_types.get(key)]
            self._put_append(log_widget, f"判断依据: {', '.join(active_time_types)}\n")
            
            # 显示目录过滤信息
            self._put_append(log_widget, f"跳过目录: {', '.join(SKIP_DIRECTORIES)}\n")
            
            # 确定时间筛选模式
            if days is not None:
                now = datetime.now()
                self._put_append(log_widget, f"时间筛选: 过去 {days} 天\n")
                use_date_range = False
            elif start_date is not None and end_date is not None:
                self._put_append(log_widget, f"时间筛选: {start_date.strftime('%Y-%m-%d')} 至 {end_date.strftime('%Y-%m-%d')}\n")
                use_date_range = True
            else:
                raise ValueError("必须指定 days 或 (start_date, end_date)")
            
            self._put_append(log_widget, "\n开始扫描文件...\n\n")
            
            count = 0
            failed_count = 0
            skipped_count = 0
            skipped_dirs_count = 0
            
            for root, dirs, files in os.walk(source_dir):
                # 检查并过滤黑名单目录
                current_dir_name = os.path.basename(root)
                if current_dir_name in SKIP_DIRECTORIES:
                    skipped_dirs_count += 1
                    self._put_append(log_widget, f"⊘ 跳过黑名单目录: {os.path.relpath(root, source_dir)}\n")
                    dirs.clear()
                    continue
                
                original_dirs = list(dirs)
                dirs[:] = [d for d in dirs if d not in SKIP_DIRECTORIES]
                if len(dirs) < len(original_dirs):
                    skipped_dirs_count += len(original_dirs) - len(dirs)
                
                for file in files:
                    file_path = os.path.join(root, file)
                    
                    try:
                        stat = os.stat(file_path)
                        times = []
                        if time_types.get('create_time'): 
                            times.append(datetime.fromtimestamp(stat.st_ctime))
                        if time_types.get('modify_time'): 
                            times.append(datetime.fromtimestamp(stat.st_mtime))
                        if time_types.get('access_time'): 
                            times.append(datetime.fromtimestamp(stat.st_atime))
                         
                        match_time = False
                        for t in times:
                            if use_date_range:
                                if start_date <= t <= end_date:
                                    match_time = True
                                    break
                            else:
                                if (now - t).days <= days:
                                    match_time = True
                                    break
                        
                        if not match_time:
                            skipped_count += 1
                            continue
                            
                        ext = os.path.splitext(file)[1].lower()
                        match_type = False
                        
                        for type_name, enabled in (file_types_cfg or {}).items():
                            if enabled and type_name in FILE_TYPE_EXTENSIONS:
                                if ext in FILE_TYPE_EXTENSIONS[type_name]:
                                    match_type = True
                                    break
                        
                        if not file_types_cfg or not any(file_types_cfg.values()):
                            match_type = True
                        
                        if not match_type:
                            skipped_count += 1
                            continue
                        
                        if stat.st_size == 0:
                            skipped_count += 1
                            continue
                        
                        rel_path = os.path.relpath(file_path, source_dir)
                        target_path = os.path.join(target_base, rel_path)
                        
                        target_dir_path = os.path.dirname(target_path)
                        if not os.path.exists(target_dir_path):
                            os.makedirs(target_dir_path, exist_ok=True)
                        
                        try:
                            shutil.copy2(file_path, target_path)
                            self._put_append(log_widget, f"✓ 复制: {rel_path}\n")
                            count += 1
                        except PermissionError:
                            self._put_append(log_widget, f"⚠️ 跳过（文件被占用）: {rel_path}\n")
                            failed_count += 1
                        except Exception as e:
                            self._put_append(log_widget, f"⚠️ 跳过（{str(e)}）: {rel_path}\n")
                            failed_count += 1
                            
                    except Exception as e:
                        rel_path = os.path.relpath(file_path, source_dir)
                        self._put_append(log_widget, f"⚠️ 无法读取文件信息: {rel_path}\n")
                        failed_count += 1
                        continue
            
            summary = f"\n========================================\n"
            summary += f"整理完成！\n"
            summary += f"========================================\n"
            summary += f"✓ 成功复制: {count} 个文件\n"
            if failed_count > 0:
                summary += f"⚠️ 失败/跳过: {failed_count} 个文件\n"
            if skipped_count > 0:
                summary += f"⊘ 不匹配条件: {skipped_count} 个文件\n"
            if skipped_dirs_count > 0:
                summary += f"⊘ 跳过黑名单目录: {skipped_dirs_count} 个\n"
            summary += f"📁 输出位置: {target_base}\n"
            summary += f"========================================\n"
            
            self._put_append(log_widget, summary)
            self._put_stop_progress(progress_widget)
            
            if count > 0:
                self.show_message('info', "完成", 
                                f"成功整理 {count} 个文件到桌面\n输出文件夹: 文件整理_{timestamp}")
            else:
                self.show_message('info', "提示", "未找到符合条件的文件")
            
        except Exception as e:
            self._put_append(log_widget, f"整理失败: {str(e)}\n")
            self._put_stop_progress(progress_widget)
            self.show_message('error', "错误", f"整理失败: {str(e)}")

    def export_filenames(self, directory, output_file):
        """导出文件名到Excel"""
        try:
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws.append(['目录', '文件名'])
            
            count = 0
            for root, dirs, files in os.walk(directory):
                for file in files:
                    rel_dir = os.path.relpath(root, directory)
                    ws.append([rel_dir, file])
                    count += 1
            
            wb.save(output_file)
            self.log(f"导出文件名完成: {output_file}")
            self.show_message('info', "完成", f"已导出 {count} 个文件名")
            return True
        except Exception as e:
            self.log(f"导出失败: {e}")
            self.show_message('error', "错误", f"导出失败: {e}")
            return False

    def merge_txt_word(self, directory):
        """合并TXT和Word"""
        pass
