
import os
import json
import queue
import platform
import ctypes
from datetime import datetime

class ConfigManager:
    """配置管理器 - 增强版，支持配置备份和恢复"""
    def __init__(self, config_file="config.json"):
        self.config_file = config_file
        self.backup_file = config_file + ".bak"
        self.config = self.load_config()

    def load_config(self):
        """加载配置文件，支持备份恢复"""
        # 尝试加载主配置文件
        config = self._load_from_file(self.config_file)
        if config is not None:
            return config
        
        # 主配置失败，尝试加载备份
        print(f"主配置文件损坏，尝试从备份恢复...")
        config = self._load_from_file(self.backup_file)
        if config is not None:
            # 恢复成功，重新保存主配置
            try:
                with open(self.config_file, 'w', encoding='utf-8') as f:
                    json.dump(config, f, indent=4, ensure_ascii=False)
                print("配置已从备份恢复")
            except:
                pass
            return config
        
        # 备份也失败，返回默认配置
        print("使用默认配置")
        return self._get_default_config()
    
    def _load_from_file(self, filepath):
        """从指定文件加载配置"""
        try:
            if os.path.exists(filepath):
                with open(filepath, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except json.JSONDecodeError as e:
            print(f"配置文件格式错误 ({filepath}): {e}")
        except Exception as e:
            print(f"加载配置失败 ({filepath}): {e}")
        return None
    
    def _get_default_config(self):
        """返回默认配置"""
        return {
            'last_organizer_days': 7,  # 默认7天
            'org_create_time': True,
            'org_modify_time': True,
            'org_access_time': False,
        }

    def save_config(self):
        """保存配置文件，自动创建备份"""
        try:
            # 先保存到临时文件
            temp_file = self.config_file + ".tmp"
            with open(temp_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            
            # 创建备份（如果主配置文件存在）
            if os.path.exists(self.config_file):
                try:
                    import shutil
                    shutil.copy2(self.config_file, self.backup_file)
                except:
                    pass  # 备份失败不影响保存
            
            # 用临时文件替换主配置文件
            if os.path.exists(self.config_file):
                os.remove(self.config_file)
            os.rename(temp_file, self.config_file)
            
        except Exception as e:
            print(f"保存配置失败: {e}")
            # 清理临时文件
            temp_file = self.config_file + ".tmp"
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except:
                    pass

    def get(self, key, default=None):
        """获取配置项"""
        return self.config.get(key, default)

    def set(self, key, value):
        """设置配置项"""
        self.config[key] = value


class MessageQueue:
    """线程安全消息队列封装"""
    def __init__(self):
        self.queue = queue.Queue()

    def put(self, task, *args):
        """
        发送消息到队列
        task: 任务类型 ('log', 'clear_log', 'append_text', 'update_progress', 'stop_progress', 'show_message')
        args: 任务参数
        """
        self.queue.put((task, args))

    def get_nowait(self):
        """非阻塞获取消息"""
        return self.queue.get_nowait()
    
    def empty(self):
        """判断队列是否为空"""
        return self.queue.empty()

def enable_high_dpi():
    """启用高DPI支持"""
    try:
        if platform.system() == "Windows":
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
