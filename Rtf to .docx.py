'''
用Python 3.8.7实现批量修改一个文件夹里面的“.Rtf文件”，并且生成一个GUI界面进行操作。
1 把“.Rtf文件”修改为“.docx文件”
'''
import os
import sys
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import traceback
import win32com.client
import pythoncom
import psutil  # 用于强制清理Word进程

class RtfToDocxConverterWin:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows专用 - RTF批量转DOCX工具 (Python 3.8.7)")
        self.root.geometry("850x650")
        self.root.resizable(False, False)
        
        # 初始化变量
        self.folder_path = tk.StringVar()
        self.word_instances = []  # 跟踪Word实例，防止泄漏
        # 定义Word常量（直接用数值，避免常量引用错误）
        self.WD_ALERTS_NONE = 0
        self.WD_FORMAT_XML_DOCUMENT = 16
        self.WD_WORD_2016 = 15
        self.MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
        self.WD_DO_NOT_SAVE_CHANGES = 0
        
        self._create_widgets()
        
    def _create_widgets(self):
        """创建Windows风格GUI，优化交互体验"""
        self.root.option_add("*Font", "微软雅黑 9")
        
        # 1. 标题区域
        title_frame = tk.Frame(self.root, bg="#0078D7", padx=10, pady=8)
        title_frame.pack(fill=tk.X)
        tk.Label(
            title_frame, text="RTF → DOCX 批量转换工具", 
            font=("微软雅黑", 12, "bold"), bg="#0078D7", fg="white"
        ).pack(anchor=tk.W)
        
        # 2. 文件夹选择区域
        folder_frame = tk.Frame(self.root, padx=15, pady=10)
        folder_frame.pack(fill=tk.X)
        
        tk.Label(
            folder_frame, text="目标文件夹：", 
            font=("微软雅黑", 10, "bold")
        ).pack(side=tk.LEFT)
        
        # 只读输入框，显示选中的文件夹
        folder_entry = tk.Entry(
            folder_frame, textvariable=self.folder_path, width=65,
            font=("微软雅黑", 10), state="readonly", bd=1, relief=tk.SUNKEN
        )
        folder_entry.pack(side=tk.LEFT, padx=8)
        
        # 选择文件夹按钮
        tk.Button(
            folder_frame, text="选择文件夹",
            command=self.select_folder,
            font=("微软雅黑", 10), bg="#4CAF50", fg="white",
            relief=tk.FLAT, padx=12, pady=2
        ).pack(side=tk.LEFT)
        
        # 3. 操作按钮区域
        btn_frame = tk.Frame(self.root, padx=15, pady=5)
        btn_frame.pack(fill=tk.X)
        
        # 转换按钮（核心操作）
        self.convert_btn = tk.Button(
            btn_frame, text="开始批量转换",
            command=self.batch_convert,
            font=("微软雅黑", 11, "bold"), bg="#2196F3", fg="white",
            relief=tk.FLAT, padx=25, pady=5
        )
        self.convert_btn.pack(side=tk.LEFT, padx=5)
        
        # 清空日志按钮
        tk.Button(
            btn_frame, text="清空日志",
            command=self.clear_log,
            font=("微软雅黑", 10), bg="#f44336", fg="white",
            relief=tk.FLAT, padx=12, pady=2
        ).pack(side=tk.LEFT, padx=5)
        
        # 清理进程按钮（应急用）
        tk.Button(
            btn_frame, text="清理残留Word进程",
            command=self.clean_word_processes,
            font=("微软雅黑", 10), bg="#FF9800", fg="white",
            relief=tk.FLAT, padx=12, pady=2
        ).pack(side=tk.LEFT, padx=5)
        
        # 4. 日志区域
        log_frame = tk.Frame(self.root, padx=15, pady=10)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(
            log_frame, text="转换日志（可滚动查看）：",
            font=("微软雅黑", 10, "bold")
        ).pack(anchor=tk.W)
        
        # 带滚动条的日志文本框（只读）
        self.log_text = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, height=30, font=("Consolas", 9),
            bg="#F8F9FA", bd=1, relief=tk.SUNKEN, state=tk.DISABLED
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 初始日志提示
        self.log("📢 工具已就绪！请选择包含RTF文件的文件夹开始转换")
        self.log("💡 提示：转换后的DOCX文件与原RTF文件同目录，确保Word已安装且可正常运行")
        
    def select_folder(self):
        """选择目标文件夹，验证有效性"""
        folder = filedialog.askdirectory(title="选择包含RTF文件的文件夹")
        if folder:
            # 验证文件夹是否存在且可访问
            if os.path.exists(folder) and os.access(folder, os.W_OK):
                self.folder_path.set(folder)
                self.log(f"✅ 已选择有效文件夹：{folder}")
            else:
                messagebox.showerror("错误", "所选文件夹不可写，请选择其他文件夹！")
                self.log(f"❌ 文件夹不可写：{folder}")
                
    def log(self, message):
        """线程安全的日志输出，保证日志区域只读"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)  # 自动滚动到最新日志
        self.root.update_idletasks()  # 强制刷新界面
        self.log_text.config(state=tk.DISABLED)
        
    def clear_log(self):
        """清空日志内容"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.log("📝 日志已清空，工具就绪")
        
    def clean_word_processes(self):
        """强制清理残留的WinWord.exe进程（防止内存泄漏）"""
        try:
            self.log("🔍 开始清理残留Word进程...")
            killed = 0
            for proc in psutil.process_iter(['pid', 'name']):
                if proc.info['name'] and 'WINWORD.EXE' in proc.info['name'].upper():
                    proc.kill()
                    killed += 1
                    self.log(f"🗑️  终止Word进程 PID: {proc.info['pid']}")
            self.log(f"✅ 共清理 {killed} 个Word残留进程")
            messagebox.showinfo("完成", f"已清理 {killed} 个Word残留进程")
        except Exception as e:
            self.log(f"❌ 清理进程失败：{str(e)}")
            messagebox.showerror("错误", f"清理进程失败：{str(e)}")
        
    def convert_single_file(self, rtf_path, docx_path):
        """
        核心转换函数：修复常量引用问题，保证DOCX可正常打开
        1. 移除易出错的constants引用，直接用数值
        2. 跳过临时文件（~$开头的文件）
        3. 增强异常处理
        """
        # 跳过Word临时文件（~$开头），这类文件无法正常转换
        if os.path.basename(rtf_path).startswith("~$"):
            self.log(f"  ⚠️  跳过Word临时文件：{os.path.basename(rtf_path)}")
            return True
        
        word = None
        doc = None
        try:
            # 初始化COM（解决多次调用问题）
            pythoncom.CoInitialize()
            
            # 创建独立的Word实例（DispatchEx），避免影响现有Word窗口
            word = win32com.client.DispatchEx("Word.Application")
            self.word_instances.append(word)  # 跟踪实例
            
            # 关键设置：禁用所有弹窗和可见性（直接用数值，避免常量错误）
            word.Visible = False
            word.DisplayAlerts = self.WD_ALERTS_NONE  # 0 = 禁用所有提示
            word.AutomationSecurity = self.MSO_AUTOMATION_SECURITY_FORCE_DISABLE  # 3 = 强制禁用宏
            
            # 打开RTF文件（禁用转换确认、只读模式打开）
            doc = word.Documents.Open(
                FileName=rtf_path,
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False
            )
            
            # 另存为DOCX（使用数值指定格式，确保兼容性）
            doc.SaveAs2(
                FileName=docx_path,
                FileFormat=self.WD_FORMAT_XML_DOCUMENT,  # 16 = DOCX格式
                CompatibilityMode=self.WD_WORD_2016  # 15 = 兼容Word 2016+
            )
            
            # 验证生成的DOCX文件是否有效
            if os.path.exists(docx_path) and os.path.getsize(docx_path) > 0:
                self.log(f"  ✅ 转换成功：{os.path.basename(rtf_path)} → {os.path.basename(docx_path)}")
                return True
            else:
                self.log(f"  ❌ 转换后文件无效：{os.path.basename(rtf_path)}")
                return False
                
        except Exception as e:
            self.log(f"  ❌ 转换失败：{os.path.basename(rtf_path)}")
            self.log(f"  📋 错误原因：{str(e)}")
            self.log(f"  📜 错误详情：{traceback.format_exc()[:600]}")  # 截断过长日志
            return False
        finally:
            # 强制释放资源（关键：防止Word进程残留）
            if doc:
                try:
                    doc.Close(SaveChanges=self.WD_DO_NOT_SAVE_CHANGES)  # 0 = 不保存更改
                except:
                    pass
            if word:
                try:
                    word.Quit(SaveChanges=self.WD_DO_NOT_SAVE_CHANGES)
                    self.word_instances.remove(word)
                except:
                    pass
            # 释放COM资源
            pythoncom.CoUninitialize()
            
    def batch_convert(self):
        """批量转换主逻辑，防重复点击、完整统计"""
        # 禁用按钮防止重复触发
        self.convert_btn.config(state=tk.DISABLED)
        
        # 验证文件夹
        folder = self.folder_path.get()
        if not folder or not os.path.exists(folder):
            messagebox.showerror("错误", "请先选择有效的文件夹！")
            self.convert_btn.config(state=tk.NORMAL)
            return
        
        # 查找所有RTF文件（不区分大小写）
        rtf_files = []
        for f in os.listdir(folder):
            if f.lower().endswith(".rtf") and os.path.isfile(os.path.join(folder, f)):
                rtf_files.append(f)
        
        if not rtf_files:
            messagebox.showinfo("提示", "文件夹中未找到任何RTF文件！")
            self.log("ℹ️  未检测到RTF文件，转换终止")
            self.convert_btn.config(state=tk.NORMAL)
            return
        
        # 开始转换
        self.log(f"\n🚀 开始批量转换 - 共检测到 {len(rtf_files)} 个RTF文件")
        self.log("-" * 70)
        
        success_count = 0
        fail_count = 0
        
        for filename in rtf_files:
            rtf_path = os.path.join(folder, filename)
            docx_filename = os.path.splitext(filename)[0] + ".docx"
            docx_path = os.path.join(folder, docx_filename)
            
            # 跳过已存在的DOCX文件（可选：可删除此判断）
            if os.path.exists(docx_path):
                self.log(f"  ⚠️  跳过已存在文件：{docx_filename}")
                continue
            
            self.log(f"\n🔄 正在处理：{filename}")
            if self.convert_single_file(rtf_path, docx_path):
                success_count += 1
            else:
                fail_count += 1
        
        # 转换完成统计
        self.log("\n" + "="*70)
        self.log(f"🏁 批量转换完成！")
        self.log(f"✅ 成功转换：{success_count} 个文件")
        self.log(f"❌ 转换失败：{fail_count} 个文件")
        self.log(f"📁 输出路径：{folder}")
        
        # 弹窗提示结果
        messagebox.showinfo(
            "转换完成",
            f"批量转换结束！\n\n✅ 成功：{success_count} 个\n❌ 失败：{fail_count} 个\n\n📁 所有DOCX文件已保存至原文件夹"
        )
        
        # 恢复按钮状态
        self.convert_btn.config(state=tk.NORMAL)
        
        # 最后清理可能的Word进程
        self.clean_word_processes()

if __name__ == "__main__":
    # 检查Python版本
    if sys.version_info[:3] != (3, 8, 7):
        messagebox.showwarning("版本警告", f"当前Python版本：{sys.version[:5]}，建议使用3.8.7！")
    
    # 安装依赖提示
    print("="*70)
    print("【Windows专用RTF转DOCX工具 - 环境准备】")
    print("1. 安装依赖（Python 3.8.7）：")
    print("   pip install pywin32==227 psutil==5.8.0")
    print("2. 确保已安装Microsoft Word（2010及以上版本）")
    print("3. 运行前关闭所有Word窗口，避免冲突")
    print("="*70 + "\n")
    
    # 启动GUI
    root = tk.Tk()
    app = RtfToDocxConverterWin(root)
    
    # 程序退出时清理Word进程
    def on_closing():
        app.clean_word_processes()
        root.destroy()
        
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
