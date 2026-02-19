'''

重命名后的文件会复制到新文件夹中，原始文件保留在原位置不变，同时保留了备份、GUI 交互和日志记录功能


用Python 3.8.7实现把一个文件夹里面的.docx文件批量修改名字, 并且生成一个GUI界面进行操作：
“M1_”移动到“P1_”后面；
“M2_”移动到“P1_”后面；
“M3_”移动到“P1_”后面；
“M4_”移动到“P1_”后面；
“M5_”移动到“P1_”后面；
“Ambient_”移动到“P1_”后面；
'''

'''
修改代码：Python批量重命名docx文件（原始文件保留，新文件存入新文件夹）

用Python 3.8.7实现把一个文件夹里面的.docx文件批量修改名字, 并且生成一个GUI界面进行操作：
“M1_”移动到“P1_”后面；
“M2_”移动到“P1_”后面；
“M3_”移动到“P1_”后面；
“M4_”移动到“P1_”后面；
“M5_”移动到“P1_”后面；
“Ambient_”移动到“P1_”后面；
'''
import tkinter as tk
from tkinter import messagebox, filedialog
import os
import sys
import shutil
from datetime import datetime

# 版本校验：确保使用Python 3.8及以上
assert sys.version_info >= (3, 8), "请使用Python 3.8及以上版本运行此程序"

class BatchRenameTool:
    def __init__(self, root):
        self.root = root
        self.root.title("批量重命名docx文件工具（保留原文件）")
        self.root.geometry("700x350")  # 扩大窗口适配新组件
        
        # 需要移动的前缀列表
        self.target_prefixes = ["M1_", "M2_", "M3_", "M4_", "M5_", "Ambient_"]
        self.target_key = "P1_"  # 目标位置前缀
        
        # 初始化变量
        self.source_folder = tk.StringVar()  # 源文件夹（原始文件）
        self.output_folder = tk.StringVar()  # 输出文件夹（重命名后的文件）
        self.process_result = tk.StringVar(value="等待处理...")
        self.rename_log = []  # 重命名日志
        
        # 创建GUI界面元素
        self.create_widgets()
    
    def create_widgets(self):
        """创建所有GUI组件"""
        # 1. 源文件夹选择区域
        tk.Label(self.root, text="源文件夹（原始文件）:", font=("Arial", 10)).grid(
            row=0, column=0, padx=10, pady=15, sticky="e")
        
        source_entry = tk.Entry(self.root, textvariable=self.source_folder, width=50)
        source_entry.grid(row=0, column=1, padx=10, pady=15)
        
        source_btn = tk.Button(self.root, text="选择源文件夹", command=self.select_source_folder,
                              bg="#2196F3", fg="white")
        source_btn.grid(row=0, column=2, padx=5, pady=15)
        
        # 2. 输出文件夹选择区域
        tk.Label(self.root, text="输出文件夹（新文件）:", font=("Arial", 10)).grid(
            row=1, column=0, padx=10, pady=10, sticky="e")
        
        output_entry = tk.Entry(self.root, textvariable=self.output_folder, width=50)
        output_entry.grid(row=1, column=1, padx=10, pady=10)
        
        output_btn = tk.Button(self.root, text="选择输出文件夹", command=self.select_output_folder,
                              bg="#9C27B0", fg="white")
        output_btn.grid(row=1, column=2, padx=5, pady=10)
        
        # 3. 功能说明标签
        desc_label = tk.Label(
            self.root, 
            text="功能：将文件名中的 M1_/M2_/M3_/M4_/M5_/Ambient_ 移动到 P1_ 后面（原始文件保留）",
            font=("Arial", 9),
            fg="#666666"
        )
        desc_label.grid(row=2, column=1, padx=10, pady=5, sticky="w")
        
        # 4. 执行批量重命名按钮
        run_btn = tk.Button(self.root, text="批量处理（复制并重命名）", command=self.batch_rename,
                            bg="#4CAF50", fg="white", font=("Arial", 11, "bold"),
                            width=25, height=1)
        run_btn.grid(row=3, column=1, padx=10, pady=15)
        
        # 5. 处理结果显示区域
        tk.Label(self.root, text="处理状态:", font=("Arial", 10)).grid(
            row=4, column=0, padx=10, pady=10, sticky="e")
        
        result_label = tk.Label(self.root, textvariable=self.process_result, 
                                fg="#FF5722", font=("Arial", 9))
        result_label.grid(row=4, column=1, padx=10, pady=10, sticky="w")
        
        # 6. 查看日志按钮
        log_btn = tk.Button(self.root, text="查看处理日志", command=self.show_log,
                            bg="#FF9800", fg="white", font=("Arial", 9))
        log_btn.grid(row=5, column=1, padx=10, pady=5)
    
    def select_source_folder(self):
        """选择源文件夹（原始docx文件所在目录）"""
        folder = filedialog.askdirectory(title="选择包含原始docx文件的文件夹")
        if folder:
            self.source_folder.set(folder)
            self.process_result.set("已选源文件夹，请选择输出文件夹")
            self.rename_log.clear()
    
    def select_output_folder(self):
        """选择输出文件夹（重命名后的文件保存目录）"""
        folder = filedialog.askdirectory(title="选择重命名后文件的保存文件夹")
        if folder:
            self.output_folder.set(folder)
            self.process_result.set("已选输出文件夹，点击按钮开始处理")
    
    def get_new_filename(self, old_name):
        """生成新文件名（核心逻辑：移动指定前缀到P1_后）"""
        name_without_ext = os.path.splitext(old_name)[0]
        ext = os.path.splitext(old_name)[1]
        
        new_name = name_without_ext
        has_changed = False
        
        # 遍历需要移动的前缀
        for prefix in self.target_prefixes:
            if prefix in new_name and self.target_key in new_name:
                # 移除目标前缀
                new_name = new_name.replace(prefix, "")
                # 将目标前缀插入到P1_后面
                p1_index = new_name.find(self.target_key)
                if p1_index != -1:
                    insert_pos = p1_index + len(self.target_key)
                    new_name = new_name[:insert_pos] + prefix + new_name[insert_pos:]
                    has_changed = True
        
        return new_name + ext if has_changed else old_name, has_changed
    
    def rename_and_copy_file(self, old_path, output_folder):
        """复制文件到输出文件夹并修改名称（原始文件保留）"""
        try:
            old_name = os.path.basename(old_path)
            ext = os.path.splitext(old_name)[1]
            
            # 仅处理docx文件
            if ext.lower() != ".docx":
                return False, f"跳过：非docx文件 - {old_name}"
            
            # 生成新文件名
            new_name, has_changed = self.get_new_filename(old_name)
            new_path = os.path.join(output_folder, new_name)
            
            # 避免输出文件夹中重复命名
            counter = 1
            temp_new_path = new_path
            while os.path.exists(temp_new_path):
                name_without_ext = os.path.splitext(new_name)[0]
                temp_new_name = f"{name_without_ext}_{counter}{ext}"
                temp_new_path = os.path.join(output_folder, temp_new_name)
                counter += 1
            
            # 复制文件（保留原文件，新文件到输出目录）
            shutil.copy2(old_path, temp_new_path)  # copy2保留文件元数据
            
            if has_changed:
                return True, f"成功：{old_name} → {os.path.basename(temp_new_path)}"
            else:
                return True, f"复制完成（无需重命名）：{old_name}"
        
        except Exception as e:
            return False, f"失败：{old_name} - {str(e)}"
    
    def batch_rename(self):
        """批量处理：复制并重命名文件到输出文件夹（原始文件保留）"""
        # 校验文件夹路径
        source_folder = self.source_folder.get().strip()
        output_folder = self.output_folder.get().strip()
        
        if not source_folder or not os.path.isdir(source_folder):
            messagebox.showwarning("警告", "请选择有效的源文件夹！")
            return
        
        if not output_folder or not os.path.isdir(output_folder):
            messagebox.showwarning("警告", "请选择有效的输出文件夹！")
            return
        
        # 遍历源文件夹中的文件
        all_files = [f for f in os.listdir(source_folder) 
                     if os.path.isfile(os.path.join(source_folder, f))]
        
        if not all_files:
            messagebox.showinfo("提示", "源文件夹内未找到任何文件！")
            self.process_result.set("无文件可处理")
            return
        
        # 开始批量处理
        success_count = 0
        skip_count = 0
        fail_count = 0
        self.rename_log.clear()
        
        self.process_result.set(f"正在处理...共{len(all_files)}个文件")
        self.root.update()  # 刷新GUI
        
        for filename in all_files:
            old_path = os.path.join(source_folder, filename)
            success, msg = self.rename_and_copy_file(old_path, output_folder)
            self.rename_log.append(msg)
            
            if "成功" in msg or "复制完成" in msg:
                success_count += 1
            elif "跳过" in msg:
                skip_count += 1
            elif "失败" in msg:
                fail_count += 1
        
        # 汇总结果
        result_summary = f"处理完成！成功：{success_count} | 跳过：{skip_count} | 失败：{fail_count}"
        self.process_result.set(result_summary)
        messagebox.showinfo("处理结果", result_summary)
    
    def show_log(self):
        """显示详细处理日志"""
        if not self.rename_log:
            messagebox.showinfo("日志", "暂无处理日志")
            return
        
        log_text = "\n".join(self.rename_log)
        # 创建日志窗口
        log_window = tk.Toplevel(self.root)
        log_window.title("处理日志")
        log_window.geometry("650x450")
        
        # 滚动条
        scrollbar = tk.Scrollbar(log_window)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 日志文本框（只读）
        log_textbox = tk.Text(log_window, yscrollcommand=scrollbar.set, font=("Arial", 9))
        log_textbox.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        log_textbox.insert(tk.END, log_text)
        log_textbox.config(state=tk.DISABLED)
        
        scrollbar.config(command=log_textbox.yview)

if __name__ == "__main__":
    # 启动GUI程序
    root = tk.Tk()
    app = BatchRenameTool(root)
    root.mainloop()
