import tkinter as tk
from tkinter import messagebox, filedialog
import os
import sys

# 版本校验：确保使用Python 3.8及以上
assert sys.version_info >= (3, 8), "请使用Python 3.8及以上版本运行此程序"

class BatchRenameTool:
    def __init__(self, root):
        self.root = root
        self.root.title("批量重命名docx文件工具")
        self.root.geometry("650x280")  # 窗口大小
        
        # 需要移动的前缀列表
        self.target_prefixes = ["M1_", "M2_", "M3_", "M4_", "M5_", "Ambient_"]
        self.target_key = "P1_"  # 目标位置前缀
        
        # 初始化变量
        self.folder_path = tk.StringVar()  # 存储选中的文件夹路径
        self.process_result = tk.StringVar(value="等待处理...")
        self.rename_log = []  # 重命名日志
        
        # 创建GUI界面元素
        self.create_widgets()
    
    def create_widgets(self):
        """创建所有GUI组件"""
        # 1. 文件夹选择区域
        tk.Label(self.root, text="目标文件夹:", font=("Arial", 10)).grid(
            row=0, column=0, padx=10, pady=15, sticky="e")
        
        # 文件夹路径输入框
        folder_entry = tk.Entry(self.root, textvariable=self.folder_path, width=55)
        folder_entry.grid(row=0, column=1, padx=10, pady=15)
        
        # 选择文件夹按钮
        browse_btn = tk.Button(self.root, text="选择文件夹", command=self.select_folder,
                              bg="#2196F3", fg="white")
        browse_btn.grid(row=0, column=2, padx=5, pady=15)
        
        # 2. 功能说明标签
        desc_label = tk.Label(
            self.root, 
            text="功能：将文件名中的 M1_/M2_/M3_/M4_/M5_/Ambient_ 移动到 P1_ 后面",
            font=("Arial", 9),
            fg="#666666"
        )
        desc_label.grid(row=1, column=1, padx=10, pady=5, sticky="w")
        
        # 3. 执行批量重命名按钮
        run_btn = tk.Button(self.root, text="批量重命名", command=self.batch_rename,
                            bg="#4CAF50", fg="white", font=("Arial", 11, "bold"),
                            width=20, height=1)
        run_btn.grid(row=2, column=1, padx=10, pady=10)
        
        # 4. 处理结果显示区域
        tk.Label(self.root, text="处理状态:", font=("Arial", 10)).grid(
            row=3, column=0, padx=10, pady=10, sticky="e")
        
        result_label = tk.Label(self.root, textvariable=self.process_result, 
                                fg="#FF5722", font=("Arial", 9))
        result_label.grid(row=3, column=1, padx=10, pady=10, sticky="w")
        
        # 5. 查看日志按钮
        log_btn = tk.Button(self.root, text="查看重命名日志", command=self.show_log,
                            bg="#FF9800", fg="white", font=("Arial", 9))
        log_btn.grid(row=4, column=1, padx=10, pady=5)
    
    def select_folder(self):
        """打开文件夹选择对话框，获取目标文件夹路径"""
        folder = filedialog.askdirectory(title="选择包含docx文件的文件夹")
        if folder:
            self.folder_path.set(folder)
            self.process_result.set("已选中文件夹，点击按钮开始重命名")
            self.rename_log.clear()  # 清空历史日志
    
    def rename_file(self, old_path):
        """单个文件重命名逻辑"""
        try:
            # 获取文件名和扩展名
            old_name = os.path.basename(old_path)
            name_without_ext = os.path.splitext(old_name)[0]
            ext = os.path.splitext(old_name)[1]
            
            # 仅处理docx文件
            if ext.lower() != ".docx":
                return False, f"跳过：非docx文件 - {old_name}"
            
            # 检查是否包含目标前缀和P1_
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
            
            # 如果文件名有变化，执行重命名
            if has_changed and new_name != name_without_ext:
                new_full_name = new_name + ext
                new_path = os.path.join(os.path.dirname(old_path), new_full_name)
                
                # 避免重复命名（如果新文件名已存在，添加序号）
                counter = 1
                temp_new_path = new_path
                while os.path.exists(temp_new_path):
                    temp_new_name = f"{new_name}_{counter}{ext}"
                    temp_new_path = os.path.join(os.path.dirname(old_path), temp_new_name)
                    counter += 1
                
                # 执行重命名
                os.rename(old_path, temp_new_path)
                return True, f"成功：{old_name} → {os.path.basename(temp_new_path)}"
            else:
                return False, f"跳过：无需修改 - {old_name}"
        
        except Exception as e:
            return False, f"失败：{os.path.basename(old_path)} - {str(e)}"
    
    def batch_rename(self):
        """批量重命名文件夹内的docx文件"""
        # 校验文件夹路径
        target_folder = self.folder_path.get().strip()
        if not target_folder or not os.path.isdir(target_folder):
            messagebox.showwarning("警告", "请选择有效的文件夹！")
            return
        
        # 遍历文件夹，筛选文件
        all_files = [f for f in os.listdir(target_folder) 
                     if os.path.isfile(os.path.join(target_folder, f))]
        
        if not all_files:
            messagebox.showinfo("提示", "选中的文件夹内未找到任何文件！")
            self.process_result.set("无文件可处理")
            return
        
        # 开始批量重命名
        success_count = 0
        skip_count = 0
        fail_count = 0
        self.rename_log.clear()
        
        self.process_result.set(f"正在处理...共{len(all_files)}个文件")
        self.root.update()  # 刷新GUI，显示处理状态
        
        for filename in all_files:
            file_path = os.path.join(target_folder, filename)
            success, msg = self.rename_file(file_path)
            self.rename_log.append(msg)
            
            if "成功" in msg:
                success_count += 1
            elif "跳过" in msg:
                skip_count += 1
            elif "失败" in msg:
                fail_count += 1
        
        # 汇总结果并提示
        result_summary = f"处理完成！成功：{success_count} | 跳过：{skip_count} | 失败：{fail_count}"
        self.process_result.set(result_summary)
        
        # 显示简要结果
        messagebox.showinfo("批量重命名结果", result_summary)
    
    def show_log(self):
        """显示重命名详细日志"""
        if not self.rename_log:
            messagebox.showinfo("日志", "暂无重命名日志")
            return
        
        log_text = "\n".join(self.rename_log)
        # 创建新窗口显示日志
        log_window = tk.Toplevel(self.root)
        log_window.title("重命名日志")
        log_window.geometry("600x400")
        
        # 添加滚动条
        scrollbar = tk.Scrollbar(log_window)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 日志文本框
        log_textbox = tk.Text(log_window, yscrollcommand=scrollbar.set, font=("Arial", 9))
        log_textbox.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        log_textbox.insert(tk.END, log_text)
        log_textbox.config(state=tk.DISABLED)  # 设置为只读
        
        scrollbar.config(command=log_textbox.yview)

if __name__ == "__main__":
    # 启动GUI程序
    root = tk.Tk()
    app = BatchRenameTool(root)
    root.mainloop()
