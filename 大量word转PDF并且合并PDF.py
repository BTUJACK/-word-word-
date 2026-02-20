#用Python 3.8.7实现把一个文件夹里面的多个word转为PDF格式，并且把转化的PDF进行合并，并且生成一个GUI界面进行操作。
#路径不能有括号
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import sys
import datetime
import win32com.client
from PyPDF2 import PdfMerger
import pythoncom

# 适配Python 3.8.7的依赖安装命令（终端执行）：
# pip install pywin32==227 PyPDF2==2.12.1

class Word2PdfMergerGUI:
    def __init__(self, root):
        # 主窗口配置
        self.root = root
        self.root.title("Word转PDF并合并工具 (Python 3.8.7)")
        self.root.geometry("750x450")
        self.root.resizable(False, False)
        
        # 存储路径变量
        self.word_folder = tk.StringVar()
        self.pdf_output_folder = tk.StringVar()
        self.merge_output_path = tk.StringVar()
        
        # 默认路径初始化
        default_pdf_folder = os.path.join(os.getcwd(), "转换后的PDF")
        default_merge_path = os.path.join(os.getcwd(), f"合并后的PDF_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.pdf")
        self.pdf_output_folder.set(default_pdf_folder)
        self.merge_output_path.set(default_merge_path)
        
        # ========== 1. 选择Word文件夹区域 ==========
        frame1 = tk.Frame(root, padx=20, pady=10)
        frame1.pack(fill=tk.X)
        tk.Label(frame1, text="待转换Word文件夹：", font=("微软雅黑", 11)).grid(row=0, column=0, sticky=tk.W)
        entry_word = tk.Entry(frame1, textvariable=self.word_folder, width=45, font=("微软雅黑", 10))
        entry_word.grid(row=0, column=1, padx=10)
        btn_word = tk.Button(
            frame1, text="选择文件夹", command=self.select_word_folder,
            font=("微软雅黑", 10), bg="#409EFF", fg="white", width=10
        )
        btn_word.grid(row=0, column=2)
        
        # ========== 2. PDF保存路径区域 ==========
        frame2 = tk.Frame(root, padx=20, pady=10)
        frame2.pack(fill=tk.X)
        tk.Label(frame2, text="PDF临时保存路径：", font=("微软雅黑", 11)).grid(row=0, column=0, sticky=tk.W)
        entry_pdf = tk.Entry(frame2, textvariable=self.pdf_output_folder, width=45, font=("微软雅黑", 10))
        entry_pdf.grid(row=0, column=1, padx=10)
        btn_pdf = tk.Button(
            frame2, text="选择路径", command=self.select_pdf_folder,
            font=("微软雅黑", 10), bg="#409EFF", fg="white", width=10
        )
        btn_pdf.grid(row=0, column=2)
        
        # ========== 3. 合并PDF保存路径 ==========
        frame3 = tk.Frame(root, padx=20, pady=10)
        frame3.pack(fill=tk.X)
        tk.Label(frame3, text="最终合并PDF路径：", font=("微软雅黑", 11)).grid(row=0, column=0, sticky=tk.W)
        entry_merge = tk.Entry(frame3, textvariable=self.merge_output_path, width=45, font=("微软雅黑", 10))
        entry_merge.grid(row=0, column=1, padx=10)
        btn_merge = tk.Button(
            frame3, text="选择路径", command=self.select_merge_path,
            font=("微软雅黑", 10), bg="#409EFF", fg="white", width=10
        )
        btn_merge.grid(row=0, column=2)
        
        # ========== 4. 执行按钮区域 ==========
        frame4 = tk.Frame(root, padx=20, pady=15)
        frame4.pack(fill=tk.X)
        self.btn_execute = tk.Button(
            frame4, text="开始转换并合并", command=self.execute_all,
            font=("微软雅黑", 14, "bold"), bg="#67C23A", fg="white",
            width=20, height=2
        )
        self.btn_execute.pack()
        
        # ========== 5. 日志显示区域 ==========
        frame5 = tk.Frame(root, padx=20, pady=5)
        frame5.pack(fill=tk.BOTH, expand=True)
        tk.Label(frame5, text="操作日志：", font=("微软雅黑", 10)).pack(anchor=tk.W)
        self.log_text = scrolledtext.ScrolledText(frame5, height=10, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 初始化日志
        self.log("✅ 工具已就绪（Python 3.8.7适配版）")
        self.log("📌 仅支持.docx/.doc格式，需确保已安装Microsoft Word")

    # 日志添加方法（带时间戳）
    def log(self, content):
        time_str = datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
        self.log_text.insert(tk.END, f"{time_str} {content}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    # 选择Word文件夹
    def select_word_folder(self):
        folder = filedialog.askdirectory(title="选择存放Word文件的文件夹")
        if folder:
            self.word_folder.set(folder)
            # 统计Word文件数量
            word_count = len([f for f in os.listdir(folder) if f.lower().endswith((".docx", ".doc"))])
            self.log(f"📂 已选择Word文件夹：{folder}")
            self.log(f"🔍 检测到 {word_count} 个Word文件(.docx/.doc)")

    # 选择PDF临时保存文件夹
    def select_pdf_folder(self):
        folder = filedialog.askdirectory(title="选择PDF临时保存文件夹")
        if folder:
            self.pdf_output_folder.set(folder)
            self.log(f"💾 已选择PDF临时保存路径：{folder}")

    # 选择合并后PDF保存路径
    def select_merge_path(self):
        file = filedialog.asksaveasfilename(
            title="选择合并后PDF的保存位置",
            defaultextension=".pdf",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if file:
            self.merge_output_path.set(file)
            self.log(f"📁 已选择合并PDF保存路径：{file}")

    # Word转PDF核心函数（适配Python 3.8.7）
    def word_to_pdf(self, word_path, pdf_path):
        """
        将单个Word文件转为PDF
        :param word_path: Word文件路径
        :param pdf_path: 输出PDF路径
        """
        try:
            # 初始化COM组件（解决多线程/重入问题）
            pythoncom.CoInitialize()
            
            # 启动Word应用
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False  # 后台运行
            word.DisplayAlerts = 0  # 禁用弹窗
            
            # 打开文档并另存为PDF
            doc = word.Documents.Open(word_path)
            doc.SaveAs(pdf_path, FileFormat=17)  # 17 = PDF格式
            doc.Close()
            word.Quit()
            
            # 释放COM组件
            pythoncom.CoUninitialize()
            
            self.log(f"✅ 转换成功：{os.path.basename(word_path)} → {os.path.basename(pdf_path)}")
            return True
        except Exception as e:
            self.log(f"❌ 转换失败：{os.path.basename(word_path)} - {str(e)}")
            # 确保Word进程退出
            try:
                word.Quit()
            except:
                pass
            pythoncom.CoUninitialize()
            return False

    # 合并PDF核心函数
    def merge_pdfs(self, pdf_files, output_path):
        """
        合并多个PDF文件
        :param pdf_files: PDF文件路径列表
        :param output_path: 合并后输出路径
        """
        try:
            merger = PdfMerger()
            # 按顺序合并PDF
            for pdf_file in pdf_files:
                if os.path.exists(pdf_file):
                    merger.append(pdf_file)
                    self.log(f"🔗 已加入合并队列：{os.path.basename(pdf_file)}")
            
            # 保存合并后的PDF
            merger.write(output_path)
            merger.close()
            self.log(f"🎉 PDF合并完成：{output_path}")
            return True
        except Exception as e:
            self.log(f"❌ PDF合并失败：{str(e)}")
            return False

    # 主执行函数：转换+合并
    def execute_all(self):
        try:
            # 1. 路径校验
            word_folder = self.word_folder.get().strip()
            pdf_folder = self.pdf_output_folder.get().strip()
            merge_path = self.merge_output_path.get().strip()
            
            if not word_folder or not os.path.exists(word_folder):
                messagebox.showerror("错误", "请选择有效的Word文件夹！")
                return
            if not pdf_folder:
                messagebox.showerror("错误", "请选择PDF临时保存路径！")
                return
            if not merge_path:
                messagebox.showerror("错误", "请选择合并PDF保存路径！")
                return
            
            # 2. 创建PDF临时文件夹（不存在则创建）
            if not os.path.exists(pdf_folder):
                os.makedirs(pdf_folder)
                self.log(f"📁 创建PDF临时文件夹：{pdf_folder}")
            
            # 3. 获取所有Word文件（.docx/.doc）
            word_files = [
                os.path.join(word_folder, f)
                for f in os.listdir(word_folder)
                if f.lower().endswith((".docx", ".doc")) and os.path.isfile(os.path.join(word_folder, f))
            ]
            if not word_files:
                messagebox.showwarning("警告", "所选文件夹内无Word文件(.docx/.doc)！")
                return
            
            self.log("="*60)
            self.log(f"🚀 开始执行Word转PDF并合并（共{len(word_files)}个文件）")
            self.log("="*60)
            
            # 4. 批量转换Word到PDF
            pdf_files = []
            success_count = 0
            for word_file in word_files:
                # 生成PDF文件名（与Word同名）
                pdf_name = os.path.splitext(os.path.basename(word_file))[0] + ".pdf"
                pdf_path = os.path.join(pdf_folder, pdf_name)
                
                # 转换
                if self.word_to_pdf(word_file, pdf_path):
                    pdf_files.append(pdf_path)
                    success_count += 1
            
            # 5. 校验转换结果
            if not pdf_files:
                messagebox.showerror("错误", "所有Word文件转换失败！")
                return
            self.log(f"📊 转换统计：成功{success_count}个 / 总{len(word_files)}个")
            
            # 6. 合并PDF
            if not self.merge_pdfs(pdf_files, merge_path):
                messagebox.showerror("错误", "PDF合并失败！")
                return
            
            # 7. 执行完成
            self.log("="*60)
            self.log(f"✅ 全部操作完成！")
            self.log(f"📄 转换后的PDF存放：{pdf_folder}")
            self.log(f"📄 合并后的PDF：{merge_path}")
            self.log("="*60)
            
            messagebox.showinfo("操作完成",
                f"✅ 执行完成！\n"
                f"📄 Word转PDF：成功{success_count}个 / 总{len(word_files)}个\n"
                f"📁 转换后PDF路径：{pdf_folder}\n"
                f"🔗 合并后PDF路径：{merge_path}")
        
        except Exception as e:
            self.log(f"❌ 执行异常：{str(e)}")
            messagebox.showerror("执行失败", f"操作过程出错：\n{str(e)}")

if __name__ == "__main__":
    # 适配Windows高分屏（Python 3.8.7兼容）
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    
    # 启动GUI
    root = tk.Tk()
    app = Word2PdfMergerGUI(root)
    root.mainloop()
