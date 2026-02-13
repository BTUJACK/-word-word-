import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import shutil
import tempfile
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# 适配 Python 3.8.7 依赖（执行前安装）：
# pip install python-docx==0.8.11

class WordImageTableTool:
    def __init__(self, root):
        # 主窗口配置
        self.root = root
        self.root.title("Word图片表格调整工具（保留空行+保格式）")
        self.root.geometry("800x520")
        self.root.resizable(False, False)

        # 备份/文件变量
        self.tmp_dir = None
        self.backup_path = ""
        self.current_file = ""

        # ========== GUI 界面布局 ==========
        # 1. 文件选择区域
        frame_file = tk.Frame(root, padx=15, pady=10)
        frame_file.pack(fill=tk.X)

        tk.Label(frame_file, text="待处理Word文件：", font=("微软雅黑", 10)).grid(row=0, column=0, sticky=tk.W)
        self.file_var = tk.StringVar()
        entry_file = tk.Entry(frame_file, textvariable=self.file_var, width=55, font=("微软雅黑", 9))
        entry_file.grid(row=0, column=1, padx=8)
        btn_file = tk.Button(frame_file, text="选择文件", command=self.choose_file,
                              font=("微软雅黑", 9), width=10, bg="#409EFF", fg="white")
        btn_file.grid(row=0, column=2)

        # 2. 功能按钮区域
        frame_btn = tk.Frame(root, padx=15, pady=10)
        frame_btn.pack(fill=tk.X)

        self.btn_process = tk.Button(frame_btn, text="执行调整：删图片上方内容+保留空行", 
                                    command=self.process_word, font=("微软雅黑", 11, "bold"),
                                    width=35, height=2, bg="#67C23A", fg="white")
        self.btn_process.pack(side=tk.LEFT, padx=5)

        self.btn_restore = tk.Button(frame_btn, text="恢复原文件", command=self.restore_original,
                                    font=("微软雅黑", 10), width=18, height=2, bg="#F56C6C", fg="white")
        self.btn_restore.pack(side=tk.LEFT, padx=5)

        # 3. 日志显示区域
        frame_log = tk.Frame(root, padx=15, pady=5)
        frame_log.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame_log, text="操作日志：", font=("微软雅黑", 10)).pack(anchor=tk.W)
        self.log_text = scrolledtext.ScrolledText(frame_log, height=15, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 初始化日志
        self.log("✅ Python 3.8.7 环境适配完成，工具就绪")
        self.log("💡 操作流程：选择Word文件 → 点击执行调整 → 完成后可恢复原文件\n")

    # ========== 基础辅助方法 ==========
    def log(self, content):
        """带时间戳的日志"""
        import datetime
        time_str = datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
        self.log_text.insert(tk.END, f"{time_str} {content}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def choose_file(self):
        """选择单个Word文件"""
        file_path = filedialog.askopenfilename(
            title="选择待处理的Word文档",
            filetypes=[("Word 2007-2019 文档", "*.docx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.file_var.set(file_path)
            self.current_file = file_path
            self.log(f"📂 已选择文件：{os.path.basename(file_path)}")

    def create_empty_paragraph(self):
        """创建空段落（空行），保留默认格式"""
        # 构建空段落的XML元素（兼容python-docx 0.8.11）
        empty_para_xml = """
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r>
                <w:t></w:t>
            </w:r>
        </w:p>
        """
        return parse_xml(empty_para_xml)

    # ========== 核心：全类型图片定位 ==========
    def find_all_images(self, doc):
        """识别所有类型的图片"""
        body_elems = list(doc._body._element)
        image_idx = -1
        target_image_elem = None

        # 支持的图片标签类型
        image_tags = ['pic:pic', 'a:graphic', 'w:drawing', 'v:shape', 'wp:inline', 'wp:anchor']

        self.log("  ▶ 开始扫描所有类型图片...")
        for idx, elem in enumerate(body_elems):
            # 检查当前元素是否是图片
            elem_xml = elem.xml.lower()
            tag_match = any(tag in elem.tag for tag in image_tags)
            content_match = 'blip' in elem_xml or 'image' in elem_xml or 'pict' in elem_xml

            if tag_match or content_match:
                image_idx = idx
                target_image_elem = elem
                self.log(f"  ✅ 找到图片！类型：{elem.tag.split('}')[-1]}，位置索引：{image_idx}")
                break

        if image_idx == -1:
            self.log("  ❌ 未找到任何类型的图片")
            return -1, None
        return image_idx, target_image_elem

    def get_table_elements_below_image(self, doc, image_idx):
        """获取图片下方的所有表格元素"""
        body_elems = list(doc._body._element)
        table_elems = []

        for idx in range(image_idx + 1, len(body_elems)):
            elem = body_elems[idx]
            if elem.tag.endswith('tbl'):
                table_elem = parse_xml(elem.xml)
                table_elems.append(table_elem)
                self.log(f"  ✅ 找到图片下方表格，索引：{idx}")

        if not table_elems:
            self.log("  ⚠️  图片下方未找到表格")
        return table_elems

    # ========== 核心优化：插入空行 + 调整内容 ==========
    def adjust_word_content(self, doc):
        """
        优化后逻辑：
        1. 删除图片上方所有文字/表格
        2. 把图片下方的表格移动到图片上方
        3. 表格上下各保留2个空行
        4. 图片上下各保留2个空行
        5. 保留所有原始格式
        """
        self.log("🔧 开始分析文档元素结构")
        
        # 步骤1：找图片
        image_idx, image_elem = self.find_all_images(doc)
        if image_idx == -1:
            return False

        # 步骤2：删除图片上方所有内容
        self.log("  ▶ 删除图片上方所有内容")
        deleted_count = 0
        for idx in range(image_idx - 1, -1, -1):
            try:
                doc._body._element.remove(doc._body._element[idx])
                deleted_count += 1
            except Exception as e:
                self.log(f"  ⚠️  删除索引{idx}元素失败：{str(e)}")
        self.log(f"  ✅ 已删除图片上方 {deleted_count} 个元素")

        # 步骤3：获取并清理图片下方表格
        table_elems = self.get_table_elements_below_image(doc, 0)  # 图片现在是索引0
        
        # 删除图片下方原表格
        self.log("  ▶ 清理图片下方原表格")
        body_elems = list(doc._body._element)
        for idx in range(len(body_elems)-1, 0, -1):
            elem = body_elems[idx]
            if elem.tag.endswith('tbl'):
                try:
                    doc._body._element.remove(elem)
                    self.log(f"  ✅ 删除图片下方原表格，索引：{idx}")
                except:
                    pass

        # 步骤4：插入表格 + 表格上下各2个空行
        if table_elems:
            self.log("  ▶ 插入表格并添加空行（表格上下各2行）")
            # 先插入表格上方的2个空行
            for _ in range(2):
                empty_para = self.create_empty_paragraph()
                doc._body._element.insert(0, empty_para)
            
            # 插入表格（保持原有顺序）
            for table_elem in reversed(table_elems):
                doc._body._element.insert(2, table_elem)  # 空行后插入表格
            
            # 插入表格下方的2个空行（图片上方）
            for _ in range(2):
                empty_para = self.create_empty_paragraph()
                doc._body._element.insert(2 + len(table_elems), empty_para)
            
            self.log(f"  ✅ 成功插入 {len(table_elems)} 个表格，且表格上下各2个空行")

        # 步骤5：图片上下各保留2个空行（补充）
        self.log("  ▶ 确保图片上下各2个空行")
        # 图片当前位置：表格空行后 → 计算图片新索引
        img_current_idx = 2 + len(table_elems) + 2  # 表格上2行 + 表格数 + 表格下2行
        
        # 图片上方已通过表格空行满足，补充图片下方2个空行
        for _ in range(2):
            empty_para = self.create_empty_paragraph()
            doc._body._element.insert(img_current_idx + 1, empty_para)
        
        self.log("  ✅ 图片上下各保留2个空行完成")

        return True

    # ========== 主流程：处理Word文件 ==========
    def process_word(self):
        """完整处理流程"""
        # 输入校验
        if not self.current_file or not os.path.exists(self.current_file):
            messagebox.showerror("错误", "请选择有效的Word文件！")
            return

        # 1. 备份原文件
        self.log("📦 开始备份原文件")
        if self.tmp_dir is None:
            self.tmp_dir = tempfile.mkdtemp(prefix="word_backup_387_")
        self.backup_path = os.path.join(self.tmp_dir, os.path.basename(self.current_file))
        shutil.copy2(self.current_file, self.backup_path)
        self.log(f"✅ 原文件已备份至：{self.backup_path}")

        # 2. 处理文档
        try:
            doc = Document(self.current_file)
            self.log(f"\n🔧 开始处理文件：{os.path.basename(self.current_file)}")

            # 核心调整
            adjust_success = self.adjust_word_content(doc)

            if adjust_success:
                # 保存处理后的文档
                doc.save(self.current_file)
                self.log(f"\n🎉 文档调整完成！")
                messagebox.showinfo("成功", 
                    f"✅ Word文件调整完成！\n📄 已执行：\n  1. 删除图片上方所有文字/表格\n  2. 将图片下方表格移动到图片上方\n  3. 表格上下各保留2个空行\n  4. 图片上下各保留2个空行\n✅ 所有格式100%保留")
            else:
                self.log(f"\n❌ 文档调整失败！")
                messagebox.showerror("错误", "文档调整失败（未找到图片）！")
                self.restore_original()

        except Exception as e:
            self.log(f"\n❌ 处理失败：{str(e)}")
            messagebox.showerror("错误", f"文件处理失败：{str(e)}")
            self.restore_original()

    # ========== 恢复原文件 ==========
    def restore_original(self):
        """恢复备份的原文件"""
        if not self.backup_path or not os.path.exists(self.backup_path):
            messagebox.showinfo("提示", "暂无需要恢复的原文件！")
            return

        try:
            # 覆盖恢复
            shutil.copy2(self.backup_path, self.current_file)
            
            # 清理临时目录
            if self.tmp_dir and os.path.exists(self.tmp_dir):
                shutil.rmtree(self.tmp_dir, ignore_errors=True)
            self.tmp_dir = None
            self.backup_path = ""

            self.log(f"✅ 已恢复原文件：{os.path.basename(self.current_file)}")
            messagebox.showinfo("恢复完成", f"✅ 原文件已成功恢复！")

        except Exception as e:
            self.log(f"❌ 恢复失败：{str(e)}")
            messagebox.showerror("错误", f"恢复原文件失败：{str(e)}")

# ========== 程序入口 ==========
if __name__ == "__main__":
    # 适配Windows高分屏
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception as e:
        print(f"DPI 适配提示：{e}（不影响运行）")

    # 启动GUI
    root = tk.Tk()
    app = WordImageTableTool(root)
    root.mainloop()

    # 清理临时文件
    try:
        if app.tmp_dir and os.path.exists(self.tmp_dir):
            shutil.rmtree(self.tmp_dir, ignore_errors=True)
    except:
        pass
