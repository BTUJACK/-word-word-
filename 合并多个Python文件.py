
'''
自动合并脚本（高效批量）
新建一个merge_files.py脚本，运行后会自动按顺序读取 4 个文件并合并为merged_all.py，避免手动复制粘贴：
'''
# 要合并的文件列表（按顺序）
file_list = ["3.py", "4.py", "5.py", "8-merge-word.py"]
# 合并后的目标文件
output_file = "merged_all.py"

# 以写入模式打开目标文件
with open(output_file, "w", encoding="utf-8") as out_f:
    for file_name in file_list:
        try:
            # 读取每个文件的内容
            with open(file_name, "r", encoding="utf-8") as in_f:
                # 写入文件分隔注释
                out_f.write(f"# ==================== {file_name} 内容开始 ====================\n")
                # 写入文件内容
                out_f.write(in_f.read())
                # 写入结束注释并换行，避免文件内容粘连
                out_f.write(f"\n# ==================== {file_name} 内容结束 ====================\n\n")
            print(f"成功读取并合并：{file_name}")
        except FileNotFoundError:
            print(f"警告：未找到文件 {file_name}，已跳过")
        except Exception as e:
            print(f"处理 {file_name} 时出错：{str(e)}")

print(f"\n所有文件合并完成，结果已保存到：{output_file}")
