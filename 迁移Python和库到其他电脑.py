你想知道如何在当前电脑上查询Python版本和已安装的所有第三方库版本，并且把这些环境配置完整地迁移到另一台电脑上，确保两台电脑的Python运行环境完全一致。

一、查询Python版本和库版本
1. 查询Python版本
有两种简单的方式可以查看Python版本：

方式1：命令行直接查询
打开终端/命令提示符，执行以下命令：

# 查询Python版本（注意：如果安装了Python2和Python3，可能需要用python3或python --version）
python --version
# 或
python -V

# 如果是Windows系统，可能需要区分python和py命令
py -3 --version
输出示例：Python 3.8.7

方式2：Python代码查询
创建一个.py文件执行，或在Python交互环境中运行：

import sys
print(f"Python版本：{sys.version}")
print(f"Python版本详细信息：{sys.version_info}")
输出示例：

Python版本：3.8.7 (default, Feb 14 2026, 10:00:00) 
[GCC 9.4.0]
Python版本详细信息：sys.version_info(major=3, minor=8, micro=7, releaselevel='final', serial=0)
2. 查询所有已安装库的版本
方式1：pip命令查询（推荐）

# 查询所有已安装库及版本（简洁版）
pip list

# 查询所有已安装库及版本（详细版，包含安装路径）
pip freeze

# 将库版本信息保存到文件（关键：用于迁移）
pip freeze > requirements.txt
执行pip freeze > requirements.txt后，会在当前目录生成requirements.txt文件，内容示例：

python-docx==0.8.11
numpy==1.24.3
pandas==2.0.1
tkinter==8.6  # 注意：tkinter通常是Python内置库，可能不会出现在freeze结果中
方式2：Python代码查询

import pkg_resources

# 获取所有已安装的第三方库
installed_packages = pkg_resources.working_set
# 格式化输出库名和版本
package_list = sorted([f"{pkg.key}=={pkg.version}" for pkg in installed_packages])

# 打印所有库版本
for pkg in package_list:
    print(pkg)

# 保存到文件
with open("requirements.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(package_list))
二、迁移环境到另一台电脑
迁移的核心是：先在新电脑安装相同版本的Python，再通过requirements.txt安装相同版本的库。

步骤1：复制版本信息文件
将生成的requirements.txt文件复制到另一台电脑（U盘、网盘、邮件等方式）。

步骤2：安装相同版本的Python
查看旧电脑的Python版本（比如3.8.7），到Python官网下载对应版本的安装包。
在新电脑安装Python，注意：
勾选「Add Python to PATH」（添加到环境变量）；
安装路径尽量与旧电脑一致（可选，非必须）。
步骤3：在新电脑安装相同版本的库
打开新电脑的终端/命令提示符，进入requirements.txt所在目录；
执行以下命令安装所有库：
# 基础安装（使用官方源）
pip install -r requirements.txt

# 国内用户建议使用清华镜像源，速度更快
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
步骤4：验证环境是否一致
在新电脑执行以下命令，检查版本是否匹配：

# 验证Python版本
python --version

# 验证库版本（以python-docx为例）
pip show python-docx
三、特殊情况处理
内置库（如tkinter）：

tkinter是Python内置库，只要安装Python时选择了「Install Tcl/Tk and IDLE」，就会自动安装，无需手动pip安装。
如果新电脑缺少tkinter，Windows系统可重新运行Python安装包，选择「Modify」→ 勾选「Tcl/Tk and IDLE」→ 「Install」。
系统相关库（如pywin32）：

部分库与操作系统（Windows/Linux/macOS）相关，requirements.txt中可能需要区分，安装时若报错，可单独安装对应系统版本：
pip install pywin32==227  # Windows专属
pip版本不一致：

如果pip版本过低导致安装失败，先升级pip：
python -m pip install --upgrade pip
总结
查询环境：用python --version查Python版本，用pip freeze > requirements.txt导出所有库版本；
迁移Python：在新电脑安装与旧电脑完全相同的Python版本；
迁移库：复制requirements.txt到新电脑，执行pip install -r requirements.txt安装相同版本的库；
验证：安装完成后检查Python和核心库的版本，确保与原环境一致。
通过这种方式，能最大程度保证两台电脑的Python运行环境完全一致，避免因版本差异导致代码运行出错。
