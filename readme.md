## 项目说明

### 项目目标
该项目旨在提供一个自动控制鼠标键盘的解决方案，并实现获取操作界面坐标的功能。

### 功能特点
* 可获取操作界面坐标；
* 根据步骤文件，自动控制鼠标键盘。

### 安装指南
1. 克隆项目：`https://github.com/chen-huai/Auto_Operation.git`
2. 安装依赖库：`pip install -r requirements.txt`
3. 运行主程序： 直接运行主程序`Auto_Operate.py`即可
4. 配置文件： 运行成功后，将在桌面生成一个`config`文件夹，并生成`config_auto.csv`的配置文件。可根据`config_sap.csv`文件设置自己需要的参数
5. 打包成exe程序：
   * 安装第三方库pyinstaller：
   `pip install pyinstaller`
   * 执行打包命令：
   `pyinstaller -F -w 主程序绝对路径`

### 源码结构说明
* `Auto_Operate.py`：主程序，包含逻辑操作
* `Auto_Operate_Ui.ui`：UI界面
* `Windows_Operate.py`：操作界面具体操作
* `File_Operate.py`：文件处理模块，用于创建文件夹和获取文件名称
* `Data_Table.py`：表格处理模块，实现表格基础设置


## Project Description

### Project Objective
The project aims to provide a solution for automatic control of mouse and keyboard inputs, as well as the functionality to retrieve coordinates on the operating interface.

### Features
* Ability to retrieve coordinates on the operating interface
* Automated control of mouse and keyboard based on step files
### Installation Guide
1. Clone the project: `https://github.com/chen-huai/Auto_Operation.git`
2. Install dependencies: `pip install -r requirements.txt`
3. Run the main program: Simply run the main program `Auto_Operate.py`
4. Configuration file: Upon successful execution, a config folder will be generated on the desktop, along with the configuration file config_auto.csv. Adjust the parameters according to the config_sap.csv file to suit your needs.
5. Packaging into an exe program:
* Install the third-party library pyinstaller: `pip install pyinstaller`
* Execute the packaging command:
`pyinstaller -F -w absolute_path_to_main_program`
### Source Code Structure
`Auto_Operate.py`: Main program containing the logical operations
`Auto_Operate_Ui.ui`: UI interface
`Windows_Operate.py`: Specific operations on the interface
`File_Operate.py`: File handling module for creating folders and obtaining file names
`Data_Table.py`: Table handling module for basic table settings