import sys
import os
import shutil
from datetime import datetime
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QGridLayout, QGroupBox, QLabel, QLineEdit, 
                            QPushButton, QTextEdit, QFileDialog, QComboBox)
import pandas as pd
import json

class ExcelProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config_file = 'config.json'
        self.match_options = ["请选择", "医保", "社保"]
        self.load_config()
        self.init_ui()
        self.setup_logging()

    def load_config(self):
        """加载配置文件"""
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                self.config = json.load(f)
        except FileNotFoundError:
            self.config = {
                'source_file1': '',
                'source_file2': '',
                'output_file': '',
                'match_field1': '',
                'match_field2': '',
                'match_type': '请选择'
            }
            self.show_config_warning = True

    def save_config(self):
        """保存配置到文件"""
        config = {
            'source_file1': self.source_file1.text(),
            'source_file2': self.source_file2.text(),
            'output_file': self.output_file.text(),
            'match_field1': self.match_field1.currentText(),
            'match_field2': self.match_field2.currentText(),
            'match_type': self.match_type_combo.currentText()
        }
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
            self.log("配置已保存")
        except Exception as e:
            self.log(f"保存配置时发生错误: {str(e)}")

    def init_ui(self):
        self.setWindowTitle('Excel数据处理工具')
        self.setMinimumSize(600, 400)

        # 创建中央部件和主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 文件选择组
        group_box = QGroupBox("文件选择")
        grid_layout = QGridLayout()
        group_box.setLayout(grid_layout)

        # 输入文件（原源文件1）
        grid_layout.addWidget(QLabel("输入文件:"), 0, 0)
        self.source_file1 = QLineEdit()
        grid_layout.addWidget(self.source_file1, 0, 1)
        browse_btn1 = QPushButton("浏览")
        browse_btn1.clicked.connect(lambda: self.browse_file(self.source_file1))
        grid_layout.addWidget(browse_btn1, 0, 2)

        # 小士兵反馈结果（原源文件2）
        grid_layout.addWidget(QLabel("小士兵反馈结果:"), 1, 0)
        self.source_file2 = QLineEdit()
        grid_layout.addWidget(self.source_file2, 1, 1)
        browse_btn2 = QPushButton("浏览")
        browse_btn2.clicked.connect(lambda: self.browse_file(self.source_file2))
        grid_layout.addWidget(browse_btn2, 1, 2)

        # 小士兵新任务文件（原输出文件）
        grid_layout.addWidget(QLabel("小士兵新任务文件:"), 2, 0)
        self.output_file = QLineEdit()
        grid_layout.addWidget(self.output_file, 2, 1)
        browse_btn3 = QPushButton("浏览")
        browse_btn3.clicked.connect(lambda: self.browse_file(self.output_file))
        grid_layout.addWidget(browse_btn3, 2, 2)

        layout.addWidget(group_box)

        # 添加匹配字段配置组
        match_group = QGroupBox("匹配字段配置")
        match_layout = QGridLayout()
        match_group.setLayout(match_layout)

        # 输入文件匹配字段（改为下拉选择框）
        match_layout.addWidget(QLabel("输入文件匹配字段:"), 0, 0)
        self.match_field1 = QComboBox()
        match_layout.addWidget(self.match_field1, 0, 1)

        # 小士兵反馈结果匹配字段（改为下拉选择框）
        match_layout.addWidget(QLabel("小士兵反馈结果匹配字段:"), 1, 0)
        self.match_field2 = QComboBox()
        match_layout.addWidget(self.match_field2, 1, 1)

        # 添加刷新字段按钮
        refresh_btn = QPushButton("刷新字段列表")
        refresh_btn.clicked.connect(self.refresh_fields)
        match_layout.addWidget(refresh_btn, 2, 1)

        # 保存配置按钮
        match_layout.addWidget(QLabel(""), 3, 0)  # 空行
        save_config_btn = QPushButton("保存配置")
        save_config_btn.clicked.connect(self.save_config)
        match_layout.addWidget(save_config_btn, 3, 1)

        # 从配置文件加载上次的配置
        self.source_file1.setText(self.config.get('source_file1', ''))
        self.source_file2.setText(self.config.get('source_file2', ''))
        self.output_file.setText(self.config.get('output_file', ''))
        self.match_field1.setCurrentText(self.config.get('match_field1', ''))
        self.match_field2.setCurrentText(self.config.get('match_field2', ''))

        layout.addWidget(match_group)

        # 处理按钮
        process_btn = QPushButton("处理数据")
        process_btn.clicked.connect(self.process_excel)
        layout.addWidget(process_btn)

        # 日志视图
        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        layout.addWidget(self.log_view)

    def setup_logging(self):
        # 创建logs目录
        os.makedirs('logs', exist_ok=True)
        
        # 创建日志文件
        self.log_file = f'logs/excel_processor_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'

    def log(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}"
        self.log_view.append(log_message)
        
        # 写入日志文件
        with open(self.log_file, 'a', encoding='utf-8') as f:
            f.write(log_message + '\n')

    def browse_file(self, line_edit):
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if filename:
            line_edit.setText(filename)

    def backup_file(self, file_path):
        # 创建backups目录
        os.makedirs('backups', exist_ok=True)
        
        # 创建备份文件名
        filename = os.path.basename(file_path)
        backup_name = f"backups/{datetime.now().strftime('%Y%m%d_%H%M%S')}_{filename}"
        
        # 复制文件
        shutil.copy2(file_path, backup_name)
        return backup_name

    def process_excel(self):
        # 检查文件是否都已选择
        if not all([self.source_file1.text(), self.source_file2.text(), self.output_file.text()]):
            self.log("请选择所有必需的文件")
            return

        try:
            # 备份源文件
            self.log("开始备份源文件...")
            self.backup_file(self.source_file1.text())
            self.backup_file(self.source_file2.text())
            self.log("文件备份完成")

            # 处理Excel文件
            self.log("开始处理Excel文件...")
            self.process_excel_files(
                self.source_file1.text(),
                self.source_file2.text(),
                self.output_file.text()
            )
            self.log("数据处理完成！")

        except Exception as e:
            self.log(f"错误: {str(e)}")

    def process_excel_files(self, source1, source2, output):
        # 获取匹配字段
        match_field1 = self.match_field1.currentText()
        match_field2 = self.match_field2.currentText()

        # 检查匹配字段是否已选择
        if not match_field1 or not match_field2:
            self.log("请选择匹配字段")
            return

        # 读取Excel文件
        df1 = pd.read_excel(source1)
        df2 = pd.read_excel(source2)

        # 检查匹配字段是否存在于数据框中
        if match_field1 not in df1.columns:
            self.log(f"错误：源文件1中不存在字段 '{match_field1}'")
            return
        if match_field2 not in df2.columns:
            self.log(f"错误：源文件2中不存在字段 '{match_field2}'")
            return

        try:
            # 记录处理前的数据量
            original_count = len(df1)
            
            # 使用inner join只保留匹配的数据
            result = pd.merge(df1, df2[[match_field2]], 
                            left_on=match_field1, 
                            right_on=match_field2, 
                            how='inner')
            
            # 记录处理后的数据量
            filtered_count = len(result)
            
            # 保存结果
            result.to_excel(output, index=False)
            self.log(f"数据处理完成！原始数据：{original_count}条，匹配数据：{filtered_count}条，过滤掉：{original_count - filtered_count}条")
        except Exception as e:
            self.log(f"合并数据时发生错误: {str(e)}")

    def on_match_type_changed(self, text):
        """当匹配类型改变时自动填充匹配字段"""
        if text == "医保":
            self.match_field1.setText("医保号")
            self.match_field2.setText("医保号")
        elif text == "社保":
            self.match_field1.setText("社保号")
            self.match_field2.setText("社保号")
        else:
            self.match_field1.setText("")
            self.match_field2.setText("")

    def refresh_fields(self):
        """从Excel文件中读取字段列表"""
        try:
            # 检查文件是否已选择
            if not self.source_file1.text() or not self.source_file2.text():
                self.log("请先选择输入文件和小士兵反馈结果文件")
                return

            # 读取输入文件的字段
            df1 = pd.read_excel(self.source_file1.text())
            self.match_field1.clear()
            self.match_field1.addItems(df1.columns.tolist())

            # 读取小士兵反馈结果文件的字段
            df2 = pd.read_excel(self.source_file2.text())
            self.match_field2.clear()
            self.match_field2.addItems(df2.columns.tolist())

            # 从配置文件加载上次选择的值
            saved_field1 = self.config.get('match_field1', '')
            saved_field2 = self.config.get('match_field2', '')
            
            # 设置上次选择的值
            index1 = self.match_field1.findText(saved_field1)
            if index1 >= 0:
                self.match_field1.setCurrentIndex(index1)
            
            index2 = self.match_field2.findText(saved_field2)
            if index2 >= 0:
                self.match_field2.setCurrentIndex(index2)

            self.log("字段列表已更新")
        except Exception as e:
            self.log(f"读取字段列表时发生错误: {str(e)}")

def main():
    app = QApplication(sys.argv)
    window = ExcelProcessor()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()