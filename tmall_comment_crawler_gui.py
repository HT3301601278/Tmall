#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import time
from functools import partial

from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt5.QtGui import QIcon, QFont, QPixmap, QColor
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QLabel, QLineEdit, QPushButton, QSpinBox, QProgressBar, 
                            QTextEdit, QCheckBox, QGroupBox, QScrollArea, QFileDialog,
                            QMessageBox, QFrame, QSplitter, QTabWidget, QGridLayout)

# 导入爬虫核心类
from tmall_comment_crawler import TmallCommentCrawler

# 定义样式表
STYLE = """
QMainWindow {
    background-color: #f5f6fa;
}
QTabWidget::pane {
    border: 1px solid #dcdde1;
    border-radius: 5px;
    background-color: white;
}
QTabBar::tab {
    background-color: #f1f2f6;
    border: 1px solid #dcdde1;
    border-bottom: none;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    padding: 8px 16px;
    margin-right: 1px;
}
QTabBar::tab:selected {
    background-color: white;
    border-bottom: 1px solid white;
}
QGroupBox {
    border: 1px solid #dcdde1;
    border-radius: 5px;
    margin-top: 1.5ex;
    font-weight: bold;
    background-color: white;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top center;
    padding: 0 5px;
    background-color: white;
}
QLineEdit, QSpinBox {
    border: 1px solid #dcdde1;
    border-radius: 3px;
    padding: 5px;
    background-color: #f8f9fa;
}
QLineEdit:focus, QSpinBox:focus {
    border: 1px solid #a29bfe;
    background-color: white;
}
QPushButton {
    background-color: #6c5ce7;
    color: white;
    border: none;
    border-radius: 3px;
    padding: 8px 16px;
    font-weight: bold;
}
QPushButton:hover {
    background-color: #a29bfe;
}
QPushButton:pressed {
    background-color: #5f48ea;
}
QPushButton:disabled {
    background-color: #b2bec3;
}
QProgressBar {
    border: 1px solid #dcdde1;
    border-radius: 3px;
    text-align: center;
    background-color: #f8f9fa;
}
QProgressBar::chunk {
    background-color: #6c5ce7;
    width: 10px;
    margin: 0.5px;
}
QTextEdit {
    border: 1px solid #dcdde1;
    border-radius: 3px;
    background-color: #f8f9fa;
    font-family: "Consolas", "Microsoft YaHei";
}
QCheckBox {
    spacing: 5px;
}
QCheckBox::indicator {
    width: 16px;
    height: 16px;
}
QScrollArea {
    border: none;
}
"""

class CrawlerThread(QThread):
    """爬虫线程类，避免界面卡顿"""
    update_signal = pyqtSignal(str)  # 日志信号
    progress_signal = pyqtSignal(int)  # 进度信号
    finished_signal = pyqtSignal(list)  # 完成信号，传递爬取的评论列表
    
    def __init__(self, item_id, page_num, cookie=None):
        super().__init__()
        self.item_id = item_id
        self.page_num = page_num
        self.cookie = cookie
        self.crawler = TmallCommentCrawler()
        
        # 如果提供了自定义Cookie，则更新爬虫的Cookie
        if self.cookie:
            self.crawler.headers['Cookie'] = self.cookie
            # 重新提取token
            self.crawler._extract_token_from_cookie()
        
    def run(self):
        self.update_signal.emit("开始爬取评论数据...")
        all_comments = []
        
        for page in range(1, self.page_num + 1):
            self.update_signal.emit(f"正在爬取第 {page}/{self.page_num} 页评论...")
            
            try:
                comments = self.crawler.get_comments(self.item_id, 1)
                if comments:
                    all_comments.extend(comments)
                    self.update_signal.emit(f"成功获取第 {page} 页的 {len(comments)} 条评论")
                else:
                    # 检查爬虫对象中是否有错误信息
                    if hasattr(self.crawler, 'last_error') and self.crawler.last_error:
                        self.update_signal.emit(f"第 {page} 页评论获取失败: {self.crawler.last_error}")
                    else:
                        self.update_signal.emit(f"第 {page} 页没有找到评论数据")
                
                # 更新进度
                progress = int((page / self.page_num) * 100)
                self.progress_signal.emit(progress)
                
                # 暂停一下，避免请求过快
                time.sleep(1)
                
            except Exception as e:
                error_msg = str(e)
                self.update_signal.emit(f"爬取第 {page} 页评论时出错: {error_msg}")
                # 如果是API错误，显示更详细的信息
                if "API调用失败" in error_msg:
                    self.update_signal.emit(f"API错误详情: {error_msg}")
        
        if all_comments:
            self.update_signal.emit(f"爬取完成，共获取 {len(all_comments)} 条评论")
        else:
            self.update_signal.emit("爬取完成，但未获取到任何评论，请检查商品ID是否正确或者查看日志获取详细错误信息")
        
        self.finished_signal.emit(all_comments)

class SaveThread(QThread):
    """保存数据线程类"""
    update_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(self, comments, selected_fields, output_file):
        super().__init__()
        self.comments = comments
        self.selected_fields = selected_fields
        self.output_file = output_file
        self.crawler = TmallCommentCrawler()
        
    def run(self):
        try:
            self.update_signal.emit(f"正在将数据保存到 {self.output_file}...")
            
            # 筛选字段并保存
            import pandas as pd
            
            data = []
            for comment in self.comments:
                item = {}
                for field, field_name in self.selected_fields.items():
                    # 处理嵌套字段，如 interactInfo.likeCount
                    if '.' in field:
                        parts = field.split('.')
                        value = comment
                        for part in parts:
                            value = value.get(part, {}) if isinstance(value, dict) else {}
                        item[field_name] = value or ''
                    else:
                        item[field_name] = comment.get(field, '')
                data.append(item)
            
            df = pd.DataFrame(data)
            df.to_excel(self.output_file, index=False)
            
            self.update_signal.emit(f"数据已成功保存到 {self.output_file}")
            self.finished_signal.emit(True, self.output_file)
            
        except Exception as e:
            error_msg = f"保存数据时出错: {str(e)}"
            self.update_signal.emit(error_msg)
            self.finished_signal.emit(False, error_msg)

class TmallCommentCrawlerGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.comments = []  # 存储爬取的评论数据
        self.field_mappings = {}  # 存储字段映射
        self.setup_ui()
        
    def setup_ui(self):
        """设置UI界面"""
        self.setWindowTitle("天猫评论爬虫 - 高级版")
        self.setMinimumSize(900, 700)
        self.setStyleSheet(STYLE)
        
        # 创建主窗口部件
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)
        self.setCentralWidget(main_widget)
        
        # 创建标题标签
        title_layout = QHBoxLayout()
        title_label = QLabel("天猫商品评论爬虫工具")
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #6c5ce7;")
        title_layout.addWidget(title_label)
        title_layout.addStretch()
        main_layout.addLayout(title_layout)
        
        # 创建分隔线
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet("background-color: #dcdde1;")
        main_layout.addWidget(line)
        
        # 创建选项卡窗口
        tabs = QTabWidget()
        tabs.setFont(QFont("Microsoft YaHei", 9))
        
        # 创建爬虫设置选项卡
        crawler_tab = QWidget()
        crawler_layout = QVBoxLayout(crawler_tab)
        crawler_layout.setContentsMargins(10, 10, 10, 10)
        
        # 参数设置区域
        settings_group = QGroupBox("爬虫设置")
        settings_layout = QGridLayout(settings_group)
        settings_layout.setContentsMargins(15, 20, 15, 15)
        settings_layout.setSpacing(10)
        
        # 商品ID输入
        id_label = QLabel("商品ID:")
        id_label.setFont(QFont("Microsoft YaHei", 9))
        self.id_input = QLineEdit()
        self.id_input.setPlaceholderText("请输入天猫商品ID，例如：714871191114")
        self.id_input.setFont(QFont("Microsoft YaHei", 9))
        settings_layout.addWidget(id_label, 0, 0)
        settings_layout.addWidget(self.id_input, 0, 1, 1, 3)
        
        # Cookie输入
        cookie_label = QLabel("Cookie:")
        cookie_label.setFont(QFont("Microsoft YaHei", 9))
        self.cookie_input = QTextEdit()
        self.cookie_input.setPlaceholderText("必须提供天猫网站的Cookie，从浏览器开发者工具中复制，包含_m_h5_tk字段")
        self.cookie_input.setFont(QFont("Microsoft YaHei", 9))
        self.cookie_input.setMaximumHeight(80)
        settings_layout.addWidget(cookie_label, 1, 0)
        settings_layout.addWidget(self.cookie_input, 1, 1, 1, 3)
        
        # 页数设置
        page_label = QLabel("爬取页数:")
        page_label.setFont(QFont("Microsoft YaHei", 9))
        self.page_spin = QSpinBox()
        self.page_spin.setRange(1, 100)
        self.page_spin.setValue(5)
        self.page_spin.setFont(QFont("Microsoft YaHei", 9))
        self.page_spin.setToolTip("每页显示20条评论")
        settings_layout.addWidget(page_label, 2, 0)
        settings_layout.addWidget(self.page_spin, 2, 1)
        
        # 页数说明
        page_tip = QLabel("(每页20条评论)")
        page_tip.setFont(QFont("Microsoft YaHei", 8))
        page_tip.setStyleSheet("color: #7f8c8d;")
        settings_layout.addWidget(page_tip, 2, 2)
        
        # 开始爬取按钮
        self.start_btn = QPushButton("开始爬取")
        self.start_btn.setFont(QFont("Microsoft YaHei", 9, QFont.Bold))
        self.start_btn.clicked.connect(self.start_crawling)
        settings_layout.addWidget(self.start_btn, 2, 3)
        
        crawler_layout.addWidget(settings_group)
        
        # 进度条
        progress_group = QGroupBox("爬取进度")
        progress_layout = QVBoxLayout(progress_group)
        progress_layout.setContentsMargins(15, 20, 15, 15)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFont(QFont("Microsoft YaHei", 9))
        progress_layout.addWidget(self.progress_bar)
        
        crawler_layout.addWidget(progress_group)
        
        # 日志区域
        log_group = QGroupBox("运行日志")
        log_layout = QVBoxLayout(log_group)
        log_layout.setContentsMargins(15, 20, 15, 15)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        self.log_text.setStyleSheet("color: #2d3436;")
        log_layout.addWidget(self.log_text)
        
        crawler_layout.addWidget(log_group)
        
        # 添加爬虫选项卡
        tabs.addTab(crawler_tab, "评论爬取")
        
        # 创建数据导出选项卡
        export_tab = QWidget()
        export_layout = QVBoxLayout(export_tab)
        export_layout.setContentsMargins(10, 10, 10, 10)
        
        # 字段选择区域
        fields_group = QGroupBox("选择要导出的字段")
        fields_layout = QVBoxLayout(fields_group)
        fields_layout.setContentsMargins(15, 20, 15, 15)
        
        # 使用滚动区域显示字段复选框
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QGridLayout(scroll_content)
        scroll_layout.setContentsMargins(5, 5, 5, 5)
        scroll_layout.setSpacing(8)
        
        # 定义可选字段
        self.field_checkboxes = {}
        self.setup_field_checkboxes(scroll_layout)
        
        scroll_area.setWidget(scroll_content)
        fields_layout.addWidget(scroll_area)
        
        # 全选/取消全选按钮
        select_layout = QHBoxLayout()
        select_all_btn = QPushButton("全选")
        select_all_btn.clicked.connect(lambda: self.toggle_all_fields(True))
        
        select_none_btn = QPushButton("取消全选")
        select_none_btn.clicked.connect(lambda: self.toggle_all_fields(False))
        
        select_common_btn = QPushButton("选择常用字段")
        select_common_btn.clicked.connect(self.select_common_fields)
        
        select_layout.addWidget(select_all_btn)
        select_layout.addWidget(select_none_btn)
        select_layout.addWidget(select_common_btn)
        select_layout.addStretch()
        
        fields_layout.addLayout(select_layout)
        export_layout.addWidget(fields_group)
        
        # 导出设置区域
        export_settings_group = QGroupBox("导出设置")
        export_settings_layout = QGridLayout(export_settings_group)
        export_settings_layout.setContentsMargins(15, 20, 15, 15)
        
        # 导出路径设置
        export_path_label = QLabel("导出路径:")
        self.export_path_input = QLineEdit()
        self.export_path_input.setReadOnly(True)
        self.export_path_input.setPlaceholderText("点击浏览选择保存位置")
        
        browse_btn = QPushButton("浏览...")
        browse_btn.clicked.connect(self.browse_export_path)
        
        export_settings_layout.addWidget(export_path_label, 0, 0)
        export_settings_layout.addWidget(self.export_path_input, 0, 1)
        export_settings_layout.addWidget(browse_btn, 0, 2)
        
        # 导出按钮
        self.export_btn = QPushButton("导出到Excel")
        self.export_btn.setEnabled(False)  # 初始状态禁用，等爬取完成后启用
        self.export_btn.clicked.connect(self.export_to_excel)
        
        export_settings_layout.addWidget(self.export_btn, 1, 2)
        
        export_layout.addWidget(export_settings_group)
        
        # 添加导出选项卡
        tabs.addTab(export_tab, "字段选择与导出")
        
        main_layout.addWidget(tabs)
        
        # 状态栏
        self.statusBar().showMessage("就绪")
        
        # 初始化日志
        self.log("天猫评论爬虫工具已启动，请输入商品ID并设置爬取页数")
        self.log("请从浏览器复制最新的天猫Cookie并粘贴到Cookie输入框，爬取前必须提供Cookie")
        
    def setup_field_checkboxes(self, layout):
        """设置字段复选框"""
        # 定义字段映射 (接口字段名 -> 显示名称)
        field_mapping = {
            # 基本信息
            'userNick': '用户昵称',
            'feedback': '评论内容',
            'createTime': '评论时间',
            'createTimeInterval': '评论时间间隔',
            'feedbackDate': '评价日期',
            'id': '评论ID',
            'auctionNumId': '商品ID',
            'auctionTitle': '商品标题',
            'skuId': 'SKUID',
            
            # 商品规格
            'skuValueStr': '规格字符串',
            
            # 评价相关
            'rateType': '评价类型',
            'annoy': '是否匿名',
            'topRate': '是否置顶',
            'hasDetail': '是否有详情',
            'repeatBusiness': '是否复购',
            'goldUser': '是否金牌用户',
            'formalBlackUser': '是否黑名单用户',
            'copy': '是否复制',
            'own': '是否本人',
            'structTagEndSize': '结构标签结束大小',
            
            # 互动信息
            'interactInfo.likeCount': '点赞数',
            'interactInfo.commentCount': '评论数',
            'interactInfo.readCount': '阅读数',
            'interactInfo.alreadyLike': '是否已点赞',
            'interactInfo.enableComment': '可否评论',
            'interactInfo.enableLike': '可否点赞',
            'interactInfo.enableShare': '可否分享',
            
            # 商家回复
            'reply': '商家回复',
            
            # 用户信息
            'userId': '用户ID',
            'creditLevel': '用户信用等级',
            'userStar': '用户星级',
            'headPicUrl': '用户头像URL',
            'headFrameUrl': '用户头像框URL',
            'userIndexURL': '用户主页URL',
            'userMark': '用户标记',
            'reduceUserNick': '减少用户昵称',
            
            # 分享信息
            'share.shareURL': '分享URL',
            'share.detailUrl': '详情URL',
            'share.detailShareUrl': '详情分享URL',
            'share.shareSupport': '支持分享',
            
            # 添加购物车URL
            'addCartUrl': '添加购物车URL',
            
            # 权限信息
            'allowComment': '允许评论',
            'allowInteract': '允许互动',
            'allowNote': '允许笔记',
            'allowReportReview': '允许举报评论',
            'allowReportUser': '允许举报用户',
            'allowShieldReview': '允许屏蔽评论',
            'allowShieldUser': '允许屏蔽用户',
            
            # 额外信息
            'extraInfoMap.userGrade': '用户等级',
            'extraInfoMap.report_url': '举报URL',
        }
        
        self.field_mappings = field_mapping
        
        # 创建复选框并添加到布局
        row, col = 0, 0
        max_cols = 3
        
        for api_field, display_name in field_mapping.items():
            checkbox = QCheckBox(display_name)
            checkbox.setFont(QFont("Microsoft YaHei", 9))
            self.field_checkboxes[api_field] = checkbox
            
            layout.addWidget(checkbox, row, col)
            
            col += 1
            if col >= max_cols:
                col = 0
                row += 1
    
    def toggle_all_fields(self, state):
        """全选或取消全选所有字段"""
        for checkbox in self.field_checkboxes.values():
            checkbox.setChecked(state)
    
    def select_common_fields(self):
        """选择常用字段"""
        # 先取消全选
        self.toggle_all_fields(False)
        
        # 定义常用字段列表
        common_fields = [
            'userNick', 'feedback', 'createTime', 'feedbackDate',
            'auctionTitle', 'skuValueStr', 'rateType', 'reply',
            'userStar', 'interactInfo.likeCount', 'repeatBusiness'
        ]
        
        # 选中常用字段
        for field in common_fields:
            if field in self.field_checkboxes:
                self.field_checkboxes[field].setChecked(True)
    
    def browse_export_path(self):
        """浏览并选择导出路径"""
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(
            self, "选择保存位置", "", 
            "Excel文件 (*.xlsx);;所有文件 (*)", 
            options=options
        )
        
        if file_path:
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
            self.export_path_input.setText(file_path)
    
    def log(self, message):
        """添加日志信息"""
        timestamp = time.strftime("%H:%M:%S", time.localtime())
        log_msg = f"[{timestamp}] {message}"
        self.log_text.append(log_msg)
        # 滚动到底部
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def start_crawling(self):
        """开始爬取评论"""
        # 获取输入参数
        item_id = self.id_input.text().strip()
        page_num = self.page_spin.value()
        cookie = self.cookie_input.toPlainText().strip()
        
        # 检查参数有效性
        if not item_id:
            QMessageBox.warning(self, "参数错误", "请输入有效的商品ID")
            return
        
        # 检查是否提供了Cookie
        if not cookie:
            QMessageBox.warning(self, "参数错误", "请提供有效的Cookie")
            return
            
        # 禁用开始按钮，避免重复点击
        self.start_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.log(f"准备爬取商品ID: {item_id}，共 {page_num} 页")
        
        self.log("使用自定义Cookie进行爬取")
        
        # 创建并启动爬虫线程
        self.crawler_thread = CrawlerThread(item_id, page_num, cookie)
        self.crawler_thread.update_signal.connect(self.log)
        self.crawler_thread.progress_signal.connect(self.progress_bar.setValue)
        self.crawler_thread.finished_signal.connect(self.on_crawl_finished)
        self.crawler_thread.start()
    
    def on_crawl_finished(self, comments):
        """爬取完成后的处理"""
        self.comments = comments
        self.start_btn.setEnabled(True)
        
        if comments:
            self.export_btn.setEnabled(True)
            comment_count = len(comments)
            self.statusBar().showMessage(f"爬取完成，共获取 {comment_count} 条评论数据")
            
            # 自动生成默认文件名
            if comments:
                item_id = comments[0].get('auctionNumId', '')
                item_title = comments[0].get('auctionTitle', '')
                if item_title:
                    import re
                    item_title = re.sub(r'[\\/:*?"<>|]', '', item_title)
                    if len(item_title) > 30:
                        item_title = item_title[:30] + '...'
                
                current_date = time.strftime("%Y%m%d", time.localtime())
                default_filename = f"{item_id}_{item_title}_{len(comments)}条评论_{current_date}.xlsx"
                
                # 获取应用程序所在目录
                if getattr(sys, 'frozen', False):
                    # 如果是打包后的exe，使用可执行文件所在目录
                    app_dir = os.path.dirname(sys.executable)
                else:
                    # 否则使用脚本所在目录
                    app_dir = os.path.dirname(os.path.abspath(__file__))
                
                default_path = os.path.join(app_dir, default_filename)
                self.export_path_input.setText(default_path)
        else:
            self.statusBar().showMessage("爬取完成，但未获取到评论数据")
    
    def export_to_excel(self):
        """导出评论数据到Excel"""
        if not self.comments:
            QMessageBox.warning(self, "导出错误", "没有可导出的评论数据")
            return
        
        # 获取导出路径
        output_file = self.export_path_input.text().strip()
        if not output_file:
            QMessageBox.warning(self, "导出错误", "请选择保存文件路径")
            return
        
        # 获取选中的字段
        selected_fields = {}
        for api_field, checkbox in self.field_checkboxes.items():
            if checkbox.isChecked():
                selected_fields[api_field] = self.field_mappings[api_field]
        
        if not selected_fields:
            QMessageBox.warning(self, "导出错误", "请至少选择一个要导出的字段")
            return
        
        # 禁用导出按钮
        self.export_btn.setEnabled(False)
        
        # 创建并启动保存线程
        self.save_thread = SaveThread(self.comments, selected_fields, output_file)
        self.save_thread.update_signal.connect(self.log)
        self.save_thread.finished_signal.connect(self.on_save_finished)
        self.save_thread.start()
    
    def on_save_finished(self, success, message):
        """保存完成后的处理"""
        self.export_btn.setEnabled(True)
        
        if success:
            QMessageBox.information(self, "导出成功", f"数据已成功导出到:\n{message}")
            self.statusBar().showMessage(f"数据已成功导出")
        else:
            QMessageBox.critical(self, "导出失败", f"导出数据时出错:\n{message}")
            self.statusBar().showMessage("导出失败")

def main():
    app = QApplication(sys.argv)
    window = TmallCommentCrawlerGUI()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main() 