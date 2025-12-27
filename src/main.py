"""应用入口"""
import sys

from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt

from .gui.main_window import MainWindow


def main():
    """主函数"""
    # 启用高 DPI 支持
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    app = QApplication(sys.argv)
    app.setApplicationName("批量文档裁剪工具")
    app.setApplicationVersion("1.0.0")
    
    # 设置样式
    app.setStyle("Fusion")
    
    # 创建主窗口
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
