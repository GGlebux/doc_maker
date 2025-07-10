import logging
import sys
import traceback

from PyQt6.QtWidgets import QMessageBox, QCheckBox, QLabel, QPushButton, QMainWindow


def setup_logging():
    """Настройка логгера"""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        datefmt='%H:%M:%S %d-%m-%Y',
        handlers=[
            logging.FileHandler("app.log", encoding='utf-8'),  # Логи в файл
            logging.StreamHandler(sys.stdout),  # Логи в консоль
        ]
    )


def simple_toggle(check_box: QCheckBox, label: QLabel, button: QPushButton, set_flag_func):
    """Статический переключатель для активации виджетов"""
    if check_box.isChecked():
        label.setEnabled(True)
        button.setEnabled(True)
        set_flag_func(True)
    else:
        label.setEnabled(False)
        button.setEnabled(False)
        set_flag_func(False)


def log_exception(parent: QMainWindow, exception: Exception, err_place: str):
    error_info = "".join(traceback.format_exception(exception))
    parent.logger.error(f'Возникла непредвиденная ошибка при {err_place}:\n{error_info}')
    QMessageBox.warning(parent, 'Критическая ошибка',
                        f'Возникла непредвиденная ошибка при {err_place}:\n(обратитесь к разработчику)')
