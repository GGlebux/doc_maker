import json

import httpx
from PyQt6.QtWidgets import (QComboBox)

from python_files.conf import API_URL, API_HEADERS
from python_files.static import log_exception


class AddressHelper:
    def __init__(self, parent):
        super().__init__()
        self.parent = parent

        self.client = httpx.Client()

        self.last_request: str = ''

    def load_hints(self, address: QComboBox):
        """Обновляем подсказки адресов c задержкой"""
        try:
            self.last_request = address.currentText()
            if len(self.last_request) > 3:
                self.update_suggestions(self.last_request, address)
        except Exception as e:
            log_exception(self.parent, e, f"попытке загрузить подсказки для адреса-'{self.last_request}'",
                          do_waring=False)

    def update_suggestions(self, request: str, address: QComboBox):
        """Обновление списка подсказок"""
        suggestions = self.get_suggestions(request)
        if not suggestions:
            return

        # Сохраняем текущее состояние
        current_text = address.currentText()
        cursor_pos = address.lineEdit().cursorPosition()

        address.blockSignals(True)

        # Обновляем список
        address.clear()
        address.addItems(suggestions)
        address.setCurrentText(current_text)
        address.lineEdit().setCursorPosition(cursor_pos)
        address.showPopup()

        address.blockSignals(False)

    def get_suggestions(self, query: str):
        """Делаем запрос для получения валидных адресов"""

        data = json.dumps({"query": query}).encode("utf-8")
        response: list[dict[str, str]] = self.client.post(API_URL,
                                                          json={'query': query},
                                                          headers=API_HEADERS).json()['suggestions']
        suggestions: list[str] = [e['value'] for e in response]
        return suggestions
