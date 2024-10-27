from cx_Freeze import setup, Executable

setup(
    name="приЁмка",
    version="2.0",
    description="Заполнит документы без рутины",
    executables=[Executable("python_files/main.py",
                            target_name="приЁмка.AppImage",
                            icon="icon.ico")],  # base="Win32GUI"
    packages=["PyQt6", "sys", "os", "docxtpl", "openpyxl"],  # Добавьте все необходимые пакеты
)
