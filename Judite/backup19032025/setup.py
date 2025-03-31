from cx_Freeze import setup, Executable

setup(
    name="Importador de Extratos Bancários",
    version="1.0",
    description="Importador de extratos bancários por meio de .xls e .xlsx, com classificação e exportação.",
    executables=[Executable("principal.py", base="Win32GUI")],
)
