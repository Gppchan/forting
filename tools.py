import copy
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from pathlib import Path

from openpyxl.cell.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

import arrow

type ExCell = Cell | tuple[Cell] | tuple[tuple[Cell]]


def is_cell(excell: ExCell) -> bool:
    return isinstance(excell, Cell)


def is_tuple_cell(excell: ExCell) -> bool:
    return isinstance(excell, tuple) and isinstance(excell[0], Cell)


def is_tuple_tuple_cell(excell: ExCell) -> bool:
    return isinstance(excell, tuple) and isinstance(excell[0], tuple) and isinstance(excell[0][0], Cell)


def to_row(index: int) -> int:
    return index + 1


def to_column(index: int) -> str:
    return chr(ord("A") + index)


def copy_cell(src: ExCell, tar: ExCell | Worksheet) -> bool:
    if isinstance(tar, Worksheet):
        if is_cell(src):
            tar_cell = tar[src.coordinate]
            tar_cell.value = src.value
            tar_cell._style = copy.deepcopy(src._style)
        elif is_tuple_cell(src):
            for _src_cell in src:
                tar_cell = tar[_src_cell.coordinate]
                tar_cell.value = _src_cell.value
                tar_cell._style = copy.deepcopy(_src_cell._style)
        elif is_tuple_tuple_cell(src):
            for _src_column in src:
                for __src_cell in _src_column:
                    tar_cell = tar[__src_cell.coordinate]
                    tar_cell.value = __src_cell.value
                    tar_cell._style = copy.deepcopy(__src_cell._style)
    elif is_cell(src) and is_cell(tar):
        tar.value = src.value
        tar._style = copy.deepcopy(src._style)
        return True
    elif is_tuple_cell(src) and is_tuple_cell(tar) and len(src) == len(tar):
        for _src_cell, _tar_cell in zip(src, tar):
            _tar_cell.value = _src_cell.value
            _tar_cell._style = copy.deepcopy(_src_cell._style)
        return True
    elif is_tuple_tuple_cell(src) and is_tuple_tuple_cell(tar) and len(src) == len(tar) and len(src[0]) == len(tar[0]):
        for _src_column, _tar_column in zip(src, tar):
            for __src_cell, __tar_cell in zip(_src_column, _tar_column):
                __tar_cell.value = __src_cell.value
                __tar_cell._style = copy.deepcopy(__src_cell._style)
        return True
    else:
        return False


def get_value(src: ExCell, axis: int):
    if is_cell(src):
        return src.value
    elif is_tuple_cell(src):
        return [i.value for i in src]
    elif is_tuple_tuple_cell(src):
        values: list[list] = []
        column_num = len(src)
        row_num = len(src[0])
        if axis:
            for i in range(row_num):
                row_values: list = []
                for j in range(column_num):
                    row_values.append(src[j][i].value)
                values.append(row_values)
        else:
            for i in range(column_num):
                column_values: list = []
                for j in range(row_num):
                    column_values.append(src[i][j].value)
                values.append(column_values)
        return values


def get_file_path() -> Path:
    window = tk.Tk()
    window.withdraw()
    path: Path = Path(filedialog.askopenfilenames()[0])
    if path.suffix not in [".xlsx"]:
        messagebox.showerror("forting", "小宝，这个文件不是表格哦")
    return path


def timer(func):
    def wrapper(*arg, **kwargs):
        start: float = arrow.now().timestamp()
        func(*arg, **kwargs)
        end: float = arrow.now().timestamp()
        print(f"运行时间：{end - start: .3f}s")

    return wrapper
