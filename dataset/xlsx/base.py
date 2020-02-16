# coding: utf-8
from abc import ABCMeta, abstractmethod

from ._static_data import (
    HUE_DICT,
    ALPHABET_NUM,
)


class XlsxIOBase(metaclass=ABCMeta):
    def get_sheetnames(self):
        """
        シート名を返す関数
        
        Returns
        -------
        list of str
            シート名のリスト
        """
        return self.wb.sheetnames

    def _convert_column_str(self, col_str):
        """
        列の英語を番号に変換して返す関数
        
        Parameters
        ----------
        col_str : str
            列の英語
        
        Returns
        -------
        int
            列番号
        """
        for i, s in enumerate(col_str[::-1]):
            if i == 0:
                col = ord(s) - ord("@")  # ascii_codeでAのひとつ前
            else:
                col += (ord(s) - ord("@")) * (ALPHABET_NUM * i) # ascii_codeでAのひとつ前
        return col

    def _convert_column_num(self, col_num):
        """
        列番号を英語に変換して返す関数
        
        Parameters
        ----------
        col_num : int
            列番号
        
        Returns
        -------
        str
            列の英語
        """
        col_str = ""
        _num = col_num
        if col_num > ALPHABET_NUM:
            _num = col_num % ALPHABET_NUM
            div_num = col_num // ALPHABET_NUM
            if _num == 0:
                _num = ALPHABET_NUM
                div_num -= 1
            col_str += self._convert_column_num(div_num)
        return col_str + chr(_num + ord("@"))

    def _convert_cell_str(self, cell_str):
        """
        セルの座標値(A1等)を行・列で別々に数値にして返す関数
        
        Parameters
        ----------
        cell_str : str
            セルの座標値
        
        Returns
        -------
        int
            行番号
        int
            列番号
        """
        for i, s in enumerate(cell_str):
            if s.isdigit():
                div_num = i
                break
        row = int(cell_str[div_num:])
        col = self._convert_column_str(cell_str[:div_num])
        return row, col

    def _convert_cell_num(self, row_num, col_num):
        """
        セルの行・列の数値を座標値(A1等)にして返す関数
        
        Parameters
        ----------
        row_num : int
            行番号
        col_num : int
            列番号
        
        Returns
        -------
        str
            セルの座標値
        """
        return "{}{}".format(
            self._convert_column_num(col_num), row_num,
        )

    def _convert_cm2width(self, cm):
        return (cm / 0.2)

    def _get_cell_widths(self, sheet, start_col, end_col):
        """
        複数セル行の横幅を取得する関数(取得横幅は終了行番号も含む)
        
        Parameters
        ----------
        sheet : openpyxl.Sheet
            エクセルのシート名
        start_col : int
            開始行番号
        end_col : int
            終了行番号
        
        Returns
        -------
        dict of {int: int}
            行番号をキー、セル幅を値とする辞書
        """
        cell_width_dict = {}
        bef_width = -1
        col_num = 1
        while col_num <= end_col:
            col_str = self._convert_column_num(col_num)
            width = sheet.column_dimensions[col_str].width
            if width is None:
                width = bef_width
            else:
                bef_width = width
            if start_col <= col_num and col_num <= end_col:
                cell_width_dict[col_num] = width
            col_num += 1
        return cell_width_dict

    @staticmethod
    def judge_color(hue):
        """
        色相から大まかにどの色かを判定する関数
        
        Parameters
        ----------
        hue : float
            色相(hue)の値(0 ~ 1)
        
        Returns
        -------
        str or None
            背景色("red", "yellow", "green", "cyan", "blue", "purple", Noneのいずれか)
        """
        _hue = int(hue * 360)  # 0 ~ 360に変化
        if _hue > 330:
            _hue -= 360  # RED用に330以上の場合はマイナス表記にする
        for color, (lower, upper) in HUE_DICT.items():
            if lower < _hue and _hue < upper:
                return color
        return None
