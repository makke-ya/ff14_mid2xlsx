# coding: utf-8
from abc import ABCMeta, abstractmethod
from math import ceil
from copy import deepcopy
from ._static_data import (
    DICT_FOR_PROGRAM_onlyFF14,
)


class MidiIOBase(metaclass=ABCMeta):
    def __init__(
        self, type=1, tempo=60, ticks_per_beat=480, rhythm_dict=None
    ):
        self.mid = None
        self.ticks_per_beat = ticks_per_beat
        self.tempo = tempo
        self.time_in_measures = {}
        self.rhythm_dict = rhythm_dict
        self.use_channel_list = [False] * 16

    def fwrite(self, filename):
        """
        ファイル出力する関数
        
        Parameters
        ----------
        filename : str
            出力ファイル名
        """
        # for track in self.mid.tracks:
        #     for msg in track:
        #         print(msg)
        self.mid.save(filename)

    def pprint(self):
        for i, track in enumerate(self.mid.tracks):
            print("Track {}: {}".format(i, track.name))
            for msg in track:
                print(msg)

    def get_program_names(self):
        """
        楽器リスト（プログラム名）を返す関数
        
        Returns
        -------
        list of tuple of str
            プログラムの候補名のタプルのリスト
        """
        return list(DICT_FOR_PROGRAM_onlyFF14.keys())

    def convert_program_str2num(self, program_str):
        """
        FF14の楽器種類名（もしくはmidiプログラム名）をmidiのプログラム番号に変える関数
        
        Parameters
        ----------
        program_str : str
            FF14の楽器種類名、もしくはmidiプログラム名
        
        Returns
        -------
        int
            midi形式のプログラム番号
        int
            FF14での C のピッチ番号
        """
        for keys, (program_num, base_pitch) in DICT_FOR_PROGRAM_onlyFF14.items():
            if program_str in keys:
                return program_num, base_pitch
        return None, None

    def convert_program_num2str(self, program_num):
        """
        midiのプログラム番号を楽器種類（プログラム名）に変える関数
        
        Parameters
        ----------
        program_num : int
            midi形式のプログラム番号
        
        Returns
        -------
        int
            midi形式のプログラム番号
        int
            FF14での C のピッチ番号
        """
        for keys, (_program_num, base_pitch) in DICT_FOR_PROGRAM_onlyFF14.items():
            if program_num == _program_num:
                return keys[-1], base_pitch
        return "ハープ", 60

    def _get_time_in_measure(self, numerator, denominator):
        """
        拍子から1小節の時間を割り出す関数
        
        Parameters
        ----------
        numerator : int
            拍数
        denominator : int
            音価
        
        Returns
        -------
        int
            1小節の時間
        """
        return int(((self.ticks_per_beat * 4) // denominator) * numerator)
