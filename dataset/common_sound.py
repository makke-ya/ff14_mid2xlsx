# coding: utf-8
from copy import deepcopy
from .midi._static_data import (
    DICT_FOR_PROGRAM_onlyFF14,
    DICT_FOR_PITCH_CONVERT,
)
from ._static_data import(
    KEY_NAME_LIST,
    KEY_PITCH_DICT,
    KEY_NOTE_NAME_DICT,
    DICT_FOR_OCTAVE,
    DICT_FOR_NOTE_CONVERT,
)


class CommonSoundData(object):
    def __init__(self):
        self.player_idx = None
        self.note_list = None
        self.pitch_list = None
        self.rate_list = None
        self.channel_num = None
        self.flatten_pitches = None
        self.key_dict = None
        self.shift_pitch = 0
        self.rhythm_dict = None
        self.program_str = None

    def add_pitch_list(self, pitch_list, delete_unnecessary_mark=True):
        self.pitch_list = pitch_list
        if delete_unnecessary_mark is True:
            self.pitch_list = self._delete_unnecessary_mark(self.pitch_list)

    def add_channel_num(self, channel_num):
        self.channel_num = channel_num
    
    def add_program_str(self, program_str):
        self.program_str = program_str
    
    def add_rhythm_dict(self, rhythm_dict):
        if self.rhythm_dict is None:
            self.rhythm_dict = {}
        for measure_num, rhythm in rhythm_dict.items():
            self.rhythm_dict[measure_num] = rhythm

    def set_player_idx(self, idx):
        self.player_idx = idx

    def set_key(self, key):
        self.key_dict = key

    def set_shift_pitch(self, shift_pitch):
        self.shift_pitch = shift_pitch

    def set_recommended_octave(self):
        _, base_pitch = self.convert_program_str2num(self.program_str)
        if base_pitch is None:
            raise ValueError

        if self.flatten_pitches is None:
            self._flatten()
        flatten_pitches = self.flatten_pitches
        if len(flatten_pitches) == 0:
            return False
        max_pitch = max(flatten_pitches)
        min_pitch = min(flatten_pitches)

        if (base_pitch - 12) <= min_pitch and max_pitch <= (base_pitch + 24):
            return True

        avg_pitch = sum(flatten_pitches) / len(flatten_pitches)
        recommended_octave = int(avg_pitch / 12)  # 切り捨て
        recommended_base_pitch = recommended_octave * 12
        self.shift_pitch = base_pitch - recommended_base_pitch
        return True

    def get_flatten_pitches(self):
        if self.flatten_pitches is None:
            self._flatten()
        return self.flatten_pitches

    def get_note_list(self):
        return deepcopy(self.note_list)

    def get_key_dict(self):
        return self.key_dict

    def get_pitch_list(self):
        return self.pitch_list

    def get_rate_list(self):
        return deepcopy(self.rate_list)

    def get_shift_pitch(self):
        return self.shift_pitch

    def get_player_idx(self):
        return self.player_idx

    def get_program_str(self):
        return self.program_str

    def get_channel_num(self):
        return self.channel_num

    def get_rhythm_dict(self):
        return self.rhythm_dict

    def get_difficulty(self, tempo, bias=.3):
        if self.flatten_pitches is None:
            self._flatten()

        # とりあえずnoteの多さだけで計算
        note_list = self.note_list if self.note_list is not None else self.pitch_list
        sum_beats = 0
        max_note_num = 0
        for notes_in_measure in note_list:
            all_rest = True
            note_num = 0
            for notes_in_beat in notes_in_measure:
                for notes in notes_in_beat:
                    if notes != "r":
                        all_rest = False
                        if notes != "-":
                            note_num += len(notes)
            if all_rest:
                continue
            sum_beats += len(notes_in_measure)
            if note_num > max_note_num:
                max_note_num = note_num
                max_notes = notes_in_measure

        flatten_note_num = len(self.flatten_pitches)
        average_note_num = (flatten_note_num / sum_beats)
        tempo_difficulty = tempo / 80.0

        # (平均の一小節の音数 + 最大の一小節の音数) * テンポの高さ * バイアス
        difficulty = (
            (
                (average_note_num * tempo_difficulty * 0.5) +
                (max_note_num * tempo_difficulty * 2.0)
            ) * bias
        )
        return "{:.1f}".format(difficulty)

    def _create_dummy_rates(self):
        """
        横幅による1拍ごとのレート（長さ）を固定間隔で登録する関数
        """
        all_rates = []
        note_list = self.note_list if self.note_list is not None else self.pitch_list
        for notes_in_measure in note_list:
            rates_in_measure = [1 for notes_in_beat in notes_in_measure]
            rates_in_beats = [
                [1 for note_in_cell in notes_in_beat]
                for notes_in_beat in notes_in_measure
            ]
            all_rates.append((rates_in_measure, rates_in_beats))
        self.rate_list = all_rates

    def _delete_unnecessary_mark(self, note_list):
        rev_note_list = deepcopy(note_list)
        for num_measure, notes_in_measure in enumerate(note_list):
            for num_beat, notes_in_beat in enumerate(notes_in_measure):
                bef_note = None
                bef_note_length = None
                delete_flag = True
                for notes in notes_in_beat:
                    if bef_note is None:
                        bef_note = notes
                        note_length = 1
                    elif notes == "-" and (bef_note == "-" or type(bef_note) is list):
                        note_length += 1
                    elif notes == "r" and bef_note == "r":
                        note_length += 1
                    else:
                        if bef_note_length is not None and bef_note_length != note_length:
                            delete_flag = False
                            break
                        bef_note = notes
                        bef_note_length = note_length
                        note_length = 1
                if bef_note_length is not None and bef_note_length != note_length:
                    delete_flag = False
                if delete_flag is True:
                    rev_note_list[num_measure][num_beat] = notes_in_beat[::note_length]
        return rev_note_list

    def _note2pitch(self, note, base_pitch):
        """
        共通音リストの音文字列をピッチ（数値）に変換する関数
        
        Parameters
        ----------
        note : str or list
            音文字列 or 音文字列のリスト(和音)
        base_pitch : int
            FF14での C のピッチ番号
        
        Returns
        -------
        int
            音のピッチ
        """
        split_note = note.split("_")
        pitch = DICT_FOR_PITCH_CONVERT[split_note[0]] + base_pitch
        if len(split_note) > 1:
            if split_note[1] == "minus1":
                pitch -= 12
            elif split_note[1] == "plus1":
                pitch += 12
            elif split_note[1] == "plus2":
                pitch += 24
        return pitch

    def _flatten(self):
        flatten_pitches = [
            pitches
            for pitches_in_measure in self.pitch_list
            for pitches_in_beat in pitches_in_measure
            for pitches in pitches_in_beat if type(pitches) is not str
        ]
        self.flatten_pitches = sum(flatten_pitches, [])

    def _pitch2note(self, pitch, base_pitch, key_pitches, key_notes, bef_pitch=None):
        """
        共通音リストのピッチ（数値）を音文字列に変換する関数
        
        Parameters
        ----------
        pitch: int
            音のピッチ
        base_pitch : int
            FF14での C のピッチ番号
        key_pitches : list of int
            調の音階の高さ
        key_notes : list of str
            調の音階名
        bef_pitch : int or None
            前の音
        
        Returns
        -------
        note : str
            音文字列
        """
        _pitch = (pitch + self.shift_pitch) % 12
        octave = (pitch - base_pitch + self.shift_pitch + 12) // 12
        if (octave < 0 or octave > 3 or (octave == 3 and _pitch != 0)):
            return None
        octave_str = DICT_FOR_OCTAVE[octave]
        print(key_pitches, pitch, _pitch, octave)
        try:
            key_note = key_notes[key_pitches.index(_pitch)]
        except ValueError:
            key_note = None
        print(pitch, bef_pitch)
        if key_note is None:
            if (bef_pitch is None or bef_pitch <= pitch):
                key_note = DICT_FOR_NOTE_CONVERT[_pitch][0]  # 上がり傾向の時は＃
            elif bef_pitch > pitch:
                key_note = DICT_FOR_NOTE_CONVERT[_pitch][1]  # 下がり傾向の時は♭
        return key_note + octave_str

    def convert_pitch2note(self, enable_chord=False):
        if self.key_dict is None:
            self.estimate_key(self.pitch_list)
        key_pitches = None
        key_notes = None
        print(self.key_dict)

        _, base_pitch = self.convert_program_str2num(self.program_str)
        if base_pitch is None:
            raise ValueError

        note_list = []
        bef_pitch = None
        max_pitch = 0
        min_pitch = 127
        outlier_list = []
        for num_measure, pitches_in_measure in enumerate(self.pitch_list, start=1):
            notes_in_measure = []
            if num_measure in self.key_dict.keys():
                key_pitches = KEY_PITCH_DICT[self.key_dict[num_measure]]
                key_notes = KEY_NOTE_NAME_DICT[self.key_dict[num_measure]]
            for num_beat, pitches_in_beat in enumerate(pitches_in_measure):
                notes_in_beat = []
                for pitches in pitches_in_beat:
                    if type(pitches) is str:
                        notes = pitches
                    else:
                        _max_pitch = max(pitches)
                        _min_pitch = min(pitches)
                        if enable_chord is False:
                            notes = [self._pitch2note(
                                _max_pitch, base_pitch, key_pitches, key_notes, bef_pitch
                            )]
                        else:
                            notes = [
                                self._pitch2note(
                                    pitch, base_pitch, key_pitches, key_notes, bef_pitch
                                ) for pitch in sorted(pitches)
                            ]
                        bef_pitch = _max_pitch
                        if max_pitch < _max_pitch:
                            max_pitch = _max_pitch 
                        if min_pitch > _min_pitch:
                            min_pitch = _min_pitch 
                        if None in notes:
                            outlier_list.append([num_measure, num_beat, pitches])
                            notes = [note for note in notes if note is not None]
                            if len(notes) == 0:
                                notes = "r"
                    notes_in_beat.append(notes)
                notes_in_measure.append(notes_in_beat)
            note_list.append(notes_in_measure)
        self.note_list = note_list
        return max_pitch, min_pitch, outlier_list

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

    def estimate_key(self, pitch_list):
        pitch_list = [pitch % 12 for pitch in pitch_list]

        # dur/majorコードでまずは推定
        max_match_num = 0
        match_key_name = None
        for key_name, pitches in KEY_PITCH_DICT.items():
            if "moll" in key_name or key_name == "FF14":
                continue
            match_num = 0
            for pitch in pitches:
                match_num += pitch_list.count(pitch)
            if max_match_num < match_num:
                max_match_num = match_num
                match_key_name = key_name

        # dur/majorかmoll/minorか判定
        pitches = KEY_PITCH_DICT[match_key_name]
        dur_picth_num = pitch_list.count(pitches[4])
        moll_picth_num = pitch_list.count(pitches[4] + 1)
        if moll_picth_num < dur_picth_num:
            idx = (KEY_NAME_LIST.index(match_key_name) + 1) % len(KEY_NAME_LIST)
            match_key_name = KEY_NAME_LIST[idx]
        self.key_dict = {1: match_key_name}
        return self.key_dict
