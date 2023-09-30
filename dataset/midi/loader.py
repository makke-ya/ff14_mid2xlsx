# coding: utf-8
import os
from math import ceil
from copy import deepcopy
from mido import Message, MetaMessage, MidiFile, MidiTrack, bpm2tempo, tempo2bpm
from .base import MidiIOBase
from .reviser import MidiReviser
from dataset.common_sound import CommonSoundData
from ._static_data import (
    DICT_FOR_QUANTIZED_UNIT_TIMES,
    DICT_FOR_DRUM,
)


class MidiLoader(MidiIOBase):
    """
    SMF(Standard Midi Files)の読み込み用データとヘルパー関数
    """
    def __init__(self, filename=None):
        """
        Parameters
        ----------
        filename : str, optional
            ファイル名, by default None
        """
        self.reviser = MidiReviser(filename)
        print("output:", self.reviser.is_revised)
        if self.reviser.is_revised:
            base, ext = os.path.splitext(filename)
            filename = base + "_revised" + ext
            self.reviser.fwrite_revised_midi(filename)
        self.filename = filename
        self.mid = MidiFile(filename=filename, debug=True)
        self.ticks_per_beat = self.mid.ticks_per_beat
        self.tempo = self.get_tempo()
        if self.tempo is None:  # テンポ設定が無い場合は60
            self.tempo = 60
        self.time_in_measures, self.rhythm_dict = self._get_measures_and_rhythms()
        self.use_channel_list = self._get_use_channel_list()
        self.quantized_unit_times = self._get_quantized_unit_times()

    def get_common_data_list(
        self, quantize_num=64, drum_modes=[
            "バスドラム", "スネアドラム", "シンバル",
            # "バスドラム+スネアドラム", "スネアドラム+シンバル", "バスドラム+シンバル", "all"
        ]
    ):
        """
        共通音楽データ(CommonSoundData)のリストを取得する関数
        
        Parameters
        ----------
        quantize_num : int (16, 32, 64のいずれか)
            量子化の数（未実装）
        drum_modes : list of str
            出力するドラムモード
            "バスドラム", "スネアドラム", "シンバル", "バスドラム+スネアドラム" のいずれか

        Returns
        -------
        list of CommonSoundData
            共通音楽データリスト
        """
        common_data_list = []
        for i, track in enumerate(self.mid.tracks):
            print(i, track.name)
            channel_num = self._get_channel_num(track)
            if channel_num is None:
                continue
            program_num = self._get_program_num(track)
            program_str, _ = self.convert_program_num2str(program_num)
            measure_pitches = self._divide_track_into_measures(track, channel_num)
            print("nonadj", measure_pitches)
            adj_measure_pitches = self._adjust_measures_time(measure_pitches)
            print("adj", adj_measure_pitches)
            pitch_list = self._convert_measures_to_pitch_list(adj_measure_pitches, channel_num)
            print("pitch")
            for n in pitch_list:
                print(n)
            if channel_num != 9:  # not drum
                common_data = CommonSoundData()
                common_data.add_pitch_list(pitch_list, delete_unnecessary_mark=True)
                common_data._create_dummy_rates()
                common_data.add_channel_num(channel_num)
                common_data.add_program_str(program_str)
                common_data.add_rhythm_dict(self.rhythm_dict)
                common_data_list.append(common_data)
            else:  # drum
                common_data_list.extend(
                    self._get_drum_common_datas(pitch_list, drum_modes)
                )
        for i, common_data in enumerate(common_data_list, start=1):
            common_data.set_player_idx(i)
        return common_data_list

    def _get_drum_common_datas(
        self, pitch_list, drum_modes=[
            "バスドラム", "スネアドラム", "シンバル",
            # "バスドラム+スネアドラム", "スネアドラム+シンバル", "バスドラム+シンバル", "all"
        ]
    ):
        """
        Ch10（ドラム）の共通音楽データ(CommonSoundData)を取得する関数

        Parameters
        ----------
        pitch_list : 
            Ch10のピッチリスト
        drum_modes : list of str
            出力するドラムモード
            "バスドラム", "スネアドラム", "シンバル", "バスドラム+スネアドラム" のいずれか
        
        Returns
        -------
        list of CommonSoundData
            共通音楽データリスト
        """
        channel_num = 9
        drum_common_datas = []
        for mode in drum_modes:
            pitch_conversion_dict = DICT_FOR_DRUM[mode]
            new_pitch_list = deepcopy(pitch_list)
            for num_measure, pitches_in_measure in enumerate(pitch_list):
                for num_beat, pitches_in_beat in enumerate(pitches_in_measure):
                    for num_sound, pitches in enumerate(pitches_in_beat):
                        new_pitches = new_pitch_list[num_measure][num_beat][num_sound]
                        if pitches == "r":
                            continue
                        else:
                            pop_nums = []
                            for num_pitch, pitch in enumerate(sorted(pitches)):
                                if pitch in pitch_conversion_dict.keys():
                                    new_pitch = pitch_conversion_dict[pitch]
                                    new_pitches[num_pitch] = new_pitch
                                else:
                                    pop_nums.append(num_pitch)
                            for num_pitch in pop_nums[::-1]:
                                new_pitches.pop(num_pitch)
                            if len(new_pitches) == 0:
                                new_pitch_list[num_measure][num_beat][num_sound] = "r"
            
            common_data = CommonSoundData()
            common_data.add_pitch_list(new_pitch_list, delete_unnecessary_mark=True)
            common_data._create_dummy_rates()
            common_data.add_channel_num(channel_num)
            common_data.add_program_str(mode)
            common_data.add_rhythm_dict(self.rhythm_dict)
            drum_common_datas.append(common_data)
        return drum_common_datas

    def _divide_track_into_measures(self, track, channel_num):
        """
        SMFの1トラックを1小節ごとのmidi命令(note_on, note_off)に変換する関数
        
        Parameters
        ----------
        track : mido.Track
            SMFのトラック
        channel_num : int
            SMFのチャネル番号
        
        Returns
        -------
        measure_pitches : list of dict {tick: [note_on, note_off]}
            小節ごとに区切ったデータ、キーは小節内の時間、
            値はその時の[note_onになった音の高さ, note_offになった時の〃]
        """
        now_time = 0
        now_measure = 1
        time_in_measure = self.time_in_measures[now_measure]
        print(time_in_measure)
        bef_measure_time = 0
        next_measure_time = time_in_measure

        measure_pitches = []
        pitch_dict = {}
        for msg in track:
            now_time += msg.time
            while next_measure_time <= now_time:
                now_measure += 1
                if now_measure in self.time_in_measures.keys():
                    time_in_measure = self.time_in_measures[now_measure]
                bef_measure_time = next_measure_time
                next_measure_time += time_in_measure
                measure_pitches.append(deepcopy(pitch_dict))
                pitch_dict = {}
            if msg.type == "note_on" or (msg.type == "note_off" and channel_num != 9):
                _time = now_time - bef_measure_time
                if _time not in pitch_dict.keys():
                    pitch_dict[_time] = [[], []]
                if msg.type == "note_on":
                    if msg.velocity > 0:
                        pitch_dict[_time][0].append(msg.note)
                    else:
                        pitch_dict[_time][1].append(msg.note)
                elif msg.type == "note_off":
                    pitch_dict[_time][1].append(msg.note)
        if len(pitch_dict) != 0:
            measure_pitches.append(pitch_dict)
        return measure_pitches

    def _adjust_measures_time(self, measure_pitches):
        """
        1小節ごとのmidi命令(note_on, note_off)の時間を量子化（微調整）する関数
        
        Parameters
        ----------
        measure_pitches : list of dict {tick: [note_on, note_off]}
            小節ごとに区切ったデータ、キーは小節内の時間、
            値はその時の[note_onになった音の高さ, note_offになった時の〃]
        
        Returns
        -------
        rev_measure_pitches : list of dict {tick: [note_on, note_off]}
            量子化によって調整された後のmeasure_pitches
        """
        time_in_measure = self.time_in_measures[1]
        quantized_units = list(self.quantized_unit_times.values())
        next_measure_pitches = []
        rev_measure_pitches = deepcopy(measure_pitches)
        for num_measure, pitch_dict in enumerate(measure_pitches):
            rev_list = []
            for val in next_measure_pitches:
                if 0 in pitch_dict.keys():
                    rev_measure_pitches[num_measure][0][0].extend(val[0])
                    rev_measure_pitches[num_measure][0][1].extend(val[1])
                else:
                    rev_measure_pitches[num_measure][0] = val
            next_measure_pitches = []
            if num_measure in self.time_in_measures.keys():
                time_in_measure = self.time_in_measures[num_measure]
            for time in pitch_dict.keys():
                odds = [time % quantized_unit for quantized_unit in quantized_units]
                min_odd_val = min(odds)
                if min_odd_val != 0:
                    # min_odd_idx = odds.index(min(odds))
                    # rev_time = (
                    #     round(time / quantized_units[min_odd_idx])
                    #     * quantized_units[min_odd_idx]
                    # )
                    # min_unit = min(DICT_FOR_QUANTIZED_UNIT_TIMES.values())
                    min_unit = DICT_FOR_QUANTIZED_UNIT_TIMES[(4, 48)]
                    rev_time = (
                        ceil(time / min_unit) * min_unit
                    )
                    rev_list.append([time, rev_time])
            for bef_time, aft_time in rev_list:
                print("[{}] ADJTIME: ".format(num_measure), bef_time, "->", aft_time)
                val = rev_measure_pitches[num_measure].pop(bef_time)
                if aft_time == time_in_measure:
                    next_measure_pitches.append(val)
                elif aft_time in pitch_dict.keys():
                    rev_measure_pitches[num_measure][aft_time][0].extend(val[0])
                    rev_measure_pitches[num_measure][aft_time][1].extend(val[1])
                else:
                    rev_measure_pitches[num_measure][aft_time] = val
        return rev_measure_pitches

    def _convert_measures_to_pitch_list(self, measure_pitches, channel_num):
        """
        1小節ごとのmidi命令(note_on, note_off)をピッチリストに変換する関数
        
        Parameters
        ----------
        measure_pitches : list of dict {tick: [note_on, note_off]}
            小節ごとに区切ったデータ、キーは小節内の時間、
            値はその時の[note_onになった音の高さ, note_offになった時の〃]
        channel_num : int
            SMFのチャネル番号
        
        Returns
        -------
        [[notes_in_measure], [〃], ... ]
            ピッチリスト
        """
        rsorted_qunit_times = sorted(
            self.quantized_unit_times.values(), key=lambda x: x, reverse=True
        )
        pitch_list = []
        pitch_list_append = pitch_list.append
        time_in_measure = self.time_in_measures[1]
        sounding_pitch = -1
        for num_measure, pitch_dict in enumerate(measure_pitches, start=1):
            print("NOTE_DICT", num_measure, pitch_dict)
            if num_measure in self.time_in_measures.keys():
                time_in_measure = self.time_in_measures[num_measure]
                rhythm = self.rhythm_dict[num_measure]
                num_beats = int(rhythm[0])
                beat_time = time_in_measure // int(num_beats)
            times = list(pitch_dict.keys())
            if len(times) == 0:
                if sounding_pitch < 0:
                    pitch_list_append([["r"]] * num_beats)
                else:
                    pitch_list_append([["-"]] * num_beats)
                continue
            no_odd_unit = 0
            for qunit_time in rsorted_qunit_times:
                odds = [time % qunit_time for time in times]
                max_odd_val = max(odds)
                if max_odd_val == 0:
                    no_odd_unit = qunit_time
                    break
            if no_odd_unit == 0:
                print("Warning: measure={} has no divisible unit.".format(num_measure))
                print(times)
                pitch_list_append([["-"]] * num_beats)
                continue
            print(num_measure)
            pitches_in_measure = [[] for _ in range(num_beats)]
            # pitches_in_measure = []
            num_beat = -1
            for i in range(ceil(time_in_measure / no_odd_unit)):
                _time = i * no_odd_unit
                if _time % beat_time == 0:
                    num_beat += 1
                    # pitches_in_measure.append([])
                add_flag = False
                if _time in times:
                    on_pitches, off_pitches = pitch_dict[_time]
                    if sounding_pitch in off_pitches:
                        sounding_pitch = -1
                    sorted_pitches = [_on_pitch for _on_pitch in sorted(on_pitches)]
                    if len(sorted_pitches) > 0:
                        pitches_in_measure[num_beat].append(
                            [_on_pitch for _on_pitch in sorted_pitches]
                        )
                        if channel_num != 9:
                            sounding_pitch = max(sorted_pitches)
                        add_flag = True
                if add_flag is False:
                    if sounding_pitch < 0:
                        pitches_in_measure[num_beat].append("r")
                    else:
                        pitches_in_measure[num_beat].append("-")
            pitch_list.append(pitches_in_measure)
        return pitch_list

    def _get_quantized_unit_times(self):
        """
        量子化する音符の長さと時間を得る関数
        
        Returns
        -------
        dict {[numerator, denominator]: time}
            キーが音価、値が時間の辞書
        """
        quantized_unit_times = {}
        for (num_mul, num_div), qtime in DICT_FOR_QUANTIZED_UNIT_TIMES.items():
            add_flag = True
            for num_measure, time_in_measure in self.time_in_measures.items():
                numerator, _ = self.rhythm_dict[num_measure]
                beat_time = time_in_measure // numerator
                if (beat_time % qtime) != 0:
                    add_flag = False
                    break
            if add_flag:
                quantized_unit_times[(num_mul, num_div)] = qtime
        return quantized_unit_times

    def _get_measures_and_rhythms(self):
        """
        SMFのコンダクタートラックから小節ごとの時間と拍子を得る関数
        
        Returns
        -------
        dict
            小節ごとの時間
        dict
            小節ごとの拍子
        """
        now_time = 0
        bef_time = 0
        now_measure = 1
        remain_measure = 0
        time_in_measure = self._get_time_in_measure(4, 4)  # temporary
        rhythm_dict = {}
        time_in_measures = {1: time_in_measure}

        for msg in self.mid.tracks[0]:
            now_time += msg.time
            if msg.type == "time_signature":
                temp_measure = float(now_time - bef_time) / time_in_measure
                now_measure += int(temp_measure) + remain_measure
                remain_measure = temp_measure - int(temp_measure)
                time_in_measure = self._get_time_in_measure(
                    msg.numerator, msg.denominator
                )
                rhythm_dict[int(now_measure)] = (msg.numerator, msg.denominator)
                time_in_measures[int(now_measure)] = time_in_measure
                bef_time = now_time
        print("TIMEIN_MEASURES", time_in_measures)

        return time_in_measures, rhythm_dict

    def _get_use_channel_list(self):
        """
        現在の使用チャネルを得る関数
        
        Returns
        -------
        list of bool
            使用しているチャネルリスト
        """
        use_channel_list = [False] * 16
        for track in self.mid.tracks:
            channel_num = self._get_channel_num(track)
            if channel_num is not None:
                use_channel_list[channel_num] = True
        return use_channel_list

    def _get_channel_num(self, track):
        """
        トラック内のチャネル番号を得る関数
        
        Parameters
        ----------
        track : mido.Track
            SMFのトラック
        
        Returns
        -------
        int or None
            チャネル番号
        """
        for msg in track:
            if msg.type == "note_on":
                return msg.channel
        return None

    def _get_program_num(self, track):
        """
        トラック内のプログラム番号を得る関数
        
        Parameters
        ----------
        track : mido.Track
            SMFのトラック
        
        Returns
        -------
        int
            プログラム番号
        """
        for msg in track:
            if msg.type == "program_change":
                program_num = msg.program
                return program_num
        return None

    def get_tempo(self, search_beats=16):
        """
        コンダクタートラックから曲のテンポを得る関数
        search_beatsの拍数だけサーチして、最後に設定されたテンポを返す
        
        Parameters
        ----------
        search_beats : int, optional
            サーチする拍の数, by default 16
        
        Returns
        -------
        int
            曲のテンポ
        """
        search_ticks = search_beats * self.ticks_per_beat
        tempo = None
        for msg in self.mid.tracks[0]:
            if msg.time > search_ticks:
                break
            if msg.type == "set_tempo":
                tempo = int(tempo2bpm(msg.tempo))
        return tempo
