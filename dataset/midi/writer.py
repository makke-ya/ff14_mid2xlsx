# coding: utf-8
from mido import Message, MetaMessage, MidiFile, MidiTrack, bpm2tempo, tempo2bpm
from .base import MidiIOBase
from ._static_data import (
    DICT_FOR_PITCH_CONVERT,
)


class MidiWriter(MidiIOBase):
    """
    SMF(Standard Midi Files)のデータとヘルパー関数
    """

    def __init__(
        self, type=1, tempo=60, ticks_per_beat=480, rhythm_dict=None
    ):
        self.mid = MidiFile(type=type, ticks_per_beat=ticks_per_beat)
        self.ticks_per_beat = ticks_per_beat
        self.tempo = tempo
        self.time_in_measures = {}
        self.mid.tracks.append(MidiTrack())
        self.set_tempo(tempo)
        self.rhythm_dict = rhythm_dict
        if rhythm_dict is None:
            self.rhythm_dict = {1: (4, 4)}
        for message in self._convert_rhythm_messages(self.rhythm_dict):
            self.mid.tracks[0].append(message)
        self.use_channel_list = [False] * 16

    def add_common_data(self, common_data):
        """
        共通音楽データ(CommonSoundData)を追加する関数
        
        Parameters
        ----------
        common_data_list : CommonSoundData
            共通音楽データ
        program : str, optional
            楽器種類（プログラム名）, by default "Acoustic Piano"
        """
        program_str = common_data.get_program_str()
        sound_list = common_data.get_note_list()
        rate_list = common_data.get_rate_list()
        program_num, base_pitch = self.convert_program_str2num(program_str)
        if rate_list is None:
            rate_list = common_data._create_dummy_rates()
        channel_num = self._get_empty_channel_num()
        if channel_num is None:
            print("Error: already using All 16 Channels!")
            return
        print("Channel num: {}")
        self.use_channel_list[channel_num] = True

        track = MidiTrack()
        track.append(
            Message("program_change", channel=channel_num, program=program_num, time=0)
        )
        bef_pitches = []
        bef_time = 0
        interval_time = 0
        now_time = 0
        time_in_measure = 1920
        basetime_in_beat = 480
        print(self.time_in_measures)
        # TODO: erase multiple for loop
        for (
            measure_num,
            (notes_in_measure, (rates_in_measure, rates_in_beats)),
        ) in enumerate(zip(sound_list, rate_list), start=1):
            # [[[c], [c, d]], [], [], []]
            print(measure_num, time_in_measure, now_time, interval_time)
            if measure_num in self.time_in_measures:
                time_in_measure = self.time_in_measures[measure_num]
            basetime_in_beat = time_in_measure // sum(rates_in_measure)
            bef_measure_time = now_time
            for beat_num, (notes_in_beat) in enumerate(notes_in_measure):
                # [[c], [c, d]]
                time_in_beat = rates_in_measure[beat_num] * basetime_in_beat
                basetime_in_cell = time_in_beat // sum(rates_in_beats[beat_num])
                for cell_num, notes in enumerate(notes_in_beat):
                    # [c] or [c, d]
                    time_in_cell = rates_in_beats[beat_num][cell_num] * basetime_in_cell
                    time_in_str = time_in_cell // len(notes)
                    for note in notes:
                        if note == "-":
                            pass
                        elif note == "r":
                            if bef_pitches:
                                for bef_pitch in bef_pitches:
                                    track.append(
                                        Message(
                                            "note_off",
                                            channel=channel_num,
                                            note=bef_pitch,
                                            velocity=0,
                                            time=interval_time,
                                        )
                                    )
                                    bef_time = now_time
                                    interval_time = 0
                                bef_pitches = []
                        else:
                            if bef_pitches:
                                for bef_pitch in bef_pitches:
                                    track.append(
                                        Message(
                                            "note_off",
                                            channel=channel_num,
                                            note=bef_pitch,
                                            velocity=0,
                                            time=interval_time,
                                        )
                                    )
                                    bef_time = now_time
                                    interval_time = 0
                                bef_pitches = []
                            if type(note) is str:
                                _note = [note]
                            pitches = []
                            for n in _note:
                                pitch = self._note2pitch(n, base_pitch)
                                track.append(
                                    Message(
                                        "note_on",
                                        channel=channel_num,
                                        note=pitch,
                                        velocity=100,
                                        time=interval_time,
                                    )
                                )
                                pitches.append(pitch)
                                bef_time = now_time
                                interval_time = 0
                            bef_pitches = pitches
                        interval_time += time_in_str
                        now_time += time_in_str
            if (bef_measure_time + time_in_measure) != (bef_time + interval_time):
                now_time = bef_measure_time + time_in_measure
                interval_time = now_time - bef_time
        if bef_pitches:
            for bef_pitch in bef_pitches:
                track.append(
                    Message(
                        "note_off",
                        channel=channel_num,
                        note=bef_pitch,
                        velocity=0,
                        time=interval_time,
                    )
                )
                interval_time = 0
        # for msg in track:
        #     print(msg)
        self.mid.tracks.append(track)

    def add_sound_list(self, sound_list, rate_list=None, program="Acoustic Piano"):
        """
        共通音リスト（、共通音レートリスト）を追加する関数
        
        Parameters
        ----------
        sound_list : list of list of list of list of str
            共通音リスト
        rate_list : list of tuple, optional
            共通音レートリスト, by default None
        program : str, optional
            楽器種類（プログラム名）, by default "Acoustic Piano"
        """
        program_num, base_pitch = self.convert_program_str2num(program)
        if rate_list is None:
            rate_list = self._create_dummy_rates(sound_list)
        channel_num = self._get_empty_channel_num()
        if channel_num is None:
            print("Error: already using All 16 Channels!")
            return
        # print("Channel num: {}")
        self.use_channel_list[channel_num] = True

        track = MidiTrack()
        print(program_num, channel_num)
        track.append(
            Message("program_change", channel=channel_num, program=program_num, time=0)
        )
        bef_pitches = []
        bef_time = 0
        interval_time = 0
        now_time = 0
        time_in_measure = 1920
        basetime_in_beat = 480
        print(self.time_in_measures)
        # TODO: erase multiple for loop
        for (
            measure_num,
            (notes_in_measure, (rates_in_measure, rates_in_beats)),
        ) in enumerate(zip(sound_list, rate_list), start=1):
            # [[[c], [c, d]], [], [], []]
            print(measure_num, time_in_measure, now_time, interval_time)
            if measure_num in self.time_in_measures:
                time_in_measure = self.time_in_measures[measure_num]
            basetime_in_beat = time_in_measure // sum(rates_in_measure)
            bef_measure_time = now_time
            for beat_num, (notes_in_beat) in enumerate(notes_in_measure):
                # [[c], [c, d]]
                time_in_beat = rates_in_measure[beat_num] * basetime_in_beat
                basetime_in_cell = time_in_beat // sum(rates_in_beats[beat_num])
                for cell_num, notes in enumerate(notes_in_beat):
                    # [c] or [c, d]
                    time_in_cell = rates_in_beats[beat_num][cell_num] * basetime_in_cell
                    time_in_str = time_in_cell // len(notes)
                    for note in notes:
                        if note == "-":
                            pass
                        elif note == "r":
                            if bef_pitches:
                                for bef_pitch in bef_pitches:
                                    track.append(
                                        Message(
                                            "note_off",
                                            channel=channel_num,
                                            note=bef_pitch,
                                            velocity=0,
                                            time=int(interval_time),
                                        )
                                    )
                                    bef_time = now_time
                                    interval_time = 0
                                bef_pitches = []
                        else:
                            if bef_pitches:
                                for bef_pitch in bef_pitches:
                                    track.append(
                                        Message(
                                            "note_off",
                                            channel=channel_num,
                                            note=bef_pitch,
                                            velocity=0,
                                            time=int(interval_time),
                                        )
                                    )
                                    bef_time = now_time
                                    interval_time = 0
                                bef_pitches = []
                            if type(note) is str:
                                _note = [note]
                            pitches = []
                            for n in _note:
                                pitch = self._note2pitch(n, base_pitch)
                                track.append(
                                    Message(
                                        "note_on",
                                        channel=channel_num,
                                        note=pitch,
                                        velocity=100,
                                        time=int(interval_time),
                                    )
                                )
                                pitches.append(pitch)
                                bef_time = now_time
                                interval_time = 0
                            bef_pitches = pitches
                        interval_time += time_in_str
                        now_time += time_in_str
            if (bef_measure_time + time_in_measure) != (bef_time + interval_time):
                now_time = bef_measure_time + time_in_measure
                interval_time = now_time - bef_time
        if bef_pitches:
            for bef_pitch in bef_pitches:
                track.append(
                    Message(
                        "note_off",
                        channel=channel_num,
                        note=bef_pitch,
                        velocity=0,
                        time=int(interval_time),
                    )
                )
                interval_time = 0
        # for msg in track:
        #     print(msg)
        self.mid.tracks.append(track)

    def set_tempo(self, tempo):
        """
        曲のテンポを設定する関数（複数回指定した場合は最後の指定が保存）

        Parameters
        ----------
        tempo : int
            曲のテンポ
        """
        del_idxs = []
        # テンポ指定するメタメッセージを全て消去
        for i, msg in enumerate(self.mid.tracks[0]):
            if msg.type == "set_tempo":
                del_idxs.append(i)
        for i in del_idxs[::-1]:
            self.mid.tracks[0].pop(i)

        self.mid.tracks[0].insert(
            0, MetaMessage("set_tempo", tempo=bpm2tempo(tempo), time=0)
        )

    def _convert_rhythm_messages(self, rhythm_dict):
        """
        拍子の格納された辞書を拍子変更のメタメッセージ群に変換する関数
        
        Parameters
        ----------
        rhythm_dict : dict of {int: (int, int)}
            キーを小節数、値に(拍数, 音価)を取る辞書, by default None
        
        Returns
        -------
        list of mido.MetaMessage
            拍子変更のメタメッセージリスト
        """
        time_in_measure = 1920
        now_time = 0
        bef_time = 0
        messages = []
        for measure_num in range(1, max(rhythm_dict.keys()) + 1):
            if measure_num in rhythm_dict:
                numerator, denominator = rhythm_dict[measure_num]
                messages.append(
                    MetaMessage(
                        "time_signature",
                        numerator=numerator,
                        denominator=denominator,
                        time=(now_time - bef_time),
                    )
                )
                time_in_measure = self._get_time_in_measure(numerator, denominator)
                self.time_in_measures[measure_num] = time_in_measure
                bef_time = now_time
            now_time += time_in_measure
        return messages

    def _get_empty_channel_num(self):
        """
        使用可能な空のチャネルの番号を返す関数
        
        Returns
        -------
        int or None
            使用可能な空のチャネル番号
        """
        for i, is_using in zip(range(0, 16), self.use_channel_list):
            if is_using is True or i == 9:
                continue
            else:
                return i
        return None

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
