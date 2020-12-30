# coding: utf-8
from os import makedirs

from dataset.midi.loader import MidiLoader
from dataset.xlsx.writer import XlsxWriter
from dataset.xlsx.writer import ThreeLineXlsxWriter
from dataset.xlsx.writer import FlexibleLineXlsxWriter


class Mid2XlsxConverter(object):
    def __init__(self):
        self.midi_data = None
        self.common_data_list = None
        self.xlsx_data = None

    def get_tempo(self):
        return self.midi_data.tempo

    def get_rhythm_dict(self):
        return self.common_data_list[0].get_rhythm_dict()

    def fopen(self, filename):
        self.midi_data = MidiLoader(filename)
        self.common_data_list = self.midi_data.get_common_data_list()

        ret_dict = {}
        for common_data in self.common_data_list:
            channel_num = common_data.get_channel_num()
            print("ch", channel_num)
            if channel_num != 9:
                ret_dict[channel_num] = common_data.get_program_str()
            else:
                if channel_num not in ret_dict:
                    ret_dict[channel_num] = []
                ret_dict[channel_num].append(common_data.get_program_str())
        return ret_dict

    def key_estimate(self):
        # 調の推定
        all_pitches = [
            common_data.get_flatten_pitches() for common_data in self.common_data_list
            if common_data.get_channel_num() != 9
        ]
        all_pitches = sum(all_pitches, [])
        print("All_pitches", all_pitches)
        estimate_key = self.common_data_list[0].estimate_key(all_pitches)
        for common_data in self.common_data_list:
            common_data.set_key(estimate_key)
        print("Estimate Key: {}".format(estimate_key))
        return estimate_key
    
    def pitch_estimate(self, program_dict=None):
        ret_dict = {}
        for common_data in self.common_data_list:
            channel_num = common_data.get_channel_num()
            print("Channel-num", channel_num)
    
            if (
                program_dict is not None
                and channel_num != 9
                and channel_num in program_dict.keys()
            ):
                common_data.add_program_str(program_dict[channel_num])

            # トラックごとに適しているオクターブを示す
            if common_data.set_recommended_octave() is False:
                continue  # 音がない
            shift_pitch = common_data.get_shift_pitch()
            print("shift-pitch:", shift_pitch)
    
            # まずはそのオクターブでサウンドリストに落とし込む
            max_pitch, min_pitch, outlier_list = common_data.convert_pitch2note(
                enable_chord=False
            )
            print("note list", common_data.get_note_list())
            print("max_pitch, min_pitch", max_pitch, min_pitch)
            print("outlier_list", outlier_list)

            if channel_num != 9:
                ret_dict[channel_num] = [
                    shift_pitch, [max_pitch, min_pitch, len(outlier_list)]
                ]
            else:
                if channel_num not in ret_dict.keys():
                    ret_dict[channel_num] = []
                ret_dict[channel_num].append([
                    shift_pitch, [max_pitch, min_pitch, len(outlier_list)]
                ])
        return ret_dict
    
    def update(self, program_dict, key_dict, pitch_dict, drum_modes=[], style="1行固定"):
        _program_dict = {}
        _key_dict = {}
        _pitch_dict = {}
        enable_chord = True if style != "1行固定" else False
        self.common_data_list = self.midi_data.get_common_data_list(
            drum_modes=drum_modes
        )
        for common_data in self.common_data_list:
            channel_num = common_data.get_channel_num()
            if (
                channel_num in program_dict
                and channel_num in pitch_dict
            ):
                if channel_num != 9:
                    program = program_dict[channel_num]
                    shift_pitch, _ = pitch_dict[channel_num]
                    common_data.set_key(key_dict)
                    print(program, shift_pitch)
                    common_data.add_program_str(program)
                    common_data.set_shift_pitch(shift_pitch)
                    max_pitch, min_pitch, outlier_list = common_data.convert_pitch2note(
                        enable_chord=enable_chord
                    )

                    _program_dict[channel_num] = common_data.get_program_str()
                    _key_dict = common_data.get_key_dict()
                    _pitch_dict[channel_num] = [
                        common_data.get_shift_pitch(), (max_pitch, min_pitch, len(outlier_list))
                    ]
                else:  # channel_num == 9
                    # pitches = pitch_dict[channel_num]
                    common_data.set_key(key_dict)
                    # common_data.add_program_str(program)
                    # common_data.set_shift_pitch(shift_pitch)
                    max_pitch, min_pitch, outlier_list = common_data.convert_pitch2note(
                        enable_chord=enable_chord
                    )
                    if channel_num not in _program_dict.keys():
                        _program_dict[channel_num] = []
                        _pitch_dict[channel_num] = []
                    _program_dict[channel_num].append(common_data.get_program_str())
                    _key_dict = common_data.get_key_dict()
                    _pitch_dict[channel_num].append(
                        [
                            common_data.get_shift_pitch(),
                            (max_pitch, min_pitch, len(outlier_list))
                        ]
                    )
        return _program_dict, _key_dict, _pitch_dict
        
    def fwrite(
        self, filename, title_name, on_list=None, style="1行固定", shorten=False,
        start_measure_num=1, num_measures_in_system=4, score_width=29.76
    ):
        _common_data_list = []
        if on_list is not None:
            for common_data in self.common_data_list:
                channel_num = common_data.get_channel_num()
                if channel_num not in on_list:
                    continue
                _common_data_list.append(common_data)
        else:
            _common_data_list = self.common_data_list

        # エクセル化する
        title = title_name if title_name != "" else "test"
        if style == "1行固定":
            writer = XlsxWriter
        elif style == "3行固定":
            writer = ThreeLineXlsxWriter
        else:
            writer = FlexibleLineXlsxWriter
        self.xlsx_data = writer(
            tempo=self.midi_data.tempo,
            title=title,
            start_measure_num=start_measure_num - 1,  # この回数だけ小節を消すため1は0に
            num_measures_in_system=num_measures_in_system,
            score_width=score_width,
            shorten=shorten,
        )
        self.xlsx_data.add_common_data_list(_common_data_list)
        self.xlsx_data.fwrite(filename)


if __name__ == "__main__":
    converter = Mid2XlsxConverter()
    # program_dict = converter.fopen(filename="./out/CRY_tempo81.mid")
    # program_dict = converter.fopen(filename="./out/WAW2_tempo60.mid")
    # program_dict = converter.fopen(filename="noname.mid")
    # program_dict = converter.fopen(filename="Eorzea de Chocobo.mid")
    # program_dict = converter.fopen(filename="FF14_古アムダプール市街_(Hard).mid")
    # program_dict = converter.fopen(filename="for form.mid")
    # program_dict = converter.fopen(filename="Pharos Sirius(hard).mid")
    program_dict = converter.fopen(filename="chocobo_christmas.mid")
    key_dict = converter.key_estimate()
    pitch_dict = converter.pitch_estimate()

    drum_modes = ["スネアドラム", "シンバル"]
    style = 0
    program_dict, key_dict, pitch_dict = converter.update(
        program_dict, key_dict, pitch_dict, drum_modes, style
    )

    on_list = list(program_dict.keys())
    converter.fwrite("test.xlsx", "test", on_list=on_list, three_line=True)

    # from xlsx2mid import Xlsx2MidConverter
    # tempo = converter.midi_data.tempo
    # rhythm_dict = converter.common_data_list[0].get_rhythm_dict()
    # print(rhythm_dict)
    # input()
    # converter2 = Xlsx2MidConverter(tempo=tempo, rhythm_dict=rhythm_dict)
    # converter2.add_common_data_list(converter.common_data_list, on_list)
    # converter2.fwrite("test.mid")


    # TODO: 変更があればここで再度回す
    # for common_data in common_data_list:
    #     common_data.set_shift_pitch()
    #     # まずはそのオクターブでサウンドリストに落とし込む
    #     outlier_list = common_data.convert_pitch2note(disable_chord=True)
    #     print("note list", common_data.get_note_list())
    #     print("outlier_list", outlier_list)

