# coding: utf-8
from os import makedirs
from argparse import ArgumentParser
from dataset.xlsx.loader import XlsxLoader
from dataset.midi.writer import MidiWriter


def create_midi(xlsx_data, tempo, data_dict, mid_name):
    midi_data = MidiWriter(tempo=tempo, rhythm_dict=data_dict["rhythm"])
    for row_list, program in zip(data_dict["row_lists"], data_dict["program_list"]):
        sound_list, rate_list = xlsx_data.get_sound_list(
            data_dict["sheet_name"],
            data_dict["start_column_char"],
            data_dict["end_column_char"],
            row_list,
        )
        midi_data.add_sound_list(
            sound_list,
            rate_list,
            program=program,
        )
    midi_data.fwrite(mid_name)
    midi_data.pprint()
    

def ems_main():
    xlsx_data = XlsxLoader("EMS楽譜専用(EMS score).xlsx")
    makedirs("out", exist_ok=True)

    # TM4
    dict_TM4 = {
        "sheet_name": "TM4",
        "start_column_char": "C",
        "end_column_char": "BP",
        "row_lists": [
            list(range(11, 24, 2)),
            [26, 28, 31, 33, 36, 38, 41],
            [44, 46, 49, 51, 54, 56, 59],
            [62, 64, 67, 69, 72, 74, 77],
            [80, 82, 85, 87, 90, 92, 95],
        ],
        "program_list": [
            "オーボエ", "トランペット", "トロンボーン", "チューバ", "ティンパニー",
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 92
    create_midi(xlsx_data, tempo, dict_TM4, "out/TM4_tempo{}.mid".format(tempo))
    tempo = 60
    create_midi(xlsx_data, tempo, dict_TM4, "out/TM4_tempo{}.mid".format(tempo))

    # TOT5
    dict_TOT5 = {
        "sheet_name": "TOT5",
        "start_column_char": "D",
        "end_column_char": "BO",
        "row_lists": [
            list(range(8, 25, 2)),
            list(range(28, 45, 2)),
            list(range(48, 65, 2)),
            list(range(68, 85, 2)),
            list(range(88, 105, 2)),
        ],
        "program_list": [
            "グランドピアノ", "グランドピアノ", "フルート", "フルート", "パンパイプ",
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 72
    create_midi(xlsx_data, tempo, dict_TOT5, "out/TOT5_tempo{}.mid".format(tempo))
    tempo = 60
    create_midi(xlsx_data, tempo, dict_TOT5, "out/TOT5_tempo{}.mid".format(tempo))

    # OVL4
    dict_OVL4 = {
        "sheet_name": "OVL4",
        "start_column_char": "D",
        "end_column_char": "BO",
        "row_lists": [
            [11, 13, 14, 16, 18, 20, 21, 23, 25, 27, 28, 30, 32, 34],
            [41, 42, 44, 45, 47, 48, 50, 51, 53, 54, 56, 57, 59, 60],
            [65, 66, 68, 69, 71, 72, 74, 75, 77, 78, 80, 81, 83, 84],
            [89, 90, 92, 93, 95, 96, 98, 99, 101, 102, 104, 105, 107, 108],
        ],
        "program_list": [
            "グランドピアノ", "グランドピアノ", "グランドピアノ", "スチールギター",
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 93
    create_midi(xlsx_data, tempo, dict_OVL4, "out/OVL4_tempo{}.mid".format(tempo))
    tempo = 60
    create_midi(xlsx_data, tempo, dict_OVL4, "out/OVL4_tempo{}.mid".format(tempo))

    # NIR
    dict_NIR = {
        "sheet_name": "NIR",
        "start_column_char": "D",
        "end_column_char": "BO",
        "row_lists": [
            list(range(9, 34, 2)),
            list(range(37, 62, 2)),
        ],
        "program_list": [
            "グランドピアノ", "グランドピアノ",
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 81
    create_midi(xlsx_data, tempo, dict_NIR, "out/NIR_tempo{}.mid".format(tempo))
    tempo = 60
    create_midi(xlsx_data, tempo, dict_NIR, "out/NIR_tempo{}.mid".format(tempo))

    # MAT2
    dict_MAT2 = {
        "sheet_name": "MAT2",
        "start_column_char": "D",
        "end_column_char": "BO",
        "row_lists": [
            list(range(11, 42, 3)),
            list(range(45, 66, 2)),
        ],
        "program_list": [
            "グランドピアノ", "スチールギター",
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 74
    create_midi(xlsx_data, tempo, dict_MAT2, "out/MAT2_tempo{}.mid".format(tempo))
    tempo = 55
    create_midi(xlsx_data, tempo, dict_MAT2, "out/MAT2_tempo{}.mid".format(tempo))

    # AMA
    dict_AMA = {
        # "sheet_name": "AMA",
        "sheet_name": "AMA(調がずれてる表記)",
        "start_column_char": "E",
        "end_column_char": "BP",
        "row_lists": [
            list(range(9, 32, 2)),
            list(range(34, 57, 2)),
            list(range(59, 82, 2)),
            list(range(84, 107, 2)),
            list(range(109, 132, 2)),
            list(range(134, 157, 2)),
        ],
        "program_list": [
            "グランドピアノ", "グランドピアノ", "グランドピアノ", "グランドピアノ", "スチールギター", "パンパイプ",
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 78
    create_midi(xlsx_data, tempo, dict_AMA, "out/AMA_tempo{}.mid".format(tempo))
    tempo = 60
    create_midi(xlsx_data, tempo, dict_AMA, "out/AMA_tempo{}.mid".format(tempo))

    # WAW2
    dict_WAW2 = {
        "sheet_name": "WAW2",
        "start_column_char": "D",
        "end_column_char": "AY",
        "row_lists": [
            [8, 10, 12, 14, 16, 18, 20, 22],
            [31, 33, 35, 37, 39, 41, 43, 45]
        ],
        "program_list": [
            "グランドピアノ", "グランドピアノ",
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 67
    create_midi(xlsx_data, tempo, dict_WAW2, "out/WAW2_tempo{}.mid".format(tempo))
    tempo = 60
    create_midi(xlsx_data, tempo, dict_WAW2, "out/WAW2_tempo{}.mid".format(tempo))

    # LIM下書き
    dict_LIM = {
        "sheet_name": "LIM7",
        "start_column_char": "D",
        "end_column_char": "BO",
        "row_lists": [
            [12, 13, 15, 16, 18, 19, 21, 22, 24, 26, 27, 30, 31, 33, 34, 36, 37],
            [41, 42, 44, 45, 47, 48, 50, 51, 53, 55, 56, 58, 59, 61, 62, 64],
            [68, 69, 71, 72, 74, 75, 77, 78, 80, 82, 83, 85, 86, 88, 89, 91],
            [95, 96, 98, 99, 101, 102, 104, 105, 107, 109, 110, 112, 113, 115, 116, 118],
            [122, 123, 125, 126, 128, 129, 131, 132, 134, 136, 137, 139, 140, 142, 143, 145],
            [149, 150, 152, 153, 155, 156, 158, 159, 161, 163, 164, 166, 167, 169, 170, 172, 173],
            [176, 177, 179, 180, 182, 183, 185, 186, 188, 190, 191, 193, 194, 196, 197, 199, 200],
        ],
        "program_list": [
            "トランペット", "トランペット", "ホルン", "ホルン", "トロンボーン", "トロンボーン", "チューバ",
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 112
    create_midi(xlsx_data, tempo, dict_LIM, "out/LIM_tempo{}.mid".format(tempo))
    tempo = 80
    create_midi(xlsx_data, tempo, dict_LIM, "out/LIM_tempo{}.mid".format(tempo))

    # CRY4_R2
    dict_CRY = {
        "sheet_name": "CRY4_R2",
        "start_column_char": "D",
        "end_column_char": "BO",
        "row_lists": [
            list(range(8, 39, 2)),
            list(range(41, 72, 2)),
            list(range(74, 105, 2)),
            list(range(107, 138, 2)),
            list(range(140, 171, 2)),
        ],
        "program_list": [
            "フルート", "オーボエ", "クラリネット", "クラリネット", "バスドラム",
        ],
        "rhythm": { 
            1: (3, 4),
            8: (4, 4),
            53: (3, 4),
            60: (4, 4),
        }
    }
    tempo = 81
    create_midi(xlsx_data, tempo, dict_CRY, "out/CRY_tempo{}.mid".format(tempo))
    tempo = 70
    create_midi(xlsx_data, tempo, dict_CRY, "out/CRY_tempo{}.mid".format(tempo))

    # BRN5
    dict_BRN5 = {
        "sheet_name": "BRN5",
        "start_column_char": "E",
        "end_column_char": "BP",
        "row_lists": [
            list(range(12, 37, 3)),
            list(range(41, 74, 4)),
            list(range(77, 102, 3)),
            list(range(105, 130, 3)),
            list(range(134, 159, 3)),
        ],
        "program_list": [
            "グランドピアノ", "グランドピアノ", "グランドピアノ", "グランドピアノ", "スチールギター",
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 70
    create_midi(xlsx_data, tempo, dict_BRN5, "out/BRN5_tempo{}.mid".format(tempo))
    tempo = 60
    create_midi(xlsx_data, tempo, dict_BRN5, "out/BRN5_tempo{}.mid".format(tempo))


def parse_args():
    parser = ArgumentParser(description="EMS形式のxlsxをmidiに変換するツール")
    parser.add_argument(
        "-i", "--xlsx_name", dest="xlsx_name", type=str, required=True,
        help="エクセルのファイル名"
    )
    parser.add_argument(
        "-s", "--sheet_name", dest="sheet_name", type=str, required=True,
        help="エクセルのシート名"
    )
    parser.add_argument(
        "-t", "--tempo", dest="tempo", type=int, required=True,
        help="テンポ"
    )
    parser.add_argument(
        "--sc", "--start_column", dest="start_column", type=str, required=True,
        help="楽譜が開始する列の英番号",
    )
    parser.add_argument(
        "--ec", "--end_column", dest="end_column", type=str, required=True,
        help="楽譜が終わる列の英番号",
    )
    parser.add_argument(
        "--rs", "--rows", dest="rows", type=str, nargs="+", required=True,
        help="各楽器の行番号リスト ex) 1,2,3,4 <-一つ目の楽器の行 5,6,7,8<-二つ目の楽器の行",
    )
    parser.add_argument(
        "-p", "--programs", dest="programs", type=str, nargs="+", required=True,
        help="各楽器の種類リスト ex) グランドピアノ スチールギター\n"
             + "対応楽器リスト: ハープ, グランドピアノ, スチールギター, ピチカート, \n"
             + "               フルート, オーボエ, クラリネット, ピッコロ, バンパイプ, \n"
             + "               ティンパニー, ボンゴ, バスドラム, スネアドラム, シンバル, \n"
             + "               トランペット, トロンボーン, チューバ, ホルン, サックス"
    )
    parser.add_argument(
        "-r", "--rhythms", dest="rhythms", type=str, nargs="+", required=True,
        help="拍子リスト ex) 1,4,4 <-1小節目は4/4 10,3,4 <-10小節目から3/4"
    )
    return parser.parse_args()


class Xlsx2MidConverter(object):
    def __init__(self, tempo, rhythm_dict):
        self.xlsx_data = None
        self.common_data_list = []
        self.tempo = tempo
        self.rhythm_dict = rhythm_dict
        self.midi_data = MidiWriter(tempo=tempo, rhythm_dict=rhythm_dict)

    def fopen(self, filename):
        self.xlsx_data = XlsxLoader(filename)

    # def create_midi(self, data_dict, mid_name):
    #     for row_list, program in zip(data_dict["row_lists"], data_dict["program_list"]):
    #         sound_list, rate_list = self.xlsx_data.get_sound_list(
    #             data_dict["sheet_name"],
    #             data_dict["start_column_char"],
    #             data_dict["end_column_char"],
    #             row_list,
    #         )
    
    def add_common_data_list(self, common_data_list, on_list):
        _common_data_list = []
        if on_list is not None:
            for common_data in common_data_list:
                channel_num = common_data.get_channel_num()
                if channel_num not in on_list:
                    continue
                _common_data_list.append(common_data)
        else:
            _common_data_list = common_data_list
        self.common_data_list = _common_data_list

    def fwrite(self, filename, on_list=None):
        midi_data = MidiWriter(tempo=self.tempo, rhythm_dict=self.rhythm_dict)
        for common_data in self.common_data_list:
            print(common_data.get_program_str())
            midi_data.add_common_data(common_data)
        midi_data.fwrite(filename)


def main():
    args = parse_args()
    xlsx_data = XlsxLoader(args.xlsx_name)
    data_dict = {
        "sheet_name": args.sheet_name,
        "start_column_char": args.start_column.upper(),
        "end_column_char": args.end_column.upper(),
        "row_lists": [[int(r) for r in rs.strip().split(",")] for rs in args.rows],
        "program_list": [p.strip() for p in args.programs],
        "rhythm": { 
            int(rhythm.strip().split(",")[0]): (
                int(rhythm.strip().split(",")[1]),
                int(rhythm.strip().split(",")[2]),
            ) for rhythm in args.rhythms
        }
    }
    print(data_dict)
    tempo = int(args.tempo)
    create_midi(xlsx_data, tempo, data_dict, "{}_tempo{}.mid".format(args.sheet_name, tempo))


if __name__ == '__main__':
    # ems_main()
    main()
