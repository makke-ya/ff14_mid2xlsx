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
        print(sound_list)
        print(rate_list)
        midi_data.add_sound_list(
            sound_list,
            rate_list,
            program=program,
        )
    midi_data.fwrite(mid_name)
    midi_data.pprint()
    

def ems_main():
    # xlsx_data = XlsxLoader("EMS楽譜専用(EMS score).xlsx")
    # xlsx_data = XlsxLoader("EMS作曲・下書き専用_WOT.xlsx", force_same_width=True)
    # xlsx_data = XlsxLoader("UU.xlsx", force_same_width=True)
    # xlsx_data = XlsxLoader("CR2下書き.xlsx", force_same_width=True)
    xlsx_data = XlsxLoader("EMS作曲・下書き専用_CON.xlsx", force_same_width=True)
    makedirs("out", exist_ok=True)

    # CON下書き
    dict_CON = {
        "sheet_name": "CON下書き",
        "start_column_char": "E",
        "end_column_char": "BP",
        "row_lists": [
            list(range(15, 94, 8)),
            list(range(16, 94, 8)),
            list(range(17, 94, 8)),
            list(range(18, 94, 8)),
            list(range(19, 94, 8)),
            list(range(20, 94, 8)),
            list(range(21, 94, 8)),
        ],
        "program_list": [
            "グランドピアノ", 
            "グランドピアノ", 
            "グランドピアノ", 
            "グランドピアノ", 
            "ヴァイオリン", 
            "ヴァイオリン", 
            "コントラバス", 
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 60
    create_midi(xlsx_data, tempo, dict_CON, "out/CON_tempo{}.mid".format(tempo))
    return

    # CR2下書き
    dict_CR2 = {
        "sheet_name": "CR2下書き",
        "start_column_char": "E",
        "end_column_char": "AZ",
        "row_lists": [
            list(range(11, 40, 3)),
            list(range(12, 40, 3)),
        ],
        "program_list": [
            "グランドピアノ", "スチールギター"
        ],
        "rhythm": { 
            1: (6, 4),
        }
    }
    tempo = 128
    create_midi(xlsx_data, tempo, dict_CR2, "out/CR2_tempo{}.mid".format(tempo))
    return

    # WOT下書き
    dict_WOT = {
        "sheet_name": "WOT",
        "start_column_char": "E",
        "end_column_char": "BP",
        "row_lists": [
            list(range(13, 31, 3)),
            list(range(14, 32, 3)),
        ],
        "program_list": [
            "パンパイプ", "スチールギター"
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 75
    create_midi(xlsx_data, tempo, dict_WOT, "out/WOT_tempo{}.mid".format(tempo))
    return

    # UU
    dict_UU = {
        "sheet_name": "UU下書き",
        "start_column_char": "E",
        "end_column_char": "BP",
        "row_lists": [
            list(range(14, 337, 9)),
            list(range(15, 337, 9)),
            list(range(16, 337, 9)),
            list(range(17, 337, 9)),
            list(range(18, 337, 9)),
            list(range(19, 337, 9)),
            list(range(20, 337, 9)),
            list(range(21, 337, 9)),
        ],
        "program_list": [
            "サックス", "ヴァイオリン", "ホルン", "パンパイプ",
            "オーバードライブギター", "ディストーションギター", "クリーンギター", "ティンパニー"
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 125
    create_midi(xlsx_data, tempo, dict_UU, "out/UU_tempo{}.mid".format(tempo))
    return

    # ALPv4下書き
    dict_BCD = {
        "sheet_name": "ALPv4下書き",
        "start_column_char": "E",
        "end_column_char": "BP",
        "row_lists": [
            list(range(12, 434, 9)),
            list(range(13, 434, 9)),
            list(range(14, 434, 9)),
            list(range(15, 434, 9)),
            list(range(16, 434, 9)),
            list(range(17, 434, 9)),
            list(range(18, 434, 9)),
            list(range(19, 434, 9)),
        ],
        "program_list": [
            "トランペット", "サックス", "トランペット", "チューバ",
            "トロンボーン", "トロンボーン", "ハープ", "ティンパニー"
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 150
    create_midi(xlsx_data, tempo, dict_BCD, "out/ALPv4_tempo{}.mid".format(tempo))
    return

    # BCD下書き
    dict_BCD = {
        "sheet_name": "BCD下書き",
        "start_column_char": "E",
        "end_column_char": "AZ",
        "row_lists": [
            list(range(9, 52, 4)),
            list(range(10, 52, 4)),
            list(range(11, 52, 4)),
        ],
        "program_list": [
            "パンパイプ", "パンパイプ", "ハープ"
        ],
        "rhythm": { 
            1: (3, 4),
        }
    }
    tempo = 84
    create_midi(xlsx_data, tempo, dict_BCD, "out/BCD_tempo{}.mid".format(tempo))
    return

    # DBA下書き
    dict_DBA = {
        "sheet_name": "DBA下書き",
        "start_column_char": "I",
        "end_column_char": "CJ",
        "row_lists": [
            list(range(11, 218, 9)),
            list(range(12, 218, 9)),
            list(range(13, 218, 9)),
            list(range(14, 218, 9)),
            list(range(15, 218, 9)),
            list(range(16, 218, 9)),
            list(range(17, 218, 9)),
            list(range(18, 218, 9)),
        ],
        "program_list": [
            "サックス", "サックス", "サックス", "グランドピアノ",
            "サックス", "サックス", "サックス", "ティンパニー"
        ],
        "rhythm": { 
            1: (5, 4),
        }
    }
    tempo = 163
    create_midi(xlsx_data, tempo, dict_DBA, "out/DBA_tempo{}.mid".format(tempo))
    return

    # MAT改下書き
    dict_MAT = {
        "sheet_name": "MAT改下書き",
        "start_column_char": "E",
        "end_column_char": "BP",
        "row_lists": [
            list(range(11, 102, 9)),
            list(range(12, 103, 9)),
            list(range(13, 104, 9)),
            list(range(14, 105, 9)),
            list(range(15, 106, 9)),
            list(range(16, 107, 9)),
            list(range(17, 108, 9)),
            list(range(18, 109, 9)),
        ],
        "program_list": [
            "グランドピアノ", "オーボエ", "パンパイプ", "フルート",
            "スチールギター", "ホルン", "ハープ", "ピッチカート"
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 74
    create_midi(xlsx_data, tempo, dict_MAT, "out/MAT改_tempo{}.mid".format(tempo))
    return

    # GRI下書き
    dict_GRI = {
        "sheet_name": "GRI下書き",
        "start_column_char": "E",
        "end_column_char": "BP",
        "row_lists": [
            list(range(11, 102, 9)),
            list(range(12, 103, 9)),
            list(range(13, 104, 9)),
            list(range(14, 105, 9)),
            list(range(15, 106, 9)),
            list(range(16, 107, 9)),
            list(range(17, 108, 9)),
            list(range(18, 109, 9)),
        ],
        "program_list": [
            "パンパイプ", "パンパイプ", "ピッチカート", "ピッチカート",
            "ピッチカート", "ハープ", "ボンゴ", "スネアドラム"
        ],
        "rhythm": { 
            1: (2, 2),
        }
    }
    tempo = 142
    create_midi(xlsx_data, tempo, dict_GRI, "out/GRI_tempo{}.mid".format(tempo))
    return

    # KUG下書き
    dict_KUG = {
        "sheet_name": "KUG下書き",
        "start_column_char": "E",
        "end_column_char": "BP",
        "row_lists": [
            list(range(11, 80, 5)),
            list(range(12, 80, 5)),
            list(range(13, 80, 5)),
            list(range(14, 80, 5)),
        ],
        "program_list": [
            "グランドピアノ", "グランドピアノ", "グランドピアノ", "グランドピアノ",
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 80
    create_midi(xlsx_data, tempo, dict_KUG, "out/KUG_tempo{}.mid".format(tempo))
    return

    # WTH下書き
    dict_WTH = {
        "sheet_name": "WTH下書き",
        "start_column_char": "E",
        "end_column_char": "BP",
        "row_lists": [
            list(range(11, 84, 8)),
            list(range(12, 85, 8)),
            list(range(13, 86, 8)),
            list(range(14, 87, 8)),
            list(range(15, 88, 8)),
            list(range(16, 89, 8)),
            list(range(17, 90, 8)),
        ],
        "program_list": [
            "グランドピアノ", "グランドピアノ", "グランドピアノ", "グランドピアノ",
            "グランドピアノ", "グランドピアノ", "グランドピアノ"
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 80
    create_midi(xlsx_data, tempo, dict_WTH, "out/WTH_tempo{}.mid".format(tempo))
    return

    # WOS下書き
    dict_WOS = {
        "sheet_name": "WOS下書き",
        "start_column_char": "E",
        "end_column_char": "BP",
        "row_lists": [
            list(range(11, 102, 6)),
            list(range(12, 103, 6)),
            list(range(13, 104, 6)),
            list(range(14, 105, 6)),
            list(range(15, 106, 6)),
        ],
        "program_list": [
            "グランドピアノ", "グランドピアノ", "グランドピアノ", "グランドピアノ", "グランドピアノ"
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 150
    create_midi(xlsx_data, tempo, dict_WOS, "out/WOS_tempo{}.mid".format(tempo))
    return

    # EW下書き
    dict_EW = {
        "sheet_name": "EW下書き",
        "start_column_char": "E",
        "end_column_char": "BP",
        "row_lists": [
            list(range(11, 78, 6)),
            list(range(12, 79, 6)),
            list(range(13, 80, 6)),
            list(range(14, 81, 6)),
            # list(range(15, 82, 6)),
        ],
        "program_list": [
            "フルート", "フルート", "ハープ", "ピッチカート"
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 100
    create_midi(xlsx_data, tempo, dict_EW, "out/EW_tempo{}.mid".format(tempo))

    # TIT下書き
    dict_TIT = {
        "sheet_name": "TIT下書き",
        "start_column_char": "E",
        "end_column_char": "AZ",
        "row_lists": [
            list(range(11, 62, 5)),
            list(range(12, 63, 5)),
            list(range(13, 64, 5)),
            list(range(14, 65, 5)),
        ],
        "program_list": [
            "フルート", "オーボエ", "クラリネット", "ホルン",
        ],
        "rhythm": { 
            1: (6, 4),
        }

    }
    tempo = 180
    create_midi(xlsx_data, tempo, dict_TIT, "out/TIT_tempo{}.mid".format(tempo))
    return

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

    # NIR2
    dict_NIR2 = {
        "sheet_name": "NIR2",
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
    create_midi(xlsx_data, tempo, dict_NIR2, "out/NIR2_tempo{}.mid".format(tempo))
    tempo = 60
    create_midi(xlsx_data, tempo, dict_NIR2, "out/NIR2_tempo{}.mid".format(tempo))

    # NIR7
    dict_NIR7 = {
        "sheet_name": "NIR7",
        "start_column_char": "D",
        "end_column_char": "AM",
        "row_lists": [
            list(range(10, 43, 2)) + list(range(72, 77, 2)),
            list(range(79, 112, 2)) + list(range(141, 146, 2)),
            list(range(148, 181, 2)) + list(range(210, 215, 2)),
            list(range(217, 250, 2)) + list(range(279, 284, 2)),
            list(range(286, 319, 2)) + list(range(348, 353, 2)),
            list(range(355, 388, 2)) + list(range(417, 422, 2)),
            list(range(424, 457, 2)) + list(range(486, 491, 2)),
        ],
        "program_list": [
            "サックス", "トランペット", "オーボエ", "トロンボーン",
            "チューバ", "スネアドラム", "ハープ"
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 162
    create_midi(xlsx_data, tempo, dict_NIR7, "out/NIR7_tempo{}.mid".format(tempo))
    tempo = 120
    create_midi(xlsx_data, tempo, dict_NIR7, "out/NIR7_tempo{}.mid".format(tempo))

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
        "start_column_char": "D",
        "end_column_char": "BO",
        "row_lists": [
            list(range(12, 35, 2)),
            list(range(37, 60, 2)),
            list(range(62, 85, 2)),
            list(range(87, 110, 2)),
            list(range(112, 135, 2)),
            list(range(137, 160, 2)),
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

    # NIB4
    dict_NIB4 = {
        "sheet_name": "NIB4",
        "start_column_char": "D",
        "end_column_char": "DA",
        "row_lists": [
            list(range(10, 31, 2)),
            list(range(33, 54, 2)),
            list(range(56, 77, 2)),
            list(range(79, 100, 2)),
        ],
        "program_list": [
            "フルート", "グランドピアノ", "グランドピアノ", "スチールギター",
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 60
    create_midi(xlsx_data, tempo, dict_NIB4, "out/NIB4_tempo{}.mid".format(tempo))

    # ILM4
    dict_ILM4 = {
        "sheet_name": "ILM4",
        "start_column_char": "D",
        "end_column_char": "AY",
        "row_lists": [
            list(range(11, 24, 2)),
            list(range(26, 39, 2)),
            list(range(41, 54, 2)),
            list(range(56, 69, 2)),
        ],
        "program_list": [
            "グランドピアノ", "グランドピアノ", "グランドピアノ", "グランドピアノ",
        ],
        "rhythm": { 
            1: (2, 4),
        }
    }
    tempo = 25
    create_midi(xlsx_data, tempo, dict_ILM4, "out/ILM4_tempo{}.mid".format(tempo))

    # ULD下書き
    dict_ULD = {
        "sheet_name": "ULD下書き",
        "start_column_char": "E",
        "end_column_char": "AZ",
        "row_lists": [
            list(range(9, 86, 5)),
            list(range(10, 86, 5)),
            list(range(11, 86, 5)),
            list(range(12, 86, 5)),
        ],
        "program_list": [
            "グランドピアノ", "グランドピアノ", "グランドピアノ", "グランドピアノ",
        ],
        "rhythm": { 
            1: (4, 4),
        }
    }
    tempo = 82
    create_midi(xlsx_data, tempo, dict_ULD, "out/ULD_tempo{}.mid".format(tempo))


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
    ems_main()
    # main()
