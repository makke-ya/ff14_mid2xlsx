## FOR WRITER
DICT_FOR_PITCH_CONVERT = {
    "Ces": -1, "C": 0, "Cis": 1,
    "Des": 1, "D": 2, "Dis": 3,
    "Es": 3, "E": 4, "Eis": 5,
    "Fes": 4, "F": 5, "Fis": 6,
    "Ges": 6, "G": 7, "Gis": 8,
    "As": 8, "A": 9, "Ais": 10,
    "B": 10, "H": 11, "His": 12,
}
DICT_FOR_PROGRAM = {
    ("Acoustic Piano",	"アコースティックピアノ"): 0,
    ("Bright Piano",	"ブライトピアノ"): 1,
    ("Electric Grand Piano",	"エレクトリック・グランドピアノ"): 2,
    ("Honky-tonk Piano",	"ホンキートンクピアノ"): 3,
    ("Electric Piano",	"エレクトリックピアノ"): 4,
    ("Electric Piano 2",	"FMエレクトリックピアノ"): 5,
    ("Harpsichord",	"ハープシコード"): 6,
    ("Clavi",	"クラビネット"): 7,
    ("Celesta",	"チェレスタ"): 8,
    ("Glockenspiel",	"グロッケンシュピール"): 9,
    ("Musical box",	"オルゴール"): 10,
    ("Vibraphone",	"ヴィブラフォン"): 11,
    ("Marimba",	"マリンバ"): 12,
    ("Xylophone",	"シロフォン"): 13,
    ("Tubular Bell",	"チューブラーベル"): 14,
    ("Dulcimer",	"ダルシマー"): 15,
    ("Drawbar Organ",	"ドローバーオルガン"): 16,
    ("Percussive Organ",	"パーカッシブオルガン"): 17,
    ("Rock Organ",	"ロックオルガン"): 18,
    ("Church organ",	"チャーチオルガン"): 19,
    ("Reed organ",	"リードオルガン"): 20,
    ("Accordion",	"アコーディオン"): 21,
    ("Harmonica",	"ハーモニカ"): 22,
    ("Tango Accordion",	"タンゴアコーディオン"): 23,
    ("Acoustic Guitar (nylon)",	"アコースティックギター（ナイロン弦）"): 24,
    ("Acoustic Guitar (steel)",	"アコースティックギター（スチール弦）"): 25,
    ("Electric Guitar (jazz)",	"ジャズギター"): 26,
    ("Electric Guitar (clean)",	"クリーンギター"): 27,
    ("Electric Guitar (muted)",	"ミュートギター"): 28,
    ("Overdriven Guitar",	"オーバードライブギター"): 29,
    ("Distortion Guitar",	"ディストーションギター"): 30,
    ("Guitar harmonics",	"ギターハーモニクス"): 31,
    ("Acoustic Bass",	"アコースティックベース"): 32,
    ("Electric Bass (finger)",	"フィンガー・ベース"): 33,
    ("Electric Bass (pick)",	"ピック・ベース"): 34,
    ("Fretless Bass",	"フレットレスベース"): 35,
    ("Slap Bass 1",	"スラップベース 1"): 36,
    ("Slap Bass 2",	"スラップベース 2"): 37,
    ("Synth Bass 1",	"シンセベース 1"): 38,
    ("Synth Bass 2",	"シンセベース 2"): 39,
    ("Violin",	"ヴァイオリン"): 40,
    ("Viola",	"ヴィオラ"): 41,
    ("Cello",	"チェロ"): 42,
    ("Double bass",	"コントラバス"): 43,
    ("Tremolo Strings",	"トレモロ"): 44,
    ("Pizzicato Strings",	"ピッチカート"): 45,
    ("Orchestral Harp",	"ハープ"): 46,
    ("Timpani",	"ティンパニ"): 47,
    ("String Ensemble 1",	"ストリングアンサンブル"): 48,
    ("String Ensemble 2",	"スローストリングアンサンブル"): 49,
    ("Synth Strings 1",	"シンセストリングス"): 50,
    ("Synth Strings 2",	"シンセストリングス2"): 51,
    ("Voice Aahs",	"声「アー」"): 52,
    ("Voice Oohs",	"声「ドゥー」"): 53,
    ("Synth Voice",	"シンセヴォイス"): 54,
    ("Orchestra Hit",	"オーケストラヒット"): 55,
    ("Trumpet",	"トランペット"): 56,
    ("Trombone",	"トロンボーン"): 57,
    ("Tuba",	"チューバ"): 58,
    ("Muted Trumpet",	"ミュートトランペット"): 59,
    ("French horn",	"フレンチ・ホルン"): 60,
    ("Brass Section",	"ブラスセクション"): 61,
    ("Synth Brass 1",	"シンセブラス 1"): 62,
    ("Synth Brass 2",	"シンセブラス 2"): 63,
    ("Soprano Sax",	"ソプラノサックス"): 64,
    ("Alto Sax",	"アルトサックス"): 65,
    ("Tenor Sax",	"テナーサックス"): 66,
    ("Baritone Sax",	"バリトンサックス"): 67,
    ("Oboe",	"オーボエ"): 68,
    ("English Horn",	"イングリッシュホルン"): 69,
    ("Bassoon",	"ファゴット"): 70,
    ("Clarinet",	"クラリネット"): 71,
    ("Piccolo",	"ピッコロ"): 72,
    ("Flute",	"フルート"): 73,
    ("Recorder",	"リコーダー"): 74,
    ("Pan Flute",	"パンフルート"): 75,
    ("Blown Bottle",	"ブロウンボトル（吹きガラス瓶）"): 76,
    ("Shakuhachi",	"尺八"): 77,
    ("Whistle",	"口笛"): 78,
    ("Ocarina",	"オカリナ"): 79,
    ("Lead 1 (square)",	"正方波"): 80,
    ("Lead 2 (sawtooth)",	"ノコギリ波"): 81,
    ("Lead 3 (calliope)",	"カリオペリード"): 82,
    ("Lead 4 (chiff)",	"チフリード"): 83,
    ("Lead 5 (charang)",	"チャランゴリード"): 84,
    ("Lead 6 (voice)",	"声リード"): 85,
    ("Lead 7 (fifths)",	"フィフスズリード"): 86,
    ("Lead 8 (bass + lead)",	"ベース + リード"): 87,
    ("Pad 1 (Fantasia)",	"ファンタジア"): 88,
    ("Pad 2 (warm)",	"ウォーム"): 89,
    ("Pad 3 (polysynth)",	"ポリシンセ"): 90,
    ("Pad 4 (choir)",	"クワイア"): 91,
    ("Pad 5 (bowed)",	"ボウ"): 92,
    ("Pad 6 (metallic)",	"メタリック"): 93,
    ("Pad 7 (halo)",	"ハロー"): 94,
    ("Pad 8 (sweep)",	"スウィープ"): 95,
    ("FX 1 (rain)",	"雨"): 96,
    ("FX 2 (soundtrack)",	"サウンドトラック"): 97,
    ("FX 3 (crystal)",	"クリスタル"): 98,
    ("FX 4 (atmosphere)",	"アトモスフィア"): 99,
    ("FX 5 (brightness)",	"ブライトネス"): 100,
    ("FX 6 (goblins)",	"ゴブリン"): 101,
    ("FX 7 (echoes)",	"エコー[要曖昧さ回避]"): 102,
    ("FX 8 (sci-fi)",	"サイファイ"): 103,
    ("Sitar",	"シタール"): 104,
    ("Banjo",	"バンジョー"): 105,
    ("Shamisen",	"三味線"): 106,
    ("Koto",	"琴"): 107,
    ("Kalimba",	"カリンバ"): 108,
    ("Bagpipe",	"バグパイプ"): 109,
    ("Fiddle",	"フィドル"): 110,
    ("Shanai",	"シャハナーイ"): 111,
    ("Tinkle Bell",	"ティンクルベル"): 112,
    ("Agogo",	"アゴゴ"): 113,
    ("Steel Drums",	"スチールドラム"): 114,
    ("Woodblock",	"ウッドブロック"): 115,
    ("Taiko Drum",	"太鼓"): 116,
    ("Melodic Tom",	"メロディックタム"): 117,
    ("Synth Drum",	"シンセドラム"): 118,
    ("Reverse Cymbal",	"逆シンバル"): 119,
    ("Guitar Fret Noise",	"ギターフレットノイズ"): 120,
    ("Breath Noise",	"ブレスノイズ"): 121,
    ("Seashore",	"海岸"): 122,
    ("Bird Tweet",	"鳥の囀り"): 123,
    ("Telephone Ring",	"電話のベル"): 124,
    ("Helicopter",	"ヘリコプター"): 125,
    ("Applause",	"拍手"): 126,
    ("Gunshot",	"銃声"): 127,
}
DICT_FOR_PROGRAM_onlyFF14 = {
    # Ver. Patch5.4
    ("Orchestral Harp",	"ハープ"): (46, 60),
    ("Acoustic Piano",	"アコースティックピアノ", "グランドピアノ"): (0, 72),
    ("Acoustic Guitar (steel)",	"アコースティックギター（スチール弦）", "スチールギター"): (25, 48),
    ("Pizzicato Strings",	"ピッチカート", "ピチカート"): (45, 48),

    ("Flute",	"フルート"): (73, 72),
    ("Oboe",	"オーボエ"): (68, 72),
    ("Clarinet",	"クラリネット"): (71, 60),
    ("Piccolo",	"ピッコロ"): (72, 84),
    ("Pan Flute",	"パンフルート", "パンパイプ"): (75, 72),

    ("Timpani",	"ティンパニ", "ティンパニー"): (47, 36),
    ("Woodblock",	"ウッドブロック", "ボンゴ"): (115, 48),
    ("Melodic Tom",	"メロディックタム", "バスドラム+スネアドラム", "バスドラム"): (117, 48),
    ("Synth Drum",	"シンセドラム", "スネアドラム"): (118, 48),
    ("Reverse Cymbal",	"逆シンバル", "シンバル"): (119, 48),  # 53
    # ("Tinkle Bell",	"ティンクルベル", "シンバル"): 112,

    ("Trumpet",	"トランペット"): (56, 60),
    ("Trombone",	"トロンボーン"): (57, 48),
    ("Tuba",	"チューバ"): (58, 36),
    ("French horn",	"フレンチ・ホルン", "ホルン"): (60, 48),
    ("Alto Sax",	"アルトサックス", "サックス"): (65, 60),
    # ("Soprano Sax",	"ソプラノサックス"): 64,

    ("Violin",	"ヴァイオリン"): (40, 60),
    ("Viola",	"ヴィオラ"): (41, 60),
    ("Cello",	"チェロ"): (42, 48),
    ("Double bass",	"コントラバス"): (43, 36),
}

## FOR LOADER
DICT_FOR_QUANTIZED_UNIT_TIMES = {
    (2, 48): 20,  # 3連64分音符
    (3, 48): 30,  # 64分音符
    (4, 48): 40,  # 3連32分音符
    (6, 48): 60,  # 32分音符
    (8, 48): 80,  # 3連16分音符
    # (9, 48): 90,  # 符点32分音符
    (12, 48): 120,  # 16部音符
    (16, 48): 160,  # 3連8分音符
    # (18, 48): 180,  # 符点16分音符
    (24, 48): 240,  # 8分音符
    (32, 48): 320,  # 3連4部音符
    # (36, 48): 360,  # 符点8分音符
    (48, 48): 480,  # 4部音符
}
DICT_FOR_DRUM = {
    "バスドラム": {
        50: 71,  # High Tom
        48: 67,  # Hi Mid Tom
        47: 65,  # Low Mid Tom
        45: 59,  # Low Tom
        43: 55,  # High Floor Tom
        41: 53,  # Low Floor Tom
        36: 43,  # Bass Drum 1
        35: 38,  # Acou Bass Drum
    },
    "スネアドラム": {
        37: 72,  # Side Stick (違うかも…)
        39: 60,  # Hand Clap (違うかも…)
        40: 55,  # Electoric Snare
        38: 50,  # Acoustic Snare
    },
    "シンバル": {  # 自信ありません
        44: 72,  # Pedal Hi-Hat (違うかも…)
        55: 69,  # Splash Cymbal
        59: 67,  # Ride Cymbal 2 (違うかも…)
        51: 65,  # Ride Cymbal 1 (違うかも…)
        57: 64,  # Crash Cymbal 2
        42: 62,  # Closed Hi-Hat (違うかも…)
        49: 60,  # Crash Cymbal 1
        53: 59,  # Ride Bell (違うかも…)
        46: 52,  # Open Hi-Hat
        52: 47,  # Chinese Cymbal
    },
    "バスドラム+スネアドラム": {
        50: 71,  # High Tom
        48: 67,  # Hi Mid Tom
        47: 65,  # Low Mid Tom
        45: 59,  # Low Tom
        43: 55,  # High Floor Tom
        41: 53,  # Low Floor Tom
        36: 43,  # Bass Drum 1
        35: 38,  # Acou Bass Drum

        40: 71,  # Electoric Snare
        38: 64,  # Acoustic Snare
    }
}