import pytest
from dataset.midi.reviser import MidiReviser


# midiをバイナリで読み込む
def test_midiをバイナリで読み込む():
    midi_path = "tests/data/musescore3_test1.mid"
    reviser = MidiReviser(midi_path)
    assert len(reviser.midi_bytes) == 1047  # 1047はファイルのバイト数


# midiの解析をする
def test_ヘッダーとトラックごとに分割する():
    midi_path = "tests/data/musescore3_test1.mid"
    reviser = MidiReviser(midi_path)
    num_tracks = 8
    num_bytes_header = 6
    num_bytes_tracks = [138, 123, 117, 117, 114, 114, 123, 123]
    assert len(reviser.parsed_midi_bytes) == 1 + num_tracks  # 1はHeader
    assert reviser.num_bytes_header == num_bytes_header
    assert reviser.num_bytes_tracks == num_bytes_tracks
    for chunk in reviser.chunks:
        assert (chunk[:4] == b"MThd" or chunk[:4] == b"MTrk")


# 変換したかどうかは修正フラグに記載される
class Test_変換の必要があるmidiは修正フラグがTrue():
    def test_musescore3は修正フラグTrue(self):
        midi_path = "tests/data/musescore3_test1.mid"
        reviser = MidiReviser(midi_path)
        assert reviser.is_revised is True

    def test_musescore4は修正フラグTrue(self):
        midi_path = "tests/data/musescore4_test1.mid"
        reviser = MidiReviser(midi_path)
        assert reviser.is_revised is True

    def test_MIDI編集ソフトは修正フラグFalse(self):
        midi_path = "tests/data/domino_test1.mid"
        reviser = MidiReviser(midi_path)
        assert reviser.is_revised is False


# 修正したmidiを出力する
def test_修正したmidiを出力する():
    midi_path = "tests/data/musescore3_test1.mid"
    reviser = MidiReviser(midi_path)
    assert reviser.is_revised is True
    reviser.fwrite_revised_midi()
