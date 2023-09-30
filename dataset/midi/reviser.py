# coding: utf-8
import operator
import os
from typing import List
from typing import Optional
from typing import Tuple
from typing import Union

MIDI_HEADER_CHUNK = b"MThd"
MIDI_TRACK_CHUNK = b"MTrk"
LT, LE, EQ, NE, GE, GT, AND, OR = (
    operator.lt,
    operator.le,
    operator.eq,
    operator.ne,
    operator.ge,
    operator.gt,
    operator.and_,
    operator.or_,
)


class MidiReviser:
    """
    midoライブラリや他のmidi編集ソフトウェアで読めないmidiファイルを修正する
    例えば、１トラックの中に複数チャンネルが混在する場合などが当てはまる
    """
    def __init__(self, midi_path: str):
        """
        midoライブラリや他のmidi編集ソフトウェアで読めないmidiファイルを修正する
        例えば、１トラックの中に複数チャンネルが混在する場合などが当てはまる

        Parameters
        ----------
        midi_path : str
            midiファイルのパス
        """
        self.midi_path = midi_path
        self.midi_bytes = b""
        self.num_bytes_header = -1
        self.num_bytes_tracks = []
        self.chunks = []
        self.parsed_midi_bytes = []

        # midiファイルを読み込む
        self._open_midi()
        # 解析し、修正したかどうかを返す
        self.is_revised = self._parse_midi()

    def _open_midi(self):
        """
        midiファイルをバイナリで読み込む
        """
        with open(self.midi_path, "rb") as f:
            self.midi_bytes = f.read(-1)

    def _parse_midi(self) -> bool:
        """
        midiファイルのバイナリを解析し、修正したかどうかを返す
        
        returns
        -------
        bool
            修正したかどうか
        """
        self.chunks = self._file_to_chunks()
        file_is_revised = False

        # header
        parsed_header = self._parse_header(self.chunks[0])
        self.parsed_midi_bytes.append(parsed_header)

        # tracks
        num_channel = 0
        for chunk in self.chunks[1:]:
            parsed_chunk, chunk_is_revised, exist_channel_event = self._parse_track(num_channel, chunk)
            if exist_channel_event:
                num_channel += 1
            self.parsed_midi_bytes.append(parsed_chunk)
            file_is_revised = (file_is_revised or chunk_is_revised)

        return file_is_revised
        
    def _file_to_chunks(self) -> List[bytes]:
        """
        midiファイル全体のバイナリ列をチャンク単位（ヘッダー、トラック）に分割する
        
        returns
        -------
        List[bytes]
            チャンク単位に分割したバイト列
        """
        current_num_bytes = 0
        self.num_bytes_header = -1
        self.num_bytes_tracks = []
        chunks = []
        try:
            is_header = False
            is_track = False
            bundle_bytes = b""
            while current_num_bytes < len(self.midi_bytes):
                if len(self.midi_bytes) < current_num_bytes + 4:
                    break
                target_bytes = self.midi_bytes[current_num_bytes:current_num_bytes + 4]
                next_num_bytes = current_num_bytes + 4
                if not is_header and not is_track:
                    if target_bytes == MIDI_HEADER_CHUNK:
                        is_header = True
                    elif target_bytes == MIDI_TRACK_CHUNK:
                        is_track = True
                    bundle_bytes += target_bytes
                else:
                    if is_header:
                        self.num_bytes_header = int.from_bytes(target_bytes, "big")
                        bundle_bytes += self.midi_bytes[current_num_bytes:current_num_bytes + 4 + self.num_bytes_header]
                        next_num_bytes = current_num_bytes + 4 + self.num_bytes_header
                        is_header = False
                    elif is_track:
                        track_byte_length = int.from_bytes(target_bytes, "big")
                        self.num_bytes_tracks.append(track_byte_length)
                        bundle_bytes += self.midi_bytes[current_num_bytes:current_num_bytes + 4 + track_byte_length]
                        next_num_bytes = current_num_bytes + 4 + track_byte_length
                        is_track = False
                    else:
                        raise Exception(f"Unknown bytes: {target_bytes}")
                    chunks.append(bundle_bytes)
                    bundle_bytes = b""
                current_num_bytes = next_num_bytes
        except Exception as e:
            print("Error: Midiファイルをチャンク単位に分割できませんでした。")
            import traceback
            traceback.print_exc()
        return chunks

    def _parse_header(self, chunk: bytes) -> List[bytes]:
        """
        ヘッダーを解析する

        Parameters
        ----------
        chunk : bytes
            ヘッダーチャンク

        Returns
        -------
        List[bytes]
            解析されたヘッダーチャンク
        """
        parsed_chunk = []
        try:
            parsed_chunk = [
                chunk[:4],     # chunk type
                chunk[4:8],    # chunk length
                chunk[8:10],   # format
                chunk[10:12],  # track num
                chunk[12:14],  # tick resolution
            ]
        except Exception as e:
            print("Error: Midiファイルをチャンク単位に分割できませんでした。")
            import traceback
            traceback.print_exc()
        return parsed_chunk

    def _parse_track(self, num_channel: int, chunk: bytes) -> Tuple[List[bytes], bool, bool]:
        """
        トラックを解析する。
        修正の必要があれば、解析中に修正する。
           (詳細説明)
           基本的にメッセージはそのまま
           ただし、ポート出力はAに修正する
           ただし、チャンネルボイスメッセージまたはチャンネルモードメッセージは、１トラックの中で１チャンネルにしか出さないように修正する
           ただし、ノートONのメッセージは１トラックの中では１チャンネルにしか出さないように修正する
           修正した場合は修正フラグをTrueで返す

        Parameters
        ----------
        num_channel: int
            チャンネル番号
        chunk : bytes
            トラックチャンク

        Returns
        -------
        List[bytes]
            解析されたトラックチャンク
        bool
            修正したかどうか
        bool
            チャンネル操作するイベントがあったかどうか
        """
        parsed_chunk = []
        is_revised = False
        exist_channel_event = False
        try:
            parsed_chunk = [chunk[:4], chunk[4:8]]  # chunk_type, chunk_length
            print("parsed_chunk ", parsed_chunk)
            current_num_bytes = 8
            buffer = b""
            while current_num_bytes < len(chunk):
                # print("current_num_bytes: ", current_num_bytes)
                # print("len(chunk): ", len(chunk))
                # print("parsed_chunk[-1] ", parsed_chunk[-1])
                # delta time
                num_len_bytes, _ = self._get_length_bytes(chunk, current_num_bytes)
                buffer += chunk[current_num_bytes:current_num_bytes + num_len_bytes]
                current_num_bytes += num_len_bytes

                val = chunk[current_num_bytes]
                if _op(val, EQ, 0xff):  # meta event
                    buffer += chunk[current_num_bytes:current_num_bytes + 1]
                    current_num_bytes += 1
                    if _op(chunk[current_num_bytes], EQ, 0x00):  # sequence
                        buffer += chunk[current_num_bytes:current_num_bytes + 4]
                        current_num_bytes += 4
                    elif _op(chunk[current_num_bytes], LE, 0x07):  # text, copyright, name, instrument, lylic, marker, queue
                        buffer += chunk[current_num_bytes:current_num_bytes + 1]
                        current_num_bytes += 1
                        num_len_bytes, num_val_bytes = self._get_length_bytes(chunk, current_num_bytes)
                        buffer += chunk[current_num_bytes:current_num_bytes + num_len_bytes + num_val_bytes]
                        current_num_bytes += (num_len_bytes + num_val_bytes)
                    elif _op(chunk[current_num_bytes], EQ, 0x20):  # midi channel prefix
                        buffer += chunk[current_num_bytes:current_num_bytes + 3]
                        current_num_bytes += 3
                    elif _op(chunk[current_num_bytes], EQ, 0x21):  # port
                        buffer += chunk[current_num_bytes:current_num_bytes + 2]
                        current_num_bytes += 2
                        num_port = chunk[current_num_bytes]
                        if num_port != 0:
                            buffer += b"\x00"  # 必ずportAを使う
                            print("port")
                            is_revised = True
                        else:
                            buffer += chunk[current_num_bytes:current_num_bytes + 1]
                        current_num_bytes += 1
                    elif _op(chunk[current_num_bytes], EQ, 0x2f):  # midi channel prefix
                        print("end")
                        buffer += chunk[current_num_bytes:current_num_bytes + 2]
                        current_num_bytes += 2
                        parsed_chunk.append(buffer)
                        break
                    elif _op(chunk[current_num_bytes], EQ, 0x51):  # set tempo
                        buffer += chunk[current_num_bytes:current_num_bytes + 5]
                        current_num_bytes += 5
                    elif _op(chunk[current_num_bytes], EQ, 0x54):  # SMPTE offset
                        buffer += chunk[current_num_bytes:current_num_bytes + 7]
                        current_num_bytes += 7
                    elif _op(chunk[current_num_bytes], EQ, 0x58):  # beat, metronome
                        buffer += chunk[current_num_bytes:current_num_bytes + 6]
                        current_num_bytes += 6
                    elif _op(chunk[current_num_bytes], EQ, 0x59):  # key
                        buffer += chunk[current_num_bytes:current_num_bytes + 4]
                        current_num_bytes += 4
                    elif _op(chunk[current_num_bytes], EQ, 0x7f):  # original meta event
                        buffer += chunk[current_num_bytes:current_num_bytes + 1]
                        current_num_bytes += 1
                        num_len_bytes, num_val_bytes = self._get_length_bytes(chunk, current_num_bytes)
                        buffer += chunk[current_num_bytes:current_num_bytes + num_len_bytes + num_val_bytes]
                        current_num_bytes += (num_len_bytes + num_val_bytes)
                elif (
                    _op(val >> 4, EQ, 0x8) or  # note off
                    _op(val >> 4, EQ, 0x9) or  # note on
                    _op(val >> 4, EQ, 0xa) or  # polyphonic key pressure, after touch
                    _op(val >> 4, EQ, 0xb) or  # control change
                    _op(val >> 4, EQ, 0xe)     # pitch wheel change
                ):
                    exist_channel_event = True
                    channel_num = _op(val, AND, 0x0f)
                    if channel_num != num_channel:
                        buffer += (_op(val, AND, 0xf0) + num_channel).to_bytes(1, "big")
                        is_revised = True
                    else:
                        buffer += chunk[current_num_bytes:current_num_bytes + 1]
                    current_num_bytes += 1
                    buffer += chunk[current_num_bytes:current_num_bytes + 2]
                    current_num_bytes += 2
                elif (
                    _op(val >> 4, EQ, 0xc) or  # program change
                    _op(val >> 4, EQ, 0xd)     # channel pressure
                ):
                    exist_channel_event = True
                    channel_num = _op(val, AND, 0x0f)
                    if channel_num != num_channel:
                        buffer += (_op(val, AND, 0xf0) + num_channel).to_bytes(1, "big")
                        print("PROGRAM", channel_num, num_channel)
                        is_revised = True
                    else:
                        buffer += chunk[current_num_bytes:current_num_bytes + 1]
                    current_num_bytes += 1
                    buffer += chunk[current_num_bytes:current_num_bytes + 1]
                    current_num_bytes += 1
                elif _op(val >> 4, EQ, 0xf):  # system common message, system realtime message
                    system_num = _op(val, AND, 0x0f)
                    buffer += chunk[current_num_bytes:current_num_bytes + 1]
                    current_num_bytes += 1
                    if system_num == 0:  # syetem exclusive start
                        # TODO: システムによってバイト数が1~3と幅があるらしい、1にしておく
                        buffer += chunk[current_num_bytes:current_num_bytes + 1]
                        current_num_bytes += 1
                    elif system_num == 1 or system_num == 3:  # midi time code, song select
                        buffer += chunk[current_num_bytes:current_num_bytes + 1]
                        current_num_bytes += 1
                    elif system_num == 2:  # song position
                        buffer += chunk[current_num_bytes:current_num_bytes + 2]
                        current_num_bytes += 2
                    elif system_num == 6 or system_num == 7 or system_num == 8 or system_num == 10 or system_num == 11 or system_num == 12 or system_num == 14 or system_num == 15:  # tune request, midi clock, start, continue, stop, active sensing, reset
                        pass
                    elif system_num == 4 or system_num == 5 or system_num == 9 or system_num == 13:  # undefined
                        pass
                else:  # midi 1.0規格書にも書いてない良くわからないイベント
                    # TODO: 0A 40 とか 5B 00 とか 5D 00 とか・・・経験的に2bytesのものばかりなので2bytesにしている
                    buffer += chunk[current_num_bytes:current_num_bytes + 2]
                    current_num_bytes += 2
                parsed_chunk.append(buffer)
                buffer = b""
        except Exception as e:
            print("Error: Midiファイルをチャンク単位に分割できませんでした。")
            import traceback
            traceback.print_exc()
        return parsed_chunk, is_revised, exist_channel_event

    def _get_length_bytes(self, chunk: bytes, current_num_bytes: int) -> Tuple[int, int]:
        """
        lengthの入っているバイト数と、値を得る
        8bitのうち、最上位ビットが次のbyteにも情報が続いていることを指すフラグになっている
        下位7bitには値が記載されており、すべての情報の下位7bitを連結したものが値になっている

        Parameters
        ----------
        chunk : bytes
            トラックチャンク
        current_num_bytes : int
            現在の読み進めたバイト数

        Returns
        -------
        int
            lengthの入っているバイト数
        int
            lengthの値
        """
        num_len_bytes = 0
        buffer = 0
        while _op(chunk[current_num_bytes + num_len_bytes], AND, 0x80) >> 7:
            buffer += _op(chunk[current_num_bytes + num_len_bytes], AND, 0x7f)
            buffer <<= 7
            num_len_bytes += 1
        buffer += chunk[current_num_bytes + num_len_bytes]
        num_len_bytes += 1
        return num_len_bytes, buffer

    def fwrite_revised_midi(self, filename: Optional[str] = None):
        """
        修正したmidiを書き出す

        Parameters
        ----------
        filename : Optional[str]
            midiファイル書き出し先
            Noneの場合は{ファイル名}_revised.midで書き出す
        """
        if filename is None:
            base, ext = os.path.splitext(self.midi_path)
            base += "_revised"
            filename = base + ext

        with open(filename, "wb") as f:
            for chunk in self.parsed_midi_bytes:
                for bytes in chunk:
                    f.write(bytes)


def _op(val1: Union[bytes, int], op: Union[LT, LE, EQ, NE, GE, GT, AND, OR], val2: Union[bytes, int]):
    _val1 = int.from_bytes(val1, "big") if type(val1) == bytes else val1
    _val2 = int.from_bytes(val2, "big") if type(val2) == bytes else val2
    return op(_val1, _val2)
