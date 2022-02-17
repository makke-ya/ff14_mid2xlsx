# coding: utf-8
from abc import ABCMeta, abstractmethod
import os.path as osp
from datetime import datetime
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as tkfd
from tkinter import messagebox as mess

from mid2xlsx import Mid2XlsxConverter
from xlsx2mid import Xlsx2MidConverter
from dataset._static_data import (
    DICT_FOR_PROGRAM_onlyFF14,
    KEY_NAME_LIST,
    STYLE_NAME_LIST,
)

# グローバル変数
FF14_INST_DICT = {
    instruments[-1]: "{}（{}-{}）".format(
        instruments[-1], base_pitch - 12, base_pitch + 24
    )
    for instruments, (_, base_pitch) in DICT_FOR_PROGRAM_onlyFF14.items()
}
FF14_INST_LIST = list(FF14_INST_DICT.keys())
FF14_INSTWPITCH_LIST = list(FF14_INST_DICT.values())
DRUM_DICT = {
    "-": None,
    "バスドラム系→バスドラム": "バスドラム",
    "スネアドラム系→スネアドラム": "スネアドラム",
    "シンバル系→シンバル": "シンバル",
    "バスドラム,スネアドラム系→バスドラム": "バスドラム+スネアドラム",
}
PITCH_LIST = [
    "0",
    "+12",
    "-12",
    "+24",
    "-24",
    "+36",
    "-36",
    "+48",
    "-48",
]


class FrameBase(tk.Frame):
    def __init__(self, master=None, bg="white"):
        super().__init__(master, bg=bg)
        self.bg = bg

    def add_msg(self, msg):
        self.master.add_msg(msg)


class LabelFrameBase(tk.LabelFrame):
    def __init__(self, master, text, bg="white", relief="groove", font=("MS Gothic", 9)):
        super().__init__(master, text=text, bg=bg, relief=relief, font=font)
        self.bg = bg

    @abstractmethod
    def state_on(self):
        pass

    @abstractmethod
    def state_off(self):
        pass


class RootApp(FrameBase):
    def __init__(
        self, master=None, title="test app", bg="snow",
    ):
        super().__init__(master, bg=bg)

        self.master.title(title)  # ウィンドウタイトルを指定
        self.master.resizable(0, 0)  # ウィンドウサイズの変更可否設定
        self.master.configure(bg=self.bg)  # ウィンドウの背景色

        self.configure(bg=self.bg)
        self.pack(padx=3, pady=3, fill="both", expand=2)

        self.main_frm = MainFrame(self, bg=self.bg)

        # self.pb = ttk.Progressbar(self)
        # self.pb.configure(value=8, maximum=10, orient="horizontal")
        # self.pb.pack(fill="x", pady=(0, 3))


class MainFrame(FrameBase):
    def __init__(self, master=None, bg="white", height=620):
        super().__init__(master, bg=bg)
        self.configure(height=height)
        self.pack(fil="x")
        self.left_frm = LeftFrame(self, bg=self.bg)
        self.right_frm = RightFrame(self, bg=self.bg)

    def add_msg(self, msg):
        self.right_frm.add_msg(msg)


class RightFrame(FrameBase):
    def __init__(self, master=None, bg="white"):
        super().__init__(master, bg=bg)
        self.pack(side="left", padx=5, fill="y")

        # 通知ラベル
        # self.msg_lbl = tk.Label(self.mgrid_frm)
        self.msg_lbl = tk.Label(self)
        self.msg_lbl.configure(
            text="通知（100件まで）", bg=self.bg, font=("MS Gothic", 9)
        )
        # self.msg_lbl.grid(row=0, column=0, sticky="nw")
        self.msg_lbl.pack(anchor="sw")

        # 通知リストのグリッド
        self.mgrid_frm = tk.Frame(self)
        self.mgrid_frm.configure(bg=self.bg)
        self.mgrid_frm.pack(fill="y")

        # 通知リスト
        self.msg_list = tk.Listbox(self.mgrid_frm)
        # self.msg_list = tk.Listbox(self)
        self.msg_list.configure(width=65, height=36)
        self.msg_list.grid(row=0, column=0, sticky="nwes")

        # Scrollbar(縦)
        self.vbar = ttk.Scrollbar(
            self.mgrid_frm,
            orient="vertical",
            command=self.msg_list.yview
        )
        self.msg_list['yscrollcommand'] = self.vbar.set
        self.vbar.grid(row=0, column=1, sticky="ns")

        # Scrollbar(横)
        self.hbar = ttk.Scrollbar(
            self.mgrid_frm,
            orient="horizontal",
            command=self.msg_list.xview
        )
        self.msg_list['xscrollcommand'] = self.hbar.set
        self.hbar.grid(row=1, column=0, sticky="we")


    def add_msg(self, msg):
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.msg_list.insert(tk.END, "{}  {}".format(now_str, msg))
        if msg.find("Error") != -1:
            self.msg_list.itemconfigure(tk.END, foreground="red")
        self.msg_list.yview_scroll(5, tk.UNITS)
        if self.msg_list.size() > 100:
            self.msg_list.delete(0)


class LeftFrame(FrameBase):
    def __init__(self, master=None, bg="white"):
        super().__init__(master, bg=bg)
        # 左のフレーム
        self.pack(side="left", padx=5, fill="y")

        # 枠線用のFrame(1)
        self.input_frm = InputFrame(
            self,
            bg=self.bg,
            text="1. MIDIファイルの入力",
        )
        self.midi_path = None

        # 枠線用のFrame(2)
        self.conf_frm = ConfFrame(
            self,
            bg=self.bg,
            text="2. 調・使用チャネル・楽器名・ピッチ変化量の決定",
        )
        self.estimated_dict = None

        # 枠線用のFrame(3)
        self.run_frm = RunFrame(
            self,
            bg=self.bg,
            text="3. 保存パラメータの決定・変換",
        )
        self.dirname = None

        # 変換インスタンス
        self.m2x_converter = Mid2XlsxConverter()
        self.x2m_converter = None
        self.program_dict = {0: "ハープ"}
        self.key = 0
        self.key_dict = {1: "C-dur"}
        self.pitch_dict = {0: [0, 0]}
        self.style = 0
        self.shorten = False

    def input_state_on(self):
        self.midi_path = None
        self.input_frm.state_on()

    def input_state_off(self, path):
        self.midi_path = path
        self._fopen()
        self.input_frm.state_off()

    def conf_state_on(self, program_dict=None):
        if program_dict is not None:
            self.program_dict = program_dict
        self._estimate()
        key_idx = KEY_NAME_LIST.index(self.key)
        on_dict = self._get_on_dict()
        self.conf_frm.state_on(key_idx, on_dict)

    def conf_state_off(self):
        self.program_dict = {0: "ハープ"}
        self.key = ""
        self.key_dict = {1: "C-dur"}
        self.pitch_dict = {0: [0, 0]}
        self.conf_frm.state_off()

    def run_state_on(self):
        self.run_frm.state_on()

    def run_state_off(self):
        self.run_frm._enter_filename("未選択/test", False)
        self.run_frm.state_off()

    def _fopen(self):
        try:
            self.program_dict = self.m2x_converter.fopen(self.midi_path)
            self.run_frm._enter_filename(self.midi_path)
            self.add_msg("MIDIファイル読み込み成功")
        except:
            self.add_msg("Error: MIDIファイル読み込み失敗")

    def _estimate(self):
        key = ""
        key_dict = {1: "C-dur"}
        pitch_dict = {0: [0, 0]}
        self.add_msg("自動推定中...")
        try:
            key_dict = self.m2x_converter.key_estimate()
            pitch_dict = self.m2x_converter.pitch_estimate(self.program_dict)
            key = key_dict[1]
            self.run_state_on()
            self.add_msg("自動推定成功")
            self.add_msg("    [調] {}".format(key))
            for channel_num, _pitch_list in sorted(
                pitch_dict.items(), key=lambda x: x[0]
            ):
                if channel_num != 9:
                    shift_pitch, (max_pitch, min_pitch, outlier_num) = _pitch_list
                    self.add_msg(
                        "    [channel{:02d}] midiピッチ範囲{}-{}（範囲外の音の数: {}）".format(
                            channel_num + 1, min_pitch, max_pitch, outlier_num
                        )
                    )
                else:
                    for shift_pitch, (max_pitch, min_pitch, outlier_num) in _pitch_list:
                        self.add_msg(
                            "    [channel{:02d}] midiピッチ範囲{}-{}（範囲外の音の数: {}）".format(
                                channel_num + 1, min_pitch, max_pitch, outlier_num
                            )
                        )
        except:
            import traceback
            traceback.print_exc()
            self.add_msg("Error: 自動推定が失敗")
        finally:
            self.key = key
            self.key_dict = key_dict
            self.pitch_dict = pitch_dict

    def _get_on_dict(self):
        on_dict = {}
        print(self.program_dict.items())
        for channel_num in self.program_dict.keys():
            program_str = self.program_dict[channel_num]
            if channel_num not in self.pitch_dict:
                continue
            if channel_num != 9:
                shift_pitch, _ = self.pitch_dict[channel_num]
                shift_pitch_str = "+{}".format(shift_pitch) if shift_pitch > 0 else "{}".format(shift_pitch)
                program_idx = FF14_INST_LIST.index(program_str)
                shift_pitch_idx = PITCH_LIST.index(shift_pitch_str)
                on_dict[channel_num] = [program_idx, shift_pitch_idx]
            else:
                if channel_num not in on_dict.keys():
                    on_dict[channel_num] = []
                pitches = self.pitch_dict[channel_num]
                for drum_mode, (shift_pitch, _) in zip(program_str, pitches):
                    shift_pitch_str = "+{}".format(shift_pitch) if shift_pitch > 0 else "{}".format(shift_pitch)
                    program_idx = FF14_INST_LIST.index(drum_mode)
                    shift_pitch_idx = PITCH_LIST.index(shift_pitch_str)
                    on_dict[channel_num].append([program_idx, 0])
        return on_dict

    def _update(self, program_dict, key_dict, pitch_dict, drum_modes, style, shorten):
        self.add_msg("更新中...")
        print(self.program_dict)
        try:
            _program_dict, key_dict, pitch_dict = self.m2x_converter.update(
                program_dict, key_dict, pitch_dict, drum_modes, style
            )
            key = key_dict[1]
            self.run_state_on()
            self.add_msg("更新成功")
            self.add_msg("    [調] {}".format(key))
            print(pitch_dict)
            for channel_num, _pitch_list in sorted(
                pitch_dict.items(), key=lambda x: x[0]
            ):
                if channel_num != 9:
                    shift_pitch, (max_pitch, min_pitch, outlier_num) = _pitch_list
                    self.add_msg(
                        "    [channel{:02d}] midiピッチ範囲{}-{}（範囲外の音数: {}）".format(
                            channel_num + 1, min_pitch, max_pitch, outlier_num
                        )
                    )
                else:
                    for shift_pitch, (max_pitch, min_pitch, outlier_num) in _pitch_list:
                        self.add_msg(
                            "    [channel{:02d}] midiピッチ範囲{}-{}（範囲外の音数: {}）".format(
                                channel_num + 1, min_pitch, max_pitch, outlier_num
                            )
                        )
        except:
            self.add_msg("Error: 更新失敗")
            import traceback; traceback.print_exc()
        finally:
            self.key = key
            self.key_dict = key_dict
            self.pitch_dict = pitch_dict
            self.style = style
            self.program_dict = program_dict
            self.shorten = shorten

    def _fwrite(
        self, filename, title_name=None, again_cvt=False,
        start_measure_num=1, num_measures_in_system=4, score_width=29.76,
    ):
        self.conf_frm.update()
        self.add_msg("変換中...")
        try:
            on_dict = self._get_on_dict()
            on_list = list(on_dict.keys())
            xlsx_name = filename + ".xlsx"
            self.m2x_converter.fwrite(
                xlsx_name, title_name, on_list,
                style=self.style,
                shorten=self.shorten,
                start_measure_num=start_measure_num,
                num_measures_in_system=num_measures_in_system,
                score_width=score_width,
            )
            self.master.add_msg("変換完了")
            self.master.add_msg("    Xlsx: {}".format(xlsx_name))

            if again_cvt is True:
                midi_name = filename + ".mid"
                tempo = self.m2x_converter.get_tempo()
                rhythm_dict = self.m2x_converter.get_rhythm_dict()
                self.x2m_converter = Xlsx2MidConverter(
                    tempo=tempo, rhythm_dict=rhythm_dict
                )
                self.x2m_converter.add_common_data_list(
                    self.m2x_converter.common_data_list, on_list
                )
                self.x2m_converter.fwrite(midi_name)
                self.master.add_msg("    Mid: {}".format(midi_name))
            res = mess.showinfo("Message", "変換が成功しました")
        except:
            self.master.add_msg("Error: 変換失敗")
            res = mess.showinfo("Message", "変換が失敗しました")
            import traceback; traceback.print_exc()


class InputFrame(LabelFrameBase):
    def __init__(self, master, text, bg="white", relief="groove", font=("MS Gothic", 9)):
        super().__init__(master, text=text, bg=bg, relief=relief, font=font)
        self.pack(pady=(0, 5), fill="x")

        # 読み込んだ画像リスト
        self.midi_list = ttk.Treeview(self)
        self.midi_list.configure(column=(1,), show="headings", height=1)
        self.midi_list.column(1, width=731)
        self.midi_list.heading(1, text="path/name")
        self.midi_list.pack(padx=5)

        # 入力用のフレーム
        self.input_btn_frm = tk.Frame(self)
        self.input_btn_frm.configure(bg=self.bg)
        self.input_btn_frm.pack(pady=5)

        # midi追加ボタン
        self.add_btn = tk.Button(self.input_btn_frm)
        self.add_btn.configure(
            text="MIDIファイルを追加",
            command=self.midi_add,
            font=("MS Gothic", 9),
        )
        self.add_btn.pack(side="left", padx=5)

        # midiリセットボタン
        self.reset_btn = tk.Button(self.input_btn_frm)
        self.reset_btn.configure(
            text="MIDIファイルをリセット",
            command=self.midi_reset,
            font=("MS Gothic", 9),
        )
        self.reset_btn.pack(side="left", padx=5)

    def midi_add(self):
        f_conf = [("Midi File", ("mid",))]
        path = tkfd.askopenfile(filetypes=f_conf)
        if path is not None:
            self.master.input_state_off(path.name)
            self.midi_list.insert("", "end", values=path.name)
            self.master.conf_state_on()
            self.master.add_msg("MIDIファイルをセットしました")

    def midi_reset(self):
        self.master.run_state_off()
        self.master.conf_state_off()
        for i in self.midi_list.get_children():
            self.midi_list.delete(i)
        self.add_btn.configure(state="normal")
        self.master.input_state_on()
        self.master.add_msg("MIDIファイルをリセットしました")

    def state_on(self):
        self.add_btn.configure(state="normal")

    def state_off(self):
        self.add_btn.configure(state="disabled")


class ConfFrame(LabelFrameBase):
    def __init__(self, master, text, bg="white", relief="groove", font=("MS Gothic", 9)):
        super().__init__(master, text=text, bg=bg, relief=relief, font=font)
        self.pack(pady=(0, 5), fill="x")

        # グリッド用のFrame
        self.kgrid_frm = tk.Frame(self)
        self.kgrid_frm.configure(bg=self.bg)
        self.kgrid_frm.pack(padx=3, pady=8, fill="x")

        # 記譜の調
        self.key_lbl = tk.Label(self.kgrid_frm)
        self.key_lbl.configure(text="記譜の調", bg=self.bg, font=("MS Gothic", 9))
        self.key_lbl.grid(row=0, column=0, sticky="nw", padx=(2, 0))
        self.key_combo = ttk.Combobox(self.kgrid_frm)
        self.key_combo.configure(
            width=16,
            state="disabled",
            value=tuple(KEY_NAME_LIST),
            font=("MS Gothic", 9),
        )
        self.key_combo.current(0)
        self.key_combo.grid(row=1, column=0, sticky="w", padx=(18, 0))

        # 和音の記譜法
        self.chord_lbl = tk.Label(self.kgrid_frm)
        self.chord_lbl.configure(
            text="記譜法", bg=self.bg, font=("MS Gothic", 9),
        )
        self.chord_lbl.grid(padx=(10, 0), row=0, column=1, sticky="nw")
        self.style_combo = ttk.Combobox(self.kgrid_frm)
        self.style_combo.configure(
            width=16,
            state="disabled",
            value=tuple(STYLE_NAME_LIST),
            font=("MS Gothic", 9),
        )
        self.style_combo.current(0)
        self.style_combo.grid(padx=(25, 0), row=1, column=1, sticky="w")

        # 音名
        self.shorten_lbl = tk.Label(self.kgrid_frm)
        self.shorten_lbl.configure(
            text="音名", bg=self.bg, font=("MS Gothic", 9),
        )
        self.shorten_lbl.grid(padx=(10, 0), row=0, column=2, sticky="nw")
        self.shorten_chkbox_var = tk.BooleanVar()
        self.shorten_chkbox = tk.Checkbutton(self.kgrid_frm)
        self.shorten_chkbox.configure(
            width=12,
            text="半角化",
            variable=self.shorten_chkbox_var,
            state="disabled",
            bg=self.bg,
            font=("MS Gothic", 9),
        )
        self.shorten_chkbox.grid(padx=(3, 0), row=1, column=2, sticky="w")

        # アドバンスド設定のフレーム
        # self.advanced_settings_frm = tk.Frame(self)
        # self.advanced_settings_frm.configure(bg=self.bg)
        # self.advanced_settings_frm.pack(padx=5, pady=3)

        # 変換開始小節数
        self.advanced_setting_lbl = tk.Label(self.kgrid_frm)
        self.advanced_setting_lbl.configure(
            bg=self.bg, text="アドバンスド設定",
            font=("ms gothic", 9),
        )
        self.advanced_setting_lbl.grid(row=0, column=3, sticky="w")
        self.advanced_setting_frm = tk.Frame(self.kgrid_frm)
        self.advanced_setting_frm.configure(bg=self.bg)
        self.advanced_setting_frm.grid(row=1, column=3, sticky="w", padx=18)
        self.advanced_setting_ent = tk.Entry(self.advanced_setting_frm)
        self.advanced_setting_ent.configure(width=52)
        self.advanced_setting_ent.insert(0, "")
        self.advanced_setting_ent.configure(state="disabled")
        self.advanced_setting_ent.pack(side="left")

        # グリッド用のFrame
        self.grid_frm = tk.Frame(self)
        self.grid_frm.configure(bg=self.bg)
        self.grid_frm.pack(padx=3, pady=5, fill="x")

        # 表のタイトルをつける
        self.nondrum_lbl = tk.Label(self.grid_frm)
        self.nondrum_lbl.configure(
            text="ドラム以外", bg=self.bg,
            font=("MS Gothic", 9),
        )
        self.nondrum_lbl.grid(row=0, column=0, sticky="nw")
        self.left_inst_lbl = tk.Label(self.grid_frm)
        self.left_inst_lbl.configure(
            text="楽器名", bg=self.bg, 
            font=("MS Gothic", 9),
        )
        self.left_inst_lbl.grid(row=0, column=1, sticky="s")
        self.left_pitch_lbl = tk.Label(self.grid_frm)
        self.left_pitch_lbl.configure(
            text="ピッチ変化", bg=self.bg,
            font=("MS Gothic", 9),
        )
        self.left_pitch_lbl.grid(row=0, column=2, sticky="s")
        self.right_inst_lbl = tk.Label(self.grid_frm)
        self.right_inst_lbl.configure(
            text="楽器名", bg=self.bg, 
            font=("MS Gothic", 9),
        )
        self.right_inst_lbl.grid(row=0, column=4, sticky="s")
        self.right_pitch_lbl = tk.Label(self.grid_frm)
        self.right_pitch_lbl.configure(
            text="ピッチ変化", bg=self.bg,
            font=("MS Gothic", 9),
        )
        self.right_pitch_lbl.grid(row=0, column=5, sticky="s")
        # self.over_lbl = tk.Label(self.grid_frm)
        # self.over_lbl.configure(text="範囲外の音数", bg=self.bg)
        # self.over_lbl.grid(row=0, column=3, sticky="s")

        # ドラム以外
        self.chkbox_var_list = [tk.BooleanVar() for _ in range(16)]
        self.chkbox_list = [tk.Checkbutton(self.grid_frm) for _ in range(16)]
        self.inst_combo_list = [ttk.Combobox(self.grid_frm) for _ in range(16)]
        self.pitch_combo_list = [ttk.Combobox(self.grid_frm) for _ in range(16)]
        row_margin = 0
        column_margin = 0
        for i in range(16):
            if i == 9:
                continue
            if i > 7:
                row_margin = -8
                column_margin = 3
            
            # チェックボックス設置
            self.chkbox_list[i].configure(
                width=12,
                text="Channel{:02d}".format(i + 1),
                variable=self.chkbox_var_list[i],
                state="disabled",
                bg=self.bg,
                font=("MS Gothic", 9),
            )
            self.chkbox_list[i].grid(row=row_margin + i + 1, column=column_margin, sticky="w")
            # 楽器セレクト
            self.inst_combo_list[i].configure(
                width=25,
                state="disabled",
                value=tuple(FF14_INST_DICT.values()),
                font=("MS Gothic", 9),
            )
            self.inst_combo_list[i].current(0)
            self.inst_combo_list[i].grid(row=row_margin + i + 1, column=column_margin + 1, sticky="w")
            # ピッチ変化セレクト
            self.pitch_combo_list[i].configure(
                width=12,
                state="disabled",
                value=tuple(PITCH_LIST),
                font=("MS Gothic", 9),
            )
            self.pitch_combo_list[i].current(0)
            self.pitch_combo_list[i].grid(row=row_margin + i + 1, column=column_margin + 2, sticky="w")

        # ドラム
        self.dgrid_frm = tk.Frame(self)
        self.dgrid_frm.configure(bg=self.bg)
        self.dgrid_frm.pack(padx=3, pady=3, fill="x")

        # 表のタイトルをつける
        self.nondrum_lbl = tk.Label(self.dgrid_frm)
        self.nondrum_lbl.configure(text="ドラム", bg=self.bg, font=("MS Gothic", 9))
        self.nondrum_lbl.grid(row=0, column=0, sticky="nw")
        self.dinst_lbl = tk.Label(self.dgrid_frm)
        self.dinst_lbl.configure(text="抽出方法", bg=self.bg, font=("MS Gothic", 9))
        self.dinst_lbl.grid(row=0, column=1, sticky="s")
        self.dchkbox_var = tk.BooleanVar()
        self.dchkbox = tk.Checkbutton(self.dgrid_frm)
        self.dchkbox.configure(
            width=12,
            text="Channel10",
            variable=self.dchkbox_var,
            state="disabled",
            bg=self.bg,
            font=("ms gothic", 9),
        )
        self.dchkbox.grid(row=1, column=0, sticky="w")
        self.drum_combo_list = [ttk.Combobox(self.dgrid_frm) for _ in range(3)]
        for i in range(3):
            # 抽出楽器セレクト
            self.drum_combo_list[i].configure(
                width=38,
                state="disabled",
                value=tuple(DRUM_DICT.keys()),
                font=("ms gothic", 9),
            )
            self.drum_combo_list[i].current(0)
            self.drum_combo_list[i].grid(row=i + 1, column=1, sticky="w")

        # 設定ボタンのフレーム
        self.conf_btn_frm = tk.Frame(self)
        self.conf_btn_frm.configure(bg=self.bg)
        self.conf_btn_frm.pack(pady=5)

        # 自動推定ボタン
        self.estimate_btn = tk.Button(self.conf_btn_frm)
        self.estimate_btn.configure(
            text="自動推定",
            state="disabled",
            command=self.estimate,
            font=("ms gothic", 9),
        )
        self.estimate_btn.pack(side="left", padx=5)

        # 更新ボタン
        self.update_btn = tk.Button(self.conf_btn_frm)
        self.update_btn.configure(
            text="更新",
            state="disabled",
            command=self.update,
            font=("ms gothic", 9),
        )
        self.update_btn.pack(side="left", padx=5)

    def state_on(self, key_idx=0, on_dict={i: [1, 1] for i in range(16)}, ):
        print("ON_DICT:", on_dict)
        self.key_combo.configure(state="readonly")
        self.key_combo.current(key_idx)
        self.style_combo.configure(state="readonly")
        self.shorten_chkbox.configure(state="normal")
        self.shorten_chkbox_var.set(False)
        for i in range(16):
            if i == 9 or i not in on_dict.keys():
                continue
            self.chkbox_list[i].configure(state="normal")
            self.chkbox_var_list[i].set(True)
            self.inst_combo_list[i].configure(state="readonly")
            self.inst_combo_list[i].current(on_dict[i][0])
            self.pitch_combo_list[i].configure(state="readonly")
            self.pitch_combo_list[i].current(on_dict[i][1])
        if 9 in on_dict.keys():
            self.dchkbox.configure(state="normal")
            self.dchkbox_var.set(True)
            for i in range(3):
                self.drum_combo_list[i].configure(state="readonly")
                self.drum_combo_list[i].current(0)
        self.estimate_btn.configure(state="normal")
        self.update_btn.configure(state="normal")
        self.advanced_setting_ent.configure(state="normal")

    def state_off(self):
        self.key_combo.configure(state="disabled")
        self.key_combo.current(0)
        self.style_combo.configure(state="disabled")
        self.style_combo.current(0)
        self.shorten_chkbox.configure(state="disabled")
        self.shorten_chkbox_var.set(False)
        for i in range(16):
            if i == 9:
                continue
            self.chkbox_list[i].configure(state="disabled")
            self.chkbox_var_list[i].set(False)
            self.inst_combo_list[i].configure(state="disabled")
            self.inst_combo_list[i].current(0)
            self.pitch_combo_list[i].configure(state="disabled")
            self.pitch_combo_list[i].current(0)
        self.dchkbox.configure(state="disabled")
        self.dchkbox_var.set(False)
        for i in range(3):
            self.drum_combo_list[i].configure(state="disabled")
            self.drum_combo_list[i].current(0)
        self.estimate_btn.configure(state="disabled")
        self.update_btn.configure(state="disabled")
        self.advanced_setting_ent.configure(state="disabled")

    def estimate(self):
        program_dict = {}
        for i in range(16):
            if self.chkbox_var_list[i].get() is False:
                continue
            program_str_idx = FF14_INSTWPITCH_LIST.index(self.inst_combo_list[i].get())
            program_str = FF14_INST_LIST[program_str_idx]
            if i != 9:
                program_dict[i] = program_str
            else:
                if 9 not in program_dict.keys():
                    program_dict[9] = []
                program_dict[i].append(program_str)
        self.master.conf_state_on(program_dict)

    def update(self):
        key = self.key_combo.get()
        key_dict = {1: key}
        style = self.style_combo.get()
        shorten = self.shorten_chkbox_var.get()
        advanced_settings = self.advanced_setting_ent.get()
        if advanced_settings != "":
            splitted_settings = advanced_settings.strip().split(",")
            for settings in splitted_settings:
                num_measure, _key = settings.split(":")
                key_dict[int(num_measure)] = _key
                
        program_dict = {}
        pitch_dict = {}
        for i in range(16):
            if i == 9:
                continue
            if self.chkbox_var_list[i].get() is False:
                continue
            program_str_idx = FF14_INSTWPITCH_LIST.index(self.inst_combo_list[i].get())
            program_str = FF14_INST_LIST[program_str_idx]
            pitch = self.pitch_combo_list[i].get()
            program_dict[i] = program_str
            pitch_dict[i] = [int(pitch), 0]

        drum_modes = []
        if self.dchkbox_var.get() is True:
            program_dict[9] = []
            pitch_dict[9] = []
            for i in range(3):
                drum_key = self.drum_combo_list[i].get()
                if DRUM_DICT[drum_key] is None:
                    continue
                drum_str = DRUM_DICT[drum_key]
                program_dict[9].append(drum_str)
                pitch_dict[9].append([0, 0])
                drum_modes.append(drum_str)

        print(program_dict, key_dict, pitch_dict, drum_modes, style)
        self.master._update(
            program_dict, key_dict, pitch_dict, drum_modes, style, shorten
        )


class RunFrame(LabelFrameBase):
    def __init__(self, master, text, bg="white", relief="groove", font=("MS Gothic", 9)):
        super().__init__(master, text=text, bg=bg, relief=relief, font=font)
        self.pack(pady=(0, 5), fill="x")

        # 変換パラメータ用のグリッドFrame(1)
        self.pgrid_frm = tk.Frame(self)
        self.pgrid_frm.configure(bg=self.bg)
        self.pgrid_frm.pack(padx=3, pady=3, fill="x")

        # 変換開始小節数
        self.smnum_lbl = tk.Label(self.pgrid_frm)
        self.smnum_lbl.configure(
            bg=self.bg, text="MIDIの変換開始小節",
            font=("ms gothic", 9),
        )
        self.smnum_lbl.grid(row=0, column=0, sticky="e", padx=5)
        self.smnum_frm = tk.Frame(self.pgrid_frm)
        self.smnum_frm.configure(bg=self.bg)
        self.smnum_frm.grid(row=0, column=1, sticky="s", pady=3)
        self.smnum_ent = tk.Entry(self.smnum_frm)
        self.smnum_ent.configure(width=12)
        self.smnum_ent.insert(0, "1")
        self.smnum_ent.configure(state="disabled")
        self.smnum_ent.pack(side="left")

        # 譜面の幅
        self.width_frm = tk.Frame(self.pgrid_frm)
        self.width_frm.configure(bg=self.bg)
        self.width_frm.grid(row=0, column=2, sticky="w", padx=16)
        self.width_lbl = tk.Label(self.width_frm)
        self.width_lbl.configure(
            bg=self.bg, text="譜面の横幅",
            font=("ms gothic", 9),
        )
        self.width_lbl.grid(row=0, column=0, sticky="w")
        self.width_frm2 = tk.Frame(self.width_frm)
        self.width_frm2.configure(bg=self.bg)
        self.width_frm2.grid(row=0, column=1, sticky="w", padx=5, pady=3)
        self.width_ent = tk.Entry(self.width_frm2)
        self.width_ent.configure(width=12)
        self.width_ent.insert(0, "29.76")
        self.width_ent.configure(state="disabled")
        self.width_ent.pack()

        # 変換開始小節数
        self.numms_lbl = tk.Label(self.pgrid_frm)
        self.numms_lbl.configure(
            bg=self.bg, text="Xlsx一段の小節数",
            font=("ms gothic", 9),
        )
        self.numms_lbl.grid(row=0, column=3, sticky="e", padx=5)
        self.numms_frm = tk.Frame(self.pgrid_frm)
        self.numms_frm.configure(bg=self.bg)
        self.numms_frm.grid(row=0, column=4, sticky="s", pady=3)
        self.numms_ent = tk.Entry(self.numms_frm)
        self.numms_ent.configure(width=12)
        self.numms_ent.insert(0, "4")
        self.numms_ent.configure(state="disabled")
        self.numms_ent.pack(side="left")

        self.mchkbox_var = tk.BooleanVar()
        self.mchkbox = tk.Checkbutton(self.pgrid_frm)
        self.mchkbox.configure(
            text="変換後のXlsxからMidi生成",
            variable=self.mchkbox_var,
            state="disabled",
            bg=self.bg,
            font=("ms gothic", 9),
        )
        self.mchkbox.grid(padx=(15, 0), row=0, column=5, sticky="w")

        # グリッド用のFrame
        self.sgrid_frm = tk.Frame(self)
        self.sgrid_frm.configure(bg=self.bg)
        self.sgrid_frm.pack(padx=3, pady=3, fill="x")

        # 保存先タイトル
        self.acvTtlLbl = tk.Label(self.sgrid_frm)
        self.acvTtlLbl.configure(
            bg=self.bg, text="保存先を選択",
            font=("ms gothic", 9),
        )
        self.acvTtlLbl.grid(row=1, column=0, sticky="e", padx=5)
        self.acv_frm = tk.Frame(self.sgrid_frm)
        self.acv_frm.configure(bg=self.bg)
        self.acv_frm.grid(row=1, column=1, sticky="s", pady=3)
        self.acv_ent = tk.Entry(self.acv_frm)
        self.acv_ent.configure(width=102)
        self.acv_ent.insert(0, "未選択")
        self.acv_ent.configure(state="disabled")
        self.acv_ent.pack(side="left")

        # 保存先参照ボタン
        self.acv_btn = tk.Button(self.acv_frm)
        self.acv_btn.configure(
            text="参照",
            command=self.acv_open,
            font=("ms gothic", 9),
        )
        self.acv_btn.pack(side="left")  # 描画

        # ファイル名
        self.fname_lbl = tk.Label(self.sgrid_frm)
        self.fname_lbl.configure(
            text="ファイル名", bg=self.bg,
            font=("ms gothic", 9),
        )
        self.fname_lbl.grid(row=2, column=0, sticky="e", padx=5)

        # ファイル名入力欄
        self.fname_ent = tk.Entry(self.sgrid_frm)
        self.fname_ent.configure(width=102)
        self.fname_ent.insert("end", "test")
        self.fname_ent.grid(row=2, column=1, sticky="w")

        # グリッド用のFrame
        self.rgrid_frm = tk.Frame(self)
        self.rgrid_frm.configure(bg=self.bg)
        self.rgrid_frm.pack(padx=3, pady=3, fill="x")

        # 変換実行ボタン
        self.run_btn = tk.Button(self.rgrid_frm)
        self.run_btn.configure(
            text="変換実行",
            state="disabled",
            command=self.midi_trans,
            font=("ms gothic", 9),
        )
        self.run_btn.grid(padx=(350, 20), pady=(0, 5), row=0, column=0, sticky="w")

    def state_on(self):
        self.run_btn.configure(state="normal")
        self.mchkbox.configure(state="normal")
        self.acv_ent.configure(state="normal")
        self.smnum_ent.configure(state="normal")
        self.numms_ent.configure(state="normal")
        self.width_ent.configure(state="normal")

    def state_off(self):
        self.run_btn.configure(state="disabled")
        self.mchkbox.configure(state="disabled")
        self.acv_ent.configure(state="disabled")
        self.smnum_ent.configure(state="disabled")
        self.numms_ent.configure(state="disabled")
        self.width_ent.configure(state="disabled")

    def _enter_filename(self, filename, add_msg=True):
        dirname = osp.dirname(filename)
        basename = osp.basename(filename)
        self._enter_dirname(dirname, add_msg=add_msg)
        self._enter_basename(basename)

    def _enter_dirname(self, dirname, add_msg=True):
        self.acv_ent.configure(state="normal")
        self.acv_ent.delete(0, "end")
        if dirname != "":
            self.acv_ent.insert(0, dirname)
            if add_msg:
                self.master.add_msg("保存先を{0}に設定しました".format(dirname))
            self.master.dirname = dirname
        self.acv_ent.configure(state="disabled")

    def _enter_basename(self, basename):
        base, _ = osp.splitext(basename)
        base = base + "_score"
        self.fname_ent.delete(0, "end")
        self.fname_ent.insert("end", base)

    def acv_open(self):
        dirname = tkfd.askdirectory()
        if dirname != "":
            self._enter_dirname(dirname)
        else:
            self.master.add_msg("保存先未選択")

    def midi_trans(self):
        dirname = self.master.dirname
        if dirname is None:
            self.master.add_msg("保存先を選択してください")
            return None
        start_measure_num = self.smnum_ent.get()
        if (
            start_measure_num is None 
            or start_measure_num.isdigit() is False
            or int(start_measure_num) <= 0
        ):
            start_measure_num = 1
        else:
            start_measure_num = int(start_measure_num)
        self.smnum_ent.delete(0, "end")
        self.smnum_ent.insert(0, "{}".format(start_measure_num))

        num_measures_in_system = self.numms_ent.get()
        if (
            num_measures_in_system is None 
            or num_measures_in_system.isdigit() is False
            or int(num_measures_in_system) <= 1
        ):
            num_measures_in_system = 2
        else:
            num_measures_in_system = int(num_measures_in_system)
        self.numms_ent.delete(0, "end")
        self.numms_ent.insert(0, "{}".format(num_measures_in_system))
        xlsx_name = self.fname_ent.get()
        again_cvt = self.mchkbox_var.get()
        score_width = float(self.width_ent.get())
        if xlsx_name == "":
            self.master.add_msg("ファイル名を入力してください")
            return None
        msg_box = mess.askquestion("Message", "変換を開始しますか？")
        if msg_box == "yes":
            filename = "{}/{}".format(dirname, xlsx_name)
            self.master._fwrite(
                filename, 
                again_cvt=again_cvt,
                start_measure_num=start_measure_num,
                num_measures_in_system=num_measures_in_system,
                score_width=score_width,
            )


def main():
    root = tk.Tk()
    app = RootApp(master=root, title="Midi to Xlsx Converter")
    app.mainloop()


if __name__ == "__main__":
    main()
