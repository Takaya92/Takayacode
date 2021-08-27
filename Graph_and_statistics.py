import numpy as np
import pandas as pd
from pathlib import Path
import PySimpleGUI as sg
from scipy import stats as st
from statsmodels.stats.multicomp import pairwise_tukeyhsd
from dataclasses import dataclass, field
import matplotlib.pyplot as plt
from lifelines import KaplanMeierFitter
from lifelines.statistics import logrank_test
from PIL import Image
import openpyxl
import os

path = Path.cwd()
path_file = path / "file.csv"
file_file = np.genfromtxt("file.csv", delimiter=",", dtype='str')
path_setting = path / "setting.csv"
setting_file = np.genfromtxt("setting.csv", delimiter=",", dtype='str')
path_sample = path / "sample.csv"
sample_file = np.genfromtxt("sample.csv", delimiter=",", dtype='str')


# ↓↓↓↓↓↓↓↓class,def"ここから"↓↓↓↓↓↓↓↓#
@dataclass
class SampleInfo:
    name: str
    color: str
    style: str
    marker: str
    data_list: list[float] = field(default_factory=list)

    def data_list(self):
        return self.data_list

    def return_name(self):
        return self.name

    def return_color(self):
        return self.color

    def return_style(self):
        return self.style

    def return_marker(self):
        return self.marker

    def return_relative(self, adjust_data):
        if values["-ROWDATA-"]==True:
            relative_value = self.data_list
        else:
            relative_value = self.data_list / np.mean(adjust_data) * 100
        return relative_value

    def simple_bar_cal(self, adjust_data):
        if values["-ROWDATA-"]==True:
            simple_bar_relative = np.mean(self.data_list)
            simple_bar_sem = np.std(self.data_list) / np.sqrt(len(self.data_list))
        else:
            simple_bar_relative = np.mean(self.data_list) / np.mean(adjust_data) * 100
            simple_bar_sem = np.std(self.data_list) / np.sqrt(
                len(self.data_list)) / np.mean(adjust_data) * 100  # adjust_data=引数で補正and標準誤差SEM
        ax.bar(self.name, simple_bar_relative, edgecolor="k", color=self.color, align="center", label=self.name,
                   yerr=simple_bar_sem, ecolor="k", capsize=2, hatch=self.style)
        xtickslist.append(self.name)

    def multi_bar_cal_relative(self, adjust_data):
        if values["-ROWDATA-"] == True:
            multi_bar_relative = np.mean(self.data_list)
        else:
            multi_bar_relative = np.mean(self.data_list) / np.mean(adjust_data) * 100
        return multi_bar_relative

    def multi_bar_cal_sem(self, adjust_data):
        if values["-ROWDATA-"] == True:
            multi_bar_sem = np.std(self.data_list) / np.sqrt(len(self.data_list))
        else:
            multi_bar_sem = np.std(self.data_list) / np.sqrt(
                len(self.data_list)) / np.mean(adjust_data) * 100  # adjust_data=引数で補正and標準誤差SEM
        return multi_bar_sem

def t_test(s1, s2, n):
    t, p = st.ttest_ind(s1, s2, equal_var=False)
    empty_list = []
    print("\n{0}回目実験結果".format(n))
    print("Welch's t-test\nウェルチのt検定", "t値", t, "p値", p)
    exec("t_test_list{0}=['{0} th analysis: Welch t-test: p-value=','','',{1:.10f}]".format(n, p))
    exec("figure_statistics_sheet.append(t_test_list{0})".format(n))
    exec("figure_statistics_sheet.append({0})".format(empty_list))


def one_way_anova(n, *args):
    print("\n{0}回目実験結果".format(n))
    f, p = st.f_oneway(*args)
    print("One-Way-ANOVA\n一元配置分散分析の", "F値=", f, "p値=", p)
    exec("one_way_list{0}=['{0} th analysis:  One-Way-ANOVA  p-value=','','','',{1:.10f}]".format(n, p))
    exec("figure_statistics_sheet.append(one_way_list{0})".format(n))


def multiple_comparison_test(combined_all_sample_data, sample_names_and_len):
    empty_list = []
    tukey_result = pairwise_tukeyhsd(combined_all_sample_data, sample_names_and_len)
    print("\nTukey HSD\nテューキーのHSD多重検定", "\n", tukey_result)
    df_tukey = pd.DataFrame(data=tukey_result._results_table.data[1:], columns=tukey_result._results_table.data[0])
    figure_statistics_sheet.append(list(df_tukey.columns))
    for n in range(df_tukey.shape[0]):
        exec("figure_statistics_sheet.append(df_tukey.loc[{0}].tolist())".format(n))
    exec("figure_statistics_sheet.append({0})".format(empty_list))


# ↑↑↑↑↑↑↑↑class,def"ここまで,PysimpleGUI"ここから"↓↓↓↓↓↓↓↓#
radio_dic = {"-SURVIVAL-": "Survival_rate",
             "-FAT-": "Fat_accumulation\nBody_length",
             "-AGINGBAR-": "Aging_bar",
             "-AGINGLINE-": "Aging_line"}
tab1_layout = [
    [sg.Button("How to design file", key="-HOWDESIGNFILE-"), sg.Text(":データシート入力方法  "),
    sg.Button("Out put of results", key="-OUTPUTRESULTS-"), sg.Text(":結果の出力について  ")],
    [sg.Text("[ For survival ]"), sg.Text(":寿命解析用シート")],
    [sg.Text("Number of sample types:"),
     sg.Input(default_text=file_file[1, 0], key="-NUMBERSAMPLETYPES-",
              size=(3, 1)), sg.Text(":サンプルの種類")],
    [sg.Text("Number of sample size for each sample types:"),
     sg.Input(default_text=file_file[1, 1], key="-NUMBERSAMPLESIZE-",
              size=(4, 1)), sg.Text(":それぞれの種類のサンプルサイズ")],
    [sg.Text("Sample names:")],
    [sg.Text("1:"),
     sg.Input(default_text=file_file[1, 2], key="-SAMPLE1-",
              size=(10, 1)),
     sg.Text("2:"),
     sg.Input(default_text=file_file[1, 3], key="-SAMPLE2-",
              size=(10, 1)),
     sg.Text("3:"),
     sg.Input(default_text=file_file[1, 4], key="-SAMPLE3-",
              size=(10, 1)),
     sg.Text("4:"),
     sg.Input(default_text=file_file[1, 5], key="-SAMPLE4-",
              size=(10, 1))],
    [sg.Text("5:"),
     sg.Input(default_text=file_file[1, 6], key="-SAMPLE5-",
              size=(10, 1)),
     sg.Text("6:"),
     sg.Input(default_text=file_file[1, 7], key="-SAMPLE6-",
              size=(10, 1)),
     sg.Text("7:"),
     sg.Input(default_text=file_file[1, 8], key="-SAMPLE7-",
              size=(10, 1)),
     sg.Text("8:"),
     sg.Input(default_text=file_file[1, 9], key="-SAMPLE8-",
              size=(10, 1))],
    [sg.Text("9:"),
     sg.Input(default_text=file_file[1, 10], key="-SAMPLE9-",
              size=(10, 1)),
     sg.Text("10:"),
     sg.Input(default_text=file_file[1, 11], key="-SAMPLE10-",
              size=(10, 1)),
     sg.Text("11:"),
     sg.Input(default_text=file_file[1, 12], key="-SAMPLE11-",
              size=(10, 1)),
     sg.Text("12:"),
     sg.Input(default_text=file_file[1, 13], key="-SAMPLE12-",
              size=(10, 1))],
    [sg.Text("Name for survival file"), sg.Input(default_text="", key="-FILENAMESURVIVAL-",
                                                 size=(15, 1)), sg.Text(".csv"),
     sg.Button("CREATE", key="-SURVIVALCREATE-"), sg.Text(":名前を入れないとエラーになります")
     ],
    [sg.Text("")],
    [sg.Text("##############################################################")],
    [sg.Text("")],
    [sg.Text("[ For other assay ]"), sg.Text(":その他解析用シート")],
    [sg.Text("Number of analysis:"),
     sg.Input(default_text=file_file[1, 14], key="-NUMBERANALYSIS2-",
              size=(4, 1)), sg.Text(":解析回数（1回解析で終わる実験は'1'）")],
    [sg.Text("Number of sample types:"),
     sg.Input(default_text=file_file[1, 15], key="-NUMBERSAMPLETYPES2-",
              size=(3, 1)), sg.Text(":サンプルの種類数")],
    [sg.Text("Name for other assay file"), sg.Input(default_text="",
                                                    key="-FILENAMEOTHER-",
                                                    size=(15, 1)), sg.Text(".csv"),
     sg.Button("CREATE", key="-OTHERCREATE-"), sg.Text(":名前を入れないとエラーになります")]
]
tab2_layout = [
    [sg.Text("Select your file:")],
    [sg.FileBrowse("Select", key="-FILENAME-", target="-NAMEOFFILE-"),
     sg.Input(default_text=setting_file[1, 0], size=(75, 1),
              key="-NAMEOFFILE-"), sg.Text(":ファイルを選択する")],
    [sg.Text("Select analysis:")],
    [sg.Radio(item[1], key=item[0], group_id='statistic') for item in
     radio_dic.items()],
    [sg.Checkbox("Use Row data", default=False,key="-ROWDATA-"), sg.Text(":生データを使う")],
    [sg.Text("Graph title:"),
     sg.Input(default_text=setting_file[1, 1], key="-GRAPHTITLE-",
              size=(30, 1))],
    [sg.Text("Graph title size:"),
     sg.Input(default_text=setting_file[1, 2], key="-GRAPHTITLESIZE-",
              size=(2, 1))],
    [sg.Text("Graph title style:"),
     sg.Input(default_text=setting_file[1, 3], key="-GRAPHTITLESTYLE-",
              size=(10, 1)), sg.Text(":（normal or italic）")],
    [sg.Text("Horizon title:"),
     sg.Input(default_text=setting_file[1, 4], key="-HORIZONTITLE-",
              size=(30, 1))],
    [sg.Text("Horizon title size:"),
     sg.Input(default_text=setting_file[1, 5], key="-HORIZONTITLESIZE-",
              size=(2, 1))],
    [sg.Text("Horizon tick size:"),
     sg.Input(default_text=setting_file[1, 6], key="-HORIZONTICKSIZE-",
              size=(2, 1))],
    [sg.Text("Vertical title:"),
     sg.Input(default_text=setting_file[1, 7], key="-VERTICALTITLE-",
              size=(30, 1))],
    [sg.Text("Vertical title size:"),
     sg.Input(default_text=setting_file[1, 8], key="-VERTICALTITLESIZE-",
              size=(2, 1))],
    [sg.Text("Vertical tick size:"),
     sg.Input(default_text=setting_file[1, 9], key="-VERTICALTICKSIZE-",
              size=(2, 1))],
    [sg.Text("Vertical range:"),
     sg.Input(default_text=setting_file[1, 10], key="-VERTICALRANGEMIN-",
              size=(5, 1)),
     sg.Text("-"),
     sg.Input(default_text=setting_file[1, 11], key="-VERTICALRANGEMAX-",
              size=(5, 1)),
     sg.Text(":縦軸の目盛り幅")],
    [sg.Text("Legend place:"),
     sg.Input(default_text=setting_file[1, 12], key="-LEGENDPLACE-",
              size=(15, 1)), sg.Text(":（upper right:枠内の右上,upper left:枠外の左上）")],
    [sg.Text("Legend size:"),
     sg.Input(default_text=setting_file[1, 13], key="-LEGENDSIZE-",
              size=(2, 1))],
    [sg.Text("Numbers of legend rows:"),
     sg.Input(default_text=setting_file[1, 14], key="-LEGENDROWS-",
              size=(2, 1))],
    [sg.Text("Figure size:("),
     sg.Input(default_text=setting_file[1, 15], key="-FIGURESIZEX-",
              size=(2, 1)),
     sg.Text(":"),
     sg.Input(default_text=setting_file[1, 16], key="-FIGURESIZEY-",
              size=(2, 1)),
     sg.Text(") (横X : 縦Y)")],
    [sg.Text("Numbers of sample types:"),
     sg.Input(default_text=setting_file[1, 17], key="-NUMBERSSAMPLE-",
              size=(3, 1)), sg.Text(":サンプルの種類数")],
    [sg.Text("Numbers of analysis:"),
     sg.Input(default_text=setting_file[1, 18], key="-NUMBERSANALYSIS-",
              size=(3, 1)), sg.Text(":解析回数（1回解析で終わる実験は'1'）")],
    [sg.Text("X tick label:"),
     sg.Input(default_text=setting_file[1, 19], key="-XLABEL1-",
              size=(5, 1)),
     sg.Input(default_text=setting_file[1, 20], key="-XLABEL2-",
              size=(5, 1)),
     sg.Input(default_text=setting_file[1, 21], key="-XLABEL3-",
              size=(5, 1)),
     sg.Input(default_text=setting_file[1, 22], key="-XLABEL4-",
              size=(5, 1)),
     sg.Input(default_text=setting_file[1, 23], key="-XLABEL5-",
              size=(5, 1)),
     sg.Input(default_text=setting_file[1, 24], key="-XLABEL6-",
              size=(5, 1)), sg.Text(":複数回解析の横軸の目盛り")]
]
tab3_layout = [
    [sg.Button("COLOR SAMPLE", key="-COLORSAMPLE-"), sg.Text(":色見本  "),
     sg.Button("STYLE SAMPLE", key="-STYLESAMPLE-"), sg.Text(":スタイル見本  "),
     sg.Button("MARKER SAMPLE", key="-MARKERSAMPLE-"), sg.Text(":マーカー見本  ")],
    [sg.Text("Sample information:")],
    [sg.Text("Sample name 1:"),
     sg.Input(default_text=sample_file[1, 0], key="-SAMPLENAME1-",
              size=(15, 1)),
     sg.Text("Sample color 1:"),
     sg.Input(default_text=sample_file[1, 1], key="-SAMPLECOLOR1-",
              size=(10, 1)),
     sg.Text("Bar style 1:"),
     sg.Input(default_text=sample_file[1, 2], key="-SAMPLESTYLE1-",
              size=(10, 1)),
     sg.Text("Bar marker 1:"),
     sg.Input(default_text=sample_file[1, 3], key="-SAMPLEMARKER1-",
              size=(10, 1))],
    [sg.Text("Sample name 2:"),
     sg.Input(default_text=sample_file[1, 4], key="-SAMPLENAME2-",
              size=(15, 1)),
     sg.Text("Sample color 2:"),
     sg.Input(default_text=sample_file[1, 5], key="-SAMPLECOLOR2-",
              size=(10, 1)),
     sg.Text("Bar style 2:"),
     sg.Input(default_text=sample_file[1, 6], key="-SAMPLESTYLE2-",
              size=(10, 1)),
     sg.Text("Bar marker 2:"),
     sg.Input(default_text=sample_file[1, 7], key="-SAMPLEMARKER2-",
              size=(10, 1))],
    [sg.Text("Sample name 3:"),
     sg.Input(default_text=sample_file[1, 8], key="-SAMPLENAME3-",
              size=(15, 1)),
     sg.Text("Sample color 3:"),
     sg.Input(default_text=sample_file[1, 9], key="-SAMPLECOLOR3-",
              size=(10, 1)),
     sg.Text("Bar style 3:"),
     sg.Input(default_text=sample_file[1, 10], key="-SAMPLESTYLE3-",
              size=(10, 1)),
     sg.Text("Bar marker 3:"),
     sg.Input(default_text=sample_file[1, 11], key="-SAMPLEMARKER3-",
              size=(10, 1))],
    [sg.Text("Sample name 4:"),
     sg.Input(default_text=sample_file[1, 12], key="-SAMPLENAME4-",
              size=(15, 1)),
     sg.Text("Sample color 4:"),
     sg.Input(default_text=sample_file[1, 13], key="-SAMPLECOLOR4-",
              size=(10, 1)),
     sg.Text("Bar style 4:"),
     sg.Input(default_text=sample_file[1, 14], key="-SAMPLESTYLE4-",
              size=(10, 1)),
     sg.Text("Bar marker 4:"),
     sg.Input(default_text=sample_file[1, 15], key="-SAMPLEMARKER4-",
              size=(10, 1))],
    [sg.Text("Sample name 5:"),
     sg.Input(default_text=sample_file[1, 16], key="-SAMPLENAME5-",
              size=(15, 1)),
     sg.Text("Sample color 5:"),
     sg.Input(default_text=sample_file[1, 17], key="-SAMPLECOLOR5-",
              size=(10, 1)),
     sg.Text("Bar style 5:"),
     sg.Input(default_text=sample_file[1, 18], key="-SAMPLESTYLE5-",
              size=(10, 1)),
     sg.Text("Bar marker 5:"),
     sg.Input(default_text=sample_file[1, 19], key="-SAMPLEMARKER5-",
              size=(10, 1))],
    [sg.Text("Sample name 6:"),
     sg.Input(default_text=sample_file[1, 20], key="-SAMPLENAME6-",
              size=(15, 1)),
     sg.Text("Sample color 6:"),
     sg.Input(default_text=sample_file[1, 21], key="-SAMPLECOLOR6-",
              size=(10, 1)),
     sg.Text("Bar style 6:"),
     sg.Input(default_text=sample_file[1, 22], key="-SAMPLESTYLE6-",
              size=(10, 1)),
     sg.Text("Bar marker 6:"),
     sg.Input(default_text=sample_file[1, 23], key="-SAMPLEMARKER6-",
              size=(10, 1))],
    [sg.Text("Sample name 7:"),
     sg.Input(default_text=sample_file[1, 24], key="-SAMPLENAME7-",
              size=(15, 1)),
     sg.Text("Sample color 7:"),
     sg.Input(default_text=sample_file[1, 25], key="-SAMPLECOLOR7-",
              size=(10, 1)),
     sg.Text("Bar style 7:"),
     sg.Input(default_text=sample_file[1, 26], key="-SAMPLESTYLE7-",
              size=(10, 1)),
     sg.Text("Bar marker 7:"),
     sg.Input(default_text=sample_file[1, 27], key="-SAMPLEMARKER7-",
              size=(10, 1))],
    [sg.Text("Sample name 8:"),
     sg.Input(default_text=sample_file[1, 28], key="-SAMPLENAME8-",
              size=(15, 1)),
     sg.Text("Sample color 8:"),
     sg.Input(default_text=sample_file[1, 29], key="-SAMPLECOLOR8-",
              size=(10, 1)),
     sg.Text("Bar style 8:"),
     sg.Input(default_text=sample_file[1, 30], key="-SAMPLESTYLE8-",
              size=(10, 1)),
     sg.Text("Bar marker 8:"),
     sg.Input(default_text=sample_file[1, 31], key="-SAMPLEMARKER8-",
              size=(10, 1))],
    [sg.Text("Sample name 9:"),
     sg.Input(default_text=sample_file[1, 32], key="-SAMPLENAME9-",
              size=(15, 1)),
     sg.Text("Sample color 9:"),
     sg.Input(default_text=sample_file[1, 33], key="-SAMPLECOLOR9-",
              size=(10, 1)),
     sg.Text("Bar style 9:"),
     sg.Input(default_text=sample_file[1, 34], key="-SAMPLESTYLE9-",
              size=(10, 1)),
     sg.Text("Bar marker 9:"),
     sg.Input(default_text=sample_file[1, 35], key="-SAMPLEMARKER9-",
              size=(10, 1))],
    [sg.Text("Sample name 10:"),
     sg.Input(default_text=sample_file[1, 36], key="-SAMPLENAME10-",
              size=(15, 1)),
     sg.Text("Sample color 10:"),
     sg.Input(default_text=sample_file[1, 37], key="-SAMPLECOLOR10-",
              size=(10, 1)),
     sg.Text("Bar style 10:"),
     sg.Input(default_text=sample_file[1, 38], key="-SAMPLESTYLE10-",
              size=(10, 1)),
     sg.Text("Bar marker 10:"),
     sg.Input(default_text=sample_file[1, 39], key="-SAMPLEMARKER10-",
              size=(10, 1))],
    [sg.Text("Sample name 11:"),
     sg.Input(default_text=sample_file[1, 40], key="-SAMPLENAME11-",
              size=(15, 1)),
     sg.Text("Sample color 11:"),
     sg.Input(default_text=sample_file[1, 41], key="-SAMPLECOLOR11-",
              size=(10, 1)),
     sg.Text("Bar style 11:"),
     sg.Input(default_text=sample_file[1, 42], key="-SAMPLESTYLE11-",
              size=(10, 1)),
     sg.Text("Bar marker 11:"),
     sg.Input(default_text=sample_file[1, 43], key="-SAMPLEMARKER11-",
              size=(10, 1))],
    [sg.Text("Sample name 12:"),
     sg.Input(default_text=sample_file[1, 44], key="-SAMPLENAME12-",
              size=(15, 1)),
     sg.Text("Sample color 12:"),
     sg.Input(default_text=sample_file[1, 45], key="-SAMPLECOLOR12-",
              size=(10, 1)),
     sg.Text("Bar style 12:"),
     sg.Input(default_text=sample_file[1, 46], key="-SAMPLESTYLE12-",
              size=(10, 1)),
     sg.Text("Bar marker 12:"),
     sg.Input(default_text=sample_file[1, 47], key="-SAMPLEMARKER12-",
              size=(10, 1))],
    [sg.Text("Sample name 13:"),
     sg.Input(default_text=sample_file[1, 48], key="-SAMPLENAME13-",
              size=(15, 1)),
     sg.Text("Sample color 13:"),
     sg.Input(default_text=sample_file[1, 49], key="-SAMPLECOLOR13-",
              size=(10, 1)),
     sg.Text("Bar style 13:"),
     sg.Input(default_text=sample_file[1, 50], key="-SAMPLESTYLE13-",
              size=(10, 1)),
     sg.Text("Bar marker 13:"),
     sg.Input(default_text=sample_file[1, 51], key="-SAMPLEMARKER13-",
              size=(10, 1))],
    [sg.Text("Sample name 14:"),
     sg.Input(default_text=sample_file[1, 52], key="-SAMPLENAME14-",
              size=(15, 1)),
     sg.Text("Sample color 14:"),
     sg.Input(default_text=sample_file[1, 53], key="-SAMPLECOLOR14-",
              size=(10, 1)),
     sg.Text("Bar style 14:"),
     sg.Input(default_text=sample_file[1, 54], key="-SAMPLESTYLE14-",
              size=(10, 1)),
     sg.Text("Bar marker 14:"),
     sg.Input(default_text=sample_file[1, 55], key="-SAMPLEMARKER14-",
              size=(10, 1))],
    [sg.Text("Sample name 15:"),
     sg.Input(default_text=sample_file[1, 56], key="-SAMPLENAME15-",
              size=(15, 1)),
     sg.Text("Sample color 15:"),
     sg.Input(default_text=sample_file[1, 57], key="-SAMPLECOLOR15-",
              size=(10, 1)),
     sg.Text("Bar style 15:"),
     sg.Input(default_text=sample_file[1, 58], key="-SAMPLESTYLE15-",
              size=(10, 1)),
     sg.Text("Bar marker 15:"),
     sg.Input(default_text=sample_file[1, 59], key="-SAMPLEMARKER15-",
              size=(10, 1))],
    [sg.Button("START", key="-START-"), sg.Text("", key="-SELECT-")]
]
layout = [
    [sg.TabGroup([[sg.Tab("File design", tab1_layout), sg.Tab("Graph design", tab2_layout), sg.Tab("Sample design",
                                                                                                   tab3_layout)]])],
]
window = sg.Window("Graph and Statistics", layout, resizable=True)
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        exit()
    if event == "-SURVIVALCREATE-":
        sample_name_list = []
        df_new_survival_file = pd.DataFrame(columns=['number', 'sample', 'time', 'death'])
        df_new_survival_file['number'] = range(1, int(values["-NUMBERSAMPLETYPES-"]) * int(
            values["-NUMBERSAMPLESIZE-"]) + 1)
        df_new_survival_file['death'] = np.repeat("Death", int(values["-NUMBERSAMPLETYPES-"]) * int(
            values["-NUMBERSAMPLESIZE-"]))
        for n in range(int(values["-NUMBERSAMPLETYPES-"])):
            exec('ndarray_sample= np.repeat(values["-SAMPLE{0}-"], int(values["-NUMBERSAMPLESIZE-"]))'.format(n + 1))
            list_sample = ndarray_sample.tolist()
            sample_name_list.extend(list_sample)
        df_new_survival_file['sample'] = sample_name_list
        exec('path_new_survival_file = path / "data&result/{0}.csv"'.format(values["-FILENAMESURVIVAL-"]))
        df_new_survival_file.to_csv(path_new_survival_file, index=False, na_rep='')

        values_list1 = np.array([int(values["-NUMBERSAMPLETYPES-"]),
                                 int(values["-NUMBERSAMPLESIZE-"]),
                                 values["-SAMPLE1-"],
                                 values["-SAMPLE2-"],
                                 values["-SAMPLE3-"],
                                 values["-SAMPLE4-"],
                                 values["-SAMPLE5-"],
                                 values["-SAMPLE6-"],
                                 values["-SAMPLE7-"],
                                 values["-SAMPLE8-"],
                                 values["-SAMPLE9-"],
                                 values["-SAMPLE10-"],
                                 values["-SAMPLE11-"],
                                 values["-SAMPLE12-"],
                                 int(values["-NUMBERANALYSIS2-"]),
                                 int(values["-NUMBERSAMPLETYPES2-"])
                                 ])
        delete_file_file = np.delete(file_file, 1, 0)
        file_new = np.block([[delete_file_file], [values_list1]])
        df_file = pd.DataFrame(data=file_new)
        df_file.to_csv(path_file, index=False, header=False)
        continue
    if event == "-OTHERCREATE-":
        other_sample_list = []
        for m in range(int(values["-NUMBERANALYSIS2-"])):
            other_sample_list.append('size')
            for n in range(int(values["-NUMBERSAMPLETYPES2-"])):
                exec('other_sample_list.append("N{0}Sample{1}")'.format(m + 1, n + 1))
        df_new_other_file = pd.DataFrame(columns=other_sample_list)
        df_new_other_file['size'] = range(1, 200 + 1)
        exec('path_new_other_file = path / "data&result/{0}.csv"'.format(values["-FILENAMEOTHER-"]))
        df_new_other_file.to_csv(path_new_other_file, index=False,
                                 na_rep='')

        values_list1 = np.array([int(values["-NUMBERSAMPLETYPES-"]),
                                 int(values["-NUMBERSAMPLESIZE-"]),
                                 values["-SAMPLE1-"],
                                 values["-SAMPLE2-"],
                                 values["-SAMPLE3-"],
                                 values["-SAMPLE4-"],
                                 values["-SAMPLE5-"],
                                 values["-SAMPLE6-"],
                                 values["-SAMPLE7-"],
                                 values["-SAMPLE8-"],
                                 values["-SAMPLE9-"],
                                 values["-SAMPLE10-"],
                                 values["-SAMPLE11-"],
                                 values["-SAMPLE12-"],
                                 int(values["-NUMBERANALYSIS2-"]),
                                 int(values["-NUMBERSAMPLETYPES2-"])
                                 ])
        delete_file_file = np.delete(file_file, 1, 0)
        file_new = np.block([[delete_file_file], [values_list1]])
        df_file = pd.DataFrame(data=file_new)
        df_file.to_csv(path_file, index=False, header=False)
        continue
    if event == "-HOWDESIGNFILE-":
        img = Image.open(path / "File_design.png")
        img.show()
        continue
    if event =="-OUTPUTRESULTS-":
        img = Image.open(path / "Results.png")
        img.show()
        continue
    if event == "-COLORSAMPLE-":
        img = Image.open(path / "color_sample.png")
        img.show()
        continue
    if event == "-STYLESAMPLE-":
        img = Image.open(path / "hatch_sample.png")
        img.show()
        continue
    if event == "-MARKERSAMPLE-":
        img = Image.open(path / "marker_sample.png")
        img.show()
        continue
    if event == "-START-" and values["-SURVIVAL-"] == False and values[
        "-FAT-"] == False and values["-AGINGBAR-"] == False and values["-AGINGLINE-"] == False:
        window["-SELECT-"].update("Select analysis!")
        continue
    else:
        values_list2 = np.array([values["-NAMEOFFILE-"], values["-GRAPHTITLE-"],
                                 int(values["-GRAPHTITLESIZE-"]),
                                 values["-GRAPHTITLESTYLE-"],
                                 values["-HORIZONTITLE-"],
                                 int(values["-HORIZONTITLESIZE-"]),
                                 int(values["-HORIZONTICKSIZE-"]),
                                 values["-VERTICALTITLE-"],
                                 int(values["-VERTICALTITLESIZE-"]),
                                 int(values["-VERTICALTICKSIZE-"]),
                                 int(values["-VERTICALRANGEMIN-"]),
                                 int(values["-VERTICALRANGEMAX-"]),
                                 values["-LEGENDPLACE-"], int(values["-LEGENDSIZE-"]),
                                 int(values["-LEGENDROWS-"]),
                                 int(values["-FIGURESIZEX-"]),
                                 int(values["-FIGURESIZEY-"]),
                                 int(values["-NUMBERSSAMPLE-"]),
                                 int(values["-NUMBERSANALYSIS-"]),
                                 values["-XLABEL1-"],
                                 values["-XLABEL2-"],
                                 values["-XLABEL3-"],
                                 values["-XLABEL4-"],
                                 values["-XLABEL5-"],
                                 values["-XLABEL6-"]
                                 ])
        values_list3 = np.array([values["-SAMPLENAME1-"], values["-SAMPLECOLOR1-"],
                                 values["-SAMPLESTYLE1-"], values["-SAMPLEMARKER1-"],
                                 values["-SAMPLENAME2-"], values["-SAMPLECOLOR2-"],
                                 values["-SAMPLESTYLE2-"], values["-SAMPLEMARKER2-"],
                                 values["-SAMPLENAME3-"], values["-SAMPLECOLOR3-"],
                                 values["-SAMPLESTYLE3-"], values["-SAMPLEMARKER3-"],
                                 values["-SAMPLENAME4-"], values["-SAMPLECOLOR4-"],
                                 values["-SAMPLESTYLE4-"], values["-SAMPLEMARKER4-"],
                                 values["-SAMPLENAME5-"], values["-SAMPLECOLOR5-"],
                                 values["-SAMPLESTYLE5-"], values["-SAMPLEMARKER5-"],
                                 values["-SAMPLENAME6-"], values["-SAMPLECOLOR6-"],
                                 values["-SAMPLESTYLE6-"], values["-SAMPLEMARKER6-"],
                                 values["-SAMPLENAME7-"], values["-SAMPLECOLOR7-"],
                                 values["-SAMPLESTYLE7-"], values["-SAMPLEMARKER7-"],
                                 values["-SAMPLENAME8-"], values["-SAMPLECOLOR8-"],
                                 values["-SAMPLESTYLE8-"], values["-SAMPLEMARKER8-"],
                                 values["-SAMPLENAME9-"], values["-SAMPLECOLOR9-"],
                                 values["-SAMPLESTYLE9-"], values["-SAMPLEMARKER9-"],
                                 values["-SAMPLENAME10-"], values["-SAMPLECOLOR10-"],
                                 values["-SAMPLESTYLE10-"], values["-SAMPLEMARKER10-"],
                                 values["-SAMPLENAME11-"], values["-SAMPLECOLOR11-"],
                                 values["-SAMPLESTYLE11-"], values["-SAMPLEMARKER11-"],
                                 values["-SAMPLENAME12-"], values["-SAMPLECOLOR12-"],
                                 values["-SAMPLESTYLE12-"], values["-SAMPLEMARKER12-"],
                                 values["-SAMPLENAME13-"], values["-SAMPLECOLOR13-"],
                                 values["-SAMPLESTYLE13-"], values["-SAMPLEMARKER13-"],
                                 values["-SAMPLENAME14-"], values["-SAMPLECOLOR14-"],
                                 values["-SAMPLESTYLE14-"], values["-SAMPLEMARKER14-"],
                                 values["-SAMPLENAME15-"], values["-SAMPLECOLOR15-"],
                                 values["-SAMPLESTYLE15-"], values["-SAMPLEMARKER15-"],
                                 ])

        delete_setting_file = np.delete(setting_file, 1, 0)
        setting_new = np.block([[delete_setting_file], [values_list2]])
        df_setting = pd.DataFrame(data=setting_new)
        df_setting.to_csv(path_setting, index=False, header=False, na_rep='NULL')

        delete_sample_file = np.delete(sample_file, 1, 0)
        sample_new = np.block([[delete_sample_file], [values_list3]])
        df_sample = pd.DataFrame(data=sample_new)
        df_sample.to_csv(path_sample, index=False, header=False, na_rep='NULL')

        break

# ↑↑↑↑↑↑↑↑PysimpleGUI"ここまで", Figure"共通設定ここから"↓↓↓↓↓↓↓↓#
fig = plt.figure(
    figsize=(int(values["-FIGURESIZEX-"]), int(values["-FIGURESIZEY-"])),
    dpi=240, constrained_layout=True)
ax = fig.add_subplot(1, 1, 1)
df_data = pd.read_csv(values["-NAMEOFFILE-"])
# ↑↑↑↑↑↑↑↑Figure"共通設定ここまで","グラフ作成ここから"↓↓↓↓↓↓↓↓#
# ↓↓↓↓↓↓↓↓Survival"ここから"↓↓↓↓↓↓↓↓#
sample_name_list = []
if values["-SURVIVAL-"] == True:
    for n in range(int(values["-NUMBERSSAMPLE-"])):
        exec(
            'N1_S{0} = SampleInfo(values["-SAMPLENAME{0}-"], values["-SAMPLECOLOR{0}-"],'
            'values["-SAMPLESTYLE{0}-"],values["-SAMPLEMARKER{0}-"])'.format(
                n + 1))
        exec("sample_name_list.append(N1_S{0}.return_name())".format(n + 1))
    life_file = values["-NAMEOFFILE-"]
    df_life = pd.read_csv(life_file, encoding='cp932', index_col="number")
    df_life.replace({'Death':bool(True), 'Miss': bool(False)},inplace=True)
    ax = None
    num = 0
    for sample_name, group_data in df_life.groupby("sample"):
        kmf = KaplanMeierFitter()
        num += 1
        kmf.fit(group_data["time"], event_observed=group_data["death"],
                label=str(sample_name))
        exec(
            'ax=kmf.plot(ax=ax, ci_show=False, linewidth=3, color = values["-SAMPLECOLOR{0}-"], '
            'style = values["-SAMPLESTYLE{0}-"], marker=values["-SAMPLEMARKER{0}-"])'.format(
                num))
    ax.legend(sample_name_list, loc=values["-LEGENDPLACE-"], fontsize=int(values["-LEGENDSIZE-"]), handlelength=0.8,
              handletextpad=0.1, labelspacing=0.1, ncol=int(values["-LEGENDROWS-"]),
              bbox_to_anchor=(1, 1), borderpad=0.5, shadow=False, frameon=False)  # legend
    ax.set_title(values["-GRAPHTITLE-"],
                 fontdict={"fontsize": int(values["-GRAPHTITLESIZE-"]),
                           "fontweight": "bold", "fontstyle": values["-GRAPHTITLESTYLE-"]})
    ax.set_xlabel(values["-HORIZONTITLE-"], labelpad=None,
                  fontdict={"fontsize": values["-HORIZONTITLESIZE-"], "fontweight": "bold"})

    ax.set_ylabel(values["-VERTICALTITLE-"], labelpad=None,
                  fontdict={"fontsize": int(values["-VERTICALTITLESIZE-"]),
                            "fontweight": "bold"})
    plt.setp(ax.get_yticklabels(), fontsize=values["-VERTICALTICKSIZE-"],
             fontweight='bold')
    ax.set_ylim(0, 1)
    ax.spines['right'].set(visible=False)
    ax.spines['top'].set(visible=False)
    plt.xticks(fontsize=int(values["-HORIZONTICKSIZE-"]), fontweight='bold')  # pltで貼付修正
    fig.savefig(path / "figure.png", format="png", bbox_inches="tight", dpi='figure')

    wb_data = openpyxl.Workbook()
    wb_data.create_sheet('figure&statistics', 0)
    figure_statistics_sheet = wb_data['figure&statistics']

    for n in range(int(values["-NUMBERSSAMPLE-"])):                 #統計検定始まり
        if n + 1 == int(values["-NUMBERSSAMPLE-"]):
            break
        exec('sample_A = N1_S{0}.return_name()'.format(n + 1))
        log_rankA = df_life[(df_life["sample"] == sample_A)]
        for m in range(int(values["-NUMBERSSAMPLE-"])):
            if n + m + 1 == int(values["-NUMBERSSAMPLE-"]):
                break
            exec('sample_B = N1_S{0}.return_name()'.format(n + m + 2))
            log_rankB = df_life[(df_life["sample"] == sample_B)]
            results = logrank_test(log_rankA["time"], log_rankB["time"], log_rankA["death"], log_rankB["death"])
            exec("S1 = N1_S{0}.return_name()".format(n + 1))
            exec("S2 = N1_S{0}.return_name()".format(n + m + 2))
            print("\n\n---------Log_rank({0} vs {1})---------".format(S1, S2))
            print("p-value=", "{:.10f}".format(results.p_value))
            exec(
                "log_rank_list{0}=['Log_rank({1} vs {2}): p-value=','','',{3:.10f}]".format(m, S1, S2, results.p_value))
            exec("figure_statistics_sheet.append(log_rank_list{0})".format(m))

    figure = openpyxl.drawing.image.Image(path / "figure.png")
    figure_statistics_sheet.add_image(figure, 'H1')
    wb_data.save(path / "{0}.xlsx".format(os.path.basename(values["-NAMEOFFILE-"]).split(".")[-2]))
    plt.show()
    exit()  # #
# ↑↑↑↑↑↑↑↑Survival"ここまで", Body_length,Fat_accumulationここから"↓↓↓↓↓↓↓↓#
if  values["-FAT-"] == True:
    xtickslist = []
    for n in range(int(values["-NUMBERSSAMPLE-"])):
        exec(
            'N1_S{0} = SampleInfo(values["-SAMPLENAME{0}-"], values["-SAMPLECOLOR{0}-"],'
            'values["-SAMPLESTYLE{0}-"],values["-SAMPLEMARKER{0}-"],df_data["N1Sample{0}"].dropna())'.format(
                n + 1))
        exec('N1_S{0}.simple_bar_cal(N1_S1.data_list)'.format(n + 1))

    ax.legend(loc=values["-LEGENDPLACE-"], fontsize=int(values["-LEGENDSIZE-"]),
              handlelength=0.75, handletextpad=0.1, labelspacing=0.2,
              borderpad=0.5,
              ncol=int(values["-LEGENDROWS-"]), bbox_to_anchor=(1, 1),
              shadow=False,
              frameon=False)
    ax.set_title(values["-GRAPHTITLE-"],
                 fontdict={"fontsize": int(values["-GRAPHTITLESIZE-"]),
                           "fontweight": "bold", "fontstyle": values["-GRAPHTITLESTYLE-"]})
    ax.set_xticks(xtickslist)
    ax.set_xlabel(values["-HORIZONTITLE-"], labelpad=None,
                  fontdict={"fontsize": values["-HORIZONTITLESIZE-"], "fontweight": "bold"})
    ax.set_xticklabels(xtickslist, fontsize=int(values["-HORIZONTICKSIZE-"]),
                       fontweight='bold')
    ax.set_ylabel(values["-VERTICALTITLE-"], labelpad=None,
                  fontdict={"fontsize": int(values["-VERTICALTITLESIZE-"]),
                            "fontweight": "bold"})
    plt.setp(ax.get_yticklabels(), fontsize=values["-VERTICALTICKSIZE-"],
             fontweight='bold')
    ax.set_ylim([int(values["-VERTICALRANGEMIN-"]), int(values["-VERTICALRANGEMAX-"])])
    ax.spines['right'].set(visible=False)
    ax.spines['top'].set(visible=False)
    fig.savefig(path / "figure.png", format="png", bbox_inches="tight",
                dpi='figure')

    wb_data = openpyxl.Workbook()
    wb_data.create_sheet('figure&statistics', 0)
    figure_statistics_sheet = wb_data['figure&statistics']

    if int(values["-NUMBERSSAMPLE-"]) == 2:                 #統計検定始まり
        t_test(N1_S1.return_relative(N1_S1.data_list), N1_S2.return_relative(N1_S1.data_list), 1)

    if int(values["-NUMBERSSAMPLE-"]) > 2:
        anova_list = []
        combined_all_sample_data = []
        sample_names_and_len = []
        for n in range(int(values["-NUMBERSSAMPLE-"])):
            exec('anova_list.append(N1_S{0}.return_relative(N1_S1.data_list))'.format(n + 1))
            exec('np_data{0} = np.array(N1_S{0}.return_relative(N1_S1.data_list))'.format(n + 1))
            exec('combined_all_sample_data = np.concatenate([combined_all_sample_data, np_data{0}])'.format(n + 1))
            exec('name_and_len{0} = np.repeat(values["-SAMPLENAME{0}-"], len(np_data{0}))'.format(n + 1))
            exec('sample_names_and_len = np.concatenate([sample_names_and_len, name_and_len{0}])'.format(n + 1))
        one_way_anova(1, *anova_list)
        multiple_comparison_test(combined_all_sample_data, sample_names_and_len)
    figure = openpyxl.drawing.image.Image(path / "figure.png")
    figure_statistics_sheet.add_image(figure, 'H1')
    wb_data.save(path / "{0}.xlsx".format(os.path.basename(values["-NAMEOFFILE-"]).split(".")[-2]))
    plt.show()
    exit()
if values["-AGINGBAR-"] == True:
    xtickslist = []
    label_x = []
    for n in range(int(values["-NUMBERSANALYSIS-"])):
        exec('label_x.append(values["-XLABEL{0}-"])'.format(n + 1))
    for n in range(int(values["-NUMBERSSAMPLE-"])):
        exec("errorS{0} = []".format(n + 1))
        exec("x_S{0} = []".format(n + 1))
        exec("y_S{0} = []".format(n + 1))
    for m in range(int(values["-NUMBERSANALYSIS-"])):
        for n in range(int(values["-NUMBERSSAMPLE-"])):
            exec(
                'N{0}_S{1} = SampleInfo(values["-SAMPLENAME{1}-"], values["-SAMPLECOLOR{1}-"],'
                'values["-SAMPLESTYLE{1}-"],values["-SAMPLEMARKER{1}-"],df_data["N{0}Sample{1}"].dropna())'.format(
                    m + 1, n + 1))
            exec("rel_N{0}S{1} = N{0}_S{1}.multi_bar_cal_relative(N1_S{1}.data_list)".format(m + 1,
                                                                                             n + 1))  # def relativeで縦軸サンプルごとに補正
            exec("sem_N{0}S{1} = N{0}_S{1}.multi_bar_cal_sem(N1_S{1}.data_list)".format(m + 1, n + 1))  # def semで標準誤差計算

    for n in range(int(values["-NUMBERSSAMPLE-"])):  # 縦棒の数
        for m in range(int(values["-NUMBERSANALYSIS-"])):
            exec("errorS{1}.append(sem_N{0}S{1})".format(m + 1, n + 1))  # 標準誤差格納
            if int(values["-NUMBERSSAMPLE-"]) <= 5:
                exec("x_S{0}.append({1}+{2})".format(n + 1, m + 1, 0.2 * n))
                exec("y_S{0}.append(rel_N{1}S{0})".format(n + 1, m + 1))
            else:
                exec("x_S{0}.append({1}+{2})".format(n + 1, m + 1, 0.1 * n))
                exec("y_S{0}.append(rel_N{1}S{0})".format(n + 1, m + 1))

    if int(values["-NUMBERSSAMPLE-"]) <= 5:  # 棒グラフのプロット、サンプル <= 5
        for n in range(int(values["-NUMBERSSAMPLE-"])):
            exec(
                'ax.bar(x_S{0}, y_S{0}, edgecolor="k", color=N1_S{0}.return_color(), width=0.2, label=N1_S{0}.return_name(), align="center", yerr=errorS{0},ecolor="k", capsize=2, hatch=N1_S{0}.return_style())'.format(
                    n + 1))
    else:
        for n in range(int(values["-NUMBERSSAMPLE-"])):  # 棒グラフのプロット、サンプル > 5
            exec(
                'ax.bar(x_S{0}, y_S{0}, edgecolor="k", color=N1_S{0}.return_color(), width=0.1, label=N1_S{0}.return_name(), align="center", yerr=errorS{0},ecolor="k", capsize=2, hatch=N1_S{0}.return_style())'.format(
                    n + 1))
    ax.legend(loc=values["-LEGENDPLACE-"], fontsize=int(values["-LEGENDSIZE-"]),
              handlelength=0.75, handletextpad=0.1, labelspacing=0.2,
              borderpad=0.5,
              ncol=int(values["-LEGENDROWS-"]), bbox_to_anchor=(1, 1),
              shadow=False,
              frameon=False)
    if int(values["-NUMBERSSAMPLE-"]) <= 5:
        N1 = int(values["-NUMBERSSAMPLE-"]) - 1
        N2 = N1 / 10
        for n in range(int(values["-NUMBERSANALYSIS-"])):
            exec("xtickslist.append({0}+N2)".format(n + 1))
    else:
        N1 = int(values["-NUMBERSSAMPLE-"]) - 1
        N2 = N1 / 10
        N3 = N2 / 2
        for n in range(int(values["-NUMBERSANALYSIS-"])):
            exec("xtickslist.append({0}+N3)".format(n + 1))

    ax.set_title(values["-GRAPHTITLE-"],
                 fontdict={"fontsize": int(values["-GRAPHTITLESIZE-"]),
                           "fontweight": "bold", "fontstyle": values["-GRAPHTITLESTYLE-"]})
    ax.set_xticks(xtickslist)
    ax.set_xlabel(values["-HORIZONTITLE-"], labelpad=None,
                  fontdict={"fontsize": values["-HORIZONTITLESIZE-"], "fontweight": "bold"})
    ax.set_ylabel(values["-VERTICALTITLE-"], labelpad=None,
                  fontdict={"fontsize": int(values["-VERTICALTITLESIZE-"]),
                            "fontweight": "bold"})
    ax.set_xticklabels(xtickslist, fontsize=int(values["-HORIZONTICKSIZE-"]),
                       fontweight='bold')
    plt.setp(ax.get_yticklabels(), fontsize=values["-VERTICALTICKSIZE-"], fontweight="bold")

    plt.xticks(xtickslist, label_x)  # pltで貼付修正

    ax.set_ylim([int(values["-VERTICALRANGEMIN-"]), int(values["-VERTICALRANGEMAX-"])])
    ax.spines['right'].set(visible=False)
    ax.spines['top'].set(visible=False)
    fig.savefig(path / "figure.png", format="png", bbox_inches="tight",
                dpi='figure')
    N1_S1.return_relative(N1_S1.data_list), N1_S2.return_relative(
        N1_S1.data_list)
    wb_data = openpyxl.Workbook()
    wb_data.create_sheet('figure&statistics', 0)
    figure_statistics_sheet = wb_data['figure&statistics']

    if int(values["-NUMBERSSAMPLE-"]) == 2:
        for n in range(int(values["-NUMBERSANALYSIS-"])):
            exec(
                "t_test(N{0}_S1.return_relative(N1_S1.data_list), N{0}_S2.return_relative(N1_S2.data_list), {0})".format(
                    n + 1))

    if int(values["-NUMBERSSAMPLE-"]) > 2:                  #統計検定始まり
        for m in range(int(values["-NUMBERSANALYSIS-"])):
            exec("anova_list{0} = []".format(m + 1))
            exec("combined_all_sample_data{0} = []".format(m + 1))
            exec("sample_names_and_len{0} = []".format(m + 1))
            for n in range(int(values["-NUMBERSSAMPLE-"])):
                exec('anova_list{0}.append(N{0}_S{1}.return_relative(N1_S{1}.data_list))'.format(m + 1, n + 1))
                exec('np_data{1} = np.array(N{0}_S{1}.return_relative(N1_S{1}.data_list))'.format(m + 1, n + 1))
                exec('combined_all_sample_data{0} = np.concatenate([combined_all_sample_data{0}, np_data{1}])'.format(
                    m + 1, n + 1))
                exec('name_and_len{1} = np.repeat(values["-SAMPLENAME{1}-"], len(np_data{1}))'.format(m + 1, n + 1))
                exec(
                    'sample_names_and_len{0} = np.concatenate([sample_names_and_len{0}, name_and_len{1}])'.format(m + 1,
                                                                                                                  n + 1))
        for m in range(int(values["-NUMBERSANALYSIS-"])):
            exec("one_way_anova({0},*anova_list{0})".format(m + 1))
            exec("multiple_comparison_test(combined_all_sample_data{0}, sample_names_and_len{0})".format(m + 1))
    figure = openpyxl.drawing.image.Image(path / "figure.png")
    figure_statistics_sheet.add_image(figure, 'H1')
    wb_data.save(path / "{0}.xlsx".format(os.path.basename(values["-NAMEOFFILE-"]).split(".")[-2]))
    plt.show()
    exit()
if values["-AGINGLINE-"] == True:
    label_x = []
    for n in range(int(values["-NUMBERSANALYSIS-"])):
        exec('label_x.append(values["-XLABEL{0}-"])'.format(n + 1))
    for m in range(int(values["-NUMBERSANALYSIS-"])):
        for n in range(int(values["-NUMBERSSAMPLE-"])):
            exec(
                'N{0}_S{1} = SampleInfo(values["-SAMPLENAME{1}-"], values["-SAMPLECOLOR{1}-"],'
                'values["-SAMPLESTYLE{1}-"],values["-SAMPLEMARKER{1}-"],df_data["N{0}Sample{1}"].dropna())'.format(
                    m + 1, n + 1))
            exec("rel_N{0}S{1} = N{0}_S{1}.multi_bar_cal_relative(N1_S{1}.data_list)".format(m + 1,
                                                                                             n + 1))  # def relativeで縦軸サンプルごとに補正
            exec("sem_N{0}S{1} = N{0}_S{1}.multi_bar_cal_sem(N1_S{1}.data_list)".format(m + 1, n + 1))  # def semで標準誤差計算
    for n in range(int(values["-NUMBERSSAMPLE-"])):  # 縦棒の数
        exec("x_S{0} = []".format(n + 1))
        exec("y_S{0} = []".format(n + 1))
        exec("errorS{0} = []".format(n + 1))
        for m in range(int(values["-NUMBERSANALYSIS-"])):
            exec("errorS{1}.append(sem_N{0}S{1})".format(m + 1, n + 1))  # 標準誤差格納
            exec("x_S{0}.append({1})".format(n + 1, m + 1))
            exec("y_S{0}.append(rel_N{1}S{0})".format(n + 1, m + 1))
    for n in range(int(values["-NUMBERSSAMPLE-"])):
        exec('sample_style = values["-SAMPLESTYLE{0}-"]'.format(n + 1))
        if sample_style == "":
            bar_style = "-"
        else:
            exec("bar_style = N1_S{0}.return_style()".format(n + 1))
        exec(
            'ax.errorbar(x_S{0}, y_S{0},label=N1_S{0}.return_name(), color=N1_S{0}.return_color(), yerr=errorS{0}, '
            'ecolor="k", capsize=3, linestyle=bar_style, marker=N1_S{0}.return_marker(), linewidth = 2.5)'.format(
                n + 1))
    handles, labels = ax.get_legend_handles_labels()
    handles = [h[0] for h in handles]
    ax.legend(handles, labels, loc=values["-LEGENDPLACE-"], fontsize=int(values["-LEGENDSIZE-"]), handlelength=2.0,
              handletextpad=0.1, labelspacing=0.1, ncol=int(values["-LEGENDROWS-"]), bbox_to_anchor=(1, 1),
              borderpad=0.5, shadow=False, frameon=False)  # legend

    xtickslist = []
    for n in range(int(values["-NUMBERSANALYSIS-"])):
        exec("xtickslist.append({0})".format(n + 1))

    ax.set_title(values["-GRAPHTITLE-"], fontdict={"fontsize": int(values["-GRAPHTITLESIZE-"]),
                                                   "fontweight": "bold",
                                                   "fontstyle": values["-GRAPHTITLESTYLE-"]})
    ax.set_xticks(xtickslist)
    ax.set_xlabel(values["-HORIZONTITLE-"], labelpad=None, fontdict={"fontsize": values["-HORIZONTITLESIZE-"],
                                                                     "fontweight": "bold"})
    ax.set_ylabel(values["-VERTICALTITLE-"], labelpad=None, fontdict={"fontsize": int(values["-VERTICALTITLESIZE-"]),
                                                                      "fontweight": "bold"})
    ax.set_xticklabels(xtickslist, fontsize=int(values["-HORIZONTICKSIZE-"]),
                       fontweight='bold')

    plt.setp(ax.get_yticklabels(), fontsize=values["-VERTICALTICKSIZE-"],
             fontweight="bold")  # pltで貼付修正
    plt.xticks(xtickslist, label_x)  # pltで貼付修正

    plt.rcParams["font.family"] = 'Arial'
    ax.set_ylim(
        [int(values["-VERTICALRANGEMIN-"]), int(values["-VERTICALRANGEMAX-"])])
    ax.spines['right'].set(visible=False)
    ax.spines['top'].set(visible=False)
    fig.savefig(path / "figure.png", format="png", bbox_inches="tight",
                dpi='figure')

    wb_data = openpyxl.Workbook()
    wb_data.create_sheet('figure&statistics', 0)
    figure_statistics_sheet = wb_data['figure&statistics']

    if int(values["-NUMBERSSAMPLE-"]) == 2:                     #統計検定始まり
        for n in range(int(values["-NUMBERSANALYSIS-"])):
            exec(
                "t_test(N{0}_S1.return_relative(N1_S1.data_list), N{0}_S2.return_relative(N1_S2.data_list), {0})".format(
                    n + 1))
    if int(values["-NUMBERSSAMPLE-"]) > 2:
        for m in range(int(values["-NUMBERSANALYSIS-"])):
            exec("anova_list{0} = []".format(m + 1))
            exec("combined_all_sample_data{0} = []".format(m + 1))
            exec("sample_names_and_len{0} = []".format(m + 1))
            for n in range(int(values["-NUMBERSSAMPLE-"])):
                exec('anova_list{0}.append(N{0}_S{1}.return_relative(N1_S{1}.data_list))'.format(m + 1, n + 1))
                exec('np_data{1} = np.array(N{0}_S{1}.return_relative(N1_S{1}.data_list))'.format(m + 1, n + 1))
                exec('combined_all_sample_data{0} = np.concatenate([combined_all_sample_data{0}, np_data{1}])'.format(
                    m + 1, n + 1))
                exec('name_and_len{1} = np.repeat(values["-SAMPLENAME{1}-"], len(np_data{1}))'.format(m + 1, n + 1))
                exec(
                    'sample_names_and_len{0} = np.concatenate([sample_names_and_len{0}, name_and_len{1}])'.format(m + 1,
                                                                                                                  n + 1))
        for m in range(int(values["-NUMBERSANALYSIS-"])):
            exec("one_way_anova({0},*anova_list{0})".format(m + 1))
            exec("multiple_comparison_test(combined_all_sample_data{0}, sample_names_and_len{0})".format(m + 1))
    figure = openpyxl.drawing.image.Image(path / "figure.png")
    figure_statistics_sheet.add_image(figure, 'H1')
    wb_data.save(path / "{0}.xlsx".format(os.path.basename(values["-NAMEOFFILE-"]).split(".")[-2]))
    plt.show()
    exit()
exit()
