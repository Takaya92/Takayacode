import numpy as np
import pandas as pd
from pathlib import Path
import sqlite3
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
path_dir = Path(path / "data&result")
path_dir.mkdir(exist_ok=True)

# ↓↓↓↓↓↓↓↓class,def"ここから"↓↓↓↓↓↓↓↓ #
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

def legend_bar():
    ax.legend(loc=values["-LEGENDPLACE-"],
              fontsize=int(values["-LEGENDSIZE-"]),
              handlelength=0.75, handletextpad=0.1, labelspacing=0.2,
              borderpad=0.5,
              ncol=int(values["-LEGENDROWS-"]), bbox_to_anchor=(1, 1),
              shadow=False,
              frameon=False)

def legend_kinetics_line():
    ax.legend(handles, labels, loc=values["-LEGENDPLACE-"],
              fontsize=int(values["-LEGENDSIZE-"]), handlelength=2.0,
              handletextpad=0.1, labelspacing=0.1,
              ncol=int(values["-LEGENDROWS-"]), bbox_to_anchor=(1, 1),
              borderpad=0.5, shadow=False, frameon=False)

def legend_survival():
    ax.legend(sample_name_list, loc=values["-LEGENDPLACE-"],
              fontsize=int(values["-LEGENDSIZE-"]), handlelength=0.8,
              handletextpad=0.1, labelspacing=0.1,
              ncol=int(values["-LEGENDROWS-"]),
              bbox_to_anchor=(1, 1), borderpad=0.5, shadow=False,
              frameon=False)

def graph_config():
    if sample_design_values["-SIMPLEBAR-"] == True or sample_design_values["-KINETICSBAR-"] == True:     # Barグラフ用legend
        legend_bar()
    elif sample_design_values["-KINETICSLINE-"] == True:  # Kineticsline用legend
        legend_kinetics_line()
    elif sample_design_values["-SURVIVAL-"] == True:  # Survival用legend
        legend_survival()
    ax.set_title(values["-GRAPHTITLE-"],
                 fontdict={"fontsize": int(values["-GRAPHTITLESIZE-"]),
                           "fontweight": "bold", "fontstyle": values["-GRAPHTITLESTYLE-"]})
    if sample_design_values["-SURVIVAL-"] == False: ax.set_xticks(xtickslist) #Survival以外の横軸
    ax.set_xlabel(values["-HORIZONTITLE-"], labelpad=None,
                  fontdict={"fontsize": values["-HORIZONTITLESIZE-"], "fontweight": "bold"})
    ax.set_ylabel(values["-VERTICALTITLE-"], labelpad=None,
                  fontdict={"fontsize": int(values["-VERTICALTITLESIZE-"]),
                            "fontweight": "bold"})
    if sample_design_values["-SURVIVAL-"] == False:
        ax.set_xticklabels(xtickslist, fontsize=int(values["-HORIZONTICKSIZE-"]),
                       fontweight='bold')
        if sample_design_values["-KINETICSBAR-"] == True or sample_design_values["-KINETICSLINE-"] == True:     # Kinetics用,pltで後付け
            plt.xticks(xtickslist, label_x)
    plt.setp(ax.get_yticklabels(), fontsize=values["-VERTICALTICKSIZE-"], fontweight="bold")
    plt.rcParams["font.family"] = 'Arial'
    if sample_design_values["-SURVIVAL-"] == False:
        ax.set_ylim([int(values["-VERTICALRANGEMIN-"]), int(values["-VERTICALRANGEMAX-"])])
    else:
        ax.set_ylim(0, 1)
        plt.xticks(fontsize=int(values["-HORIZONTICKSIZE-"]),fontweight='bold')  # pltで貼付修正
    ax.spines['right'].set(visible=False)
    ax.spines['top'].set(visible=False)
    fig.savefig(path / "figure.png", format="png", bbox_inches="tight",
                dpi='figure')
    wb_data = openpyxl.Workbook()
    wb_data.create_sheet('figure&statistics', 0)
    figure_statistics_sheet = wb_data['figure&statistics']
    return wb_data, figure_statistics_sheet

def save_and_plot_fig():
    figure = openpyxl.drawing.image.Image(path / "figure.png")
    figure_statistics_sheet.add_image(figure, 'H1')
    wb_data.save(path_dir / "{0}.xlsx".format(os.path.basename(values["-NAMEOFFILE-"]).split(".")[-2]))
    plt.show()
    exit()

# ↓↓↓↓↓↓↓↓設定DB↓↓↓↓↓↓↓↓ #
def call_value_tab1(id1):
    db_setting = path / "db_tab1.db"
    conn = sqlite3.connect(db_setting)
    cur = conn.cursor()
    if len(cur.execute('SELECT * FROM tab1 WHERE id = ?', (f"{id1}",)).fetchall()) == 0:
        cur.execute('INSERT INTO tab1 VALUES(?,?,?)', (f"{id1}", f"sample{id1-3}", None))   #id-3=sample
    for row in cur.execute('SELECT * FROM tab1 WHERE id = ?', (f"{id1}",)):
        id, option, value = row
    cur.close()
    conn.close()
    return value

def update_value_tab1(new_id,new_value):
    db_setting = path / "db_tab1.db"
    conn = sqlite3.connect(db_setting)
    cur = conn.cursor()
    for row in cur.execute('SELECT * FROM tab1 WHERE id =?' , (f"{new_id}",)).fetchall():
        old_id, old_option, old_value = row
    cur.execute('UPDATE tab1 SET value = ? WHERE id = ?', (f"{new_value}", f"{old_id}"))
    conn.commit()
    cur.close()
    conn.close()

def call_value_tab2(id2):
    db_setting = path / "db_tab2.db"
    conn = sqlite3.connect(db_setting)
    cur = conn.cursor()
    if len(cur.execute('SELECT * FROM tab2 WHERE id = ?', (f"{id2}",)).fetchall()) == 0:
        cur.execute('INSERT INTO tab2 VALUES(?,?,?)', (f"{id2}", f"xlabel{id2-18}", None))   #id-18=Xlabel
    for row in cur.execute('SELECT * FROM tab2 WHERE id = ?', (f"{id2}",)):
        id, option, value = row
    cur.close()
    conn.close()
    return value

def update_value_tab2(new_id,new_value):
    db_setting = path / "db_tab2.db"
    conn = sqlite3.connect(db_setting)
    cur = conn.cursor()
    for row in cur.execute('SELECT * FROM tab2 WHERE id =?' , (f"{new_id}",)).fetchall():
        old_id, old_option, old_value = row
    cur.execute('UPDATE tab2 SET value = ? WHERE id = ?', (f"{new_value}", f"{old_id}"))
    conn.commit()
    cur.close()
    conn.close()

def call_value_tab3(id3):
    db_setting = path / "db_tab3.db"
    conn = sqlite3.connect(db_setting)
    cur = conn.cursor()
    deid3 = id3 % 4
    if len(cur.execute('SELECT * FROM tab3 WHERE id = ?', (f"{id3}",)).fetchall()) == 0:
        if deid3 == 0:
            cur.execute('INSERT INTO tab3 VALUES(?,?,?)',
                        (f"{id3}", f"name{id3//4+1}", None))  # name
        elif deid3 == 1:
            cur.execute('INSERT INTO tab3 VALUES(?,?,?)',
                        (f"{id3}", f"color{id3//4+1}", None))  # name
        elif deid3 == 2:
            cur.execute('INSERT INTO tab3 VALUES(?,?,?)',
                        (f"{id3}", f"style{id3//4+1}", None))  # name
        elif deid3 == 3:
            cur.execute('INSERT INTO tab3 VALUES(?,?,?)',
                        (f"{id3}", f"marker{id3//4+1}", None))  # name
    for row in cur.execute('SELECT * FROM tab3 WHERE id = ?', (f"{id3}",)):
        id, option, value = row
    cur.close()
    conn.close()
    return value

def update_value_tab3(new_id,new_value):
    db_setting = path / "db_tab3.db"
    conn = sqlite3.connect(db_setting)
    cur = conn.cursor()
    for row in cur.execute('SELECT * FROM tab3 WHERE id =?' , (f"{new_id}",)).fetchall():
        old_id, old_option, old_value = row
    cur.execute('UPDATE tab3 SET value = ? WHERE id = ?', (f"{new_value}", f"{old_id}"))
    conn.commit()
    cur.close()
    conn.close()
# ↑↑↑↑↑↑↑↑設定DB↑↑↑↑↑↑↑↑ #


# ↑↑↑↑↑↑↑↑class,def"ここまで,PysimpleGUI"ここから"↓↓↓↓↓↓↓↓#
sg.theme("Darkblue12")
#tab1のフレーム↓
#tab1_frame1のフレーム↓
#tab1_frame1のフレーム↑
tab1_frame1_lifespan = sg.Frame("",[
    [sg.Text("Number of sample types:"),
     sg.Input(default_text=call_value_tab1(0), key="-NUMBERSAMPLETYPES-",
              size=(3, 1)), sg.Text(":サンプルの種類数"),sg.Button("SAPLE NAME\n& CREATE", key="-SAMPLESURVIVAL-", button_color=("white", "blue"),size=(10,2))],
    [sg.Text("Number of sample size\nfor each sample types:"),
     sg.Input(default_text=call_value_tab1(1), key="-NUMBERSAMPLESIZE-",
              size=(4, 1)), sg.Text(":それぞれのサンプル種のサイズ(N数)")],
], border_width=1, relief=sg.RELIEF_SUNKEN)
tab1_frame1_others = sg.Frame("",[
    [sg.Text("Number of sample types:"),
     sg.Input(default_text=call_value_tab1(2), key="-NUMBERSAMPLETYPES2-",
              size=(3, 1)), sg.Text(":サンプルの種類数")],
    [sg.Text("Number of analysis:"),
     sg.Input(default_text=call_value_tab1(3), key="-NUMBERANALYSIS2-",
              size=(4, 1)), sg.Text(":解析回数\n(Kinetics用オプション）")],
    [sg.Text("")],
    [sg.Text("Name for other assay file"), sg.Input(default_text="",
                                                    key="-FILENAMEOTHER-",
                                                    size=(15, 1)),
     sg.Text(".csv"),
     sg.Button("CREATE", key="-OTHERCREATE-", button_color=("white", "blue")), sg.Text(":押下時に名前を入れないとエラーになります\n(日本語対応済)")]
], border_width=1, relief=sg.RELIEF_SUNKEN)
#tab1のフレーム↑

tab1_layout = [
    [sg.Text("")],
    [sg.Button("How to design file", key="-HOWDESIGNFILE-"),
     sg.Text(":データシート入力方法  "),
     sg.Button("Output of results", key="-OUTPUTRESULTS-"),
     sg.Text(":結果の出力について")],
    [sg.Text("")],
    [sg.Text("For survival       \n寿命解析用\nシート作成"), tab1_frame1_lifespan],
    [sg.Text("")],
    [sg.Text("For other assay\nその他解析用\nシート作成"), tab1_frame1_others],
]

#tab2のフレーム↓
tab2_frame00_file = sg.Frame("",[
    [sg.FileBrowse("Select", key="-FILENAME-", target="-NAMEOFFILE-"),
     sg.Input(default_text=call_value_tab2(0), size=(75, 1),
              key="-NAMEOFFILE-")],
], border_width=1, relief=sg.RELIEF_FLAT)
tab2_frame1_title = sg.Frame("",[
    [sg.Text("Graph title:"),
     sg.Input(default_text=call_value_tab2(1), key="-GRAPHTITLE-",
              size=(30, 1))],
    [sg.Text("Graph title size:"),
     sg.Input(default_text=call_value_tab2(2), key="-GRAPHTITLESIZE-",
              size=(2, 1))],
    [sg.Text("Graph title style:"),
     sg.Input(default_text=call_value_tab2(3), key="-GRAPHTITLESTYLE-",
              size=(10, 1)), sg.Text(":（normal or italic）")],
], border_width=1, relief=sg.RELIEF_SUNKEN)
tab2_frame2_others = sg.Frame("",[
    [sg.Checkbox("Use Japanese", default=False, key="-JAPANESE-"),
     sg.Text(":グラフに日本語文字を使う\n(フォントが変わります)")],
    [sg.Checkbox("Use Row data", default=False, key="-ROWDATA-"),
     sg.Text(":生データを使う(Not for Survival)")],
    [sg.Text("Figure size:("),
     sg.Input(default_text=call_value_tab2(15), key="-FIGURESIZEX-",
              size=(2, 1)),
     sg.Text(":"),
     sg.Input(default_text=call_value_tab2(16), key="-FIGURESIZEY-",
              size=(2, 1)),
     sg.Text(") (横X : 縦Y)")],
    [sg.Text("Numbers of sample types:"),
     sg.Input(default_text=call_value_tab2(17), key="-NUMBERSSAMPLE-",
              size=(3, 1)), sg.Text(":サンプルの種類数")]
], border_width=1, relief=sg.RELIEF_SUNKEN)
tab2_sample_design = sg.Frame("",[[sg.Button("SAMPLE DESIGN\n& ANALYSIS", key="-SAMPLEDESIGN-", button_color=("white", "blue"),size=(14,3))
]], border_width=1, relief=sg.RELIEF_SUNKEN)
tab2_frame3_horizon = sg.Frame("",[
    [sg.Text("Horizon title:"),
     sg.Input(default_text=call_value_tab2(4), key="-HORIZONTITLE-",
              size=(30, 1))],
    [sg.Text("Horizon title size:"),
     sg.Input(default_text=call_value_tab2(5), key="-HORIZONTITLESIZE-",
              size=(2, 1))],
    [sg.Text("Horizon tick size:"),
     sg.Input(default_text=call_value_tab2(6), key="-HORIZONTICKSIZE-",
              size=(2, 1))],
], border_width=1, relief=sg.RELIEF_SUNKEN)
tab2_frame4_vertical = sg.Frame("",[
    [sg.Text("Vertical title:"),
     sg.Input(default_text=call_value_tab2(7), key="-VERTICALTITLE-",
              size=(30, 1))],
    [sg.Text("Vertical title size:"),
     sg.Input(default_text=call_value_tab2(8), key="-VERTICALTITLESIZE-",
              size=(2, 1))],
    [sg.Text("Vertical tick size:"),
     sg.Input(default_text=call_value_tab2(9), key="-VERTICALTICKSIZE-",
              size=(2, 1))],
    [sg.Text("Vertical range:"),
     sg.Input(default_text=call_value_tab2(10), key="-VERTICALRANGEMIN-",
              size=(5, 1)),
     sg.Text("-"),
     sg.Input(default_text=call_value_tab2(11), key="-VERTICALRANGEMAX-",
              size=(5, 1)),
     sg.Text(":縦軸の目盛り幅")],
], border_width=1, relief=sg.RELIEF_SUNKEN)
tab2_frame5_legend = sg.Frame("",[
    [sg.Text("Legend place:"),
     sg.Input(default_text=call_value_tab2(12), key="-LEGENDPLACE-",
              size=(15, 1)), sg.Text(":（upper right:枠内の右上,upper left:枠外の左上）")],
    [sg.Text("Legend size:"),
     sg.Input(default_text=call_value_tab2(13), key="-LEGENDSIZE-",
              size=(2, 1))],
    [sg.Text("Numbers of legend rows:"),
     sg.Input(default_text=call_value_tab2(14), key="-LEGENDROWS-",
              size=(2, 1))],
], border_width=1, relief=sg.RELIEF_SUNKEN)
tab2_frame6_kinetics = sg.Frame("",[
    [sg.Text("Numbers of analysis:"),
     sg.Input(default_text=call_value_tab2(18), key="-NUMBERSANALYSIS-",
              size=(3, 1)), sg.Text(":解析回数（1回解析で終わる実験は'1'）")],
    [sg.Text("X tick label:"),
     sg.Button("Design X tick labels", key="-XTICKLABEL-"),
     sg.Text(":複数回解析の横軸の目盛り")],
     [sg.Checkbox("Adjust data by every time", default=False, key="-EVERYTIME-"),
     sg.Text(":各解析時毎にCTで補正する")],
], border_width=1, relief=sg.RELIEF_SUNKEN)
#tab2のフレーム↑

tab2_layout = [
    [sg.Text("")],
    [sg.Text("Select your file:\nファイルを選択する"), tab2_frame00_file],
    [sg.Text("")],
    [sg.Text("#Graph title:\nタイトル"), tab2_frame1_title, sg.Text("#Others:\nその他"),
     tab2_frame2_others,tab2_sample_design],
    [sg.Text("#Horizon:    \n横軸"), tab2_frame3_horizon, sg.Text("  #Vertical:\n  縦軸"),
     tab2_frame4_vertical],
    [sg.Text("#Legend:    \n凡例"), tab2_frame5_legend],
    [sg.Text("#Options for\n kinetics:\nKinetics用"), tab2_frame6_kinetics],
]


layout = [
    [sg.TabGroup([[sg.Tab("File design", tab1_layout), sg.Tab("Graph design", tab2_layout)]])],
]



window = sg.Window("Graph and Statistics", layout, resizable=True)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        exit()
    if event == "-SAMPLESURVIVAL-":
        if int(values["-NUMBERSAMPLETYPES-"]) % 4 == 0:
            numbers_row = int(int(values["-NUMBERSAMPLETYPES-"]) / 4)
        else:
            numbers_row = int(values["-NUMBERSAMPLETYPES-"]) // 4 + 1
        #エラーを防ぐために毎回layoutを定義する
        values_list1 = [
            int(values["-NUMBERSAMPLETYPES-"]),
            int(values["-NUMBERSAMPLESIZE-"]),
            int(values["-NUMBERANALYSIS2-"]),
            int(values["-NUMBERSAMPLETYPES2-"])
        ]
        for list_len in range(len(values_list1)):
            update_value_tab1(list_len,values_list1[list_len])
        survival_layout = [[sg.Text("")],
        [[sg.Text(f"{1 + 4 * row:03}:"),
          sg.Input(default_text=call_value_tab1(4 + 4 * row),
                   key=f"-SAMPLE{1 + 4 * row}-",
                   size=(10, 1)),
          sg.Text(f"{2 + 4 * row:03}:"),
          sg.Input(default_text=call_value_tab1(5 + 4 * row),
                   key=f"-SAMPLE{2 + 4 * row}-",
                   size=(10, 1)),
          sg.Text(f"{3 + 4 * row:03}:"),
          sg.Input(default_text=call_value_tab1(6 + 4 * row),
                   key=f"-SAMPLE{3 + 4 * row}-",
                   size=(10, 1)),
          sg.Text(f"{4 + 4 * row:03}:"),
          sg.Input(default_text=call_value_tab1(7 + 4 * row),
                   key=f"-SAMPLE{4 + 4 * row}-",
                   size=(10, 1))] for row in
         range(numbers_row)],
         [sg.Text("")],
         [sg.Text("Name for survival file"),
          sg.Input(default_text="", key="-FILENAMESURVIVAL-",
                  size=(15, 1)), sg.Text(".csv"),
          sg.Button("CREATE", key="-SURVIVALCREATE-", button_color=("white", "blue")),
          sg.Text(":押下時に名前を入れないとエラーになります\n(日本語対応済)")
]]
        survival_window = sg.Window('Survival sample information', layout=survival_layout)
        while True: #寿命シート作成popupのwindow用
            survival_event, survival_values = survival_window.read()
            if survival_event == sg.WIN_CLOSED:
                break
            if survival_event == "-SURVIVALCREATE-": #survivalシート作成押下時
                sample_name_list = []
                df_new_survival_file = pd.DataFrame(
                    columns=['number', 'sample', 'time', 'death'])
                df_new_survival_file['number'] = range(1, int(
                    values["-NUMBERSAMPLETYPES-"]) * int(
                    values["-NUMBERSAMPLESIZE-"]) + 1)
                df_new_survival_file['death'] = np.repeat("Death", int(
                    values["-NUMBERSAMPLETYPES-"]) * int(
                    values["-NUMBERSAMPLESIZE-"]))
                for n in range(int(values["-NUMBERSAMPLETYPES-"])):
                    exec(
                        'ndarray_sample= np.repeat(survival_values["-SAMPLE{0}-"], int(values["-NUMBERSAMPLESIZE-"]))'.format(
                            n + 1))
                    list_sample = ndarray_sample.tolist()
                    sample_name_list.extend(list_sample)
                df_new_survival_file['sample'] = sample_name_list
                exec('path_new_survival_file = path_dir / "{0}.csv"'.format(
                    survival_values["-FILENAMESURVIVAL-"]))
                df_new_survival_file.to_csv(path_new_survival_file, index=False,
                                            na_rep='', encoding="shift_jis")
                break
        survival_window.close()
        continue

    if event == "-OTHERCREATE-":
        other_sample_list = []
        for m in range(int(values["-NUMBERANALYSIS2-"])):
            other_sample_list.append('size')
            for n in range(int(values["-NUMBERSAMPLETYPES2-"])):
                exec('other_sample_list.append("N{0}Sample{1}")'.format(m + 1, n + 1))
        df_new_other_file = pd.DataFrame(columns=other_sample_list)
        df_new_other_file['size'] = range(1, 200 + 1)
        exec('path_new_other_file = path_dir / "{0}.csv"'.format(values["-FILENAMEOTHER-"]))
        df_new_other_file.to_csv(path_new_other_file, index=False,
                                 na_rep='', encoding="shift_jis")

        values_list1 = [int(values["-NUMBERSAMPLETYPES-"]),
                                 int(values["-NUMBERSAMPLESIZE-"]),
                                 int(values["-NUMBERSAMPLETYPES2-"]),
                                 int(values["-NUMBERANALYSIS2-"])]
        for list_len in range(len(values_list1)):
            update_value_tab1(list_len,values_list1[list_len])
        continue

    if event == "-HOWDESIGNFILE-":
        img = Image.open(path / "File_design.png")
        img.show()
        continue
    if event =="-OUTPUTRESULTS-":
        img = Image.open(path / "Results.png")
        img.show()
        continue

    if event == "-XTICKLABEL-": #X_tick_label用
        x_ticks_layout = [[sg.Input(default_text=call_value_tab2(column + 19), key=f"-XLABEL{column+1}-",size=(5, 1)) for column in range(int(values["-NUMBERSANALYSIS-"]))],[sg.Button("XTICKS", key="-XTICKLABEL-", button_color=("white", "blue"))]]
        x_ticks_window = sg.Window('X tick labels information', layout=x_ticks_layout)
        while True: #Xticks作成popupのwindow用
            x_ticks_event, x_ticks_values = x_ticks_window.read()
            if x_ticks_event == sg.WIN_CLOSED:
                break
            if x_ticks_event == "-XTICKLABEL-":  # Xtickslabel作成押下時
                for column in range(int(values["-NUMBERSANALYSIS-"])):
                    update_value_tab2(column+19,str(x_ticks_values[f"-XLABEL{column+1}-"]))
                break
        x_ticks_window.close()
        continue

    if event == "-SAMPLEDESIGN-":
        radio_dic = {"-SURVIVAL-": "Survival rate",
                     "-SIMPLEBAR-": "Bar graph\n(One time)",
                     "-KINETICSBAR-": "Bar graph\n(Kinetics)",
                     "-KINETICSLINE-": "Line graph\n(Kinetics)"}
        tab3_frame0_analysis = sg.Frame("", [
            [sg.Radio(item[1], key=item[0], group_id='statistic') for item in
             radio_dic.items()],
        ], border_width=1, relief=sg.RELIEF_FLAT)
        tab3_table1 = sg.Frame("", [[
            sg.Text(f"Sample {column+1:02}:"), sg.Text("Name:"),
            sg.Input(default_text=call_value_tab3(0+column*4), key=f"-SAMPLENAME{column+1}-",
                     size=(15, 1)),
            sg.Text("Color:"),
            sg.Input(default_text=call_value_tab3(1+column*4), key=f"-SAMPLECOLOR{column+1}-",
                     size=(10, 1)),
            sg.Text("Style:"),
            sg.Input(default_text=call_value_tab3(2+column*4), key=f"-SAMPLESTYLE{column+1}-",
                     size=(10, 1)),
            sg.Text("Marker:"),
            sg.Input(default_text=call_value_tab3(3+column*4), key=f"-SAMPLEMARKER{column+1}-",
                     size=(10, 1))] for column in range(int(values["-NUMBERSSAMPLE-"]))
        ], border_width=1, relief=sg.RELIEF_SUNKEN)

        sample_design_layout = [[
            [sg.Text("")],
            [sg.Button("COLOR SAMPLE", key="-COLORSAMPLE-"),
             sg.Text(":色見本  "),
             sg.Button("STYLE SAMPLE", key="-STYLESAMPLE-"),
             sg.Text(":スタイル見本  "),
             sg.Button("MARKER SAMPLE", key="-MARKERSAMPLE-"),
             sg.Text(":マーカー見本  ")],
            [sg.Text("")],
            [sg.Text("Select analysis:\n解析項目"), tab3_frame0_analysis],
            [sg.Text("")],
            [sg.Text("Sample information"),tab3_table1],
            [sg.Button("START", key="-START-", button_color=("white", "blue")),
             sg.Text("Start analysis     ", key="-SELECT-")]
        ]]
        sample_design_window = sg.Window('Sample design',
                                         layout=sample_design_layout)
        while True: #Sample desiign popupのwindow用
            sample_design_event, sample_design_values = sample_design_window.read()
            if sample_design_event == sg.WIN_CLOSED:
                start = False
                break
            if sample_design_event == "-COLORSAMPLE-":
                img = Image.open(path / "color_sample.png")
                img.show()
                continue
            if sample_design_event == "-STYLESAMPLE-":
                img = Image.open(path / "hatch_sample.png")
                img.show()
                continue
            if sample_design_event == "-MARKERSAMPLE-":
                img = Image.open(path / "marker_sample.png")
                img.show()
                continue
            start = False
            # 解析START押下時
            if sample_design_event == "-START-" and sample_design_values[
                "-SURVIVAL-"] == False and \
                    sample_design_values[
                        "-SIMPLEBAR-"] == False and sample_design_values[
                "-KINETICSBAR-"] == False and sample_design_values[
                "-KINETICSLINE-"] == False:
                start = True
                sample_design_window["-SELECT-"].update("Select analysis!")
                continue
            else:
                start = True
                values_list2 = [values["-NAMEOFFILE-"], values["-GRAPHTITLE-"],
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
                     ]
                for list_len in range(len(values_list2)):
                    update_value_tab2(list_len, values_list2[list_len])
                values_list3 = []
                for sample in range(int(values["-NUMBERSSAMPLE-"])):
                    values_list3.extend([sample_design_values[f"-SAMPLENAME{sample+1}-"], sample_design_values[f"-SAMPLECOLOR{sample+1}-"], sample_design_values[f"-SAMPLESTYLE{sample+1}-"], sample_design_values[f"-SAMPLEMARKER{sample+1}-"]])
                for list_len in range(len(values_list3)):
                    update_value_tab3(list_len, values_list3[list_len])
                sample_design_window.close()
                break
    if start == False:
        continue
    break



# ↑↑↑↑↑↑↑↑PysimpleGUI"ここまで", Figure"共通設定ここから"↓↓↓↓↓↓↓↓#
if values["-JAPANESE-"] == True:
    import japanize_matplotlib
fig = plt.figure(
    figsize=(int(values["-FIGURESIZEX-"]), int(values["-FIGURESIZEY-"])),
    dpi=240, constrained_layout=True)
ax = fig.add_subplot(1, 1, 1)
df_data = pd.read_csv(values["-NAMEOFFILE-"])
# ↑↑↑↑↑↑↑↑Figure"共通設定ここまで","グラフ作成ここから"↓↓↓↓↓↓↓↓#
sample_name_list = []
xtickslist = []
label_x = []
# ↓↓↓↓↓↓↓↓Survival"ここから"↓↓↓↓↓↓↓↓#
if sample_design_values["-SURVIVAL-"] == True:
    for n in range(int(values["-NUMBERSSAMPLE-"])):
        exec(
            'N1_S{0} = SampleInfo(sample_design_values["-SAMPLENAME{0}-"], sample_design_values["-SAMPLECOLOR{0}-"],'
            'sample_design_values["-SAMPLESTYLE{0}-"],sample_design_values["-SAMPLEMARKER{0}-"])'.format(
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
            'ax=kmf.plot(ax=ax, ci_show=False, linewidth=3, color = sample_design_values["-SAMPLECOLOR{0}-"], '
            'style = sample_design_values["-SAMPLESTYLE{0}-"], marker=sample_design_values["-SAMPLEMARKER{0}-"])'.format(
                num))

    wb_data, figure_statistics_sheet = graph_config() # グラフ共通設定
# Survival統計検定始まり↓
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
# Survival統計検定終わり↑
    save_and_plot_fig()
# ↑↑↑↑↑↑↑↑Survival"ここまで", Simplebar,Kineticsここから"↓↓↓↓↓↓↓↓#
if  sample_design_values["-SIMPLEBAR-"] == True:
    for n in range(int(values["-NUMBERSSAMPLE-"])):
        exec(
            'N1_S{0} = SampleInfo(sample_design_values["-SAMPLENAME{0}-"], sample_design_values["-SAMPLECOLOR{0}-"],'
            'sample_design_values["-SAMPLESTYLE{0}-"],sample_design_values["-SAMPLEMARKER{0}-"],df_data["N1Sample{0}"].dropna())'.format(
                n + 1))
        exec('N1_S{0}.simple_bar_cal(N1_S1.data_list)'.format(n + 1))

    wb_data, figure_statistics_sheet = graph_config() # グラフ共通設定
# Simple_bar統計検定始まり↓
    if int(values["-NUMBERSSAMPLE-"]) == 2:
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
# Simple_bar統計検定終わり↑
    save_and_plot_fig()
if sample_design_values["-KINETICSBAR-"] == True:
    for n in range(int(values["-NUMBERSANALYSIS-"])):
        exec('label_x.append(call_value_tab2({0}))'.format(n + 19))
    for n in range(int(values["-NUMBERSSAMPLE-"])):
        exec("errorS{0} = []".format(n + 1))
        exec("x_S{0} = []".format(n + 1))
        exec("y_S{0} = []".format(n + 1))
    for m in range(int(values["-NUMBERSANALYSIS-"])):
        for n in range(int(values["-NUMBERSSAMPLE-"])):
            exec(
                'N{0}_S{1} = SampleInfo(sample_design_values["-SAMPLENAME{1}-"], sample_design_values["-SAMPLECOLOR{1}-"],'
                'sample_design_values["-SAMPLESTYLE{1}-"],sample_design_values["-SAMPLEMARKER{1}-"],df_data["N{0}Sample{1}"].dropna())'.format(
                    m + 1, n + 1))
            if values["-EVERYTIME-"] == True:
                exec(
                    "rel_N{0}S{1} = N{0}_S{1}.multi_bar_cal_relative(N{0}_S1.data_list)".format(
                        m + 1,
                        n + 1))  # def relativeで縦軸サンプルごとに補正(各解析回数で補正)
                exec(
                    "sem_N{0}S{1} = N{0}_S{1}.multi_bar_cal_sem(N{0}_S1.data_list)".format(
                        m + 1, n + 1))  # def semで標準誤差計算(各解析回数で補正)
            else:
                exec(
                    "rel_N{0}S{1} = N{0}_S{1}.multi_bar_cal_relative(N1_S{1}.data_list)".format(
                        m + 1,
                        n + 1))  # def relativeで縦軸サンプルごとに補正(最初の解析回で補正)
                exec(
                    "sem_N{0}S{1} = N{0}_S{1}.multi_bar_cal_sem(N1_S{1}.data_list)".format(
                        m + 1, n + 1))  # def semで標準誤差計算(最初の解析回で補正)

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

    wb_data, figure_statistics_sheet = graph_config() # グラフ共通設定
# Kinetics_bar統計検定始まり↓
    if int(values["-NUMBERSSAMPLE-"]) == 2:     #統計検定始まり
        for n in range(int(values["-NUMBERSANALYSIS-"])):
            exec(
                "t_test(N{0}_S1.return_relative(N1_S1.data_list), N{0}_S2.return_relative(N1_S1.data_list), {0})".format(
                    n + 1))
    if int(values["-NUMBERSSAMPLE-"]) > 2:
        for m in range(int(values["-NUMBERSANALYSIS-"])):
            exec("anova_list{0} = []".format(m + 1))
            exec("combined_all_sample_data{0} = []".format(m + 1))
            exec("sample_names_and_len{0} = []".format(m + 1))
            for n in range(int(values["-NUMBERSSAMPLE-"])):
                if values["-EVERYTIME-"] == True:
                    exec(
                        'anova_list{0}.append(N{0}_S{1}.return_relative(N{0}_S4.data_list))'.format(
                            m + 1, n + 1))
                    exec(
                        'np_data{1} = np.array(N{0}_S{1}.return_relative(N{0}_S1.data_list))'.format(
                            m + 1, n + 1))
                else:
                    exec(
                        'anova_list{0}.append(N{0}_S{1}.return_relative(N1_S{1}.data_list))'.format(
                            m + 1, n + 1))
                    exec(
                        'np_data{1} = np.array(N{0}_S{1}.return_relative(N1_S{1}.data_list))'.format(
                            m + 1, n + 1))
                exec('combined_all_sample_data{0} = np.concatenate([combined_all_sample_data{0}, np_data{1}])'.format(
                    m + 1, n + 1))
                exec('name_and_len{1} = np.repeat(values["-SAMPLENAME{1}-"], len(np_data{1}))'.format(m + 1, n + 1))
                exec(
                    'sample_names_and_len{0} = np.concatenate([sample_names_and_len{0}, name_and_len{1}])'.format(m + 1,
                                                                                                                  n + 1))
        for m in range(int(values["-NUMBERSANALYSIS-"])):
            exec("one_way_anova({0},*anova_list{0})".format(m + 1))
            exec("multiple_comparison_test(combined_all_sample_data{0}, sample_names_and_len{0})".format(m + 1))
# Kinetics_line統計検定終わり↑
    save_and_plot_fig()
if sample_design_values["-KINETICSLINE-"] == True:
    for n in range(int(values["-NUMBERSANALYSIS-"])):
        exec('label_x.append(call_value_tab2({0}))'.format(n + 19))
    for m in range(int(values["-NUMBERSANALYSIS-"])):
        for n in range(int(values["-NUMBERSSAMPLE-"])):
            exec(
                'N{0}_S{1} = SampleInfo(sample_design_values["-SAMPLENAME{1}-"], sample_design_values["-SAMPLECOLOR{1}-"],'
                'sample_design_values["-SAMPLESTYLE{1}-"],sample_design_values["-SAMPLEMARKER{1}-"],df_data["N{0}Sample{1}"].dropna())'.format(
                    m + 1, n + 1))
            if values["-EVERYTIME-"] == True:
                exec(
                    "rel_N{0}S{1} = N{0}_S{1}.multi_bar_cal_relative(N{0}_S1.data_list)".format(
                        m + 1,
                        n + 1))  # def relativeで縦軸サンプルごとに補正(各解析回数で補正)
                exec(
                    "sem_N{0}S{1} = N{0}_S{1}.multi_bar_cal_sem(N{0}_S1.data_list)".format(
                        m + 1, n + 1))  # def semで標準誤差計算(各解析回数で補正)
            else:
                exec(
                    "rel_N{0}S{1} = N{0}_S{1}.multi_bar_cal_relative(N1_S{1}.data_list)".format(
                        m + 1,
                        n + 1))  # def relativeで縦軸サンプルごとに補正(最初の解析回で補正)
                exec(
                    "sem_N{0}S{1} = N{0}_S{1}.multi_bar_cal_sem(N1_S{1}.data_list)".format(
                        m + 1, n + 1))  # def semで標準誤差計算(最初の解析回で補正)
    for n in range(int(values["-NUMBERSSAMPLE-"])):  # 縦棒の数
        exec("x_S{0} = []".format(n + 1))
        exec("y_S{0} = []".format(n + 1))
        exec("errorS{0} = []".format(n + 1))
        for m in range(int(values["-NUMBERSANALYSIS-"])):
            exec("errorS{1}.append(sem_N{0}S{1})".format(m + 1, n + 1))  # 標準誤差格納
            exec("x_S{0}.append({1})".format(n + 1, m + 1))
            exec("y_S{0}.append(rel_N{1}S{0})".format(n + 1, m + 1))
    for n in range(int(values["-NUMBERSSAMPLE-"])):
        exec('sample_style = sample_design_values["-SAMPLESTYLE{0}-"]'.format(n + 1))
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
    for n in range(int(values["-NUMBERSANALYSIS-"])):
        exec("xtickslist.append({0})".format(n + 1))

    wb_data, figure_statistics_sheet = graph_config() # グラフ共通設定
# Kinetics_line統計検定始まり↓
    if int(values["-NUMBERSSAMPLE-"]) == 2:
        for n in range(int(values["-NUMBERSANALYSIS-"])):
            if values["-EVERYTIME-"] == True:
                exec(
                    "t_test(N{0}_S1.return_relative(N{0}_S1.data_list), N{0}_S2.return_relative(N{0}_S1.data_list), {0})".format(
                        n + 1))
            else:
                exec(
                    "t_test(N{0}_S1.return_relative(N1_S1.data_list), N{0}_S2.return_relative(N1_S2.data_list), {0})".format(
                        n + 1))
    if int(values["-NUMBERSSAMPLE-"]) > 2:
        for m in range(int(values["-NUMBERSANALYSIS-"])):
            exec("anova_list{0} = []".format(m + 1))
            exec("combined_all_sample_data{0} = []".format(m + 1))
            exec("sample_names_and_len{0} = []".format(m + 1))
            for n in range(int(values["-NUMBERSSAMPLE-"])):
                if values["-EVERYTIME-"] == True:
                    exec(
                        'anova_list{0}.append(N{0}_S{1}.return_relative(N1_S{1}.data_list))'.format(
                            m + 1, n + 1))
                    exec(
                        'np_data{1} = np.array(N{0}_S{1}.return_relative(N1_S{1}.data_list))'.format(
                            m + 1, n + 1))
                else:
                    exec(
                        'anova_list{0}.append(N{0}_S{1}.return_relative(N{0}_S1.data_list))'.format(
                            m + 1, n + 1))
                    exec(
                        'np_data{1} = np.array(N{0}_S{1}.return_relative(N{0}_S1.data_list))'.format(
                            m + 1, n + 1))

                exec(
                    'combined_all_sample_data{0} = np.concatenate([combined_all_sample_data{0}, np_data{1}])'.format(
                        m + 1, n + 1))
                exec(
                    'name_and_len{1} = np.repeat(values["-SAMPLENAME{1}-"], len(np_data{1}))'.format(
                        m + 1, n + 1))
                exec(
                    'sample_names_and_len{0} = np.concatenate([sample_names_and_len{0}, name_and_len{1}])'.format(
                        m + 1,
                        n + 1))
        for m in range(int(values["-NUMBERSANALYSIS-"])):
            exec("one_way_anova({0},*anova_list{0})".format(m + 1))
            exec("multiple_comparison_test(combined_all_sample_data{0}, sample_names_and_len{0})".format(m + 1))
# Kinetics_line統計検定終わり↑
    save_and_plot_fig()
exit()
