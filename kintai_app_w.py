import sys
import pandas as pd
import re
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLineEdit, QTextEdit, QLabel,
    QFileDialog, QMessageBox, QRadioButton, QButtonGroup,
    QDialog, QTextEdit as QEdit, QDialogButtonBox
)
from PySide6.QtGui import QFont
from PySide6.QtGui import QIcon


# --------------------------------------------------------------
# NGワード編集用ダイアログ
# --------------------------------------------------------------
class EditWordsDialog(QDialog):
    def __init__(self, words):
        super().__init__()
        self.setWindowTitle("休みワード編集")
        self.resize(400, 300)

        layout = QVBoxLayout()
        layout.addWidget(QLabel("1行に1つ入力してください："))

        self.edit = QEdit()
        self.edit.setPlainText("\n".join(words))
        layout.addWidget(self.edit)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

        self.setLayout(layout)

    def get_words(self):
        lines = [w.strip() for w in self.edit.toPlainText().split("\n")]
        return [w for w in lines if w]


# --------------------------------------------------------------
# メインアプリ
# --------------------------------------------------------------
class KintaiApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("勤務者検索ツール")
        self.resize(820, 900)
        

        self.df = None

        # ✅ 初期 NG ワード
        self.NG_WORDS = ["－","深あ","夜わ","夜さ","夜こ","休", "休み", "年休"]

        layout = QVBoxLayout()

        # ----------------------------------------------------------
        # Excel 読み込み
        # ----------------------------------------------------------
        load_btn = QPushButton("Excelファイルを選択")
        load_btn.clicked.connect(self.load_excel)
        layout.addWidget(load_btn)

        # ----------------------------------------------------------
        # 検索モード
        # ----------------------------------------------------------
        mode_layout = QHBoxLayout()
        self.mode_group = QButtonGroup()

        self.radio_day = QRadioButton("日付検索")
        self.radio_day.setChecked(True)

        self.radio_name = QRadioButton("名前検索")
        self.radio_range = QRadioButton("日付範囲検索")
        self.radio_compare = QRadioButton("2人比較")

        for rb in [self.radio_day, self.radio_name, self.radio_range, self.radio_compare]:
            self.mode_group.addButton(rb)
            mode_layout.addWidget(rb)

        layout.addLayout(mode_layout)

        # ----------------------------------------------------------
        # 各入力欄（widget ごと show/hide 切替）
        # ----------------------------------------------------------
        # 日付単体
        self.single_widget = QWidget()
        sw = QHBoxLayout(self.single_widget)
        sw.addWidget(QLabel("日付："))
        self.single_input = QLineEdit()
        self.single_input.setFixedWidth(80)
        sw.addWidget(self.single_input)
        layout.addWidget(self.single_widget)

        # 名前
        self.name_widget = QWidget()
        nw = QHBoxLayout(self.name_widget)
        nw.addWidget(QLabel("名前："))
        self.name_input = QLineEdit()
        self.name_input.setFixedWidth(160)
        nw.addWidget(self.name_input)
        layout.addWidget(self.name_widget)
        self.name_widget.hide()

        # 日付範囲
        self.range_widget = QWidget()
        rw = QHBoxLayout(self.range_widget)
        rw.addWidget(QLabel("開始："))
        self.range_start = QLineEdit()
        self.range_start.setFixedWidth(80)
        rw.addWidget(self.range_start)
        rw.addWidget(QLabel("終了："))
        self.range_end = QLineEdit()
        self.range_end.setFixedWidth(80)
        rw.addWidget(self.range_end)
        layout.addWidget(self.range_widget)
        self.range_widget.hide()

        # 2人比較
        self.compare_widget = QWidget()
        cw = QHBoxLayout(self.compare_widget)
        cw.addWidget(QLabel("名前1："))
        self.comp1 = QLineEdit()
        self.comp1.setFixedWidth(150)
        cw.addWidget(self.comp1)
        cw.addWidget(QLabel("名前2："))
        self.comp2 = QLineEdit()
        self.comp2.setFixedWidth(150)
        cw.addWidget(self.comp2)
        layout.addWidget(self.compare_widget)
        self.compare_widget.hide()

        # 入力欄切替
        self.radio_day.toggled.connect(self.update_inputs)
        self.radio_name.toggled.connect(self.update_inputs)
        self.radio_range.toggled.connect(self.update_inputs)
        self.radio_compare.toggled.connect(self.update_inputs)

        # ----------------------------------------------------------
        # 結果欄（通常：1つ）
        # ----------------------------------------------------------
        self.result_single = QTextEdit()
        self.result_single.setReadOnly(True)
        self.result_single.setMinimumHeight(500)
        layout.addWidget(self.result_single)

        # ----------------------------------------------------------
        # 日付検索専用：左右 6:4 の分割
        # ----------------------------------------------------------
        self.result_split_widget = QWidget()
        split = QHBoxLayout(self.result_split_widget)

        self.result_left = QTextEdit()
        self.result_left.setReadOnly(True)

        self.result_right = QTextEdit()
        self.result_right.setReadOnly(True)

        split.addWidget(self.result_left, 6)
        split.addWidget(self.result_right, 4)

        layout.addWidget(self.result_split_widget)
        self.result_split_widget.hide()

        # ----------------------------------------------------------
        # NGワード編集
        # ----------------------------------------------------------
        edit_btn = QPushButton("休みワード編集")
        edit_btn.clicked.connect(self.edit_ng_words)
        layout.addWidget(edit_btn)

        # ----------------------------------------------------------
        # Excel 出力
        # ----------------------------------------------------------
        export_btn = QPushButton("結果をExcel保存")
        export_btn.clicked.connect(self.export_to_excel)
        layout.addWidget(export_btn)

        # ----------------------------------------------------------
        # 検索ボタン
        # ----------------------------------------------------------
        search_btn = QPushButton("検索")
        search_btn.clicked.connect(self.on_search)
        layout.addWidget(search_btn)

        self.setLayout(layout)

    # ==============================================================
    # Excel 読み込み
    # ==============================================================
    def load_excel(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Excel選択", "", "Excel (*.xlsx *.xls)"
        )
        if not file:
            return

        df = pd.read_excel(file)

        # 日付列を "xx日" に補正
        new_cols = []
        for c in df.columns:
            s = str(c)
            m = re.search(r"(\d+)", s)
            if m:
                d = int(m.group(1))
                if 1 <= d <= 31:
                    new_cols.append(f"{d}日")
                    continue
            new_cols.append(c)

        df.columns = new_cols
        self.df = df

        QMessageBox.information(self, "成功", "Excelを読み込みました。")

    # ==============================================================
    # 入力欄切替
    # ==============================================================
    def update_inputs(self):
        self.single_widget.hide()
        self.name_widget.hide()
        self.range_widget.hide()
        self.compare_widget.hide()

        if self.radio_day.isChecked():
            self.single_widget.show()
        elif self.radio_name.isChecked():
            self.name_widget.show()
        elif self.radio_range.isChecked():
            self.range_widget.show()
        elif self.radio_compare.isChecked():
            self.compare_widget.show()

        # ✅ 必ず通常結果欄
        self.result_single.show()
        self.result_split_widget.hide()

    # ==============================================================
    # 休みワード編集
    # ==============================================================
    def edit_ng_words(self):
        dlg = EditWordsDialog(self.NG_WORDS)
        if dlg.exec():
            self.NG_WORDS = dlg.get_words()
            QMessageBox.information(self, "保存", "休みワードを更新しました。")

    # ==============================================================
    # 検索ボタン押下
    # ==============================================================
    def on_search(self):
        if self.df is None:
            QMessageBox.warning(self, "警告", "先にExcelを読み込んでください。")
            return

        # ✅ 必ず初期化
        self.result_single.clear()
        self.result_split_widget.hide()
        self.result_single.show()

        if self.radio_day.isChecked():
            self.search_day()
        elif self.radio_name.isChecked():
            self.search_name()
        elif self.radio_range.isChecked():
            self.search_range()
        elif self.radio_compare.isChecked():
            self.search_compare()

    # ==============================================================
    # ① 日付検索（左右分割）
    # ==============================================================
    def search_day(self):
        day = self.single_input.text().strip()
        if not day.isdigit():
            QMessageBox.critical(self, "エラー", "日付は数字で入力してください。")
            return

        col = f"{day}日"
        if col not in self.df.columns:
            QMessageBox.critical(self, "エラー", f"{col} が見つかりません。")
            return

        # ✅ 2分割モードへ
        self.result_single.hide()
        self.result_split_widget.show()
        self.result_left.clear()
        self.result_right.clear()

        name_col = self.df.columns[0]

        # 左分類
        normal_list = []
        hyphen_list = []

        # 右分類（順番固定）
        order = ["深あ","夜わ","夜さ","夜こ","休", "休み", "年休", "nan", ""]
        buckets = {k: [] for k in order}

        for _, row in self.df.iterrows():
            name = row[name_col]
            work = str(row[col]).strip()

            # 左側
            if work not in self.NG_WORDS and work not in ["", "nan"]:
                if work == "ー":
                    normal_list.append((work, name))
                
                    
                
                else:
                    normal_list.append((work, name))
                continue

            # 右側
            matched = False
            for key in order:
                if key and key in work:
                    buckets[key].append((work, name))
                    matched = True
                    break
            if not matched:
                buckets[""].append((work, name))

        # 出力 左
        self.result_left.append(f"【{col}：通常勤務】\n")
        for w, n in normal_list:
            self.result_left.append(f"{w}    {n}")
        


        if hyphen_list:
            self.result_left.append("\n―― － の勤務 ――")
            for w, n in normal_list:
                self.result_left.append(f"{w}    {n}")
             

        # 出力 右
        self.result_right.append(f"【{col}：休・特殊勤務】\n")
        for key in order:
            for w, n in buckets[key]:
                val = w if w not in ["", "nan"] else "nan"
                self.result_right.append(f"{val}    {n}")

    # ==============================================================
    # ② 名前検索
    # ==============================================================
    def search_name(self):
        name = self.name_input.text().strip()
        col0 = self.df.columns[0]

        if name not in list(self.df[col0]):
            QMessageBox.critical(self, "エラー", "名前が存在しません。")
            return

        row = self.df[self.df[col0] == name].iloc[0]

        self.result_single.append(f"【{name} の勤務一覧】\n")
        for d in self.df.columns[1:]:
            w = str(row[d]).strip()
            if w not in ["", "nan", ""]:
                self.result_single.append(f"{d}: {w}")

    # ==============================================================
    # ③ 日付範囲検索
    # ==============================================================
    def search_range(self):
        s = self.range_start.text().strip()
        e = self.range_end.text().strip()

        if not (s.isdigit() and e.isdigit()):
            QMessageBox.critical(self, "エラー", "開始/終了は数字で入力")
            return

        s, e = int(s), int(e)
        if s > e:
            QMessageBox.critical(self, "エラー", "開始日は終了日より前")
            return

        col0 = self.df.columns[0]

        self.result_single.append(f"【{s}日〜{e}日】\n")

        for _, row in self.df.iterrows():
            name = row[col0]
            self.result_single.append(f"＜{name}＞")
            for d in range(s, e + 1):
                col = f"{d}日"
                if col in self.df.columns:
                    w = str(row[col]).strip()
                    if w not in ["", "nan", "休"]:
                        self.result_single.append(f"{col}: {w}")
            self.result_single.append("")

    # ==============================================================
    # ④ 2人比較
    # ==============================================================
    def search_compare(self):
        a = self.comp1.text().strip()
        b = self.comp2.text().strip()
        col0 = self.df.columns[0]

        if a not in list(self.df[col0]) or b not in list(self.df[col0]):
            QMessageBox.critical(self, "エラー", "名前が見つかりません。")
            return

        r1 = self.df[self.df[col0] == a].iloc[0]
        r2 = self.df[self.df[col0] == b].iloc[0]

        self.result_single.append(f"【{a} vs {b}】\n")
        hdr = f"{'日付':<8}{a:<15}{b:<15}"
        self.result_single.append(hdr)
        self.result_single.append("-" * len(hdr))

        for d in self.df.columns[1:]:
            w1 = str(r1[d]).strip()
            w2 = str(r2[d]).strip()

            w1 = w1 if w1 not in ["", "nan"] else "休"
            w2 = w2 if w2 not in ["", "nan"] else "休"

            self.result_single.append(f"{d:<8}{w1:<15}{w2:<15}")

    # ==============================================================
    # Excel 保存（左右長さを揃えて保存）
    # ==============================================================
    def export_to_excel(self):
    # 左側（通常勤務）だけを Excel に保存する
        path, _ = QFileDialog.getSaveFileName(
            self, "Excelに保存", "", "Excel Files (*.xlsx)"
        )
        if not path:
            return

        try:
            lines = self.result_left.toPlainText().split("\n")

            work_list = []
            name_list = []

            for line in lines:
                line = line.strip()
                if not line:
                    continue

            # 行のフォーマット： "work    name"
                parts = line.split()
                if len(parts) >= 2:
                   work = parts[0]
                   name = parts[-1]
                else:
                # 名前だけ / 勤務だけなどの例外処理
                    work = parts[0]
                    name = ""

                work_list.append(work)
                name_list.append(name)

            df_out = pd.DataFrame({
                "勤務": work_list,
                "名前": name_list
            })

            df_out.to_excel(path, index=False)

            QMessageBox.information(self, "保存完了", "左側の通常勤務だけをExcelに保存しました。")

        except Exception as e:
            QMessageBox.critical(self, "エラー", str(e))

# --------------------------------------------------------------
# アプリ起動
# --------------------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)

    # 全体を1.1倍
    font = QFont()
    font.setPointSize(int(font.pointSize() * 1.1))
    app.setFont(font)

    w = KintaiApp()
    w.show()
    sys.exit(app.exec())