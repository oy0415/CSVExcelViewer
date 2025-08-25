# CSVExcelViewer
# CSVExcelViewer — CSV/Excel結合 & ビューア (TkEasyGUI)

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](#license)
[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](#%E5%BF%85%E8%A6%81%E7%92%B0%E5%A2%83)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-lightgrey.svg)]()

CSV と Excel（.xlsx/.xlsm）を **まとめて読み込み→結合→テーブルでプレビュー→CSV保存** まで行える Windows デスクトップツールです。  
Tkinter ベースの軽量 GUI フレームワーク **TkEasyGUI** を採用し、非エンジニアにも扱いやすい操作性を意識しています。

> 📦 すぐ使いたい方は **Releases** から `CSVExcelViewer.exe` をダウンロードしてください。

---

## 主な機能

- CSV / Excel（.xlsx / .xlsm）の**同時選択 & 結合**
- 文字コード（utf-8-sig / UTF-8 / CP932 / EUC-JP）と区切り文字（`,` `;` `\t` `|`）の**自動判定**
- ヘッダー行の**重複スキップ**、**空行スキップ**
- Excel の日付・日時を `"YYYY-MM-DD"` / `"YYYY-MM-DD HH:MM:SS"` に整形
- GUI テーブルで**プレビュー**し、そのまま **UTF-8(BOM) でCSV保存**
- **ポータブル運用**（レジストリ書き込みなし。任意フォルダで実行可能）

> 実装の要点は `src/csv_excel_viewer.py` を参照してください。

---

## 画面イメージ（例）

> `screenshots/` に PNG / GIF を配置してください（例: ScreenToGif 等で録画）

---

## 必要環境

- Windows 10 / 11
- Python 3.10 以上（実行のみなら EXE も可）

---

## セットアップ（ソースから実行）

```bash
# 1) 仮想環境は任意。ここでは venv の例
python -m venv .venv
.venv\Scripts\activate

# 2) 依存ライブラリのインストール
pip install -r requirements.txt

# 3) 実行
python src/csv_excel_viewer.py
```

---

## 使い方

1. **ファイル選択**を押して、CSV と Excel を複数選択（ドラッグ選択可）  
2. プレビュー画面で内容を確認（列幅は自動調整）  
3. **保存** を押して結合結果を CSV で出力（UTF-8 BOM）

---

## フォルダ構成

```
CSVExcelViewer/
├─ src/
│   └─ csv_excel_viewer.py
├─ screenshots/        # スクショやGIF（任意）
├─ examples/           # サンプルCSV/Excel（任意）
├─ .gitignore
├─ requirements.txt
├─ LICENSE
└─ README.md
```


---

## EXE の作り方（PyInstaller）

PyInstaller を使えば、Python 未導入の PC でも配布できます。

```bash
pip install pyinstaller
cd src

# 1ファイル（持ち運び最小）
pyinstaller --noconsole --onefile csv_excel_viewer.py

# フォルダ配布（起動が速い）
pyinstaller --noconsole --onedir csv_excel_viewer.py
```

- 生成物は `dist/` に出力されます。
- `--onefile` 版を **Releases** にアップロードし、`CSVExcelViewer.exe` として配布するのがおすすめです。
- 相対パスで動く実装のため、**任意フォルダへ移動しても動作**します。

---

## よくある質問（FAQ）

- **Q. 文字化けします。**  
  A. 主要な文字コードを自動判定しますが、稀に判定できない場合があります。CSV を UTF-8(BOM) に変換してからお試しください。

- **Q. 大きなファイルで固まります。**  
  A. 現在はメモリ上で結合しています。将来的に分割読み込みに対応予定です（Issue で要望ください）。

---

## 開発者向けメモ

- UI は Tkinter をラップする **TkEasyGUI** を使用
- CSV の**区切り文字**は `csv.Sniffer` による推定を優先、失敗時は Excel 既定の `,` を採用
- Excel 読み込みは **openpyxl**（`read_only=True`, `data_only=True`）を使用

---

## ライセンス

本ソフトウェアは MIT License の下で配布します。詳細は [LICENSE](./LICENSE) をご覧ください。

---

## Author

- name:小熊優介(oguma yuusuke)
- お仕事のご相談は Issues またはプロフィールの連絡先へ
