# from function import *  # 関数を名前空間に再登録
from function import (
    load_target_iraisho_files,  
    load_target_hokokusho_files,
    get_cell_value_by_cell_reference,
    process_date_value,
    calc_formula,
    add_prefix,
    split_kikaku_series,
    make_lot_no,
)

import os
import pandas as pd
import openpyxl
import re
from pathlib import Path

import function  # モジュール全体をインポート
import importlib

importlib.reload(function)  # 変更を反映して再読み込み

user_path = os.path.expanduser("~")


# 入力してもらったマッピング表を読み込む
mapping_file_path = rf"{user_path}\OneDrive - トオカツフーズ株式会社\TKDX推進室\01_develop\【冷凍生産】_簡易ツール関連\（小杉）計画外移送のフォーマット作成\作成中\計画外移送登録_マッピング表.xlsx"
mapping_wb = openpyxl.load_workbook(mapping_file_path, data_only=True)
mapping_ws = mapping_wb["自動化用_マッピング"]
mapping_df = pd.DataFrame(mapping_ws.values)
mapping_df.columns = mapping_df.iloc[0]  # 1行目をヘッダーにセット
temp_file_path = rf"{user_path}\OneDrive - トオカツフーズ株式会社\TKDX推進室\01_develop\【冷凍生産】_簡易ツール関連\（小杉）計画外移送のフォーマット作成\作成中\計画外移送登録_フォーマット_テンプレート.xlsx"


mapping_df = pd.read_excel(mapping_file_path, sheet_name="自動化用_マッピング")
temp_df = pd.read_excel(temp_file_path)


# mapping_dfから必要な項目を抽出し、辞書型で取得する()
hissu_mappings = mapping_df.columns.tolist()

# mapping_dfから行ごとにhissu_mappingsに含まれる項目を抽出し、辞書型で取得する
hissu_mapping_dicts = []
for _, row in mapping_df.iterrows():
    mapping_dict = {}
    for col in hissu_mappings:
        mapping_dict[col] = row[col]
    hissu_mapping_dicts.append(mapping_dict)
hissu_mapping_dicts


# 対象の拡張子リスト
extensions = ["xlsx", "xls", "csv"]


# マッピングから配送依頼書のデータを取得してくる

# 最終的に書類パターンごとのDataFrameを格納するリスト
merged_iraisho_list = []


# 書類パターンごとに処理
for mapping in hissu_mapping_dicts:
    # 書類名がNaNの場合はスキップ
    if mapping["書類名"] is None or pd.isna(mapping["書類名"]):
        continue

    # 出庫報告書は後ほど処理する
    elif mapping["書類種類"] == "配送依頼書＋出庫報告書":
        continue

    else:
        print(f"{mapping['書類パターン']}のファイル処理を開始します。")

        # 参照フォルダのパスの指定
        reference_folder_path = Path(
            rf"{user_path}\OneDrive - トオカツフーズ株式会社\TKDX推進室\01_develop\【冷凍生産】_簡易ツール関連\（小杉）計画外移送のフォーマット作成\作成中\計画外移送登録データ作成ツール\出庫兼配送依頼書および出庫報告書"
        )

        # 書類名パターンに基づき配送依頼書を検索、読み込んで加工し、値の取得に用いるデータフレームを取得する
        results = load_target_iraisho_files(mapping, reference_folder_path, extensions)

        for target_df, target_wb, target_ws, create_temp_df in results:
            print(f"{len(target_df)} 行を読み込みました。")

            create_temp_df["書類名"] = mapping["書類名"]
            create_temp_df["取消"] = mapping["取消"]  # 空白
            create_temp_df["工場CD"] = mapping["工場CD"]  # 空白
            create_temp_df["移送No"] = mapping["移送No"]  # 固定値
            create_temp_df["移送行"] = mapping["移送行"]  # 固定値
            create_temp_df["移送区分CD"] = mapping["移送区分CD"]  # 固定値
            create_temp_df["移送区分"] = mapping["移送区分"]  # 空白
            create_temp_df["状態CD"] = mapping["状態CD"]  # 固定値
            create_temp_df["状態"] = mapping["状態"]  # 空白

            create_temp_df["移送日(マッピング)"] = mapping["移送日"]  # セル参照
            create_temp_df["移送日(形式統一前)"] = get_cell_value_by_cell_reference(
                target_ws, mapping["移送日"]
            )  # セル参照による取得後
            create_temp_df["移送日"] = process_date_value(
                create_temp_df["移送日(形式統一前)"], 1
            )  # YYYY/MM/DDに統一後

            create_temp_df["納入日(マッピング)"] = mapping["納入日"]  # セル参照
            create_temp_df["納入日(形式統一前)"] = get_cell_value_by_cell_reference(
                target_ws, mapping["納入日"]
            )  # セル参照による取得後
            create_temp_df["納入日"] = process_date_value(
                create_temp_df["納入日(形式統一前)"], 1
            )  # YYYY/MM/DDに統一後
            create_temp_df["在庫参照日(マッピング)"] = mapping["在庫参照日"]  # セル参照

            create_temp_df["在庫参照日(形式統一前)"] = get_cell_value_by_cell_reference(
                target_ws, mapping["在庫参照日"]
            )  # セル参照による取得後
            create_temp_df["在庫参照日"] = process_date_value(
                create_temp_df["在庫参照日(形式統一前)"], 1
            )  # YYYY/MM/DDに統一後

            # TypeAとTypeCで分岐が必要な箇所
            if mapping["書類種類"] == "配送依頼書":
                create_temp_df["元保管場所CD(マッピング)"] = mapping[
                    "元保管場所CD"
                ]  # セル参照
                create_temp_df["元保管場所CD"] = get_cell_value_by_cell_reference(
                    target_ws, mapping["元保管場所CD"]
                )  # セル参照による取得後
            else:  # 出庫報告書の場合
                create_temp_df["元保管場所CD"] = mapping["元保管場所CD"]  # 固定値

            create_temp_df["元保管場所"] = mapping["元保管場所"]  # 空白
            create_temp_df["元保管棚CD"] = mapping["元保管棚CD"]  # 固定値
            create_temp_df["元保管棚"] = mapping["元保管棚"]  # 空白

            # create_temp_df['品番(36追加前)'] = target_df[mapping['品番']] #カラム参照
            # create_temp_df['品番(36追加前)_str'] = create_temp_df['品番(36追加前)'].astype(str) # 品番が入力されていない行を削除する作業。load_target_iraisho_files関数内で処理したはずだができなかったのでここで再度実行。
            # create_temp_df = create_temp_df[create_temp_df['品番(36追加前)_str'].str.isdigit()]
            create_temp_df["品番"] = target_df[mapping["品番"]].apply(
                lambda x: add_prefix(str(x))
            )  # 先頭に36を追加

            create_temp_df["品名"] = mapping["品名"]  # 空白
            create_temp_df["版"] = mapping["版"]  # 固定値
            create_temp_df["規格"] = mapping["規格"]  # 空白
            create_temp_df["先工場CD"] = mapping["先工場CD"]  # 空白
            create_temp_df["先工場"] = mapping["先工場"]  # 空白

            # 元保管場所CDと同様。カラムの並びをこの通りにしたいだけなので、処理を統一して後からカラム並び替えに変更するか検討
            if mapping["書類種類"] == "配送依頼書":
                create_temp_df["先保管場所CD(マッピング)"] = mapping[
                    "先保管場所CD"
                ]  # セル参照
                create_temp_df["先保管場所CD"] = get_cell_value_by_cell_reference(
                    target_ws, mapping["先保管場所CD"]
                )  # セル参照による取得後
            else:  # 出庫報告書の場合
                create_temp_df["先保管場所CD"] = mapping["先保管場所CD"]  # 固定値

            create_temp_df["先保管場所"] = mapping["先保管場所"]  # 空白
            create_temp_df["先保管棚CD"] = mapping["先保管棚CD"]  # 固定値
            create_temp_df["先保管棚"] = mapping["先保管棚"]  # 空白
            create_temp_df["輸送便CD"] = mapping["輸送便CD"]  # 固定値
            create_temp_df["輸送便"] = mapping["輸送便"]  # 空白

            lot_no = target_df[mapping["ロットNo"]]  # カラム参照
            # create_temp_df['ロットNo（型変換）'] = lot_no.apply(
            #     lambda x: str(int(float(x))) if str(x).replace('.', '', 1).isdigit() else str(x)
            # )
            # create_temp_df['ロットNo'] = create_temp_df['ロットNo（型変換）'].apply(lambda x: process_date_value(x, 2))
            # まず lot_no から「型変換済み」列を作成
            create_temp_df["ロットNo（型変換）"] = lot_no.apply(
                lambda x: str(x).strip()
                if re.fullmatch(r"\d{8}-\d{4}", str(x).strip())
                else (
                    str(int(float(x)))
                    if str(x).replace(".", "", 1).isdigit()
                    else str(x)
                )
            )

            # 「8桁-4桁」形式はそのまま、「それ以外」は process_date_value(x, 2) に通す
            create_temp_df["ロットNo"] = create_temp_df["ロットNo（型変換）"].apply(
                lambda x: x
                if re.fullmatch(r"\d{8}-\d{4}", str(x).strip())
                else process_date_value(x, 2)
            )

            create_temp_df["ロット枝番"] = mapping["ロット枝番"]  # 固定値
            create_temp_df["入数"] = mapping["入数"]  # 空白

            create_temp_df["移送数(マッピング)"] = mapping["移送数"]  # 計算式
            create_temp_df["移送数"] = calc_formula(
                target_df, mapping["移送数"]
            )  # 計算式による取得後

            create_temp_df["単位区分CD"] = mapping["単位区分CD"]  # 固定値
            create_temp_df["単位"] = mapping["単位"]  # 空白
            create_temp_df["個数"] = mapping["個数"]  # 空白
            create_temp_df["個単位"] = mapping["個単位"]  # 空白
            create_temp_df["換算数"] = mapping["換算数"]  # 空白
            create_temp_df["備考"] = mapping["備考"]  # 空白
            create_temp_df["丸め数（個）"] = mapping["丸め数（個）"]  # 空白
            create_temp_df["最小手配（個）"] = mapping["最小手配（個）"]  # 空白
            create_temp_df["担当者CD"] = mapping["担当者CD"]  # 空白
            create_temp_df["担当者"] = mapping["担当者"]  # 空白
            create_temp_df["担当部門CD"] = mapping["担当部門CD"]  # 空白
            create_temp_df["担当部門"] = mapping["担当部門"]  # 固定値
            create_temp_df["出庫受払No"] = mapping["出庫受払No"]  # 空白

            # display(create_temp_df.head())

            merged_iraisho_list.append(create_temp_df)

        print(f"{mapping['書類パターン']}の処理が完了しました。")


# 依頼書データをすべてまとめたDataFrameを作成
merged_iraisho_df = pd.concat(merged_iraisho_list, ignore_index=True)


# # 値取得後の配送依頼書データフレームのExcel出力
merged_iraisho_df.to_excel(
    rf"{user_path}\Desktop\計画外移送登録データ_配送依頼書の取得結果.xlsx", index=False
)


# 出庫報告書を読みこみ、データフレームとして保持する

# 最終的に書類パターンごとのDataFrameを格納するリスト
merged_hokokusho_list = []

# 書類パターンごとに処理
for mapping in hissu_mapping_dicts:
    # 書類名がNaNの場合はスキップ
    if mapping["書類名"] is None or pd.isna(mapping["書類名"]):
        continue

    # 出庫報告書（配送依頼書との突合に用いるもの）が処理対象
    elif mapping["書類種類"] != "配送依頼書＋出庫報告書":
        continue

    else:
        print(f"{mapping['書類パターン']}のファイル処理を開始します。")

        # 参照フォルダのパスの指定
        reference_folder_path = Path(
            rf"{user_path}\OneDrive - トオカツフーズ株式会社\TKDX推進室\01_develop\【冷凍生産】_簡易ツール関連\（小杉）計画外移送のフォーマット作成\作成中\計画外移送登録データ作成ツール\出庫兼配送依頼書および出庫報告書"
        )

        # 書類名パターンに基づき出庫報告書を検索、読み込んで、値の取得に用いるデータフレームを取得する
        results = load_target_hokokusho_files(
            mapping, reference_folder_path, extensions
        )

        for target_df, target_wb, target_ws, create_temp_df in results:
            # display(target_df.head())

            create_temp_df["書類名"] = mapping["書類名"]  # 固定値

            # 日付形式をYYYY/MM/DDに統一（ゼロ埋め）
            create_temp_df["移送日"] = pd.to_datetime(
                target_df[mapping["移送日"]], errors="coerce"
            ).dt.strftime("%Y/%m/%d")

            create_temp_df["渡し先名"] = target_df[
                mapping["渡し先名"]
            ]  # マッピング表には存在しないが報告書サマリには載せる項目、仮にマッピング済み

            create_temp_df["品番（マッピング）"] = target_df[
                mapping["品番"]
            ]  # カラム参照
            create_temp_df["品番"] = create_temp_df["品番（マッピング）"].apply(
                lambda x: add_prefix(str(x))
            )  # 先頭に36を追加

            # 第一倉庫冷蔵の出庫報告書には規格列がない
            if target_df.get("規格") is None:
                create_temp_df["規格"] = None
            else:
                create_temp_df["規格"] = target_df[
                    "規格"
                ]  # 移送数列のマッピングで指定される項目、マッピングしたい

            # 規格列から4桁ロット、入数、合を分離して取得
            count_series, total_series, lot_series = split_kikaku_series(target_df)
            create_temp_df["入数"] = count_series  # 規格列から分離する項目
            create_temp_df["合"] = total_series  # 規格列から分離する項目

            # 出庫個数・出庫端数列があるのは京都冷蔵のみ（標準化課題）
            for special_col in ["出庫個数", "出庫端数"]:
                if target_df.get(special_col) is None:
                    create_temp_df[special_col] = None
                else:
                    create_temp_df[special_col] = target_df[
                        special_col
                    ]  # 移送数列のマッピングで指定される項目、マッピングしたい

            create_temp_df["移送数"] = calc_formula(
                target_df, mapping["移送数"]
            )  # 独自に計算する項目

            create_temp_df["賞味期限（マッピング）"] = target_df[
                mapping["賞味期限"]
            ]  # マッピング表には存在しないが報告書サマリには載せる項目、仮にマッピング済み
            create_temp_df["賞味期限"] = process_date_value(
                create_temp_df["賞味期限（マッピング）"], 1
            )  # YYYY/MM/DDに統一後

            create_temp_df["4桁ロット"] = lot_series  # 規格列から分離する項目

            create_temp_df["ロットNo"] = create_temp_df.apply(
                lambda create_temp_df: make_lot_no(
                    create_temp_df["4桁ロット"], create_temp_df["賞味期限"]
                ),
                axis=1,
            )  # 独自に計算する項目

            # display(create_temp_df.head())

            merged_hokokusho_list.append(create_temp_df)

        print(f"{mapping['書類パターン']}の処理が完了しました。")


# 依頼書データをすべてまとめたDataFrameを作成
merged_hokokusho_df = pd.concat(merged_hokokusho_list, ignore_index=True)
# display(merged_hokokusho_df)


# 値取得後の配送依頼書データフレームのExcel出力
merged_hokokusho_df.to_excel(
    rf"{user_path}\Desktop\計画外移送登録データ_出庫報告書サマリ.xlsx", index=False
)


# ① merged_iraisho_dfのうちロットNoの有無でフィルタリングし、記載のあるもの（iraisho_df_with_lot_no）とないもの（iraisho_df_without_lot_no）を分ける

# ロットNoが数字のみ（13桁のロットを通過させるため、ハイフンを除外する）で構成されているかどうかでフィルタリング
merged_iraisho_df["ロットNo_str"] = merged_iraisho_df["ロットNo"].astype(str)

lot_lacking_iraisho_df = merged_iraisho_df[
    ~merged_iraisho_df["ロットNo_str"].str.replace("-", "", regex=False).str.isdigit()
].copy()

lot_designated_iraisho_df = merged_iraisho_df[
    merged_iraisho_df["ロットNo_str"].str.replace("-", "", regex=False).str.isdigit()
].copy()

lot_lacking_iraisho_df.drop(columns=["ロットNo_str"], inplace=True)
lot_designated_iraisho_df.drop(columns=["ロットNo_str"], inplace=True)


# ②（iraisho_df_with_lot_no）=タイプA・Cの場合、不要カラムを削除し、そのままアウトプット出力
required_columns = []
for col in lot_designated_iraisho_df.columns:
    if col in hissu_mapping_dicts[0] and hissu_mapping_dicts[0][col] == "必須":
        required_columns.append(col)

iraisho_df_with_lot_no = lot_designated_iraisho_df[required_columns].copy()
iraisho_df_with_lot_no.to_excel(
    rf"{user_path}\Desktop\計画外移送登録データ_ロット記入あり.xlsx", index=False
)


# タイプBの場合も、不要カラムを削除し、そのままアウトプット出力
required_columns = []
for col in lot_lacking_iraisho_df.columns:
    if col in hissu_mapping_dicts[0] and hissu_mapping_dicts[0][col] == "必須":
        required_columns.append(col)

iraisho_df_without_lot_no = lot_lacking_iraisho_df[required_columns].copy()
iraisho_df_without_lot_no.to_excel(
    rf"{user_path}\Desktop\計画外移送登録データ_ロット記入なし、処理前.xlsx",
    index=False,
)


# ③（lot_lacking_iraisho_df）＝タイプBの場合、merged_hokokusho_dfに対し、移送日＋品番でフィルターした結果に応じて、以下の処理をおこなう。
# （１）1種類のロットならそのロットを補填	→（iraisho_df_with_lot_no）に追加
# （２）2種類以上のロットがあったら、,(カンマ)区切りでロットを列挙し、ユーザーはそれを見てどのロットを補填するか判断


single_lot_filled_list = []
several_lot_filled_list = []
lot_unfilled_list = []

for _, row in lot_lacking_iraisho_df.iterrows():
    temp_iraisho_date = row["移送日"]
    temp_iraisho_item = row["品番"]

    # 突合
    matched_hokokusho_rows = merged_hokokusho_df[
        (merged_hokokusho_df["移送日"] == temp_iraisho_date)
        & (merged_hokokusho_df["品番"] == temp_iraisho_item)
    ]

    unique_lot_nos = matched_hokokusho_rows["ロットNo"].unique()

    # --- ① 1種類のロットが見つかった場合 ---
    if len(unique_lot_nos) == 1:
        lot_to_fill = unique_lot_nos[0]
        lot_lacking_iraisho_df.at[row.name, "ロットNo"] = lot_to_fill

        # この行のみ抽出してリストに追加
        single_lot_filled_list.append(
            lot_lacking_iraisho_df.loc[[row.name], required_columns]
        )

        print(f"１種類のロットが見つかりました: {lot_to_fill}")

    # --- ② 複数ロットが見つかった場合 ---
    elif len(unique_lot_nos) > 1:
        lot_list_str = ", ".join(map(str, unique_lot_nos))
        lot_lacking_iraisho_df.at[row.name, "ロットNo"] = lot_list_str

        # この行のみ抽出してリストに追加
        several_lot_filled_list.append(
            lot_lacking_iraisho_df.loc[[row.name], required_columns]
        )

        print(f"複数種類のロットが見つかりました: {lot_list_str}")

    else:
        lot_unfilled_list.append(
            lot_lacking_iraisho_df.loc[[row.name], required_columns]
        )
        print(f"ロットNoの検出に失敗しました: {unique_lot_nos}")

# --- 出力処理（各リストを結合してからExcel保存） ---
if single_lot_filled_list:
    single_lot_filled_df = pd.concat(single_lot_filled_list, ignore_index=True)
    single_lot_filled_df.to_excel(
        rf"{user_path}\Desktop\計画外移送登録データ_ロット記入なし（補填成功）.xlsx",
        index=False,
    )

if several_lot_filled_list:
    several_lot_filled_df = pd.concat(several_lot_filled_list, ignore_index=True)
    several_lot_filled_df.to_excel(
        rf"{user_path}\Desktop\計画外移送登録データ_ロット記入なし（候補出力）.xlsx",
        index=False,
    )


if lot_unfilled_list:
    lot_unfilled_df = pd.concat(lot_unfilled_list, ignore_index=True)
    lot_unfilled_df.to_excel(
        rf"{user_path}\Desktop\計画外移送登録データ_ロット記入なし（補填失敗）.xlsx",
        index=False,
    )
