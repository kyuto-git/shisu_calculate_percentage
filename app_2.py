import pandas as pd
import os

def process_csv_files(csv_dir, output_dir, output_file_name):
    # Outputディレクトリが存在しない場合は作成
    os.makedirs(output_dir, exist_ok=True)

    # 出力ファイルのフルパスを設定
    output_file = os.path.join(output_dir, output_file_name)

    # Excelライターを作成
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # CSVファイルのリストを取得
        csv_files = [f for f in os.listdir(csv_dir) if f.endswith('.csv')]

        # 各CSVファイルを読み込んでExcelの異なるシートに書き出す
        for csv_file in csv_files:
            df = pd.read_csv(os.path.join(csv_dir, csv_file))
            # CSVファイル名（拡張子なし）をシート名として使用
            sheet_name = os.path.splitext(csv_file)[0]
            df.to_excel(writer, sheet_name=sheet_name, index=False)

# 使用例
csv_dir = 'csv_files'  # CSVファイルが格納されているディレクトリ
output_dir = 'output'   # 出力ディレクトリ
output_file_name = 'output.xlsx'  # 出力するExcelファイル名
process_csv_files(csv_dir, output_dir, output_file_name)


# if __name__ == "__main__":
#     # process_csv_files()
#     process_csv_files(csv_dir, output_dir, output_file_name)

