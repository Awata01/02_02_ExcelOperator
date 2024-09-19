import excel_operator as eo

#  01, 02
# trg_path = r"C:\Users\awata\　awata\Programing\CheckOut\02_ExcelOperator\Doccument"
# excel_files = eo.ExcelOperator.get_files_in_path(trg_path, True, True)
# print(excel_files)


# 03
# trg_path = r"C:\Users\awata\　awata\Programing\CheckOut\00_Common\Doccument\02_SourceTreeインストール手順.xlsx"
# trg_path = r"C:\Users\awata\　awata\Programing\CheckOut\02_ExcelOperator\Doccument\01_実装機能一覧.xlsx"
# trg_path = r"C:\Users\awata\　awata\Programing\CheckOut\02_ExcelOperator\Doccument"
# excel_files = eo.ExcelOperator.get_sheets_name(trg_path)
# print(excel_files)

# # 04
# trg_path = r"C:\Users\awata\　awata\Programing\CheckOut\02_ExcelOperator\Doccument\新しいフォルダー\aaa.xlsx"
# excel_files = eo.ExcelOperator.sort_sheet(trg_path, 'desc')
# print(excel_files)

# # 05
# trg_path = r"C:\Users\awata\　awata\Programing\CheckOut\02_ExcelOperator\Doccument\新しいフォルダー\aaa.xlsx"
# trg_path2 = r"C:\Users\awata\　awata\Programing\CheckOut\02_ExcelOperator\Doccument\新しいフォルダー\abc.csv"
# eo.ExcelOperator.convert_csv(trg_path, '1', trg_path2)

# 09
trg_path = r"C:\Users\awata\　awata\Programing\CheckOut\02_ExcelOperator\Doccument\新しいフォルダー\aaa.xlsx"
# eo.ExcelOperator.change_font(trg_path, '4')
eo.ExcelOperator.change_font(trg_path, '4', 'ＭＳ ゴシック')