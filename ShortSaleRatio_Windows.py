import os
import xlrd
import pprint

flist = os.listdir(path='xls')      # ディレクトリ(xls)の全ファイル名(パス含まず)のリストを返す
print( len(flist) )                 # リストのサイズを返す
xlslist = []                        # リストを定義
meigara = 9432                      # ★銘柄番号

# 拡張子xls のファイルリスト作成
for fname in flist:                 # 現在パスの「全ファイル」によるループ
    if fname.endswith('.xls'):      # ファイル名の末尾(拡張子)が'xls'なら真
        xlslist.append( fname )     # xlsファイル名をリストに追加


# 拡張子xls のファイルリスト作成
for fname in xlslist:                               # 現在パスの「全xlsファイル」によるループ
    tmp_list = []
    wb = xlrd.open_workbook( "xls/" + fname )       # xlsファイルのBookオブジェクトを取得
    tmp_list.append( wb._sheet_names[0] )           # リストにシート名を追加

    sheet = wb.sheet_by_name( wb._sheet_names[0] )  # 指定シートを取得

    lineno = 8
    hit = 0
    value = 0
    while lineno < sheet.nrows:                 # 最終行まで
        cell = sheet.cell( lineno, 2 )          # セルを読む
        if cell.ctype == 0:                     # セルが空白の場合
            break                               # ループを抜ける
        if cell.value == meigara:               # セルが銘柄番号に該当？
            hit = 1                             # 該当銘柄があったフラグ
            cell = sheet.cell( lineno, 10 )     # セルを読む
            value += cell.value                 # 「空売り残高割合」を加算
        lineno += 1                             # 行+1

    if hit == 1:                                # 該当銘柄があった場合
        tmp_list.append( value * 100)           # 「空売り残高割合」をリストに追加
        print( tmp_list )                       # 「空売り残高割合」を出力

