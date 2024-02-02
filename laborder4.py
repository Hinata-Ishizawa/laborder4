import pandas as pd
import os
import sys
import configparser

def mail_writing(who, how_many, name, maker, code, volume):
    print('本文:\n\nお世話になっております。京都大学雑草学研究室の'+ who + 'です。\nこの度は商品発注をお願いしたく連絡させていただきました。\n以下の商品'+ how_many + '個の発注をお願いいたします。\n')
    if 'Series([], )' in name:
            pass
    else:
            print('商品名:' + name)
    if 'Series([], )' in maker:
            pass
    else:
            print('メーカー:' + maker)
    if 'Series([], )' in code:
            pass
    else:
            print('コード:' + code)
    if 'Series([], )' in volume:
            pass
    else:
            print('容量:' + volume)
    print('\nよろしくお願いいたします。\n')
    print(who + '\n')
        
def extract_rows_with_string(file_path, target_string):
    # エクセルファイルを読み込んでデータフレームに格納
    df = pd.read_excel(file_path, skiprows=7)

    # 特定の文字列を含む行を抽出
    selected_rows = df[df.apply(lambda row: target_string in str(row), axis=1)]

    # 空白のセルを除外
    selected_rows_2 = selected_rows.dropna(axis=1)

    # 行番号の振り直し
    selected_rows_3 = selected_rows_2.reset_index(drop=True)

    # 複数行が抽出された場合、ユーザーに選択させる
    if len(selected_rows_3) > 1:
        while True:
            print("複数の行が抽出されました。どの行を表示しますか？")
            print(selected_rows_3)

            # ユーザーの選択を取得
            user_choice = int(input("行番号を入力してください: "))

            # 注文数を取得
            how_many = input("いくつ注文しますか？: ")

            # 注文内容を確認
            selected_rows_4 = selected_rows_3.iloc[user_choice:user_choice+1]
            print(selected_rows_4)
            print("を" + how_many + "個注文で良いですか？")
            kakunin = input("どちらかを選んでください (yまたはn): ")

            #
            if kakunin == "y":
                break
            elif kakunin == "n":
                continue
            else:
                print("無効な選択です。'y' または 'n' を入力してください。")

        # 必要な情報の抽出
        selected_rows_5 = selected_rows.iloc[user_choice:user_choice+1]

        name = selected_rows_5['商品名'].dropna().to_string(index=False)
        maker = selected_rows_5['メーカー'].dropna().to_string(index=False)
        code = selected_rows_5['コード'].dropna().to_string(index=False)
        volume = selected_rows_5['容量'].dropna().to_string(index=False)
        
        # config.iniからユーザー名を抽出
        config = configparser.ConfigParser()
        import os
        dir = os.path.dirname(sys.argv[0])
        if dir == "":
            dir = "."
        config_path = os.path.join(dir, './output/config.ini')
        config.read(config_path)
        who = config['User']['user_name']

        # メール文の記述
        mail_writing(who, how_many, name, maker, code, volume)
        
        # dictionaryの作成
        data_dict = {
            "発注者": [who],
            "個数": [how_many]
        }
        
        # 抽出した情報をdictionaryに格納
        if 'Series([], )' in name:
            data_dict["品名"] = [""]  # "品名"に対応するvalueは無し
        else:
            data_dict["品名"] = [name]  # "品名"のvalueとしてnameを格納
            
        if 'Series([], )' in volume:
            data_dict["容量"] = [""]
        else:
            data_dict["容量"] = [volume]
            
        if 'Series([], )' in code:
            data_dict["コード"] = [""]
        else:
            data_dict["コード"] = [code]
        
        # datetimeモジュールによって日付情報を取得する
        from datetime import datetime
        date = datetime.now().strftime("%-m月%-d日")  # strftime()でフォーマットを指定
        
        # 日付情報をdictionaryに格納
        data_dict["発注日"] = [date]
        
        # 下書きエクセルファイルの読み込み
        import os
        dir = os.path.dirname(sys.argv[0])
        if dir == "":
            dir = "."
        order_info_path = os.path.join(dir, './order_info.xlsx')
        order_info_df = pd.read_excel(order_info_path)
        
        # dictionaryの各keyに対応するvalueがあればエクセルに記入
        for key, value in data_dict.items():
            if key in order_info_df.columns:
                order_info_df[key] = value
        
        # 書き込んだデータをorder_infoに保存
        order_info_df.to_excel(order_info_path, index=False)
        
        print('エンターキーを押してください...')
        _ = input()
        
        # order_info.xlxsを開く
        os.system(f'open "{order_info_path}"')
        
        # 発注履歴.xlsxのパスを取得
        history_path = config['Files']['history']
        # 発注履歴.xlsxを開く
        os.system(f'open "{history_path}"')

    elif len(selected_rows_3) == 1:
        # 1行だけ抽出された場合はその行を表示
        print(selected_rows_3)

        # 注文数を取得
        how_many = input("いくつ注文しますか？: ")

        # 必要な情報の抽出
        name = selected_rows_3['商品名'].dropna().to_string(index=False)
        maker = selected_rows_3['メーカー'].dropna().to_string(index=False)
        code = selected_rows['コード'].dropna().to_string(index=False)
        volume = selected_rows_3['容量'].dropna().to_string(index=False)
        
        # config.iniからユーザー名を抽出
        config = configparser.ConfigParser()
        import os
        dir = os.path.dirname(sys.argv[0])
        if dir == "":
            dir = "."
        config_path = os.path.join(dir, './output/config.ini')
        config.read(config_path)
        who = config['User']['user_name']

        # メール文の記述
        mail_writing(who, how_many, name, maker, code, volume)
        
        # dictionaryの作成
        data_dict = {
            "発注者": [who],
            "個数": [how_many]
        }
        
        # 抽出した情報をdictionaryに格納
        if 'Series([], )' in name:
            data_dict["品名"] = [""]  # "品名"に対応するvalueは無し
        else:
            data_dict["品名"] = [name]  # "品名"のvalueとしてnameを格納
            
        if 'Series([], )' in volume:
            data_dict["容量"] = [""]
        else:
            data_dict["容量"] = [volume]
            
        if 'Series([], )' in code:
            data_dict["コード"] = [""]
        else:
            data_dict["コード"] = [code]
        
        # datetimeモジュールによって日付情報を取得する
        from datetime import datetime
        date = datetime.now().strftime("%-m月%-d日")  # strftime()でフォーマットを指定
        
        # 日付情報をdictionaryに格納
        data_dict["発注日"] = [date]
        
        # 下書きエクセルファイルの読み込み
        import os
        dir = os.path.dirname(sys.argv[0])
        if dir == "":
            dir = "."
        order_info_path = os.path.join(dir, './order_info.xlsx')
        order_info_df = pd.read_excel(order_info_path)
        
        # dictionaryの各keyに対応するvalueがあればエクセルに記入
        for key, value in data_dict.items():
            if key in order_info_df.columns:
                order_info_df[key] = value
        
        # 書き込んだデータをorder_infoに保存
        order_info_df.to_excel(order_info_path, index=False)
        
        print('エンターキーを押してください...')
        _ = input()
        
        # order_info.xlxsを開く
        os.system(f'open "{order_info_path}"')
        
        # 発注履歴.xlsxのパスを取得
        history_path = config['Files']['history']
        # 発注履歴.xlsxを開く
        os.system(f'open "{history_path}"')

    else:
        print("該当する行はありませんでした。")


# config.iniのパスを取得する
dir = os.path.dirname(sys.argv[0])
if dir == "":
    dir = "."
config_path = os.path.join(dir, './output/config.ini')

# config.iniが存在することを確認
# 存在しなければ作成
if os.path.exists(config_path):
    pass
else:
    print("Input your name: ")
    user_name = input()
    print("Input path to '消耗品価格表.xlsx': ")
    price_path = input()
    print("Input path to '発注履歴.xlsx': ")
    history_path = input()
    
    config = configparser.ConfigParser()
    config['User'] = {
    'user_name': user_name
    }
    config['Files'] = {
    'price': price_path,
    'history': history_path
    }
    
    # config.iniを保存
    with open(config_path, 'w') as f:
        config.write(f)
        print("wrote config")

# 消耗品価格表のパスを読み込む
config = configparser.ConfigParser()
config.read(config_path)
price_excel_path = config['Files']['price']

# 特定の文字列を入力
target_string = input("検索ワードを入力してください: ")

# 関数を呼び出して処理を開始
extract_rows_with_string(price_excel_path, target_string)
