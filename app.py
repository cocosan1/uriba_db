import streamlit as st
import pandas as pd
import os
import shutil
from PIL import Image

from google.oauth2 import service_account
from googleapiclient.discovery import build

# pip install streamlit pandas pillow openpyxl google-api-python-client

st.set_page_config('売り場画像抽出app', layout='wide')
st.markdown('# 売り場画像抽出app')

#current working dir
cwd = os.path.dirname(__file__)

# 関数 driveからファイルの取得と保存
# 関数get_file_from_gdrive内で使用
def get_and_save_file(service, folder, file_name, mimeType):
    #Google Drive上のファイルを検索するためのクエリ
    # ファイル名とMIMEタイプを指定。
    # query = f"name='画像情報.xlsx' and mimeType='application/vnd.ms-excel'"
    query = f"name='{file_name}' and mimeType= '{mimeType}'"

    #files().list(q=query)　指定されたクエリに一致するファイルのリストを取得
    #execute() APIリクエストを実行し、結果（ファイルのリスト含む）を返す。
    results = service.files().list(q=query).execute()
    #ファイルのリストを抽出。無い場合は空のリストを返す
    items = results.get("files", [])

    if not items:
        st.warning(f"No files found with name: {file_name}")

    else:
        file_id = items[0]["id"]
        # get_media(fileId=file_id) ファイルを取得するためのメソッド
        file_content = service.files().get_media(fileId=file_id).execute()

        # ファイルを保存する
        file_path = os.path.join(cwd, folder, file_name)
        with open(file_path, "wb") as f:
            f.write(file_content)

# 関数　画像ファイルの圧縮



def get_file_from_gdrive(cwd, folder, file_name_list, mimeType):

    # Google Drive APIを使用するための認証情報を生成
    creds_dict = st.secrets["gcp_service_account"]
    creds = service_account.Credentials.from_service_account_info(creds_dict)

    # Drive APIとやり取りするクライアントを作成
    #API名（ここでは"drive"）、APIのバージョン（ここでは"v3"）、および認証情報を指定
    service = build("drive", "v3", credentials=creds)

    ## フォルダの新規作成　存在する場合一度削除
    # 作成したいフォルダの名前
    folder_name = os.path.join(cwd, folder)

    # フォルダが存在するかどうかを確認
    if os.path.exists(folder_name):
        # フォルダが存在する場合は、shutil.rmtree()で削除
        shutil.rmtree(folder_name)

    # フォルダを作成
    os.mkdir(folder_name)

    # ファイルのdriveからの取り出しと保存
    for file_name in file_name_list:
        get_and_save_file(service, folder, file_name, mimeType)


df = pd.read_excel('./xlsx/画像情報.xlsx')

with st.expander('画像情報', expanded=False):
    st.write(df)


# 全選択オプションを追加
all_option = "全て選択"

## 項目選択
# 店舗名
op_shopname_list1 = ['選択なし', all_option] + df["店舗名"].unique().tolist()
op_shopname_list2 = st.sidebar.multiselect(\
    "店舗名", op_shopname_list1, default=['選択なし']
    )

if op_shopname_list2 == ['選択なし']:
    st.info('上から順に項目を選択してください')
    st.stop()
    
elif op_shopname_list2 == [all_option]:
    filtered_df = df

else:
    filtered_df = df[df["店舗名"].isin(op_shopname_list2)]

# 床の色
op_floorcolor_list1 = [all_option] + filtered_df["床の色"].unique().tolist()
op_floorcolor_list2 = st.sidebar.multiselect(\
    "床の色", op_floorcolor_list1, default=[all_option]
    )

if op_floorcolor_list2 == [all_option]:
    filtered_df = filtered_df
else:
    filtered_df = filtered_df[filtered_df["床の色"].isin(op_floorcolor_list2)]

# 壁の色
op_wallcolor_list1 = filtered_df["壁の色"].unique().tolist()
op_wallcolor_list1 = sorted(op_wallcolor_list1)
op_wallcolor_list1 = [all_option] + op_wallcolor_list1
op_wallcolor_list2 = st.sidebar.multiselect(\
    "壁の色", op_wallcolor_list1, default=[all_option]
    )

if op_wallcolor_list2 == [all_option]:
    filtered_df = filtered_df
else:
    filtered_df = filtered_df[filtered_df["壁の色"].isin(op_wallcolor_list2)]

# シリーズ
op_series_list1 = filtered_df["シリーズ"].unique().tolist()
op_series_list1 = sorted(op_series_list1)
op_series_list1 = [all_option] + op_series_list1
op_series_list2 = st.sidebar.multiselect(\
    "シリーズ", op_series_list1, default=[all_option]
    )

if op_series_list2 == [all_option]:
    filtered_df = filtered_df
else:
    filtered_df = filtered_df[filtered_df["シリーズ"].isin(op_series_list2)]

# 塗色
op_woodcolor1_list1 = filtered_df["塗色"].unique().tolist()
op_woodcolor1_list1 = sorted(op_woodcolor1_list1)
op_woodcolor1_list1 = [all_option] + op_woodcolor1_list1
op_woodcolor1_list2 = st.sidebar.multiselect(\
    "塗色", op_woodcolor1_list1, default=[all_option]
    )

if op_woodcolor1_list2 == [all_option]:
    filtered_df = filtered_df
else:
    filtered_df = filtered_df[filtered_df["塗色"].isin(op_woodcolor1_list2)]

# 張地
op_fabric_list1 = filtered_df["張地"].unique().tolist()
op_fabric_list1 = sorted(op_fabric_list1)
op_fabric_list1 = [all_option] + op_fabric_list1 
op_fabric_list2 = st.sidebar.multiselect(\
    "張地", op_fabric_list1, default=[all_option]
    )

if op_fabric_list2 == [all_option]:
    filtered_df = filtered_df
else:
    filtered_df = filtered_df[filtered_df["張地"].isin(op_fabric_list2)]

## driveからファイル取得dataに保存
#画像ファイル
file_name_list = filtered_df['ファイル名']
mimeType='image/jpeg'
get_file_from_gdrive(
    cwd=cwd, 
    folder='img', 
    file_name_list=file_name_list, 
    mimeType=mimeType
    )

# 画像情報.xlsx
file_name_list = ["画像情報.xlsx"]
mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
# mimeType='application/vnd.ms-excel'
get_file_from_gdrive(
    cwd=cwd, 
    folder='xlsx', 
    file_name_list=file_name_list, 
    mimeType=mimeType
    )

# カラム設定
col1, col2 = st.columns(2)

# 検索結果の画像を表示
for i in range(len(filtered_df)):
    image_path = os.path.join('img', filtered_df["ファイル名"].iloc[i])
    image = Image.open(image_path)

    # caption
    shop_name = filtered_df["店舗名"].iloc[i]
    series_name = filtered_df["シリーズ"].iloc[i]
    wood_color = filtered_df["塗色"].iloc[i]
    fabric_name = filtered_df["張地"].iloc[i]
    file_name = filtered_df["ファイル名"].iloc[i]
    text = f'{shop_name} {series_name} {wood_color} {fabric_name} {file_name}'

    # 偶数はcol1 奇数はcol2
    if i == 0:
        with col1:
            st.image(image,caption=text)
    elif i % 2 == 0:
        with col1:
            st.image(image,caption=text)
    else:
        with col2:
            st.image(image,caption=text)
    
    # 何枚でstopするか
    if i == 10:
        break