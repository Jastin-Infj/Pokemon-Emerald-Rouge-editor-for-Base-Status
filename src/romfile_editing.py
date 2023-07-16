import pandas as pd

import requests
import bs4

file_name = 'baseStatus.xlsx'
rom_name = 'Pokemon Emerald Rogue EX (v1.3.2a).gba'

# 通常 + メガシンカ + ゲンシ + リージョン(アローラ) + ガラル + おきがえピカチュウ + サトピカ + ギザみみピチュー
select_row1 = 898 + 48 + 2 + 18 + 19 + 5 + 8 + 1 + 1
# アンノーン + ポワルン + デオキシス + [すな[ミノムッチ,ミノマダム],ゴミ[ミノムッチ,ミノマダム]] + [カラナクシ,トリトドン] + ロトム + ギラティナ + シェイミ + アルセウス
select_row2 = 28 + 3 + 3 + 4 + 2 + 5 + 1 + 1 + 17
# バスラオ(白なし) + ヒヒダルマ ダルマモード[原種,ガラル] + シキジカ + メブキジカ + 霊獣 + キュレム + ケルディオ + メロエッタ + ゲノセクト
select_row3 = 1 + 2 + 3 + 3 + 3 + 2 + 1 + 1 + 4
# ゲッコウガ[通常,きずなへんげ] + ビビヨン + フラベベ + フラエッテ[,AZ] + フラージェス + トリミアン + ニャオニクス + ギルガルド + バケッチャ + パンプジン + ゼルネアス + ジガルデ[10%,50%,50%,100%] + フーパ
select_row4 = 2 + 19 + 4 + 5 + 4 + 9 + 1 + 1 + 3 + 3 + 1 + 4 + 1
# オドリドリ + イワンコ + ルガルガン[まよなか,たそがれ] + ヨワシ + シルヴァディ + メテノ[流星] + メテノ[コア] + ミミッキュ + ネクロズマ[日食,月食,ウルトラ] + マギアナ + 
select_row5 = 3 + 1 + 2 + 1 + 17 + 6 + 7 + 1 + 3 + 1
# ウッウ + ストリンダー + [ヤバチャ,ポットデス] + マホイップ + コオリッポ + イエッサン + モルペコ + [ザシアン,ザマゼンタ,ムゲンダイナ] + ウーラオス + ザルード + バドレックス[はくばじょう,こくばじょう]
select_row6 = 2 + 1 + 2 + 8 + 1 + 1 + 1 + 3 + 1 + 1 + 2

read_len = select_row1 + select_row2 + select_row3 + select_row4 + select_row5 + select_row6
write_len = read_len

sheet_Name = {
  "Master": 'BaseMaster',
  "Edit": 'Edit'
}

row_Name = {
  "Name": [],
  "HP": [],
  "攻撃": [],
  "防御": [],
  "特攻": [],
  "特防": [],
  "素早さ": []
}

write_option = {
  "MAX_Name_Pokemon": 898
}

html_urls = {
  "Name": "https://wiki.xn--rckteqa2e.com/wiki/%E3%83%9D%E3%82%B1%E3%83%A2%E3%83%B3%E4%B8%80%E8%A6%A7"
}

# 種族値データ読み取り
def rom_fetch_baseStatus():
  with open(rom_name,'rb') as f:

    # 開始位置
    start_address = 0x3A4310
    # Base Byte
    len_pokemonBase = 0x24

    f.seek(start_address)
    data_all = f.read(len_pokemonBase * read_len)

    rom_fetch_datas = []
    # 全体のデータを 36byte に分割
    for i in range(0,len(data_all),len_pokemonBase):
      # 36byte
      base_param = data_all[i: i + len_pokemonBase]
      rom_fetch_datas.append(base_param)
    
    # 種族値データ
    export_datas = []
    for vals in rom_fetch_datas:
      ex_data = []
      start_pointer = 0x04
      i_h = int(vals[start_pointer + 0])
      i_a = int(vals[start_pointer + 1])
      i_b = int(vals[start_pointer + 2])
      i_s = int(vals[start_pointer + 3])
      i_c = int(vals[start_pointer + 4])
      i_d = int(vals[start_pointer + 5])

      ex_data.append(i_h)
      ex_data.append(i_a)
      ex_data.append(i_b)
      ex_data.append(i_c)
      ex_data.append(i_d)
      ex_data.append(i_s)

      export_datas.append(ex_data)

    f.close()

    return export_datas

# 種族値データの適用
def attach_BaseStatus(base_status,length):
  for i in range(length):
    row_Name["HP"].append(base_status[i][0])
    row_Name["攻撃"].append(base_status[i][1])
    row_Name["防御"].append(base_status[i][2])
    row_Name["特攻"].append(base_status[i][3])
    row_Name["特防"].append(base_status[i][4])
    row_Name["素早さ"].append(base_status[i][5])

# ポケモン名の適用
def attach_PokemonName(pokemonList,length):
  for i in range(length):
    if i < write_option["MAX_Name_Pokemon"]:
      row_Name["Name"].append(pokemonList[i])
    else:
      row_Name["Name"].append(None)

# HTML ポケモン名を取得
def html_fetch_pokemonList():
  req = requests.get(html_urls["Name"])
  html_text = bs4.BeautifulSoup(req.text,'html.parser')
  
  select = "#mw-content-text > div.mw-parser-output > table > tbody > tr"

  divs = {
    "index": html_text.select(select + " " + "td:nth-of-type(1)"),
    "Name": html_text.select(select + " " + "td:nth-of-type(2) > a"),
    "Type1": html_text.select(select + " " + "td:nth-of-type(3) > a"),
    "Type2": html_text.select(select + " " + "td:nth-of-type(4) > a")
  }

  export_data = {
    "Name": [],
    "Type1": [],
    "Type2": []
  }

  # 名前
  for name in divs["Name"]:
    text = name.get_text()
    export_data["Name"].append(text)

  # 重複している要素は削除
  export_data["Name"] = list(dict.fromkeys(export_data["Name"]))
  
  # タイプ1
  for type1 in divs["Type1"]:
    text = type1.get_text()
    export_data["Type1"].append(text)

  # タイプ2
  for type2 in divs["Type2"]:
    text = type2.get_text()
    export_data["Type2"].append(text)

  return export_data

def read_excel():
  print()
  print('read')

  excel_file = pd.ExcelFile(file_name)

  df_master = pd.read_excel(excel_file,sheet_name=sheet_Name["Master"],index_col=0)
  df_edit = pd.read_excel(excel_file,sheet_name=sheet_Name["Edit"],index_col=0)

  value_master = df_master.values
  value_edit = df_edit.values

  return {
    sheet_Name["Master"]: value_master,
    sheet_Name["Edit"]: value_edit
  }

def write_excel():

  # 編集用 保存
  export_file = pd.DataFrame(row_Name)
  export_file.to_excel(file_name,index=True,sheet_name=sheet_Name["Master"])

if __name__ == '__main__':

  # ROM read 種族値データ
  base_status = rom_fetch_baseStatus()
  attach_BaseStatus(base_status,write_len)

  # HTML read ポケモン名
  pokemonlist = html_fetch_pokemonList()
  attach_PokemonName(pokemonlist["Name"],write_len)

  # 書き出し
  write_excel()

  # Excel_data = read_excel()
  # Pokemon_Context = Pokemon_Fetch_List()
  # write_excel(Excel_data,Pokemon_Context)

  print('fin rom edited!')
  print()