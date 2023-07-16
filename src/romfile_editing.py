import pandas as pd

import requests
import bs4

import json

file_name = 'baseStatus.xlsx'
rom_name = 'Pokemon Emerald Rogue EX (v1.3.2a).gba'
json_name = './src/pokemonForum.json'

# ROM 情報
START_ADDRESS = 0x3A4310
LEN_POKEMON_BASE = 0x24

write_option = {
  "MAX_Name_Pokemon": 898,
  "MAX_Name_Pokemon_F": 308
}

read_len = write_option["MAX_Name_Pokemon"] + write_option["MAX_Name_Pokemon_F"]

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

html_urls = {
  "Name": "https://wiki.xn--rckteqa2e.com/wiki/%E3%83%9D%E3%82%B1%E3%83%A2%E3%83%B3%E4%B8%80%E8%A6%A7"
}

# 種族値データ読み取り
def fetch_rom_baseStatus():
  with open(rom_name,'rb') as f:

    # 開始位置
    f.seek(START_ADDRESS)
    data_all = f.read(LEN_POKEMON_BASE * read_len)

    rom_fetch_datas = []
    # 全体のデータを 36byte に分割
    for i in range(0,len(data_all),LEN_POKEMON_BASE):
      # 36byte
      base_param = data_all[i: i + LEN_POKEMON_BASE]
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

# ポケモンの別フォルム名を取得
def fetch_PokemonName_Forum():
  f = open(json_name)
  data = json.load(f)
  f.close()
  
  return data

# 種族値データの適用
def attach_BaseStatus(base_status):
  for i in range(len(base_status)):
    row_Name["HP"].append(base_status[i][0])
    row_Name["攻撃"].append(base_status[i][1])
    row_Name["防御"].append(base_status[i][2])
    row_Name["特攻"].append(base_status[i][3])
    row_Name["特防"].append(base_status[i][4])
    row_Name["素早さ"].append(base_status[i][5])
  
  return

# ポケモン名の適用
def attach_PokemonName(pokemonList):
  for i in range(read_len):
    if i < write_option["MAX_Name_Pokemon"]:
      row_Name["Name"].append(pokemonList[i])
    else:
      row_Name["Name"].append(None)
  
  return

# ポケモン名の適用すべて
def attach_PokemonNameAll(pokemonList,forumdata):
  attach_PokemonName(pokemonList)

  diff = len(row_Name["Name"]) - write_option["MAX_Name_Pokemon"]

  # 登録数が足らない
  if len(forumdata) < diff:
    for i in range(len(forumdata)):
      row_Name["Name"][write_option["MAX_Name_Pokemon"] + i] = forumdata[i]
    return
  elif len(forumdata) > diff:
    print("Error __PokemonNameAll__")
    return
  else:
    # 登録数 ぴったり
    for i in range(diff):
      row_Name["Name"][write_option["MAX_Name_Pokemon"] + i] = forumdata[i]
  
  return

# HTML ポケモン名を取得
def fetch_html_pokemonData():
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

# ROM 書き出す パラメータ
def edit_rom_param(rom_data,excel_data):

  # 開始位置

  # 配列にすることで 編集を可能にする
  rom_data = list(bytes(rom_data))

  # 編集項目
  for i in range(read_len):
    pokemonData = excel_data[i]

    start_pointer = START_ADDRESS + i * LEN_POKEMON_BASE
    for row in range(LEN_POKEMON_BASE):
      # HP
      if row == 0x04:
        rom_data[start_pointer + row] = int(pokemonData[1])
      # 攻撃
      elif row == 0x05:
        rom_data[start_pointer + row] = int(pokemonData[2])
      # 防御
      elif row == 0x06:
        rom_data[start_pointer + row] = int(pokemonData[3])
      # 素早さ
      elif row == 0x07:
        rom_data[start_pointer + row] = int(pokemonData[6])
      # 特攻
      elif row == 0x08:
        rom_data[start_pointer + row] = int(pokemonData[4])
      # 特防
      elif row == 0x09:
        rom_data[start_pointer + row] = int(pokemonData[5])

  # bytes に戻す
  rom_data = bytes(rom_data)

  return rom_data

# Excel 読み込み
def read_excel(sheet_Name):
  excel_file = pd.ExcelFile(file_name)
  df_master = pd.read_excel(excel_file,sheet_Name,index_col=0)

  values = df_master.values
  return values

# Excel 書き出し
def write_excel():
  # 編集用 保存
  export_file = pd.DataFrame(row_Name)
  export_file.to_excel(file_name,index=True,sheet_name=sheet_Name["Master"])

# ROM 書き出し
def write_rom(excel_data):

  rom_data = None
  with open(rom_name,'rb') as f:

    # 全体
    rom_data = f.read()
    print(hex(len(rom_data)))

    rom_data = edit_rom_param(rom_data,excel_data)

    f.close()
    pass
  
  with open(rom_name,'wb') as f:
    f.write(rom_data)
    f.close()
    pass

  return

if __name__ == '__main__':

  # 初期書き出し
  def write_first():
    # ROM read 種族値データ
    base_status = fetch_rom_baseStatus()

    attach_BaseStatus(base_status)

    # HTML read ポケモン名
    pokemondata = fetch_html_pokemonData()

    pokemon_names = pokemondata["Name"]
    pokemonlist_names_f = fetch_PokemonName_Forum()

    attach_PokemonNameAll(pokemon_names,pokemonlist_names_f)

    # 書き出し
    write_excel()

    return
  
  # ROM書き出し
  def writing_rom():
    # 読み込み
    excel_data = read_excel(sheet_Name["Master"])
    write_rom(excel_data)
    return
  

  # 読み込み

  # write_first()

  # 書き出し
  writing_rom()
  print('fin rom edited!')
  print()