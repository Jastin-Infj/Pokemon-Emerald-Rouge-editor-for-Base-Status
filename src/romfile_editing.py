import pandas as pd

import requests
import bs4

import json , re

file_name = 'baseStatus.xlsx'
rom_name = 'Pokemon Emerald Rogue EX (v1.3.2a).gba'
filepath_pokemon_f_list = './src/pokemonForum.json'
filePath_pokemon_f_indexs = './src/pokemon_hpCopylist.jsonc'
filePath_pokemon_f_paramCopy = './src/pokemon_paramCopy.jsonc'

# ROM 情報
START_ADDRESS = 0x3A4310
LEN_POKEMON_BASE = 0x24

# 書き出しオプション
write_option = {
  "MAX_Name_Pokemon": 898,
  "MAX_Name_Pokemon_F": 308
}

WRITE_OPTION_Dependence = []
WRITE_OPTION_ParamCopy = []

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

# jsonc 形式ファイルの読み取り
def fileRead_jsonc(filepath):
  datas = None
  with open(filepath,'r') as f:
    datas = f.read()
    f.close()
  
  format_data = re.sub(r'/\*[\s\S]*?\*/|//.*','',datas)
  json_obj = json.loads(format_data)
  return json_obj

# 依存先のリストを作成
def create_pokemon_f_HPCopyList(meIndex,copyIndex):
  copystyle = {
    'Me': meIndex,
    'Copy': copyIndex
  }
  return copystyle

# 依存関係リスト作成
def append_writeOption_forum_dependence(forum_data,start,length):
  start_pointer = write_option["MAX_Name_Pokemon"]

  for i in range(length):
    l_index = start + i
    c_index = start_pointer + l_index
    WRITE_OPTION_Dependence.append(create_pokemon_f_HPCopyList(c_index,forum_data[str(l_index)]))
    pass

  return

# 書き出しオプションの初期化
def init_write_option_dependence():

  # json から読み取り
  copy_indexs = fileRead_jsonc(filePath_pokemon_f_indexs)

  # HP 依存関係 あり
  start = None

  # メガシンカ or ゲンシカイキ
  start = 0
  append_writeOption_forum_dependence(copy_indexs,start,50)

  # ポワルン
  start = 129
  append_writeOption_forum_dependence(copy_indexs,start,3)

  # チェリム
  start = 139
  append_writeOption_forum_dependence(copy_indexs,start,1)

  # ギラティナ オリジン
  start = 147
  append_writeOption_forum_dependence(copy_indexs,start,1)

  # ヒヒダルマ
  start = 167
  append_writeOption_forum_dependence(copy_indexs,start,2)
  
  # メロエッタ
  start = 181
  append_writeOption_forum_dependence(copy_indexs,start,1)

  # ゲノセクト
  start = 182
  append_writeOption_forum_dependence(copy_indexs,start,4)
  
  # ゲッコウガ
  start = 186
  append_writeOption_forum_dependence(copy_indexs,start,2)
  
  # ギルガルド
  start = 230
  append_writeOption_forum_dependence(copy_indexs,start,1)

  # ゼルネアス
  start = 237
  append_writeOption_forum_dependence(copy_indexs,start,1)

  # ヨワシ
  start = 249
  append_writeOption_forum_dependence(copy_indexs,start,1)

  # メテノ
  start = 267
  append_writeOption_forum_dependence(copy_indexs,start,13)
  
  # ミミッキュ
  start = 280
  append_writeOption_forum_dependence(copy_indexs,start,1)

  # ネクロズマ
  start = 283
  append_writeOption_forum_dependence(copy_indexs,start,1)

  # マギアナ ?
  start = 284
  append_writeOption_forum_dependence(copy_indexs,start,1)

  # ウッウ
  start = 285
  append_writeOption_forum_dependence(copy_indexs,start,2)
  
  # コオリッポ
  start = 298
  append_writeOption_forum_dependence(copy_indexs,start,1)

  # モルペコ
  start = 300
  append_writeOption_forum_dependence(copy_indexs,start,1)

  # ザシアン
  start = 301
  append_writeOption_forum_dependence(copy_indexs,start,1)

  # ザマゼンタ
  start = 302
  append_writeOption_forum_dependence(copy_indexs,start,1)

  # ブリザポス
  start = 306
  append_writeOption_forum_dependence(copy_indexs,start,2)
  
  return

# 依存関係 パラメータコピー オブジェクト作成
def create_pokemon_f_paramCopy(meIndex,copyIndex,param):
  return {
    "F_ID": meIndex,
    "C_ID": copyIndex,
    "Copy": param["Copy"],
    "Params": param
  }

# 依存関係 パラメータコピー フラグ設定
def append_writeOption_forum_paramCopy(start,length,param):
  start_pointer = write_option["MAX_Name_Pokemon"]

  for i in range(length):
    l_index = start + i
    c_index = start_pointer + l_index
    WRITE_OPTION_ParamCopy.append(create_pokemon_f_paramCopy(l_index,c_index,param))
  return

# 依存関係 パラメータコピーフラグ
def init_write_option_paramCopy():
  # json から読み取り
  paramCopy = fileRead_jsonc(filePath_pokemon_f_paramCopy)

  # 依存関係
  start = None

  # ピカチュウ
  start = 87
  append_writeOption_forum_paramCopy(start,14,paramCopy[str(start)])

  # ピチュー
  start = 101
  append_writeOption_forum_paramCopy(start,1,paramCopy[str(start)])

  # アンノーン
  start = 102
  append_writeOption_forum_paramCopy(start,27,paramCopy[str(start)])

  # ポワルン
  start = 129
  append_writeOption_forum_paramCopy(start,3,paramCopy[str(start)])

  # チェリム
  start = 139
  append_writeOption_forum_paramCopy(start,1,paramCopy[str(start)])

  # カラナクシ
  start = 140
  append_writeOption_forum_paramCopy(start,1,paramCopy[str(start)])

  # トリトドン
  start = 141
  append_writeOption_forum_paramCopy(start,1,paramCopy[str(start)])

  # ロトム
  start = 142
  append_writeOption_forum_paramCopy(start,5,paramCopy[str(start)])

  # アルセウス
  start = 149
  append_writeOption_forum_paramCopy(start,17,paramCopy[str(start)])

  # シキジカ
  start = 169
  append_writeOption_forum_paramCopy(start,3,paramCopy[str(start)])

  # メブキジカ
  start = 172
  append_writeOption_forum_paramCopy(start,3,paramCopy[str(start)])

  # メロエッタ
  start = 181
  append_writeOption_forum_paramCopy(start,1,paramCopy[str(start)])

  # ゲノセクト
  start = 182
  append_writeOption_forum_paramCopy(start,4,paramCopy[str(start)])

  # ゲッコウガ
  start = 186
  append_writeOption_forum_paramCopy(start,2,paramCopy[str(start)])

  # ビビヨン
  start = 188
  append_writeOption_forum_paramCopy(start,19,paramCopy[str(start)])

  # フラベベ
  start = 207
  append_writeOption_forum_paramCopy(start,4,paramCopy[str(start)])

  # フラエッテ
  start = 211
  append_writeOption_forum_paramCopy(start,4,paramCopy[str(start)])

  # フラージェス
  start = 216
  append_writeOption_forum_paramCopy(start,4,paramCopy[str(start)])

  # トリミアン
  start = 220
  append_writeOption_forum_paramCopy(start,9,paramCopy[str(start)])

  # ゼルネアス
  start = 237
  append_writeOption_forum_paramCopy(start,1,paramCopy[str(start)])

  # ジガルデ
  start = 239
  append_writeOption_forum_paramCopy(start,2,paramCopy[str(start)])

  # オドリドリ
  start = 243
  append_writeOption_forum_paramCopy(start,3,paramCopy[str(start)])

  # イワンコ
  start = 246
  append_writeOption_forum_paramCopy(start,1,paramCopy[str(start)])

  # シルヴァディ
  start = 250
  append_writeOption_forum_paramCopy(start,17,paramCopy[str(start)])

  # メテノ
  start = 267
  append_writeOption_forum_paramCopy(start,13,paramCopy[str(start)])

  # ミミッキュ
  start = 280
  append_writeOption_forum_paramCopy(start,1,paramCopy[str(start)])

  # ネクロズマ
  start = 283
  append_writeOption_forum_paramCopy(start,1,paramCopy[str(start)])

  # マギアナ
  start = 284
  append_writeOption_forum_paramCopy(start,1,paramCopy[str(start)])

  # ウッウ
  start = 285
  append_writeOption_forum_paramCopy(start,2,paramCopy[str(start)])

  # ストリンダー
  start = 287
  append_writeOption_forum_paramCopy(start,1,paramCopy[str(start)])

  # ヤバチャ
  start = 288
  append_writeOption_forum_paramCopy(start,1,paramCopy[str(start)])

  # ポットデス
  start = 289
  append_writeOption_forum_paramCopy(start,1,paramCopy[str(start)])

  # マホイップ
  start = 290
  append_writeOption_forum_paramCopy(start,8,paramCopy[str(start)])

  # コオリッポ
  start = 298
  append_writeOption_forum_paramCopy(start,1,paramCopy[str(start)])

  # モルペコ
  start = 300
  append_writeOption_forum_paramCopy(start,1,paramCopy[str(start)])

  return

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
  f = open(filepath_pokemon_f_list)
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

# 依存関係 HP チェック
def filter_Dependence_HP_Forum(dex_id):
  match_forums = []
  for val in WRITE_OPTION_Dependence:
    isMatch = ('Copy',dex_id) in val.items()
    if isMatch == True:
      match_forums.append(val)
  
  return match_forums

# 依存関係 コピーフラグチェック
def filter_ParamCopy_Forum(dex_id):
  match_forums = []
  for val in WRITE_OPTION_ParamCopy:
    isMatch = ('Copy',dex_id) in val.items()
    if isMatch == True:
      match_forums.append(val)

  return match_forums

# ROM 書き出す パラメータ
def edit_rom_param(rom_data,excel_data):

  # 開始位置

  # 配列にすることで 編集を可能にする
  rom_data = list(bytes(rom_data))

  # 編集項目
  for i in range(read_len):
    pokemonData = excel_data[i]

    # 基準ポインタ
    start_pointer = START_ADDRESS + i * LEN_POKEMON_BASE
    # 依存先ポインタ
    forum_pointer = None

    for row in range(LEN_POKEMON_BASE):
      # HP
      if row == 0x04:
        rom_data[start_pointer + row] = int(pokemonData[1])
        match_forum = filter_Dependence_HP_Forum(i)

        # 依存しているポケモンあり
        if len(match_forum) > 0:
          for forum in match_forum:
            # コピーしたいポケモン
            me_index = forum["Me"]
            forum_pointer = START_ADDRESS + me_index * LEN_POKEMON_BASE
            # ROM 情報
            rom_data[forum_pointer + row] = int(pokemonData[1])
            # Excel 情報を上書き
            excel_data[me_index][1] = int(pokemonData[1])

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
      pass
    
    # 依存先パラメータコピーポインタ
    paramcopy_pointer = None
    match_param = filter_ParamCopy_Forum(i)

    # 依存関係あり
    if len(match_param) > 0:
      for param in match_param:
        # コピーフラグがある要素をコピー
        m_index = param["C_ID"]

        # HP
        if param["Params"]["H"] == True:
          excel_data[m_index][1] = int(pokemonData[1])
        # 攻撃
        if param["Params"]["A"] == True:
          excel_data[m_index][2] = int(pokemonData[2])
        # 防御
        if param["Params"]["B"] == True:
          excel_data[m_index][3] = int(pokemonData[3])
        # 特攻
        if param["Params"]["C"] == True:
          excel_data[m_index][4] = int(pokemonData[4])
        # 特防
        if param["Params"]["D"] == True:
          excel_data[m_index][5] = int(pokemonData[5])
        # 素早さ
        if param["Params"]["S"] == True:
          excel_data[m_index][6] = int(pokemonData[6])
        pass
    
    pass

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
  
  # 初期化処理
  init_write_option_dependence()
  init_write_option_paramCopy()

  # 初期読み込み

  write_first()

  # 書き出し

  # writing_rom()
  print('fin rom edited!')
  print()