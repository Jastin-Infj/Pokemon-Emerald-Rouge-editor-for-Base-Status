import pandas as pd

def output_BaseRom(file_name,sheet_name,fetch_datas):
  print()
  print('main rom image loading ')
  print()

  export_datas = {
    "HP": [],
    "攻撃": [],
    "防御": [],
    "特攻": [],
    "特防": [],
    "素早さ": []
  }

  for row_datas in fetch_base_datas:
    base_h = int(row_datas[4],16)
    base_a = int(row_datas[5],16)
    base_b = int(row_datas[6],16)
    base_s = int(row_datas[7],16)
    base_c = int(row_datas[8],16)
    base_d = int(row_datas[9],16)

    export_datas["HP"].append(base_h)
    export_datas["攻撃"].append(base_a)
    export_datas["防御"].append(base_b)
    export_datas["特攻"].append(base_c)
    export_datas["特防"].append(base_d)
    export_datas["素早さ"].append(base_s)

  export_file = pd.DataFrame(export_datas)
  export_file.to_excel(file_name,sheet_name,index=True)

  print()
  print("fin export complated!")
  print()

# main

if __name__ == '__main__':

  # 読み込み場所
  start_address = 0x3A4310
  # 取り込む数

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

  # 読み込み内容
  params = {
    "Length": 0x20 + 0x04,
    "space": 0x04,
    "baseStatus": 0x06,
    "types": 0x02,
    "other1": 0x10,
    "ability1": 0x01,
    "ability2": 0x01,
    "other2": 0x02,
    "ability3": 0x01,
    "other3": 0x03
  }
  read_own_byte = params["Length"]

  # print([45,49,49,45,65,65])
  flag = True

  if flag == True:
    with open('PokemonEmeraldRouge','rb') as f:

      # 開始位置
      f.seek(start_address)

      data_all = f.read(params["Length"] * read_len)

      fetch_base_datas = []
      # 全体のデータを 36byte に分割
      for cur in range(0,len(data_all),params["Length"]):
        # 36byte
        base_param = data_all[cur: cur + params["Length"]]
        fetch_base_datas.append(base_param)

      # 全体のデータを 16進数
      count = 0
      for vals in fetch_base_datas:
        h_base = []

        for val in vals:
          i_val = int(val)
          h_val = hex(i_val)
          h_base.append(h_val)
        
        fetch_base_datas[count] = h_base
        count += 1

      # 読み終えたら閉じる
      f.close()
    
    # 内部データに保管されている種族値を出力する
    output_BaseRom('baseStatus.xlsx','BaseMaster',fetch_base_datas)