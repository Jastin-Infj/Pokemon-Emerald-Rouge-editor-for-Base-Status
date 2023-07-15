import pandas as pd

file_name = 'sample.xlsx'
sheet_name = 'Master'

row1_name = 'product_name'
row2_name = 'price'

def write_sample():
  data = {
    row1_name: ['computer','printer','tablet','monitor'],
    row2_name: [1200,150,300,450]
  }

  # データ作成
  df = pd.DataFrame(data)

  # print(df)

  df.to_excel('sample.xlsx',index=True)



def write_sample2():
  read_file = pd.read_excel(file_name,index_col=0,sheet_name=sheet_name)

  # [][]
  # print(read_file.values)
  datas = read_file.values

  edit_data = {
    row1_name: [],
    row2_name: []
  }

  for val in datas:
    name = val[0]
    price = val[1]

    edit_data[row1_name].append(name)
    edit_data[row2_name].append(price)
  
  def TwoPrice(num):
    return num * 2
  
  edit_data[row2_name] = list(map(TwoPrice,edit_data[row2_name]))

  export_file = pd.DataFrame(edit_data)

  export_file.to_excel(file_name,index=True,sheet_name=sheet_name)


def read_sample():
  test = pd.read_excel('sample.xlsx')
  
  # 確認
  print(test)

  # 要素取り出し
  for val in test.values:
    name = val[0]
    price = val[1]
    print(price)

if __name__ == '__main__':
  write_sample2()