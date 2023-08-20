# Pokemon Emerald Rouge BaseStatus Editor

Pokemon Emerald Rouge の 内部データを変更することができます

![image](https://github.com/Jastin-Infj/PokemonEditBaseStatus/assets/34806778/948b9766-9299-4fe3-88af-d1a6481d111d)


**必要なもの:**

- **ROMデータ Pokemon Emerald Rouge v1.3.2a の ROM 形式**

**※ROM をネットからダウンロードするのは日本国内において違法になっております。**

**ROMデータを編集行うため、事前にバックアップを推奨します**

# Usege

## Command 1 はじめて利用するとき

ROM情報から Excel データを作成します。

Excelのシート役割

- Edit

ROMデータにExcelの値を書き出しする

- EditCalc

Editシートに対しての計算式を自由に編集する

※ セルの値渡しが必要です

値コピーやセル参照を行ってください

- BaseMaster ※データの編集をしないでください

初回でROMデータから抽出されたデータ

Excelデータには以下の情報が抽出される

- インデックス
- ポケモンの種族値
    - HP
    - 攻撃
    - 防御
    - 特攻
    - 特防
    - 素早さ

romfile_editing.py

```bash
COMMAND = 0
```

## Command 2 初回以降のExcel ファイル抽出

上記と基本的なデータは変化しない

- Edit

ROM情報から常に値を更新する

- EditCalc

個人で編集した値やExcelの計算式を維持したまま抽出される

- BaseMaster

初回時に記録した情報を維持

※ただし、誤って情報を編集したExcelデータの場合は 誤りのデータを渡す

```bash
COMMAND = 2
```

## Command 3 ROMデータに値を書き出し

Edit シートを基にROMへ値を書き出しします

```bash
COMMAND = 3
```
