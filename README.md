# MasterConverter

ExcelとYAML/JSONファイルを双方向に変換するマスターデータ管理ツールです。

**Excelで編集 → テキストファイルでGit管理** という運用を実現します。

```
Export : Excel (.xlsx)         →  レコードファイル (.record)
Import : レコードファイル (.record)  →  Excel (.xlsx)
```

---

## 目次

1. [動作要件](#動作要件)
2. [ビルド](#ビルド)
3. [クイックスタート](#クイックスタート)
4. [ファイル構成](#ファイル構成)
5. [ClassSchema.xlsx の作り方](#classschemaxlsx-の作り方)
6. [コマンドライン引数](#コマンドライン引数)
7. [settings.ini](#settingsini)
8. [出力ファイルの形式](#出力ファイルの形式)
9. [ワークフロー](#ワークフロー)
10. [トラブルシューティング](#トラブルシューティング)
11. [依存ライブラリ](#依存ライブラリ)

---

## 動作要件

- .NET 6.0 以上
- Windows / macOS / Linux

---

## ビルド

```bash
dotnet build
```

ビルド成果物は `bin/Debug/net6.0/` または `bin/Release/net6.0/` に出力されます。

### Linux / macOS (Docker)

```bash
docker run -it --rm -v `pwd`:/src --workdir=/src mono:5.20.1.19 bash
/src# mono /usr/lib/mono/msbuild/15.0/bin/MSBuild.dll

# または（ワンライナー）
docker run --rm -v `pwd`:/src --workdir=/src mono:5.20.1.19 mono /usr/lib/mono/msbuild/15.0/bin/MSBuild.dll
```

---

## クイックスタート

### 1. マスターフォルダを用意する

フォルダ名がそのままマスター名になります。

```
Enemy/
└── ClassSchema.xlsx   ← 自分で作成する
```

### 2. ClassSchema.xlsx を作成する

Excelで以下の形式で作成します（詳細は [ClassSchema.xlsx の作り方](#classschemaxlsx-の作り方) を参照）。

| 行 | A | B | C | D | E |
|----|---|---|---|---|---|
| 1 | ID | 名前 | レベル | 攻撃力 | タグ |
| **2（型）** | `int` | `string` | `int` | `float?` | `string[]` |
| **3（フィールド名）** | `id` | `name` | `level` | `attack` | `tags` |

シート名は **`Master`** にしてください。

### 3. Import を実行して編集用 Excel を生成する

```bash
MasterConverter --input ./Enemy --mode import --exit
```

`Enemy/Enemy.xlsx` が生成されます。

### 4. Enemy.xlsx にデータを入力して Export する

```bash
MasterConverter --input ./Enemy --mode export --exit
```

`Enemy/Records/` にレコードファイルが出力されます。

---

## ファイル構成

```
MasterName/
├── ClassSchema.xlsx          # クラス定義ファイル（自分で作成・管理）
├── MasterName.xlsx           # 編集用 Excel（Import 時に自動生成 / 上書き）
├── MasterName.index          # レコード順序の記録（自動生成）
└── Records/                  # レコードファイル群（Export 時に自動生成 / 上書き）
    ├── RecordName1.record    # レコードデータ（YAML または JSON）
    ├── RecordName1.option    # セルの書式情報（色・コメント等）
    ├── RecordName2.record
    └── ...
```

| ファイル | 説明 | Git 管理 |
|---------|------|----------|
| `ClassSchema.xlsx` | フィールド定義。自分で管理する。 | ✅ 管理する |
| `MasterName.xlsx` | 自動生成される編集用 Excel。 | ❌ 管理しない（.gitignore 推奨）|
| `MasterName.index` | レコードの並び順を保存する。 | ✅ 管理する |
| `Records/*.record` | レコード 1 件ごとのデータファイル。 | ✅ 管理する |
| `Records/*.option` | セルの色・コメントを保存する。 | ✅ 管理する |

---

## ClassSchema.xlsx の作り方

### シート名

**`Master`** という名前のシートを使用します。

### 行構成

| 行番号 | 内容 | 必須 |
|--------|------|------|
| 1 | 任意（日本語ラベルや説明など） | — |
| **2** | **データ型**（設定で変更可） | ✅ |
| **3** | **フィールド名**（設定で変更可） | ✅ |
| 4 以降 | テンプレート行。書式・列幅の基準として使用される。 | — |

> 行番号は `settings.ini` で変更できます。

### フィールド名のルール

- フィールド名の先頭に `#` を付けると **除外フィールド** になります。
- 除外フィールドは `string` 型として扱われ、`.record` ファイルには出力されません。
- Excel 上にメモ列や管理用の列を設けたいときに使います。

```
#memo  ← このフィールドはExcelには表示されるが出力されない
```

### 対応する型

| 分類 | 記述例 | 説明 |
|------|--------|------|
| 基本型 | `int` `long` `float` `double` `bool` `string` `DateTime` | System 名前空間の型全般が使用可 |
| 配列型 | `int[]` `string[]` `float[]` | 任意の型に `[]` を付加 |
| Null 許容型 | `int?` `float?` `DateTime?` | 任意の型に `?` を付加。空セルは `null` として出力 |

### 配列の入力形式

Excel のセルには以下の形式で入力します。

```
[値1,値2,値3]
```

例：`[fire,ice,thunder]`、`[10,20,30]`

### 入力例

| 行 | A | B | C | D | E |
|----|---|---|---|---|---|
| 1 | ID | 名前 | レベル | 攻撃力 | タグ |
| 2 | `int` | `string` | `int` | `float?` | `string[]` |
| 3 | `id` | `name` | `level` | `attack` | `tags` |
| 4 | `1` | `スライム` | `1` | `10.5` | `[weak,slime]` |

> 4 行目はテンプレート行です。実際のレコードデータは `Export` 時に書き込まれます。

---

## コマンドライン引数

```
MasterConverter --input <path> --mode <import|export> [--exit]
```

| 引数 | 必須 | 説明 |
|------|------|------|
| `--input <path>` | ✅ | 変換対象のディレクトリ。`,` 区切りで複数指定可 |
| `--mode <import\|export>` | ✅ | 動作モード |
| `--exit` | — | 処理完了後に自動でウィンドウを閉じる |

### `--input` の指定方法

指定したディレクトリ以下を再帰的に検索し、`ClassSchema.xlsx` が存在するフォルダを自動検出して一括処理します。

```bash
# 単一ディレクトリ
MasterConverter --input ./Masters --mode export --exit

# 複数ディレクトリ（, 区切り）
MasterConverter --input ./Masters/Enemy,./Masters/Item --mode export --exit
```

---

## settings.ini

実行ファイルと同じディレクトリに配置します。

```ini
[Rows]
dataTypeRow    = 2   ; ClassSchema.xlsx のデータ型行番号
fieldNameRow   = 3   ; ClassSchema.xlsx のフィールド名行番号
recordStartRow = 4   ; Excel のレコード開始行番号

[File]
format = yaml        ; 出力フォーマット: yaml または json
```

---

## 出力ファイルの形式

### `.record` ファイル（YAML の場合）

レコード 1 件につき 1 ファイル出力されます。ファイル名はレコードのキーフィールド値から自動生成されます。

```yaml
# 1_スライム.record
attack: 10.5
id: 1
level: 1
name: スライム
tags:
- weak
- slime
```

フィールドはアルファベット順にソートされます。

### `.option` ファイル

セルに色やコメントが設定されている場合のみ出力されます。

```yaml
- backgroundColor:
    rgb: FFFFFF00
  column: 3
  patternType: Solid
```

### `.index` ファイル

Excel 上のレコード順序を保存します。Import 時にこの順序を使って Excel が再構築されます。

---

## ワークフロー

### 初回セットアップ

```
1. ClassSchema.xlsx を作成する
2. import を実行 → MasterName.xlsx が生成される
3. MasterName.xlsx を開いてデータを入力する
4. export を実行 → Records/ にファイルが出力される
5. ClassSchema.xlsx と Records/ を Git にコミットする
```

### 通常の編集フロー

```
1. MasterName.xlsx を開く（なければ import で再生成）
2. データを編集する
3. MasterName.xlsx を閉じる
4. export を実行
5. 変更された .record ファイルを Git にコミットする
```

### 他のメンバーの変更を取り込む

```
1. git pull
2. import を実行 → MasterName.xlsx に最新の Records/ が反映される
```

### 更新スキップの仕組み

最終更新日時を比較して、変更がない場合は処理をスキップします。

- **Import**: `.xlsx` より新しい `.record` / `.option` / `.index` が存在する場合のみ実行
- **Export**: `.xlsx` が `.index` より新しい場合のみ実行

---

## トラブルシューティング

### `File locked!!`

```
FileLoadException: File locked!!
/path/to/MasterName.xlsx
```

**原因**: Export / Import 対象の `.xlsx` を Excel で開いたまま実行した。
**対処**: Excel を閉じてから再実行する。

---

### `Duplication records exist!`

```
Duplication records exist!
File : /path/to/MasterName.xlsx
  [Row 5, 12] id=1, name=スライム
```

**原因**: 同一内容のレコードが複数行存在する。
**対処**: エラーに表示された行番号の Excel 行を確認・修正する。

---

### `Directory not found`

```
DirectoryNotFoundException: Directory not found. /path/to/Master
```

**原因**: `--input` に指定したパスが存在しない。
**対処**: パスを確認する。

---

### ClassSchema.xlsx のフィールドが Excel に反映されない

**原因**: ClassSchema.xlsx のシート名が `Master` になっていない。
**対処**: シート名を `Master` に変更する。

---

### 数式セルに値が入力されない

仕様です。数式が設定されているセルへはデータを書き込まず、数式の結果をそのまま表示します。

---

## 依存ライブラリ

| ライブラリ | バージョン | 用途 |
|-----------|-----------|------|
| EPPlus | 6.2.7 | Excel ファイルの読み書き |
| YamlDotNet | 13.1.1 | YAML シリアライズ / デシリアライズ |
| Newtonsoft.Json | 13.0.3 | JSON シリアライズ / デシリアライズ |
| CommandLineParser | 2.9.1 | コマンドライン引数のパース |
| ini-parser-new | 2.6.2 | settings.ini の読み込み |
| System.CodeDom | 7.0.0 | 型名から型情報の解決 |
