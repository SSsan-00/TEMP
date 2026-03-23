# add_dashboard.ps1 使い方

## 概要

`add_dashboard.ps1` は、指定した Excel ブックに `進捗ダッシュボード` シートを追加または再生成するスクリプトです。

- 参照元シートのデータをもとにダッシュボードを作成します
- 実行前に元のブックを自動でバックアップします
- 既存の `進捗ダッシュボード` があれば上書き更新します

## 実行前の注意

- 対象の Excel ブックは閉じてから実行してください
- 別プロセスで開かれていると更新できません
- 実行すると同じフォルダに `*.backup_YYYYMMDD_HHMMSS.xlsx` が作成されます

## 基本実行

`plan.xlsx` を対象にし、ダッシュボード以外の先頭シートを参照元にする場合:

```powershell
cd C:\work\Macro\planAnalyzer
powershell -ExecutionPolicy Bypass -File .\add_dashboard.ps1
```

## 対象ブックを指定して実行

```powershell
cd C:\work\Macro\planAnalyzer
powershell -ExecutionPolicy Bypass -File .\add_dashboard.ps1 -WorkbookPath .\plan.demo.xlsx
```

## 参照元シートを指定して実行

参照元シート名が分かっている場合は `-SourceSheetName` を付けて実行します。

```powershell
cd C:\work\Macro\planAnalyzer
powershell -ExecutionPolicy Bypass -File .\add_dashboard.ps1 -WorkbookPath .\plan.xlsx -SourceSheetName '本番データ'
```

## ソースコードで既定の参照元シートを固定する

`add_dashboard.ps1` 冒頭の以下を変更すると、既定の参照元シートを固定できます。

```powershell
[string]$SourceSheetName = ''
```

例えば `本番データ` に固定したい場合:

```powershell
[string]$SourceSheetName = '本番データ'
```

スクリプト内にもコメントを入れてあります。

## 列位置が違う場合

列名が違っても、列の役割と位置が同じならそのまま動きます。

列位置が違う場合は `add_dashboard.ps1` の `columnMap` を修正してください。

例:

```powershell
$columnMap = @{
    ProgramId = 'H'
    FunctionName = 'F'
    Subsystem = 'E'
    Block = 'G'
    Group = 'I'
    ...
}
```

## 実行結果

実行が成功すると、PowerShell に以下のような情報が出ます。

```text
dashboard_sheet=進捗ダッシュボード
source_sheet=Sheet1
backup_path=.\plan.backup_20260323_231232.xlsx
workbook_path=.\plan.xlsx
```

## よくある原因

- `because it is being used by another process`
  - Excel で対象ファイルが開いています。閉じてから再実行してください
- `Source worksheet not found`
  - `-SourceSheetName` の指定名が実際のシート名と一致していません
- ダッシュボードの値が想定と違う
  - `columnMap` の列位置が元データと合っているか確認してください
