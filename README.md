# VBA 配布物（.xlsm 操作用）

## 概要
この `VBA/` 配下の `.bas` をマクロ有効ブック（`.xlsm`）へ取り込み、
その `.xlsm` から外部 `.xlsx` を開いて `ソース` シートの `C5:D` を整形します。

- 対象: 外部 `.xlsx`
- 対象シート: `ソース`
- 対象列: `C`（PHP） / `D`（C#）
- 動作: 対応行を揃えるため、片側に空セル行を挿入した配列を再書き戻し
- 安全策: 保存前に `*_yyyyMMdd_HHmmss.bak.xlsx` を `SaveCopyAs` で作成

## 配置ファイル
- `VBA/modMain.bas`
- `VBA/modAlignEngine.bas`
- `VBA/modNormalize.bas`
- `VBA/modTests.bas`

## 取り込み手順（Excel）
1. 新規または既存の `.xlsm` を作成
2. VBE（`Alt+F11`）を開く
3. `標準モジュール` を追加し、各 `.bas` をインポート
4. `modMain.SetupOperationSheet` を一度実行（操作シートとボタン作成）
5. 以後は `Run` 実行、または「整形実行」ボタン押下

## 実行フロー
1. ファイルダイアログで対象 `.xlsx` を選択
2. 対象ブックを編集可能で開く（既に開いていればその参照を使用）
3. バックアップを同一フォルダへ作成
4. `ソース` シート `C5:D最終行` を配列で読込
5. diff ロジックで整形後の 2 列配列を作成
6. `C5:D` に再書き戻し、`.xlsx` 上書き保存

## diff-viewer からの移植マッピング
参照元: `src/diffEngine/`

- `normalize.ts`
  - `normalizeText` → `modNormalize.NormalizeText`
- `appendLiteral.ts`
  - `extractAppendLiteral` / `toAppendLiteralOrLine` → `modNormalize.ExtractAppendLiteral` / `modNormalize.ToAppendLiteralOrLine`
- `lineSignature.ts`
  - `extractLineKey` → `modNormalize.ExtractLineKey`
- `diffLines.ts`
  - `normalizeForMatch` → `modNormalize.NormalizeForMatch`
  - `buildMyersTrace` → `modAlignEngine.BuildMyersTrace`
  - `backtrackOps` → `modAlignEngine.BacktrackOps`
  - `buildUniquePairs` → `modAlignEngine.BuildUniquePairs`
  - `longestIncreasingPairs` → `modAlignEngine.LongestIncreasingPairs`
  - `diffLinesMyers` / `diffLinesPatience` / `diffLinesFromLines`
    → `modAlignEngine.DiffLinesMyers` / `modAlignEngine.DiffLinesPatience` / `modAlignEngine.DiffLinesFromLines`
- `pairReplace.ts`
  - `pairReplace` / `pairBlock` / `alignBracePairs`
    → `modAlignEngine.PairReplace` / `modAlignEngine.PairBlock` / `modAlignEngine.AlignBracePairs`

### 備考（lineSimilarity 相当）
本件では「行対応（line alignment）」に必要な部分のみを対象としているため、
`pairReplace.ts` が要求するスコアリング要素を `modAlignEngine.ScoreLinePairSimple` と
その補助関数群へ移植しています。

- `AppendLine` ペイロード抽出
- 行キー（`extractLineKey`）
- 識別子 / 文字列 / 数値トークンの重なり
- 近傍ウィンドウ + トークン索引での候補探索

この構成で `replace` ペア化を行い、空セル挿入を最小化します。

## テスト
即時ウィンドウから以下を実行:

- `RunAllTests`
- `Test_AppendLineAlignment`
- `Test_InsertDeleteAlignment`

`Debug.Print` 出力とエラー有無で確認できます。
