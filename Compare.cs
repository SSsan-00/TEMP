using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace ResultSetSorter;

public static class ResultSetSorterService
{
    /// <summary>
    /// IDictionary を用いて結果セットを in-place(再構築) で並び替える
    ///
    /// 期待する構造（実行時チェック）:
    /// - 外側: IDictionary（RowId -> Row）
    ///   - Key は int である必要がある（タイブレークと 0..N-1 振り直しのため）
    ///   - Value は IDictionary（列名 -> 値）である必要がある
    /// - sortKey の値は string または null のみ許可
    ///
    /// 挙動:
    /// - ASC: NULLS LAST / DESC: NULLS FIRST
    /// - 文字列比較は UTF-8 バイト列の辞書順（Culture 非依存）
    /// - 同値なら RowId 昇順でタイブレーク
    /// - 並び替え後、全体結果セットを Clear し、RowId を 0..N-1 に振り直して再構築
    /// </summary>
    /// <param name="resultSet">全体結果セット（RowId -> 行結果セット）</param>
    /// <param name="sortKey">ソート対象の列名</param>
    /// <param name="descending">降順にするか（ASC: NULLS LAST / DESC: NULLS FIRST）</param>
    /// <returns>並び替え後の同一インスタンス（IDictionary として返す）</returns>
    public static IDictionary SortInPlace(
        IDictionary resultSet,
        string sortKey,
        bool descending = false)
    {
        // resultSet = null なら例外
        ArgumentNullException.ThrowIfNull(resultSet);

        // sortKey = null/空/空白 なら例外
        if (string.IsNullOrWhiteSpace(sortKey))
        {
            throw new ArgumentException("sortKey cannot be null or whitespace.", nameof(sortKey));
        }

        // ---- UTF-8 バイト比較 -----------------
        // 同じ文字列が何度も出る場合、UTF-8 変換を毎回やるとコストが高いのでキャッシュする
        var byteCache = new Dictionary<string, byte[]>(capacity: 16);

        byte[] GetUtf8Bytes(string value)
        {
            if (byteCache.TryGetValue(value, out var cached))
            {
                return cached;
            }

            var bytes = Encoding.UTF8.GetBytes(value);
            byteCache[value] = bytes;
            return bytes;
        }

        int CompareUtf8Bytes(string x, string y)
        {
            var xBytes = GetUtf8Bytes(x);
            var yBytes = GetUtf8Bytes(y);

            // 共通部分の長さまでバイト単位で比較
            var minLength = xBytes.Length < yBytes.Length ? xBytes.Length : yBytes.Length;

            for (var i = 0; i < minLength; i++)
            {
                // バイト差が出た時点で順序が確定
                var diff = xBytes[i] - yBytes[i];
                if (diff != 0)
                {
                    return diff;
                }
            }

            // 共通部分が一致した場合は長さで決める（短い方が小さい）
            if (xBytes.Length == yBytes.Length)
            {
                return 0;
            }

            return xBytes.Length < yBytes.Length ? -1 : 1;
        }
        // --------------------------------------------------------------------

        // ソート用に詰め替え
        // - RowId: 元の行番号（タイブレークに使う）
        // - Row  : 行結果セット（IDictionary）
        // - Value: sortKey の値（string?）
        var rows = new List<(int RowId, IDictionary Row, string? Value)>(resultSet.Count);

        // sortKey が「どの行にも存在しない」場合に例外にするためのフラグ
        var sortKeyExists = false;

        // IDictionary を列挙（非ジェネリックなので DictionaryEntry を使う）
        foreach (DictionaryEntry entry in resultSet)
        {
            // 外側キーは object だが、元仕様に合わせるため int(RowId) を要求
            if (entry.Key is not int rowId)
            {
                throw new ArgumentException(
                    "Outer dictionary keys must be int (RowId) to preserve the original behavior.",
                    nameof(resultSet));
            }

            // 行（内側辞書）は null ならNG
            if (entry.Value is null)
            {
                throw new ArgumentException("Row dictionary cannot be null.", nameof(resultSet));
            }

            // 行が IDictionary でないならNG（列名→値 の辞書として扱えない）
            if (entry.Value is not IDictionary row)
            {
                throw new ArgumentException(
                    "Each row must implement IDictionary (columnName -> value).",
                    nameof(resultSet));
            }

            string? valueToSort = null;

            // sortKey の存在確認（非ジェネリック IDictionary の Contains は object キー）
            if (row.Contains(sortKey))
            {
                sortKeyExists = true;

                var rawValue = row[sortKey]; // 存在するのでインデクサで取得してOK

                // sortKey の値は string? だけ許可（元仕様と同じ）
                if (rawValue is not null && rawValue is not string)
                {
                    throw new ArgumentException(
                        $"Value for key '{sortKey}' must be string or null.",
                        nameof(resultSet));
                }

                valueToSort = (string?)rawValue;
            }

            rows.Add((rowId, row, valueToSort));
        }

        // sortKey がどの行にも存在しないなら例外（元仕様と同じ）
        if (!sortKeyExists)
        {
            throw new ArgumentException(
                $"The sort key '{sortKey}' does not exist in any row.",
                nameof(resultSet));
        }

        // 比較関数（NULLS FIRST/LAST + UTF-8 bytes + RowId タイブレーク）
        int CompareRows(
            (int RowId, IDictionary Row, string? Value) left,
            (int RowId, IDictionary Row, string? Value) right)
        {
            var leftIsNull = left.Value is null;
            var rightIsNull = right.Value is null;

            // NULL絡みの扱い
            if (leftIsNull || rightIsNull)
            {
                // 両方 NULL → RowId 昇順でタイブレーク
                if (leftIsNull && rightIsNull)
                {
                    return left.RowId.CompareTo(right.RowId);
                }

                // DESC: NULLS FIRST / ASC: NULLS LAST
                if (descending)
                {
                    return leftIsNull ? -1 : 1;
                }

                return leftIsNull ? 1 : -1;
            }

            // 両方非NULL：UTF-8 バイト列で比較
            var cmp = CompareUtf8Bytes(left.Value!, right.Value!);

            // DESC なら反転
            if (descending)
            {
                cmp = -cmp;
            }

            // 同値なら RowId 昇順でタイブレーク
            if (cmp != 0)
            {
                return cmp;
            }

            return left.RowId.CompareTo(right.RowId);
        }

        // 並び替え
        rows.Sort(CompareRows);

        // 外側辞書をクリア
        resultSet.Clear();

        // 0..N-1 に振り直して詰め直し
        var nextRowId = 0;
        foreach (var r in rows)
        {
            // IDictionary のインデクサは object を受けるが、
            // 実際には int を入れる（= 元仕様の RowId 振り直し）
            resultSet[nextRowId++] = r.Row;
        }

        return resultSet;
    }

    /// <summary>
    /// オーバーロード：
    /// 「元のカスタムクラス型のまま返したい」場合に使う
    /// 内部は IDictionary 版に委譲し、挙動は変えない
    /// </summary>
    public static T SortInPlace<T>(
        T resultSet,
        string sortKey,
        bool descending = false)
        where T : IDictionary
    {
        SortInPlace((IDictionary)resultSet, sortKey, descending);
        return resultSet;
    }
}
