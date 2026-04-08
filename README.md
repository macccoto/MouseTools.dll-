# MouseTools.dll-

Excel / Access / Word の **VBA から**マウスカーソルの**座標取得**と**移動**だけを行う、最小構成の **Windows ネイティブ DLL**（**C 言語**）です。

用途の例: 業務自動化の補助、画面位置の合わせ込み、UI テスト補助。**ユーザーがマクロ実行を明示的に選ぶ前提**のユーティリティ向けです。

- リポジトリ: [macccoto/MouseTools.dll-](https://github.com/macccoto/MouseTools.dll-)

## 方針（含まないもの）

実装しているのは次の **3 関数のみ**です。

- 現在座標の取得（`GetCursorPos`）
- 絶対座標への移動（`SetCursorPos`）
- 相対移動（取得 → 加算 → `SetCursorPos`）

**含めていません:** クリック送信、キー送信、フック、常駐、COM、.NET、MFC / ATL。

## 要件

- Windows
- ビルド: Visual Studio または **Build Tools for Visual Studio**（`cl.exe`）
- 実行: **32bit Office には 32bit DLL**、**64bit Office には 64bit DLL**（ビット数を一致させる）

## API

| 関数 | 説明 | 戻り値 |
|------|------|--------|
| `GetCursorPosXY(int* x, int* y)` | 現在のスクリーン座標を `*x`, `*y` に書き込む | 成功 `1`、失敗または `NULL` で `0` |
| `MoveCursorTo(int x, int y)` | 絶対座標 `(x, y)` に移動 | 成功 `1`、失敗 `0` |
| `MoveCursorBy(int dx, int dy)` | 現在位置から `(dx, dy)` だけ移動 | 成功 `1`、失敗 `0` |

呼び出し規約: **`__stdcall`**。マルチモニタ環境では仮想スクリーン座標になります。

## ビルド

**Developer Command Prompt**（または **x64 / x86 Native Tools**）を開き、`MouseTools.c` があるディレクトリで実行します。

### 64bit DLL（64bit Office 向け）

```bat
call "%ProgramFiles(x86)%\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat"
cd /d このリポジトリの MouseTools フォルダ
cl /LD /O2 /W3 MouseTools.c /FeMouseTools_x64.dll
```

インストールパスは環境に合わせて変更してください（例: `Community`、`18\BuildTools` など）。

### 32bit DLL（32bit Office 向け）

```bat
call "%ProgramFiles(x86)%\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars32.bat"
cd /d このリポジトリの MouseTools フォルダ
cl /LD /O2 /W3 MouseTools.c /FeMouseTools_x86.dll
```

生成例: `MouseTools_x64.dll` または `MouseTools_x86.dll`（ほかに `.lib` / `.exp` / `.obj` が出ます）。VBA には **対応するビット数の `.dll` だけ**あれば足ります。

文字コード警告（C4819）を抑えたい場合は `cl` に `/utf-8` を追加できます。

## VBA から呼ぶ（宣言の例）

`Lib` には、実際に配置した DLL の名前（またはフルパス）を指定します。64bit Office では **`PtrSafe`** が必要です。

```vb
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetCursorPosXY Lib "MouseTools_x64.dll" ( _
        ByRef x As Long, ByRef y As Long) As Long
    Private Declare PtrSafe Function MoveCursorTo Lib "MouseTools_x64.dll" ( _
        ByVal x As Long, ByVal y As Long) As Long
    Private Declare PtrSafe Function MoveCursorBy Lib "MouseTools_x64.dll" ( _
        ByVal dx As Long, ByVal dy As Long) As Long
#Else
    Private Declare Function GetCursorPosXY Lib "MouseTools_x86.dll" ( _
        ByRef x As Long, ByRef y As Long) As Long
    Private Declare Function MoveCursorTo Lib "MouseTools_x86.dll" ( _
        ByVal x As Long, ByVal y As Long) As Long
    Private Declare Function MoveCursorBy Lib "MouseTools_x86.dll" ( _
        ByVal dx As Long, ByVal dy As Long) As Long
#End If
```

- **`ByRef` / `ByVal`:** C 側が `int*` の引数は VBA では `ByRef`、値の `int` は `ByVal`。
- **`LongPtr`:** この Declare では座標は `Long` で足ります（生ポインタを値として渡す API では `LongPtr` が必要になります）。

## 動作確認のMinimal例

```vb
Sub Sample()
    Dim x As Long, y As Long
    If GetCursorPosXY(x, y) = 1 Then
        Debug.Print x, y
    End If
    MoveCursorBy 10, 0
End Sub
```

本番に近いテストでは、`MsgBox` で続行確認してから移動するなど、安全側に倒してください。

## トラブルシューティング

| 症状 | 対処のヒント |
|------|----------------|
| DLL が見つからない | `Lib` をフルパスにする、DLL を Office と**同じビット数**でビルドしたか確認 |
| エントリ ポイントが見つからない | 32/64 の取り違え、別名の DLL を読み込んでいないか確認（依存関係ビューアでエクスポート名を確認） |
| 呼び出し規約が不正（Bad DLL calling convention） | `__stdcall` と一致しているか、`GetCursorPosXY` は `ByRef`、移動系は `ByVal` か確認 |
| 型不一致 | 戻り値・座標は `Long` で統一 |
| Excel が落ちる | ビット数不一致が典型。宣言と DLL を見直す |

## 構成

| ファイル | 内容 |
|----------|------|
| `MouseTools.h` | 公開 API の宣言 |
| `MouseTools.c` | 実装（`user32.lib` を `pragma` でリンク） |

## ライセンス

未指定です。利用・再配布条件が必要な場合はリポジトリオーナーに合わせて `LICENSE` を追加してください。

## 免責

このソフトウェアは現状のまま提供されます。マウスカーソルを動かすため、**実行前に利用者が内容を理解し、組織のセキュリティ方針に従って**ご利用ください。
