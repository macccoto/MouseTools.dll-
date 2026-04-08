/**
 * MouseTools.h
 *
 * Excel / Access / Word の VBA などから呼び出すための、
 * マウスカーソル位置の取得・移動だけを行うネイティブ DLL の宣言ヘッダ。
 *
 * - 言語は C（C++ ではない）
 * - 呼び出し規約: __stdcall（VBA の Declare と相性がよい）
 * - エクスポート: __declspec(dllexport)
 *
 * ビルド時にマクロ MOUSETOOLS_BUILD を定義した翻訳単位では dllexport、
 * それ以外では dllimport として扱う（この DLL 本体では .c 側で定義する）。
 */

#ifndef MOUSETOOLS_H
#define MOUSETOOLS_H

#ifdef MOUSETOOLS_BUILD
#define MOUSETOOLS_API __declspec(dllexport)
#else
#define MOUSETOOLS_API __declspec(dllimport)
#endif

#ifdef __cplusplus
extern "C" {
#endif

/**
 * 現在のマウスカーソル位置を取得する。
 *
 * @param x 画面座標 X を書き込む先（スクリーン座標、ピクセル）。NULL 不可。
 * @param y 画面座標 Y を書き込む先。NULL 不可。
 * @return 成功時 1、失敗時または引数が不正なら 0。
 *
 * 内部では Win32 API の GetCursorPos を使用する。
 */
MOUSETOOLS_API int __stdcall GetCursorPosXY(int* x, int* y);

/**
 * 画面上の絶対座標へカーソルを移動する。
 *
 * @param x スクリーン X 座標（ピクセル）。
 * @param y スクリーン Y 座標（ピクセル）。
 * @return 成功時 1、失敗時 0。
 *
 * 内部では SetCursorPos を使用する。
 * 複数モニタ環境では仮想スクリーン上の座標となる。
 */
MOUSETOOLS_API int __stdcall MoveCursorTo(int x, int y);

/**
 * 現在位置から相対的にカーソルを移動する。
 *
 * @param dx X 方向の移動量（ピクセル）。
 * @param dy Y 方向の移動量（ピクセル）。
 * @return 成功時 1、途中の取得や移動が失敗したら 0。
 *
 * 現在位置を GetCursorPos で取得し、(x+dx, y+dy) に SetCursorPos する。
 */
MOUSETOOLS_API int __stdcall MoveCursorBy(int dx, int dy);

#ifdef __cplusplus
}
#endif

#endif /* MOUSETOOLS_H */
