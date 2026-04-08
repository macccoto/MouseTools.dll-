/**
 * MouseTools.c
 *
 * MouseTools.dll — マウスカーソル位置の取得・絶対移動・相対移動のみ。
 * Win32 API: GetCursorPos / SetCursorPos
 *
 * ビット数: 32bit Office 用には 32bit DLL、64bit Office 用には 64bit DLL を
 * それぞれ別々にビルドして配置すること（混在不可）。
 */

/* この翻訳単位ではヘッダ上の API を dllexport として扱う */
#define MOUSETOOLS_BUILD
#include "MouseTools.h"

/*
 * Windows.h:
 *   BOOL, POINT, GetCursorPos, SetCursorPos など。
 * 警告レベルが高い環境向けに WIN32_LEAN_AND_MEAN でヘッダを軽量化する。
 */
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#include <Windows.h>

/* GetCursorPos / SetCursorPos は user32 にあるため、明示的にリンクする */
#pragma comment(lib, "user32.lib")

/*
 * __stdcall では x86 ビルド時、リンカが「飾り付き」シンボル名
 * （例: _GetCursorPosXY@8）をエクスポートすることがある。
 * VBA の Declare では通常「飾りのない名前」で検索するため、
 * 同じ関数を飾りなし名でもエクスポートしておく。
 *
 * x64 では名前の飾りは付かないため、この指定は不要だが、
 * 指定しても害にはならない。
 */
#ifdef _M_IX86
#pragma comment(linker, "/EXPORT:GetCursorPosXY=_GetCursorPosXY@8")
#pragma comment(linker, "/EXPORT:MoveCursorTo=_MoveCursorTo@8")
#pragma comment(linker, "/EXPORT:MoveCursorBy=_MoveCursorBy@8")
#endif

/**
 * GetCursorPosXY — 現在のカーソル座標を取得。
 *
 * 戻り値:
 *   1 … 正常終了（*x, *y に現在位置を書き込んだ）
 *   0 … x または y が NULL、または GetCursorPos が失敗（GetLastError で理由確認可）
 */
int __stdcall GetCursorPosXY(int* x, int* y)
{
    POINT pt;

    if (x == NULL || y == NULL) {
        return 0;
    }

    /*
     * GetCursorPos:
     *   成功時は非 0、失敗時は 0 を返し SetLastError を設定する。
     */
    if (!GetCursorPos(&pt)) {
        return 0;
    }

    *x = (int)pt.x;
    *y = (int)pt.y;
    return 1;
}

/**
 * MoveCursorTo — 絶対座標へ移動。
 *
 * 戻り値:
 *   1 … SetCursorPos が成功
 *   0 … SetCursorPos が失敗
 */
int __stdcall MoveCursorTo(int x, int y)
{
    /*
     * SetCursorPos:
     *   成功時は非 0。無効な座標などでは 0 になる場合がある。
     *   仮想スクリーン座標系で動作する（マルチモニタ対応）。
     */
    if (!SetCursorPos(x, y)) {
        return 0;
    }
    return 1;
}

/**
 * MoveCursorBy — 相対移動。
 *
 * 手順: 現在位置取得 → 新座標計算 → SetCursorPos
 * いずれかが失敗すれば 0。
 */
int __stdcall MoveCursorBy(int dx, int dy)
{
    POINT pt;
    int nx;
    int ny;

    if (!GetCursorPos(&pt)) {
        return 0;
    }

    nx = (int)pt.x + dx;
    ny = (int)pt.y + dy;

    if (!SetCursorPos(nx, ny)) {
        return 0;
    }
    return 1;
}
