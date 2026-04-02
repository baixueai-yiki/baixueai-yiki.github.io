#pragma once
#include <windows.h>

// 渲染接口，由窗口过程在 WM_PAINT 时调用
void Render(HWND hwnd);
