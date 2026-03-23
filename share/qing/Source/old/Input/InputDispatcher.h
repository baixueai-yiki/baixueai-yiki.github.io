#pragma once
#include <windows.h>

// 处理输入事件（鼠标移动、点击），控制是否穿透以及拖动状态。
void QingHandleInput(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
