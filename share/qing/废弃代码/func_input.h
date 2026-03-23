#pragma once
#include <windows.h>

// 拖拽输入处理
void HandleDragMouseDown(HWND hwnd, LPARAM lParam, UINT window_width, UINT window_height);
void HandleDragMouseMove(HWND hwnd, LPARAM lParam, UINT window_width, UINT window_height);
void HandleDragMouseUp();