#pragma once
#include <windows.h>

// 窗口过程由 qingWindow.cpp 实现，用来分发消息给输入/渲染模块。
LRESULT CALLBACK QingWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

// 创建全屏透明窗口并返回 HWND（失败返回 nullptr）。
HWND QingCreateWindow(HINSTANCE hInstance);
