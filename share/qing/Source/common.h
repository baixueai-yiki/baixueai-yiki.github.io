#pragma once
#include <windows.h>

// 角色当前位置与大小（四边形）
extern int g_charX;
extern int g_charY;
extern int g_charW;
extern int g_charH;

// 拖拽状态：是否正在拖拽，以及鼠标偏移
extern bool g_isDragging;
extern int g_offsetX;
extern int g_offsetY;

// 主窗口句柄，供其他模块使用（比如聊天窗口的父窗口）
extern HWND g_hwnd;
