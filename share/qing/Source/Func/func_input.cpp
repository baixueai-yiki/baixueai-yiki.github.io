#include "func_input.h"

// ===== 外部变量（从主程序引用）=====
extern int character_x;
extern int character_y;

extern bool g_isDragging;
extern int g_offsetX;
extern int g_offsetY;

// 鼠标按下
void HandleDragMouseDown(HWND hwnd, LPARAM lParam, UINT window_width, UINT window_height)
{
    int x = LOWORD(lParam);
    int y = HIWORD(lParam);

    if (x >= character_x && x <= character_x + (int)window_width &&
        y >= character_y && y <= character_y + (int)window_height)
    {
        g_isDragging = true;
        g_offsetX = x - character_x;
        g_offsetY = y - character_y;
    }
}

// 鼠标移动
void HandleDragMouseMove(HWND hwnd, LPARAM lParam, UINT window_width, UINT window_height)
{
    if (!g_isDragging) return;

    int x = LOWORD(lParam);
    int y = HIWORD(lParam);

    character_x = x - g_offsetX;
    character_y = y - g_offsetY;

    InvalidateRect(hwnd, NULL, TRUE);
}

// 鼠标松开
void HandleDragMouseUp()
{
    g_isDragging = false;
}