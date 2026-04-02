#include "render.h"
#include "common.h"

// 渲染函数，由窗口在 WM_PAINT 时调用，负责把角色矩形绘制到屏幕
void Render(HWND hwnd)
{
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(hwnd, &ps);

    HBRUSH brush = CreateSolidBrush(RGB(255, 200, 200)); // 粉色背景，模拟角色

    RECT rect;
    rect.left = g_charX;
    rect.top = g_charY;
    rect.right = g_charX + g_charW;
    rect.bottom = g_charY + g_charH;

    FillRect(hdc, &rect, brush);

    DeleteObject(brush);
    EndPaint(hwnd, &ps);
}
