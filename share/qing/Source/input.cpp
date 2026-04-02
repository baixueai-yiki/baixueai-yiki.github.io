#include "input.h"
#include "common.h"

// 判断鼠标是否落在角色当前的矩形区域之内
bool IsInsideCharacter(int x, int y)
{
    return x >= g_charX &&
           x <= g_charX + g_charW &&
           y >= g_charY &&
           y <= g_charY + g_charH;
}

// 鼠标按下时，如果点击在角色上就准备拖拽，并记录光标相对于角色的偏移
void HandleDragMouseDown(HWND hwnd, LPARAM lParam)
{
    int x = LOWORD(lParam);
    int y = HIWORD(lParam);

    if (IsInsideCharacter(x, y))
    {
        g_isDragging = true;
        g_offsetX = x - g_charX;
        g_offsetY = y - g_charY;
        SetCapture(hwnd); // 捕获鼠标，保证即使移到窗口外仍能收到消息
    }
}

// 鼠标移动时根据偏移更新角色位置，并主动刷新窗口
void HandleDragMouseMove(HWND hwnd, LPARAM lParam)
{
    if (!g_isDragging)
        return;

    int x = LOWORD(lParam);
    int y = HIWORD(lParam);

    g_charX = x - g_offsetX;
    g_charY = y - g_offsetY;

    InvalidateRect(hwnd, NULL, FALSE); // 请求重绘，让窗口立即反映新位置
}

// 鼠标抬起时结束拖拽，释放之前的捕获状态
void HandleDragMouseUp()
{
    g_isDragging = false;
    ReleaseCapture();
}
