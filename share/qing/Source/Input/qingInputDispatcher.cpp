#include "qingInputDispatcher.h"
#include "../Core/qingPetState.h"

// 外部状态
extern QingPetState g_pet;

// HitTest
static bool IsInsidePet(int x, int y)
{
    return x >= g_pet.x &&
           x <= g_pet.x + g_pet.w &&
           y >= g_pet.y &&
           y <= g_pet.y + g_pet.h;
}

void QingHandleInput(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    int x = GET_X_LPARAM(lParam);
    int y = GET_Y_LPARAM(lParam);

    switch (msg)
    {
    case WM_MOUSEMOVE:
    {
        // 控制点击穿透
        if (IsInsidePet(x, y))
        {
            SetWindowLong(hwnd, GWL_EXSTYLE,
                GetWindowLong(hwnd, GWL_EXSTYLE) & ~WS_EX_TRANSPARENT);
        }
        else
        {
            SetWindowLong(hwnd, GWL_EXSTYLE,
                GetWindowLong(hwnd, GWL_EXSTYLE) | WS_EX_TRANSPARENT);
        }

        // 拖动
        if (g_pet.isDragging)
        {
            g_pet.x = x - g_pet.dragOffsetX;
            g_pet.y = y - g_pet.dragOffsetY;
        }

        break;
    }

    case WM_LBUTTONDOWN:
    {
        if (IsInsidePet(x, y))
        {
            g_pet.isDragging = true;
            g_pet.dragOffsetX = x - g_pet.x;
            g_pet.dragOffsetY = y - g_pet.y;
        }
        break;
    }

    case WM_LBUTTONUP:
    {
        g_pet.isDragging = false;
        break;
    }
    }
}