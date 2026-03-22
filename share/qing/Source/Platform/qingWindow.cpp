#include <windows.h>
#include "qingWindow.h"
#include "../Input/qingInputDispatcher.h"

// =======================
// 窗口过程（只做转发）
// =======================
LRESULT CALLBACK QingWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_MOUSEMOVE:
    case WM_LBUTTONDOWN:
    case WM_LBUTTONUP:
    case WM_LBUTTONDBLCLK:
    {
        // 👉 所有输入统一交给 InputDispatcher
        QingHandleInput(hwnd, msg, wParam, lParam);
        break;
    }

    case WM_PAINT:
    {
        // 👉 渲染交给 Render 模块（这里先留空）
        PAINTSTRUCT ps;
        BeginPaint(hwnd, &ps);
        EndPaint(hwnd, &ps);
        break;
    }

    case WM_DESTROY:
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }

    return 0;
}

// =======================
// 创建全屏透明窗口
// =======================
HWND QingCreateWindow(HINSTANCE hInstance)
{
    WNDCLASSW wc = {};
    wc.lpfnWndProc = QingWndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"QingPetWindow";

    RegisterClassW(&wc);

    HWND hwnd = CreateWindowExW(
        WS_EX_LAYERED | WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_TRANSPARENT,
        wc.lpszClassName,
        L"QingPet",
        WS_POPUP,
        0, 0,
        GetSystemMetrics(SM_CXSCREEN),
        GetSystemMetrics(SM_CYSCREEN),
        nullptr,
        nullptr,
        hInstance,
        nullptr
    );

    // 设置透明度（可后续改为真正的透明+渲染）
    SetLayeredWindowAttributes(hwnd, 0, 255, LWA_ALPHA);

    return hwnd;
}