#include <windows.h>
#include "Platform/qingWindow.h"
#include "Render/qingRenderer.h"

// 窗口句柄（全局或从创建函数返回）
HWND g_hwnd = nullptr;

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR, int)
{
    // 1. 创建窗口
    g_hwnd = QingCreateWindow(hInstance);

    if (!g_hwnd)
        return -1;

    // 2. 初始化渲染系统
    QingRendererInit();

    // 3. 显示窗口
    ShowWindow(g_hwnd, SW_SHOW);

    // 4. 消息循环
    MSG msg = {};
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    // 5. 清理资源
    QingRendererShutdown();

    return 0;
}