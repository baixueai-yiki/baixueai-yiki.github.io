#include <windows.h>       // Windows 核心 API：创建窗口、处理消息、管理 HDC 等
#include <gdiplus.h>       // GDI+ 图形库：更方便地加载透明 PNG 并绘制图像
#include "Func/func_chat.h" // 聊天气泡的接口（右键三次节奏时弹出）
#pragma comment(lib,"gdiplus.lib") // 链接 GDI+ 库，避免链接器提示缺少符号

using namespace Gdiplus; // 直接使用 GDI+ 命名空间下的类，避免 everywhere 都写 Gdiplus::


// ================= 全局状态（跨函数共享） =================

int window_x = 100;    // 窗口初始位置，X 轴
int window_y = 100;    // 窗口初始位置，Y 轴
UINT window_width = 200;  // 窗口宽度，实际会被图片真实宽度覆盖
UINT window_height = 300; // 窗口高度，实际会被图片真实高度覆盖
HWND hwndMain = nullptr;  // 主窗口句柄，后面所有窗口调用都会用到
ULONG_PTR gdiplusToken;   // GDI+ 的启动句柄（必须在 GdiplusShutdown 时使用）
DWORD clickTimes[6] = { 0 }; // 记录最近 6 次右键点击时间，用于判定“慢三下＋快三下”的节奏
Bitmap* g_character = nullptr; // 桌宠图片，保持全局避免重复加载


// ================= 窗口过程：负责处理所有 Windows 消息 =================
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_CREATE:
        // 创建完成时被调用，这里不做额外逻辑，图片加载在主函数里处理
        break;

    case WM_PAINT:
    {
        // 每当窗口需要重绘（首次显示、遮挡后恢复等）就进入该分支
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        if (g_character && g_character->GetLastStatus() == Ok)
        {
            Graphics graphics(hdc); // GDI+ 的绘图上下文
            graphics.DrawImage(g_character, 0, 0); // 直接绘制整张角色图
        }
        EndPaint(hwnd, &ps); // 必须调用，通知系统重绘完成
        break;
    }

    case WM_RBUTTONDOWN:
    {
        // 记录本次点击，用以判断节奏是否满足“慢三下+快三下”
        DWORD now = GetTickCount();
        for (int i = 0; i < 5; ++i)
            clickTimes[i] = clickTimes[i + 1];
        clickTimes[5] = now;

        const int slowMax = 1000; // 慢节奏 1 秒以内
        const int fastMax = 300;  // 快节奏 0.3 秒以内
        bool slowTriple = (clickTimes[3] - clickTimes[0] <= slowMax);
        bool fastTriple = (clickTimes[5] - clickTimes[3] <= fastMax);

        if (slowTriple && fastTriple)
        {
            // 条件满足，弹出聊天输入气泡
            for (int i = 0; i < 6; ++i)
                clickTimes[i] = 0;
            ShowChatInput(hwnd);
        }
        break;
    }

    case WM_DESTROY:
    {
        // 窗口准备关闭，释放资源
        delete g_character;
        g_character = nullptr;
        PostQuitMessage(0);
        break;
    }

    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }

    return 0;
}


// ================= 主入口 wWinMain（Windows GUI 程序的起点） =================
int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR, int)
{
    // GDI+ 初始化
    GdiplusStartupInput gdiplusStartupInput;
    GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, nullptr);

    // 预先加载桌宠图片，后续的窗口尺寸依赖此尺寸
    g_character = Bitmap::FromFile(L"character.png");
    if (!g_character || g_character->GetLastStatus() != Ok)
    {
        MessageBoxW(NULL, L"图片加载失败，请确保 Content\\Images\\character.png 存在。", L"错误", MB_OK);
        return 0;
    }
    window_width = g_character->GetWidth();
    window_height = g_character->GetHeight();

    // 注册窗口类（告诉系统我们要创建一个什么样的窗口）
    WNDCLASSW wc = {};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"CharacterWindow";
    RegisterClassW(&wc);

    // 创建无边框、透明、置顶的窗口，覆盖整个屏幕
    hwndMain = CreateWindowExW(
        WS_EX_TOPMOST | WS_EX_LAYERED | WS_EX_TOOLWINDOW,
        wc.lpszClassName,
        L"桌宠",
        WS_POPUP,
        window_x, window_y,
        window_width, window_height,
        nullptr, nullptr, hInstance, nullptr
    );

    SetLayeredWindowAttributes(hwndMain, 0, 255, LWA_ALPHA); // 确保窗口允许透明
    ShowWindow(hwndMain, SW_SHOW);

    // 主消息循环：不断从系统取消息，交给 WndProc 处理
    MSG msg = {};
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    // 程序退出前关闭 GDI+
    GdiplusShutdown(gdiplusToken);
    return 0;
}
