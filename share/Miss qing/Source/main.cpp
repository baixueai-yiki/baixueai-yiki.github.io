#include <windows.h>
#include <gdiplus.h>
#include "Func/func_chat.h"
#pragma comment(lib,"gdiplus.lib")

using namespace Gdiplus;

int x = 100, y = 100;// 窗口初始位置
HWND hwndMain = nullptr;// 主窗口句柄（用于全局访问）
ULONG_PTR gdiplusToken;// GDI+ 初始化句柄
DWORD clickTimes[6] = {0};// 记录点击时间，用于检测“特定节奏的右键点击”
Bitmap* g_character = nullptr;// 全局图片指针（避免每次重绘都重新加载）

// 窗口过程函数（所有窗口消息都在这里处理）
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    // 窗口创建时触发（适合做初始化）
    switch (msg)
    {
    case WM_CREATE:
    {
        g_character = Bitmap::FromFile(L"character.png");// 从文件加载桌宠图片
        if (!g_character || g_character->GetLastStatus() != Ok)// 检查图片是否加载成功
        {
            MessageBox(hwnd, L"图片加载失败", L"错误", MB_OK);
        }
        break;
    }

    // 窗口需要重绘时触发（系统或你调用 InvalidateRect）
    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);// 开始绘制，获取设备上下文 HDC
        if (g_character && g_character->GetLastStatus() == Ok)// 确保图片存在且加载成功
        {
            Graphics graphics(hdc);// GDI+ 绘图对象（封装了绘图操作）
            graphics.DrawImage(g_character, 0, 0);// 将图片绘制到窗口左上角 (0,0)
        }

        // 结束绘制
        EndPaint(hwnd, &ps);
        break;
    }

    // 鼠标右键按下时触发
    case WM_RBUTTONDOWN:
    {
        // 获取当前时间（毫秒）
        DWORD now = GetTickCount();

        // 将历史点击记录整体向前移动一位
        for (int i = 0; i < 5; i++)
            clickTimes[i] = clickTimes[i + 1];

        // 记录最新一次点击时间
        clickTimes[5] = now;

        // 定义时间阈值（用于判断点击节奏）
        const int slowMax = 1000; // 慢节奏最大间隔
        const int fastMax = 300;  // 快节奏最大间隔

        // 判断是否满足“慢三次点击”
        bool slowTriple = (clickTimes[3] - clickTimes[0] <= slowMax);

        // 判断是否满足“快三次点击”
        bool fastTriple = (clickTimes[5] - clickTimes[3] <= fastMax);

        // 如果两种节奏都满足，认为触发特殊操作
        if (slowTriple && fastTriple)
        {
            // 清空点击记录，避免重复触发
            for (int i = 0; i < 6; i++)
                clickTimes[i] = 0;

            // 调用聊天输入窗口（你在 Func/func_chat.h 里实现的）
            ShowChatInput(hwnd);
        }
        break;
    }

    // 窗口销毁时触发
    case WM_DESTROY:
        // 释放图片资源，避免内存泄漏
        if (g_character)
        {
            delete g_character;
            g_character = nullptr;
        }

        // 通知系统退出消息循环
        PostQuitMessage(0);
        break;

    // 默认消息处理（必须保留）
    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }

    return 0;
}

// 程序入口（Unicode版本）
int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR, int)
{
    // 初始化 GDI+
    GdiplusStartupInput gdiplusStartupInput;
    GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

    // 定义窗口类
    WNDCLASSW wc = {};
    wc.lpfnWndProc = WndProc;            // 指定窗口过程函数
    wc.hInstance = hInstance;            // 当前实例
    wc.lpszClassName = L"CharacterWindow"; // 窗口类名

    // 注册窗口类
    RegisterClassW(&wc);

    // 创建窗口（桌宠窗口）
    hwndMain = CreateWindowExW(
        WS_EX_TOPMOST | WS_EX_LAYERED | WS_EX_TOOLWINDOW, // 顶层 + 分层 + 工具窗口（不出现在任务栏）
        wc.lpszClassName,
        L"桌宠",
        WS_POPUP, // 无边框窗口
        x, y, 200, 300, // 位置 + 尺寸
        NULL, NULL, hInstance, NULL
    );

    // 设置窗口透明度（255=完全不透明）
    SetLayeredWindowAttributes(hwndMain, 0, 255, LWA_ALPHA);

    // 显示窗口
    ShowWindow(hwndMain, SW_SHOW);

    // 消息循环（程序主循环）
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg); // 键盘消息转换
        DispatchMessage(&msg);  // 分发到 WndProc
    }

    // 释放 GDI+
    GdiplusShutdown(gdiplusToken);

    return 0;
}