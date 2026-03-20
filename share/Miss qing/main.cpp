#include <windows.h>
#include <gdiplus.h>
#include "func\func_chat.h"
#pragma comment(lib,"gdiplus.lib")

using namespace Gdiplus; // 使用GDI+命名空间

int x=100,y=100;
HWND hwndMain=nullptr;
ULONG_PTR gdiplusToken;
DWORD clickTimes[6]={0};

LRESULT CALLBACK WndProc(HWND hwnd,UINT msg,WPARAM wParam,LPARAM lParam)
{
    switch(msg)
    {
    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd,&ps);

        Graphics graphics(hdc);
        Bitmap characterImage(L"character.png"); // 桌宠图片
        graphics.DrawImage(&characterImage,0,0);

        EndPaint(hwnd,&ps);
        break;
    }
    case WM_TIMER:
        InvalidateRect(hwnd,NULL,TRUE);
        break;
    case WM_RBUTTONDOWN:
    {
        DWORD now = GetTickCount();
        for(int i=0;i<5;i++) clickTimes[i]=clickTimes[i+1];
        clickTimes[5]=now;

        const int slowMax=1000,fastMax=300;
        bool slowTriple = (clickTimes[3]-clickTimes[0]<=slowMax);
        bool fastTriple = (clickTimes[5]-clickTimes[3]<=fastMax);

        if(slowTriple && fastTriple)
        {
            for(int i=0;i<6;i++) clickTimes[i]=0;
            ShowChatInput(hwnd);
        }
        break;
    }
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hwnd,msg,wParam,lParam);
    }
    return 0;
}

int WINAPI wWinMain(HINSTANCE hInstance,HINSTANCE, PWSTR, int)
{
    GdiplusStartupInput gdiplusStartupInput;
    GdiplusStartup(&gdiplusToken,&gdiplusStartupInput,NULL);

    WNDCLASSW wc={};
    wc.lpfnWndProc=WndProc;
    wc.hInstance=hInstance;
    wc.lpszClassName=L"CharacterWindow";
    RegisterClassW(&wc);

    hwndMain=CreateWindowExW(
        // 设置窗口为顶层、透明、工具窗口
        WS_EX_TOPMOST | WS_EX_LAYERED | WS_EX_TOOLWINDOW,
        wc.lpszClassName,
        L"桌宠",
        WS_POPUP,//无边框窗口
        x,y,200,300,//窗口位置和大小
        NULL,NULL,hInstance,NULL
    );

    SetLayeredWindowAttributes(hwndMain,0,255,LWA_ALPHA);
    ShowWindow(hwndMain,SW_SHOW);
    SetTimer(hwndMain,1,30,NULL);

    MSG msg;
    while(GetMessage(&msg,NULL,0,0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    GdiplusShutdown(gdiplusToken);
    return 0;
}