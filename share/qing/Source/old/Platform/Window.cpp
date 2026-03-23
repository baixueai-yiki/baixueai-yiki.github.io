#include <windows.h>
#include <windowsx.h>
#include "../Input/qingInputDispatcher.h"
#include "../Render/qingRenderer.h"
#include "qingWindow.h"

// The window procedure simply routes mouse events to the input dispatcher
// and asks the renderer to draw whenever Windows requests a repaint.
LRESULT CALLBACK QingWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_MOUSEMOVE:
    case WM_LBUTTONDOWN:
    case WM_LBUTTONUP:
    case WM_LBUTTONDBLCLK:
        QingHandleInput(hwnd, msg, wParam, lParam);
        break;

    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        QingRendererRender(hdc);
        EndPaint(hwnd, &ps);
        break;
    }

    case WM_ERASEBKGND:
        // Signal that we handle background ourselves (reduce flicker).
        return 1;

    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;

    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }

    return 0;
}

// Creates a layered, transparent, always-on-top window that fills the screen.
HWND QingCreateWindow(HINSTANCE hInstance)
{
    WNDCLASSW wc = {};
    wc.lpfnWndProc = QingWndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"QingPetWindow";
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = reinterpret_cast<HBRUSH>(GetStockObject(NULL_BRUSH));

    ATOM atom = RegisterClassW(&wc);
    if (atom == 0 && GetLastError() != ERROR_CLASS_ALREADY_EXISTS)
        return nullptr;

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

    if (hwnd)
    {
        SetLayeredWindowAttributes(hwnd, 0, 255, LWA_ALPHA);
    }

    return hwnd;
}
