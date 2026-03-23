#include "qingRenderer.h"
#include "../Platform/qingPath.h"
#include "../Core/qingPetState.h"
#include <gdiplus.h>
#include <windows.h>

using namespace Gdiplus;

// 静态全局变量：保存加载的图片和 GDI+ token，用于清理资源。
static Image* s_image = nullptr;
static ULONG_PTR s_token = 0;

// 初始化渲染器：启动 GDI+、加载角色 PNG、计算初始位置并填充 g_pet。
bool QingRendererInit()
{
    if (s_image != nullptr)
        return true;

    GdiplusStartupInput input;
    if (GdiplusStartup(&s_token, &input, nullptr) != Ok)
        return false;

    std::wstring path = QingGetContentPath(L"..\\Content\\Images\\character.png");
    s_image = Image::FromFile(path.c_str());
    if (!s_image || s_image->GetLastStatus() != Ok)
    {
        delete s_image;
        s_image = nullptr;
        GdiplusShutdown(s_token);
        s_token = 0;
        return false;
    }

    // 将角色尺寸写入全局状态，供输入逻辑判定区域。
    g_pet.w = s_image->GetWidth();
    g_pet.h = s_image->GetHeight();

    // 计算窗口初始位置为屏幕中心，防止超出边界。
    int screenW = GetSystemMetrics(SM_CXSCREEN);
    int screenH = GetSystemMetrics(SM_CYSCREEN);
    g_pet.x = (screenW - g_pet.w) / 2;
    if (g_pet.x < 0) g_pet.x = 0;
    g_pet.y = (screenH - g_pet.h) / 2;
    if (g_pet.y < 0) g_pet.y = 0;

    return true;
}

// 在 WM_PAINT 时调用，负责实际绘制角色图像。
void QingRendererRender(HDC hdc)
{
    if (!hdc || !s_image || s_image->GetLastStatus() != Ok)
        return;

    Graphics graphics(hdc);
    graphics.SetCompositingMode(CompositingModeSourceOver);
    graphics.Clear(Color(0, 0, 0, 0)); // 先清空为透明
    graphics.DrawImage(s_image, g_pet.x, g_pet.y, g_pet.w, g_pet.h);
}

// 退出时调用，释放 Image 和关闭 GDI+。
void QingRendererShutdown()
{
    delete s_image;
    s_image = nullptr;

    if (s_token != 0)
    {
        GdiplusShutdown(s_token);
        s_token = 0;
    }

    g_pet.w = 0;
    g_pet.h = 0;
}
