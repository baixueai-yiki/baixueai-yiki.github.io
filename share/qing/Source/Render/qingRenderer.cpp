#include "qingRenderer.h"
#include "../Platform/qingPath.h"
#include <gdiplus.h>

using namespace Gdiplus;

static Image* g_image = nullptr;
static ULONG_PTR g_token;

void QingRendererInit()
{
    GdiplusStartupInput input;
    GdiplusStartup(&g_token, &input, NULL);

    std::wstring path = QingGetContentPath(L"..\\Content\\Images\\character.png");

    g_image = new Image(path.c_str());
}