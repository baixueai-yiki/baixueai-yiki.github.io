#pragma once
#include <windows.h>

// 初始化渲染器，返回是否成功加载 GDI+ 和角色图像。
bool QingRendererInit();

// 释放渲染相关资源（图像 + GDI+）。
void QingRendererShutdown();

// 在 WM_PAINT 的 HDC 上把角色贴图画出来。
void QingRendererRender(HDC hdc);
