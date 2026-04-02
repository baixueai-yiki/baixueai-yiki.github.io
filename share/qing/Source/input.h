#pragma once
#include <windows.h>

// 鼠标按下/移动/抬起的拖拽辅助函数
void HandleDragMouseDown(HWND hwnd, LPARAM lParam);
void HandleDragMouseMove(HWND hwnd, LPARAM lParam);
void HandleDragMouseUp();

// 判断一个坐标是否落在角色当前的矩形区域内
bool IsInsideCharacter(int x, int y);
