#include "qingPetState.h"

// 默认初始值：屏幕左上角附近，尚未加载尺寸。
QingPetState g_pet = { 120, 120, 0, 0, false, 0, 0 };

// 重新设置为初始值，避免多次运行留下旧状态。
void QingPetInit()
{
    g_pet.x = 120;
    g_pet.y = 120;
    g_pet.w = 0;
    g_pet.h = 0;
    g_pet.isDragging = false;
    g_pet.dragOffsetX = 0;
    g_pet.dragOffsetY = 0;
}
