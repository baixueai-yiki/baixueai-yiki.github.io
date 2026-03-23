#pragma once

// 全局桌宠状态：位置、尺寸、当前拖拽状态及偏移。
struct QingPetState
{
    int x;
    int y;
    int w;
    int h;
    bool isDragging;
    int dragOffsetX;
    int dragOffsetY;
};

extern QingPetState g_pet;

// 供程序启动时调用，重置为默认位置/状态。
void QingPetInit();
