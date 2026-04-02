#pragma once

// 简化版的桌宠状态枚举，区分空闲、对话、拖动
enum PetState
{
    STATE_IDLE,
    STATE_TALK,
    STATE_DRAG
};

// 状态读写接口
void SetState(PetState state);
PetState GetState();
